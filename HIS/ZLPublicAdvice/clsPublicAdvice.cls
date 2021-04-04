VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public objDockInAdvices As zlPublicAdvice.clsDockInAdvices
Public objDockOutAdvices As zlPublicAdvice.clsDockOutAdvices
Public objMedRecPage As zlPublicAdvice.clsArchiveMedRec
Private WithEvents clsDis As zl9Disease.clsDisease
Attribute clsDis.VB_VarHelpID = -1

Public Event PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String)  '转科
Public Event Closed(ByVal lngFunID As Long, ByVal strTag As String) '阳性结恶果反馈单填写了诊断后关闭窗体时触发

Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级

Public Function ShowCommItem(frmParent As Object, ByVal strFind As String, ByRef blnCancle As Boolean, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal intType As Integer) As String
    ShowCommItem = gobjKernel.ShowCommItem(frmParent, strFind, blnCancle, lngLeft, lngTop, intType)
End Function

Public Sub ShowDiagHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean)
    Call gobjKernel.ShowDiagHelp(bytModal, frmParent, lngItemID, blnShowInTaskBar)
End Sub

Public Sub ShowClincHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean)
    Call gobjKernel.ShowClincHelp(bytModal, frmParent, lngItemID, blnShowInTaskBar)
End Sub

Public Function ShowBloodApply(frmParent As Object, ByVal lng输血医嘱ID As Long) As Boolean
    ShowBloodApply = gobjKernel.ShowBloodApply(frmParent, lng输血医嘱ID)
End Function

Public Function ShowConsultationApply(frmParent As Object, ByVal lng医嘱ID As Long, Optional ByVal intType As Integer = 2, Optional objfrm As Object) As Boolean
    ShowConsultationApply = gobjKernel.ShowConsultationApply(frmParent, lng医嘱ID, intType, objfrm)
End Function

Public Function ShowOperationApply(frmParent As Object, ByVal lng医嘱ID As Long) As Boolean
    ShowOperationApply = gobjKernel.ShowOperationApply(frmParent, lng医嘱ID)
End Function

Public Function ShowPacsApplication(frmParent As Object, ByVal lng检查医嘱ID As Long) As Boolean
'功能：外部调用检查申请单查看功能
    ShowPacsApplication = gobjKernel.ShowPacsApplication(frmParent, lng检查医嘱ID)
End Function

Public Function ShowSchemeEdit(frmParent As Object, ByVal int范围 As Long, Optional rsScheme As ADODB.Recordset, _
    Optional ByVal blnView As Boolean, Optional ByVal bln显示缺省列 As Boolean, Optional ByVal str使用科室 As String, _
    Optional ByVal byt场合 As Byte, Optional ByVal str诊疗分类 As String, Optional ByVal str操作类型 As String, _
    Optional ByVal str执行分类 As String) As ADODB.Recordset
'功能：编辑成套方案接口
'参数：int范围=1-门诊使用,2-住院使用,3-门诊和住院都可以使用
'      blnView=是否仅查看
'      bln显示缺省列=临床路径项目定义调用时"选择使用"
'返回：与表"诊疗项目组合"相同结构的动态记录集,如果取消则返回Nothing
    Set ShowSchemeEdit = gobjKernel.ShowSchemeEdit(frmParent, int范围, rsScheme, blnView, bln显示缺省列, str使用科室, byt场合, str诊疗分类, str操作类型, str执行分类)
End Function

Public Function ShowAdviceEdit(frmParent As Object, ByVal int场合 As Integer, ByVal bytUseType As Byte, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
      ByVal strAdivceOfPath As String, ByVal dat开始时间 As Date, arrSQL As Variant, strAdviceOfItem As String, Optional ByVal rsLastAdvice As ADODB.Recordset, _
       Optional ByVal datSendTime As Date, Optional ByRef str路径项目IDs As String, Optional ByRef objMip As Object) As Boolean
'功能：生成临床路径项目对应的医嘱
'参数：
'      int场合          =0-医生站调用,1-护士站调用
'      bytUseType       =1:路径生成,2=添加路径外项目
'      strAdivceOfPath  =bytUseType=1时，传入医嘱内容ID:婴儿序号:路径项目ID,...，例：227:0:38,335:1:69
'                       =bytUseType=2时，传入医嘱ID串,此时事务还没有提交，例：2315,2316
'      dat开始时间      =bytUseType=1时，传入生成医嘱的开始时间
'                       =bytUseType=2时，传入病人路径最后一次生成的当前日期
'      rs重选项目的医嘱 =bytUseType=1时，传入重选项目的医嘱集，字段包括：项目id,病人医嘱ID，组id,诊疗项目id,是按医嘱生成的序号升序排的
'      datSendTime    =bytUseType=1时，传入路径的生成时间
'      objMip 消息对象 zl9ComLib.clsMipModule
'返回：arrSQL           =医嘱保存和校对的SQL语句
'      strAdviceOfItem  =bytUseType=1时，路径项目与医嘱ID的对应,例：38:1983,69:1978
'                       =bytUseType=2时，传入医嘱ID串,此时事务还没有提交，例：2315,2316
'      ShowAdviceEdit   =true:保存，false:取消
'      str路径项目IDs   =路径生成时中医修改了的配方的，且超出了允许修改配方的比例的项目，对应的变异原因：项目ID1|变异编码1,项目2|变异编码2····
    ShowAdviceEdit = gobjKernel.ShowAdviceEdit(frmParent, int场合, bytUseType, lng病人ID, lng主页ID, strAdivceOfPath, dat开始时间, arrSQL, strAdviceOfItem, rsLastAdvice, datSendTime, str路径项目IDs, objMip)
End Function

Public Function ShowAdviceOperate(frmParent As Object, ByVal int场合 As Integer, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal strAdviceOfItem As String, ByVal dateStop As Date, Optional ByRef objMip As Object) As Boolean
'功能：用于临床路径执行后调用停止长嘱的功能
'      strAdviceOfItem=路径项目相关的医嘱ID
'      dateStop=停止长嘱的时间
'      objMip 消息对象 zl9ComLib.clsMipModule
    ShowAdviceOperate = gobjKernel.ShowAdviceOperate(frmParent, int场合, lng病人ID, lng主页ID, lng病区ID, strAdviceOfItem, dateStop, objMip)
End Function

Public Function AdviceOperate(frmParent As Object, ByVal strMainPrivs As String, ByVal int类型 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, _
    Optional ByVal lng医嘱ID As Long, Optional ByRef objMip As Object, Optional ByRef strPatisOut As String, Optional ByVal blnOnePati As Boolean) As Boolean
'功能：双击消息提醒列表时调用校对或者确认停止功能
    AdviceOperate = gobjKernel.AdviceOperate(frmParent, strMainPrivs, int类型, lng病人ID, lng主页ID, lng病区ID, lng医嘱ID, objMip, strPatisOut, blnOnePati)
End Function

Public Function AdviceSend(frmParent As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strMainPrivs As String, Optional ByRef objMip As Object) As Boolean
'功能：双击消息提醒列表时调用医嘱发送功能
    AdviceSend = gobjKernel.AdviceSend(frmParent, lng病区ID, lng病人ID, lng主页ID, strMainPrivs, objMip)
End Function

Public Function ChargeDelAudit(frmParent As Object, ByVal lng执行科室ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：双击消息提醒列表时调用销帐审核功能
    ChargeDelAudit = gobjKernel.ChargeDelAudit(frmParent, lng执行科室ID, lng病人ID)
End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    If gobjKernel Is Nothing Then CloseWindows = True: Exit Function
    CloseWindows = gobjKernel.CloseWindows
    If Not gobjKernel Is Nothing Then Set gobjKernel = Nothing
End Function


Public Function InitCommon(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, Optional ByVal str站点 As String, Optional ByRef frmMain As Object, _
        Optional ByVal lngModule As Long, Optional ByVal MainPrivs As String, Optional ByRef objMip As Object, Optional ByRef objEmr As Object) As Boolean
'功能：初始化医嘱公共部件
'参数：MainPrivs=主界面权限串
'     frmMain =主界面窗体对象
'     mlngModule=主界面模块号
'     objMip=消息平台对象
    Set gcnOracle = cnOracle
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    glngSys = lngSys
    glngModule = lngModule
    gMainPrivs = MainPrivs
    
    gstrNodeNo = str站点
    '初始化费用公共部件
    On Error Resume Next
    If gobjExpense Is Nothing Then
        Set gobjExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Not gobjExpense Is Nothing Then
            If gobjExpense.zlInitCommon(lngSys, cnOracle) = False Then Exit Function
        End If
    End If
    If gclsInsure Is Nothing Then
        Set gclsInsure = CreateObject("zl9Insure.clsInsure")
        If Not gclsInsure Is Nothing Then
            Call gclsInsure.InitOracle(cnOracle)
        End If
    End If
    If gobjKernel Is Nothing Then
        Set gobjKernel = New clsCISKernel
        Call gobjKernel.InitCISKernel(cnOracle, frmMain, lngSys, MainPrivs)
    End If
    Call gobjKernel.InitCISKernelPar
    
    If gobjCISJob Is Nothing Then
        Set gobjCISJob = CreateObject("zl9CISJob.clsCISJob")
        Call gobjCISJob.InitCISJob(cnOracle, frmMain, lngSys, MainPrivs, objEmr)
    End If
    Call gobjCISJob.InitCISJobPar
     
    Call InitDisease(cnOracle, lngSys, frmMain, lngModule, MainPrivs, objMip)
    
    Call objMedRecPage.InitArchiveMedRec(cnOracle, lngSys, lngModule, objMip)
    
    Call InitSysPar
    
    InitCommon = True

End Function

Public Function GetExecAdviceRecord(ByVal strIDsAndNos As String, ByRef rsAdvice As Recordset, Optional ByVal blnMoved As Boolean) As Boolean
'功能：根据指定执行的医嘱ID，返回执行医嘱信息记录集（与医生站医嘱清单的显示格式不同）(主要用于医技显示）
'参数：strIDsAndNos 医嘱ID和发送号字符串，医嘱ID1:发送号1,医嘱ID2:发送号2
'     blnMoved 是否存在数据转出的情况（可不传入，自动判断）
    Dim rsTmp As Recordset
    Dim strSQL As String, strSQL1 As String
    Dim rsReturn As Recordset
    Dim blnDateMoved As Boolean
    
    strSQL = "Select  b.Id As 医嘱ID, b.相关id,a.发送号, b.病人id, b.主页id, b.开始执行时间 as 开始时间,'' as 医嘱内容,  a.发送数次 || Nvl(D.计算单位,C.计算单位) as 数次, b.医生嘱托, b.开嘱医生, b.开嘱时间,a.发送时间, a.发送人,Nvl(D.计算单位,C.计算单位) as 计算单位, a.No as 单据号, a.记录性质, a.执行部门id, a.完成人, a.完成时间, b.序号, b.病人来源," & vbNewLine & _
            "       b.挂号单, b.婴儿, b.姓名, b.医嘱期效, b.诊疗类别, b.诊疗项目id, b.标本部位, b.检查方法, b.天数, b.单次用量, b.总给予量, b.执行频次," & vbNewLine & _
            "       b.紧急标志, b.收费细目id, b.首次用量, b.执行时间方案, b.皮试结果, b.医嘱内容 as 医嘱内容Tmp,0 as 数据转出" & vbNewLine & _
            "From 病人医嘱发送 A, 病人医嘱记录 B, 诊疗项目目录 C,收费项目目录 D" & vbNewLine & _
            "Where a.医嘱id = b.Id And b.诊疗项目id = c.Id And b.收费细目ID = D.ID(+) And" & vbNewLine & _
            "      (a.医嘱id, a.发送号) In (Select /*+cardinality(x,10)*/ x.C1 As 医嘱id, x.C2 As 发送号 From Table(f_Num2list2([1])) x)" & _
            " order by b.序号"
    If gobjComlib Is Nothing Then Call zlGetComLib
    If blnMoved = False Then
        blnMoved = gobjComlib.zlDatabase.TableDataMoved("病人医嘱发送", " (医嘱ID,发送号) IN", " (Select C1 As 医嘱id, C2 As 发送号 From Table(f_Num2list2('" & strIDsAndNos & "')))")
    End If
    If blnMoved Then
        strSQL1 = strSQL
        strSQL1 = Replace(strSQL1, "0 as 数据转出", "1 as 数据转出")
        strSQL1 = Replace(strSQL1, "病人医嘱记录", "H病人医嘱记录")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL = strSQL & " Union ALL " & strSQL1
    End If
        
    On Error GoTo errH
    
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetInAdviceRecord", strIDsAndNos)
    Set rsReturn = gobjComlib.zlDatabase.CopyNewRec(rsTmp)
    If rsReturn.RecordCount > 0 Then
        Do While Not rsReturn.EOF
            rsReturn!医嘱内容 = Get执行内容(Val(rsReturn!发送号), Val(rsReturn!医嘱ID), Val(rsReturn!相关ID & ""), rsReturn!诊疗类别 & "", rsReturn!医嘱内容Tmp & "", Val(rsReturn!数据转出 & "") = 1)
            
            rsReturn.Update
            rsReturn.MoveNext
        Loop
        rsReturn.MoveFirst
    End If
    Set rsAdvice = rsReturn
    GetExecAdviceRecord = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function AddExamineAdvice(ByRef arrSQL As Variant, ByVal objExamineAdvice As clsExamineAdvice, Optional ByRef lng医嘱ID As Long) As Boolean
'功能：新增检查医嘱、包含附项内容
'参数：arrSQL/出：返回生成医嘱的SQL
'     objExamineadvice：医嘱对象信息
'     lng医嘱ID：生成的医嘱的相关ID
    Dim str项目名称     As String
    Dim lng医嘱序号     As Long
    Dim rsTmp       As Recordset
    Dim strSQL      As String

    On Error GoTo errH

    '获取医嘱所需的相关数据

    str项目名称 = gobjComlib.Sys.RowValue("诊疗项目目录", objExamineAdvice.检查项目ID, "名称")

    lng医嘱ID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")        '获取医嘱ID
    
    '组织主医嘱插入语句
    lng医嘱序号 = GetMaxAdviceNO(objExamineAdvice.病人ID, objExamineAdvice.主页ID, objExamineAdvice.婴儿) + 1 '病人医嘱记录.序号，递增

   If TypeName(arrSQL) = "Empty" Then arrSQL = Array()

    If objExamineAdvice.修改医嘱ID <> 0 Then
        '修改医嘱，删除后重新插入
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & objExamineAdvice.修改医嘱ID & ",1)"
    End If
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "" & _
                             "ZL_病人医嘱记录_Insert(" & lng医嘱ID & _
                             ",NULL," & _
                             lng医嘱序号 & "," & _
                             objExamineAdvice.病人来源 & _
                             "," & objExamineAdvice.病人ID & "," & _
                             IIF(objExamineAdvice.主页ID = 0, "NULL", objExamineAdvice.主页ID) & "," & _
                             objExamineAdvice.婴儿 & ",1,1,'D'," & _
                             objExamineAdvice.检查项目ID & _
                             ",NULL,NULL,NULL,1," & _
                             "'" & FormatExamineAdvice(str项目名称, objExamineAdvice.部位方法, objExamineAdvice.执行类型) & _
                             "',Null,Null,'一次性',NULL,NULL,NULL,NULL,0," & _
                             IIF(objExamineAdvice.执行科室ID <= 0, "Null", objExamineAdvice.执行科室ID) & _
                             "," & IIF(objExamineAdvice.执行科室ID <= 0, "5", objExamineAdvice.执行科室性质) & "," & objExamineAdvice.紧急标志 & _
                             ",to_date('" & Format(objExamineAdvice.开始时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                             ",NULL," & _
                             objExamineAdvice.病人科室ID & "," & objExamineAdvice.开单科室ID & _
                             ",'" & objExamineAdvice.开嘱医生 & "'," & _
                             "to_date('" & Format(objExamineAdvice.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                             "'" & objExamineAdvice.挂号单 & "',Null,Null," & _
                             objExamineAdvice.执行类型 & _
                             ",NULL,NULL,'',Null,NULL,NULL,NULL," & objExamineAdvice.申请序号 & ")"

    '诊断关联信息
    If objExamineAdvice.关联诊断IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & lng医嘱ID & ",'" & objExamineAdvice.关联诊断IDs & "')"
    End If

    '组织部位插入语句
    Call GetAdvicePartSaveSql(lng医嘱ID, str项目名称, objExamineAdvice, arrSQL, lng医嘱序号, objExamineAdvice.申请序号)

    '组织申请附项插入语句
    Call GetAdviceAffixSaveSql(lng医嘱ID, arrSQL, objExamineAdvice.申请附项串)

    AddExamineAdvice = True

    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddInspectionAdvice(ByRef arrSQL As Variant, ByVal objInspectionAdvice As clsInspectionAdvice, Optional ByRef lng医嘱ID As Long) As Boolean
'功能：新增检验医嘱、包含附项内容
'参数：arrSQL/出：返回生成医嘱的SQL
'     objInspectionAdvice：医嘱对象信息
'     lng医嘱ID：生成的医嘱的相关ID
    Dim lng相关ID As Long, Y As Long
    Dim lng检验项目ID As Long, str检验项目 As String, str检验项目组合 As String
    Dim str检验计价性质 As String, str检验执行性质 As String
    Dim str医嘱内容 As String, lng序号 As Long
    Dim str采集方法 As String, str采集计价性质 As String, str采集执行性质 As String
    Dim bln是否组合项目 As Boolean
    Dim rsTmp As Recordset
    
    If TypeName(arrSQL) = "Empty" Then arrSQL = Array()
    lng相关ID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")
    lng序号 = GetMaxAdviceNO(objInspectionAdvice.病人ID, objInspectionAdvice.主页ID, objInspectionAdvice.婴儿)

    If objInspectionAdvice.修改医嘱ID <> 0 Then
        '修改医嘱，删除后重新插入
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & objInspectionAdvice.修改医嘱ID & ",1)"
    End If
    
    'a.先产生检验医嘱
    For Y = 0 To UBound(Split(objInspectionAdvice.检验项目IDs, ","))
        lng医嘱ID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")
        lng检验项目ID = Split(Split(objInspectionAdvice.检验项目IDs, ",")(Y), ":")(0)
        str检验项目 = gobjComlib.Sys.RowValue("诊疗项目目录", lng检验项目ID, "名称")
        bln是否组合项目 = False
        If UBound(Split(Split(objInspectionAdvice.检验项目IDs, ",")(Y), ":")) > 0 Then
            If Val(Split(Split(objInspectionAdvice.检验项目IDs, ",")(Y), ":")(1)) = 0 Then
                str检验项目组合 = str检验项目组合 & "," & str检验项目
            Else
                bln是否组合项目 = True
            End If
        Else
            str检验项目组合 = str检验项目组合 & "," & str检验项目
        End If
        Set rsTmp = Get诊疗项目记录(lng检验项目ID)
        str检验计价性质 = Val("" & rsTmp!计价性质)
        str检验执行性质 = IIF("" & rsTmp!执行科室 = "", "NULL", "" & rsTmp!执行科室)
        str医嘱内容 = str检验项目
        lng序号 = lng序号 + 1
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
            lng医嘱ID & "," & lng相关ID & "," & lng序号 & "," & objInspectionAdvice.病人来源 & "," & objInspectionAdvice.病人ID & "," & _
            ZVal(objInspectionAdvice.主页ID) & "," & ZVal(objInspectionAdvice.婴儿) & ",1,1,'C'," & _
            lng检验项目ID & ",Null,Null,Null,1," & _
            "'" & str医嘱内容 & "',Null," & "'" & objInspectionAdvice.标本 & "','一次性',Null," & _
            "Null,Null,Null," & str检验计价性质 & "," & objInspectionAdvice.执行科室ID & _
            "," & str检验执行性质 & "," & objInspectionAdvice.紧急标志 & ",to_date('" & Format(objInspectionAdvice.开始时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & objInspectionAdvice.病人科室ID & "," & _
            objInspectionAdvice.开单科室ID & ",'" & objInspectionAdvice.开嘱医生 & "',to_date('" & Format(objInspectionAdvice.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & objInspectionAdvice.挂号单 & "',Null," & _
            "Null,0,Null,Null,''" & _
            ",Null,Null,Null,Null," & objInspectionAdvice.申请序号 & ",NULL,NULL,NULL,NULL," & IIF(bln是否组合项目, 1, "NULL") & ")"
    Next

    'b.再产生采集方法医嘱
    str采集方法 = gobjComlib.Sys.RowValue("诊疗项目目录", objInspectionAdvice.采集项目ID, "名称")
    str检验项目组合 = Mid(str检验项目组合, 2)
    Set rsTmp = Get诊疗项目记录(objInspectionAdvice.采集项目ID)
    str采集计价性质 = Val("" & rsTmp!计价性质)
    str采集执行性质 = "" & rsTmp!执行科室
    str医嘱内容 = FormatInspectionAdvice(str检验项目组合, str采集方法, objInspectionAdvice.标本)
    lng序号 = lng序号 + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
        lng相关ID & ",Null," & lng序号 & "," & objInspectionAdvice.病人来源 & "," & objInspectionAdvice.病人ID & "," & _
        ZVal(objInspectionAdvice.主页ID) & "," & ZVal(objInspectionAdvice.婴儿) & ",1,1,'E'," & _
        objInspectionAdvice.采集项目ID & ",Null,Null,Null,1," & _
        "'" & str医嘱内容 & "','" & objInspectionAdvice.嘱托 & "'," & "'" & objInspectionAdvice.标本 & "','一次性',Null," & _
        "Null,Null,Null," & str采集计价性质 & "," & objInspectionAdvice.采集科室ID & _
        "," & str采集执行性质 & "," & objInspectionAdvice.紧急标志 & ",to_date('" & Format(objInspectionAdvice.开始时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') ,Null," & objInspectionAdvice.病人科室ID & "," & _
        objInspectionAdvice.开单科室ID & ",'" & objInspectionAdvice.开嘱医生 & "'," & "to_date('" & Format(objInspectionAdvice.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & objInspectionAdvice.挂号单 & "',Null," & _
        "Null,0,Null,Null,''" & _
        ",Null,Null,Null,Null," & objInspectionAdvice.申请序号 & ")"
    
    '组织申请附项插入语句
    Call GetAdviceAffixSaveSql(lng相关ID, arrSQL, objInspectionAdvice.申请附项串)
    
    If objInspectionAdvice.关联诊断IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & lng相关ID & ",'" & objInspectionAdvice.关联诊断IDs & "')"
    End If
    
    lng医嘱ID = lng相关ID
    AddInspectionAdvice = True

    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AddOperationAdvice(ByRef arrSQL As Variant, ByVal objOperationAdvice As clsOperationAdvice, Optional ByRef lng医嘱ID As Long) As Boolean
'功能：新增手术医嘱、包含附项内容
'参数：arrSQL/出：返回生成医嘱的SQL
'     objOperationAdvice：医嘱对象信息
'     lng医嘱ID：生成的医嘱的相关ID
    Dim lng相关ID As Long, Y As Long
    Dim lng手术项目ID As Long, str手术项目 As String, str附加手术 As String
    Dim str手术计价性质 As String, str手术执行性质 As String
    Dim str医嘱内容 As String, lng序号 As Long
    Dim str麻醉项目 As String, str麻醉计价性质 As String, str麻醉执行性质 As String
    Dim rsTmp As Recordset, lngIndex As Long, lng序号Tmp As Long '主要手术SQL索引
    
    lng相关ID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")
    lng序号 = GetMaxAdviceNO(objOperationAdvice.病人ID, objOperationAdvice.主页ID, 0)
    str麻醉项目 = gobjComlib.Sys.RowValue("诊疗项目目录", objOperationAdvice.麻醉项目ID, "名称")
    
    If TypeName(arrSQL) = "Empty" Then arrSQL = Array()
    If objOperationAdvice.修改医嘱ID <> 0 Then
        '修改医嘱，删除后重新插入
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & objOperationAdvice.修改医嘱ID & ",1)"
    End If
    
    '把主要手术的位置留起来
    lngIndex = UBound(arrSQL)
    lng序号 = lng序号 + 1
    lng序号Tmp = lng序号
    
    'a.先产生附加手术医嘱
    For Y = 1 To UBound(Split(objOperationAdvice.手术项目IDs, ","))
        lng医嘱ID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")
        lng手术项目ID = Split(objOperationAdvice.手术项目IDs, ",")(Y)
        str手术项目 = gobjComlib.Sys.RowValue("诊疗项目目录", lng手术项目ID, "名称")
        str附加手术 = str附加手术 & IIF(str附加手术 = "", "", ",") & str手术项目
        Set rsTmp = Get诊疗项目记录(lng手术项目ID)
        str手术计价性质 = Val("" & rsTmp!计价性质)
        str手术执行性质 = IIF("" & rsTmp!执行科室 = "", "NULL", "" & rsTmp!执行科室)
        str医嘱内容 = str手术项目

        lng序号 = lng序号 + 1
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
            lng医嘱ID & "," & lng相关ID & "," & lng序号 & "," & objOperationAdvice.病人来源 & "," & objOperationAdvice.病人ID & "," & _
            ZVal(objOperationAdvice.主页ID) & ",NULL,1,1,'F'," & _
            lng手术项目ID & ",Null,Null,Null,1," & _
            "'" & str医嘱内容 & "','" & objOperationAdvice.嘱托 & "'," & "'" & Format(objOperationAdvice.手术时间, "yyyy-MM-dd HH:mm:ss") & "','一次性',Null," & _
            "Null,Null,Null," & str手术计价性质 & "," & objOperationAdvice.执行科室ID & _
            "," & str手术执行性质 & "," & objOperationAdvice.紧急标志 & ",to_date('" & Format(objOperationAdvice.开始时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & objOperationAdvice.病人科室ID & "," & _
            objOperationAdvice.开单科室ID & ",'" & objOperationAdvice.开嘱医生 & "',to_date('" & Format(objOperationAdvice.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & objOperationAdvice.挂号单 & "',Null," & _
            "Null,0,Null,Null,''" & _
            ",Null,Null,Null,Null,NULL,NULL,NULL,NULL," & objOperationAdvice.手术情况 & ")"
    Next
    
    'b.产生主要手术医嘱
    lng医嘱ID = lng相关ID
    lng手术项目ID = Split(objOperationAdvice.手术项目IDs, ",")(0)
    str手术项目 = gobjComlib.Sys.RowValue("诊疗项目目录", lng手术项目ID, "名称")

    Set rsTmp = Get诊疗项目记录(lng手术项目ID)
    str手术计价性质 = Val("" & rsTmp!计价性质)
    str手术执行性质 = IIF("" & rsTmp!执行科室 = "", "NULL", "" & rsTmp!执行科室)

    str医嘱内容 = FormatOperationAdvice(str手术项目, str麻醉项目, str附加手术, Format(objOperationAdvice.手术时间, "yyyy-MM-dd HH:mm:ss"), objOperationAdvice.手术部位)

    arrSQL(lngIndex) = "ZL_病人医嘱记录_Insert(" & _
        lng医嘱ID & ",NULL," & lng序号Tmp & "," & objOperationAdvice.病人来源 & "," & objOperationAdvice.病人ID & "," & _
        ZVal(objOperationAdvice.主页ID) & ",NULL,1,1,'F'," & _
        lng手术项目ID & ",Null,Null,Null,1," & _
        "'" & str医嘱内容 & "','" & objOperationAdvice.嘱托 & "'," & "'" & Format(objOperationAdvice.手术时间, "yyyy-MM-dd HH:mm:ss") & "','一次性',Null," & _
        "Null,Null,Null," & str手术计价性质 & "," & objOperationAdvice.执行科室ID & _
        "," & str手术执行性质 & "," & objOperationAdvice.紧急标志 & ",to_date('" & Format(objOperationAdvice.开始时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & objOperationAdvice.病人科室ID & "," & _
        objOperationAdvice.开单科室ID & ",'" & objOperationAdvice.开嘱医生 & "',to_date('" & Format(objOperationAdvice.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & objOperationAdvice.挂号单 & "',Null," & _
        "Null,0,Null,Null,''" & _
        ",Null,Null,Null,Null,NULL,NULL,NULL,NULL," & objOperationAdvice.手术情况 & ")"

    'c.再产生麻醉医嘱
    Set rsTmp = Get诊疗项目记录(objOperationAdvice.麻醉项目ID)
    str麻醉计价性质 = Val("" & rsTmp!计价性质)
    str麻醉执行性质 = "" & rsTmp!执行科室
    str医嘱内容 = str麻醉项目
    lng医嘱ID = gobjComlib.zlDatabase.GetNextID("病人医嘱记录")
    lng序号 = lng序号 + 1
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
        lng医嘱ID & "," & lng相关ID & "," & lng序号 & "," & objOperationAdvice.病人来源 & "," & objOperationAdvice.病人ID & "," & _
        ZVal(objOperationAdvice.主页ID) & ",Null,1,1,'G'," & _
        objOperationAdvice.麻醉项目ID & ",Null,Null,Null,1," & _
        "'" & str医嘱内容 & "','" & objOperationAdvice.嘱托 & "'," & "'" & Format(objOperationAdvice.手术时间, "yyyy-MM-dd HH:mm:ss") & "','一次性',Null," & _
        "Null,Null,Null," & str麻醉计价性质 & "," & objOperationAdvice.麻醉科室ID & _
        "," & str麻醉执行性质 & "," & objOperationAdvice.紧急标志 & ",to_date('" & Format(objOperationAdvice.开始时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') ,Null," & objOperationAdvice.病人科室ID & "," & _
        objOperationAdvice.开单科室ID & ",'" & objOperationAdvice.开嘱医生 & "'," & "to_date('" & Format(objOperationAdvice.开嘱时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & objOperationAdvice.挂号单 & "',Null," & _
        "Null,0,Null,Null,''" & _
        ",Null,Null,Null,Null,NULL,NULL,NULL,NULL," & objOperationAdvice.手术情况 & ")"
    
    '组织申请附项插入语句
    Call GetAdviceAffixSaveSql(lng相关ID, arrSQL, objOperationAdvice.申请附项串)
    
    If objOperationAdvice.关联诊断IDs <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & lng相关ID & ",'" & objOperationAdvice.关联诊断IDs & "')"
    End If
    
    lng医嘱ID = lng相关ID
    AddOperationAdvice = True

    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function Get诊疗执行科室(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef rsDept As Recordset, ByVal lng项目ID As Long, _
    ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, _
    Optional ByVal lng当前执行ID As Long, Optional ByVal lng药品ID As Long, Optional ByVal int期效 As Integer = 1, Optional ByVal int范围 As Integer = 2, _
    Optional ByVal bln输液 As Boolean, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal bln屏蔽输液中心 As Boolean, Optional ByVal lng用法ID As Long, Optional ByVal bln静脉营养 As Boolean, Optional ByVal lng病区ID As Long) As Boolean
'功能：根据诊疗项目执行科室信息返回可用的执行科室在指定下拉框中
'参数：lng病人科室ID=病人科室ID
'      lng当前执行ID=医嘱当前的执行科室ID
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院(缺省)
'      bln输液=当前药品是否属于输液类的（给药途径为输液）
'      blnEditable=false表示当前医嘱不可编辑
'      bln屏蔽输液中心=如果一组输液其他医嘱的执行科室不是输液配置中心，则此条医嘱也不能设置为输液配置中心
'说明：对非药医嘱,当前的执行科室可能是强行选择出来的,需要显示在选择框中;另选择框中增加一个其它供选择
    Dim strSQL As String, str药房 As String, str药房IDs As String
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim bln规格 As Boolean, i As Long
    Dim strStock As String, lng操作员科室ID As Long
    Dim str输液配置科室 As String
    Dim int输液配置医嘱期效 As Integer
    Dim bln不默认配置中心 As Boolean
    Dim bln指定输液配置中心 As Boolean
    Dim rsTmp As Recordset
    Dim str类别 As String, int执行科室 As Integer
    Dim str配液给药途径 As String, strFilter As String

    
    On Error GoTo errH
    Set rsTmp = Get诊疗项目记录(lng项目ID)
    str类别 = rsTmp!类别
    int执行科室 = Val(rsTmp!执行科室 & "")
    
    If str类别 = "4" Then
        strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And A.收费细目ID=[4]" & _
            " Order by B.服务对象,C.编码"
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        '系统可以指定药品执行科室,这里提取所有可选的供再选择
        If str类别 = "5" Then
            str药房 = "西药房"
            str药房IDs = gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "可用西药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            str药房IDs = gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "可用成药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            str药房IDs = gobjComlib.zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "可用中药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达))
        End If
            
        '药品库存限制
        If bln规格 And str药房IDs <> "" And blnEditable Then
            If gblnStock Then
                strStock = " And Exists(" & _
                    " Select 1 From 药品库存" & _
                    " Where (Nvl(批次,0)=0 Or 效期 Is Null Or 效期>Trunc(Sysdate))" & _
                    " And 性质=1 And 药品ID=[4] And 库房ID=A.执行科室ID" & _
                    " And 可用数量>0 And Instr('," & str药房IDs & ",',','||库房ID||',')>0)"
            Else
                strStock = " And Instr('," & str药房IDs & ",',','||A.执行科室ID||',')>0"
            End If
        End If
        
        '是否启用输液配置中心
        If bln屏蔽输液中心 And Val(gobjComlib.zlDatabase.GetPara("特殊性质药品允许发送到配制中心", glngSys, p输液配置中心, "0")) = 0 Then
            '71101当选择自备药或其他执行性质时，也允许选择输液配置中心，因为有可能输液配置中心同时兼任普通药房的工作。
            'strStock = strStock & " And A.执行科室ID <> [12]  "
        Else
            If gstr输液配置中心 <> "" And blnEditable Then
                If bln输液 And int范围 = 2 Then
                    str输液配置科室 = gobjComlib.zlDatabase.GetPara("来源病区", glngSys, p输液配置中心, "")
                    int输液配置医嘱期效 = Val(gobjComlib.zlDatabase.GetPara("医嘱类型", glngSys, p输液配置中心, "1")) - 1
                    str配液给药途径 = gobjComlib.zlDatabase.GetPara("输液给药途径", glngSys, p输液配置中心)
                    If (str输液配置科室 = "" Or InStr("," & str输液配置科室 & ",", "," & lng病区ID & ",") > 0) And (int期效 = int输液配置医嘱期效 Or int输液配置医嘱期效 = -1) And _
                            (InStr("," & str配液给药途径 & ",", "," & lng用法ID & ",") > 0 Or str配液给药途径 = "") Or bln静脉营养 Then
                        bln指定输液配置中心 = True
                    Else
                        'bln不默认配置中心=启用了输液配置中心，但是期效\科室\给药途径参数不匹配，把配置中心放最后。
                        bln不默认配置中心 = True
                    End If
                Else
                    '不是输液类的或范围不是住院的，把输液配置中心放最后
                    bln不默认配置中心 = True
                End If
            End If
        End If
        
        '药品从系统指定的储备药房中找
        If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
        If Not bln上班安排 Then
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & strStock & _
                IIF(bln不默认配置中心, " Order By Decode(instr('," & gstr输液配置中心 & ",',',' || C.ID || ','),0,B.服务对象,9),C.编码", _
                " Order by B.服务对象,C.编码")
        Else
            bytDay = Weekday(gobjComlib.zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & strStock & _
                IIF(bln不默认配置中心, " Order By Decode(instr('," & gstr输液配置中心 & ",',',' || C.ID || ','),0,B.服务对象,9),C.编码", _
                " Order by B.服务对象,C.编码")
        End If
    Else
        Select Case int执行科室
            Case 0, 5 '0-无执行的叮嘱,5-院外执行
                Get诊疗执行科室 = True: Exit Function
            Case 1 '1-病人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([3],[6]) Order by 编码"
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([3],[6]) Order by 编码"
                Else
                    strSQL = _
                        " Select A.ID,A.编码,A.简码,A.名称" & _
                        " From 部门表 A,病案主页 B" & _
                        " Where A.ID=B.当前病区ID And B.病人ID=[9] And B.主页ID=[10]" & _
                        " Union " & _
                        " Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
            Case 3 '3-操作员所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([7],[6]) Order by 编码"
                lng操作员科室ID = Get操作员部门ID(int范围)
            Case 4 '4-指定科室
                If int期效 = 1 Then bln上班安排 = Check上班安排(False)
                If Not bln上班安排 Then
                    strSQL = _
                        " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                        " From 部门表 A,诊疗执行科室 B,部门性质说明 C" & _
                        " Where A.ID=B.执行科室ID And B.诊疗项目ID=[5] And A.ID=C.部门ID" & _
                        " And C.服务对象 IN([2],3) And (B.病人来源 is NULL Or B.病人来源=[2])" & _
                        " And (B.开单科室ID is NULL Or B.开单科室ID=[3])" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                Else
                    bytDay = Weekday(gobjComlib.zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
                        " From 诊疗执行科室 A,部门安排 B,部门表 C,部门性质说明 D" & _
                        " Where A.执行科室ID+0=B.部门ID And B.部门ID=C.ID And B.星期=[8]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                        " And C.ID=D.部门ID And D.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[3]) And A.诊疗项目ID=[5]" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
            Case 6 '6-开单人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([11],[6]) Order by 编码"
        End Select
    End If
        
    Set rsDept = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "Get诊疗执行科室", str药房, int范围, lng病人科室ID, lng药品ID, lng项目ID, _
         lng当前执行ID, lng操作员科室ID, bytDay, lng病人ID, lng主页ID, lng开嘱科室ID, gstr输液配置中心)
         
    '输液药品发送到输液配置中心
    If bln指定输液配置中心 Then
        For i = 0 To UBound(Split(gstr输液配置中心, ","))
            If str药房IDs = "" Or InStr("," & str药房IDs & ",", "," & Split(gstr输液配置中心, ",")(i) & ",") > 0 Then
                strFilter = strFilter & " Or ID=" & Split(gstr输液配置中心, ",")(i)
            End If
        Next
        rsTmp.Filter = Mid(strFilter, 5)
        If rsTmp.RecordCount = 0 Then
            '如果输液配置中心不是存储库房，则可以选择其他存储库房为发药药房(65111)
            rsTmp.Filter = 0
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
    End If
    
    Get诊疗执行科室 = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function AdviceExecuteTag(ByRef arrSQL As Variant, ByVal str医嘱串 As String, ByVal intType As Integer, ByVal bln审核费用 As Boolean) As Boolean
'功能：执行、撤销执行医嘱
'参数：arrSQL/出：返回执行的SQL
'      str医嘱串：医嘱ID1,医嘱ID2······
'      inttype：1-执行，0-撤销执行
'      bln审核费用：如果为True，如果是记账划价单据则调用费用审核记账费用接口
    Dim strSQL As String, rsTmp As Recordset
    Dim colSQL As Collection, i As Long
    
    If TypeName(arrSQL) = "Empty" Then arrSQL = Array()
    
    On Error GoTo errH
    If intType = 1 Then
        strSQL = "Select a.要求时间, a.医嘱id, a.发送号, b.病人来源,C.NO" & vbNewLine & _
                "From 医嘱执行时间 A, 病人医嘱记录 B, 病人医嘱发送 C" & vbNewLine & _
                "Where a.医嘱id = b.Id And a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And" & vbNewLine & _
                "      a.医嘱id In (Select Column_Value From Table(f_Num2list([1])))"
    Else
        strSQL = "Select a.执行时间, a.医嘱id, a.发送号" & vbNewLine & _
                "From 病人医嘱执行 A" & vbNewLine & _
                "Where " & vbNewLine & _
                "      a.医嘱id In (Select Column_Value From Table(f_Num2list([1])))"
    End If

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "AdviceExecuteTag", str医嘱串)
    
    Do While Not rsTmp.EOF
        If intType = 1 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱执行_Insert(" & rsTmp!医嘱ID & "," & rsTmp!发送号 & "," & _
                    "To_Date('" & Format(rsTmp!要求时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    1 & ",'','" & UserInfo.姓名 & "'," & _
                    "To_Date('" & Format(rsTmp!要求时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    0 & "," & "0," & intType & ",'','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            If bln审核费用 Then
                Set colSQL = New Collection
                Call gobjExpense.zlBillOperation.VerfyBillingPriceBill(Val(rsTmp!病人来源 & ""), rsTmp!NO & "", False, colSQL)
                For i = 1 To colSQL.Count
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = colSQL(i)
                Next
            End If
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱执行_Delete(" & rsTmp!医嘱ID & "," & rsTmp!发送号 & "," & _
                    "To_Date('" & Format(rsTmp!执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        End If
        rsTmp.MoveNext
    Loop
    AdviceExecuteTag = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function AdviceAduitingTag(ByRef arrSQL As Variant, ByVal str医嘱串 As String, ByVal intType As Integer) As Boolean
'功能：执行完成、撤销执行完成医嘱
'参数：arrSQL/出：返回执行的SQL
'      str医嘱串：医嘱ID1,医嘱ID2······
'      inttype：1-执行完成，0-撤销执行完成
    Dim strSQL As String, rsTmp As Recordset
    Dim colSQL As Collection, i As Long
    
    If TypeName(arrSQL) = "Empty" Then arrSQL = Array()
    
    On Error GoTo errH
    strSQL = "Select a.医嘱id, a.发送号,A.NO,B.病人来源,A.完成人" & vbNewLine & _
            "From 病人医嘱发送 A,病人医嘱记录 B" & vbNewLine & _
            "Where A.医嘱ID=B.ID And " & vbNewLine & _
            "      a.医嘱id In (Select Column_Value From Table(f_Num2list([1])))"

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "AdviceAduitingTag", str医嘱串)
    
    Do While Not rsTmp.EOF
        If intType = 1 Then
            If rsTmp!完成人 & "" = "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱执行_Finish(" & rsTmp!医嘱ID & "," & rsTmp!发送号 & ")"
                Set colSQL = New Collection
                Call gobjExpense.zlBillOperation.VerfyBillingPriceBill(Val(rsTmp!病人来源 & ""), rsTmp!NO & "", False, colSQL)
                For i = 1 To colSQL.Count
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = colSQL(i)
                Next
            End If
        Else
            If rsTmp!完成人 & "" <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病人医嘱执行_cancel(" & rsTmp!医嘱ID & "," & rsTmp!发送号 & ")"
            End If
        End If
        rsTmp.MoveNext
    Loop
    AdviceAduitingTag = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function AdviceRefuseTag(ByRef arrSQL As Variant, ByVal str医嘱串 As String, ByVal intType As Integer, Optional ByVal str拒绝原因 As String) As Boolean
'功能：拒绝、取消拒绝医嘱
'参数：arrSQL/出：返回执行的SQL
'      str医嘱串：医嘱ID1,医嘱ID2······
'      inttype：1-拒绝执行，0-撤销拒绝执行
'      str拒绝原因：如果是拒绝执行，则可传入拒绝的原因
    Dim strSQL As String, rsTmp As Recordset
    
    If TypeName(arrSQL) = "Empty" Then arrSQL = Array()
    
    On Error GoTo errH
    strSQL = "Select a.医嘱id, a.发送号,A.NO" & vbNewLine & _
            "From 病人医嘱发送 A" & vbNewLine & _
            "Where " & vbNewLine & _
            "      a.医嘱id In (Select Column_Value From Table(f_Num2list([1])))"

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "AdviceRefuseTag", str医嘱串)
    
    Do While Not rsTmp.EOF
        If intType = 1 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱执行_拒绝执行(" & rsTmp!医嘱ID & "," & rsTmp!发送号 & ",NULL,NULL,NULL,'" & str拒绝原因 & "')"
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱执行_取消拒绝(" & rsTmp!医嘱ID & "," & rsTmp!发送号 & ")"
        End If
        rsTmp.MoveNext
    Loop
    AdviceRefuseTag = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAdviceState(ByRef rsState As Recordset, ByVal str医嘱串 As String) As Boolean
'功能：获取医嘱执行状态
'参数：rsState/出：返回执行状态记录集
'      str医嘱串：医嘱ID1:发送号1,医嘱ID2:,发送号2······，发送号可不传入.
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    If InStr(str医嘱串, ":") = 0 Then
        strSQL = "Select a.医嘱id, a.发送号,A.执行状态" & vbNewLine & _
                "From 病人医嘱发送 A" & vbNewLine & _
                "Where " & vbNewLine & _
                "      a.医嘱id In (Select Column_Value From Table(f_Num2list([1])))"
    Else
        strSQL = "Select a.医嘱id, a.发送号,A.执行状态" & vbNewLine & _
                "From 病人医嘱发送 A" & vbNewLine & _
                "Where " & vbNewLine & _
                "      (a.医嘱id,A.发送号) In (Select C1 As 医嘱id, C2 As 发送号 From Table(f_Num2list2([1])))"
    End If

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetAdviceState", str医嘱串)
    
    Set rsState = rsTmp
    GetAdviceState = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckSkinResult(ByRef rsResult As Recordset, ByVal strNOs As String) As Boolean
'功能：获取有效皮试结果
'参数：Result/出：返回皮试结果记录集（病人ID,药物ID，结果(+/-)）
'      strNOs：传入的NOs逗号分隔
    Dim strSQL As String, rsTmp As Recordset
    Dim strNOsTmp As String
    
    On Error GoTo errH
    strSQL = "Select Distinct a.病人id, a.诊疗项目id As 药物id, MAX(Nvl(c.过敏时间, c.记录时间)) As 过敏时间,  Decode(c.结果,1,'(+)','(-)') AS 皮试结果" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B, 病人过敏记录 C" & vbNewLine & _
            "Where a.Id = b.医嘱id And a.诊疗项目id = c.药物id(+) And a.病人id = c.病人id(+) And Nvl(c.结果, 1) = 1 And" & vbNewLine & _
            "      b.No In" & vbNewLine & _
            "      (Select Column_Value" & vbNewLine & _
            "       From Table(f_Str2list([1]))) Group by a.病人id, a.诊疗项目id, c.结果"
    
    If Len(strNOs) > 4000 Then
        strNOsTmp = Mid(strNOs, 1, InStrRev(Mid(strNOs, 1, 4000), ",") - 1)
        strNOs = Mid(strNOs, InStrRev(Mid(strNOs, 1, 4000), ",") + 1)
    Else
        strNOsTmp = strNOs
        strNOs = ""
    End If
    
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckSkinResult", strNOsTmp)
    Set rsResult = gobjComlib.zlDatabase.CopyNewRec(rsTmp)
    Do While strNOs <> ""
        If Len(strNOs) > 4000 Then
            strNOsTmp = Mid(strNOs, 1, InStrRev(Mid(strNOs, 1, 4000), ",") - 1)
            strNOs = Mid(strNOs, InStrRev(Mid(strNOs, 1, 4000), ",") + 1)
        Else
            strNOsTmp = strNOs
            strNOs = ""
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckSkinResult", strNOsTmp)
        Do While Not rsTmp.EOF
            rsResult.AddNew
            rsResult!病人ID = rsTmp!病人ID
            rsResult!药物id = rsTmp!药物id
            rsResult!过敏时间 = rsTmp!过敏时间
            rsResult!皮试结果 = rsTmp!皮试结果
            rsResult.Update
            rsTmp.MoveNext
        Loop
    Loop
    
    '排除过敏有效时间外的
    rsResult.Filter = "过敏时间>=" & Format(gobjComlib.zlDatabase.Currentdate, "YYYY-MM-DD") & " Or 过敏时间 = NULL"
    If rsResult.RecordCount > 0 Then rsResult.MoveFirst
    CheckSkinResult = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetOutAdviceDate(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef strOutDate As String) As Boolean
'功能：获取有效出院医嘱开始时间
'参数：strOutDate/出：返回有效的出院时间
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "Select a.开始执行时间" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
            "Where a.诊疗项目id = b.Id And A.病人ID = [1] And A.主页ID = [2] And a.医嘱状态 = 8 And a.诊疗类别 = 'Z' And b.操作类型 = '5' Order By a.开始执行时间 Desc"

    
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetOutAdviceDate", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        strOutDate = Format(rsTmp!开始执行时间 & "", "yyyy-MM-dd HH:mm:ss")
    End If
    GetOutAdviceDate = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetDockInAdviceRecord(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int场合 As Integer, Optional ByVal int婴儿 As Integer = -1, Optional ByVal str期效 As String = "11", Optional ByVal str状态 As String = "111", _
       Optional ByVal dat开始时间 As Date = CDate(0), Optional ByVal dat结束时间 As Date = CDate(0), Optional ByVal bln未记帐 As Boolean, Optional ByVal bln科内 As Boolean = True, Optional ByVal bln重整 As Boolean, Optional ByVal str前提IDs As String, Optional ByRef rsAdvice As Recordset) As Boolean
'功能：返回住院病人医嘱记录集
'参数：rsAdvice/出：返回医嘱记录记录集
'      int场合:  '调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      婴儿：默认-1（显示所有医嘱），0（只显示病人医嘱），>1（婴儿序号）
'      期效：默认"11"，"10"表示只返回长嘱，"01"表示只返回临嘱
'      状态：默认"111",第一位1表示显示新下达的，第二位表示显示已校对的，第三位表示显示已停止和作废的，默认三个都显示。
'      开始时间：（时间范围起点）默认cdate(0)。
'      结束时间：(时间范围结束)默认cdate(0)。
'      BLN未记帐:包含未记帐的医嘱
'      bln科内:只显示本科医嘱
'      bln重整:只显示重整后的医嘱
'      前提IDs：默认：""，医技站查询时使用，表示关联的医嘱ID字符串，逗号分隔。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim lngTop As Long, i As Long, j As Long
    Dim strFormat As String, strTmp As String
    Dim bln给药途径 As Boolean, bln中药用法 As Boolean
    Dim bln采集方法 As Boolean, bln输血途径 As Boolean, blnFirst As Boolean
    Dim lng医嘱ID As Long, str状态SQL As String
    Dim dat重整 As Date, str未记帐 As String
    Dim blnDo As Boolean, strCurr As String, strTime As String
    Dim str医嘱期效 As String, str医嘱状态 As String
    Dim blnMoved As Boolean, strGroupBy As String

    If lng病人ID = 0 Then Exit Function

    On Error GoTo errH

    blnMoved = CheckPatiDataMoved(lng病人ID, lng主页ID)
    
    strWhere = ""
    If int婴儿 <> -1 Then
        strWhere = strWhere & " And Nvl(A.婴儿,0)=[4]"
    End If
    If str期效 <> "11" Then
        For i = 1 To Len(str期效)
            If Val(Mid(str期效, i, 1)) = 1 Then str医嘱期效 = str医嘱期效 & "," & i - 1
        Next
        str医嘱期效 = Mid(str医嘱期效, 2)
        strWhere = strWhere & " And Instr([5],','||Nvl(A.医嘱期效,0)||',')>0"
    End If
    If str状态 <> "111" Then
        For i = 1 To Len(str状态)
            If Val(Mid(str状态, i, 1)) = 1 Then
                If i = 1 Then str医嘱状态 = str医嘱状态 & ",1,2"  '新下达
                If i = 2 Then str医嘱状态 = str医嘱状态 & ",3,5,6,7"  '已校对
                If i = 3 Then str医嘱状态 = str医嘱状态 & ",4,8,9"  '已废止
            End If
        Next
        str医嘱状态 = Mid(str医嘱状态, 2)
        strWhere = strWhere & " And Instr([6],','||Nvl(A.医嘱状态,0)||',')>0"

        '未用医嘱
        If Val(Mid(str状态, 3, 1)) = 0 Then
            strWhere = strWhere & " And Nvl(A.执行标记,0)<>-1"
        End If
    End If
    If dat开始时间 <> CDate(0) And dat结束时间 <> CDate(0) Then
        strWhere = strWhere & " And A.开嘱时间+0 Between [7] And [8]"
    End If

    '求最近次重整时间：当有长嘱显示时
    If Mid(str期效, 1, 1) = "1" Then
        strSQL = "Select 医嘱重整时间 as 时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDockInAdviceRecord", lng病人ID, lng主页ID)
        If Not IsNull(rsTmp!时间) Then
            dat重整 = Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '只显示包含未记帐费用的医嘱
    If bln未记帐 Then
        str未记帐 = _
        " And Exists" & vbNewLine & _
                 " (Select 1" & vbNewLine & _
                 "       From (Select Nvl(C.相关id, C.ID) As 医嘱id" & vbNewLine & _
                 "              From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C" & vbNewLine & _
                 "              Where A.医嘱id = C.ID And A.NO = B.NO And A.记录性质 = B.记录性质 And A.记录性质 = 2 And B.记录状态 = 0 And" & vbNewLine & _
                 "                    C.病人id = [1] And C.主页id = [2]" & IIF(int婴儿 <> -1, " And Nvl(C.婴儿, 0) = [4]", "") & ")" & vbNewLine & _
                 "       Where A.ID = 医嘱id Or A.相关id = 医嘱id)"
    End If

    '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法'总量及用法计算
    str状态SQL = "Decode(A.医嘱状态,1,'新开',2,'疑问',3,'校对',4,'作废',5,'重整',6,'暂停',7,'启用',8,'停止',9,'确认停止')"
    strSQL = _
    "Select /*+ RULE */ A.ID,A.相关ID,A.序号,Nvl(A.婴儿,0) as 婴儿ID,A.医嘱状态," & _
             " Nvl(A.诊疗类别,'*') as 诊疗类别,B.操作类型,C.毒理分类,A.紧急标志 as 标志," & _
             " A.审查结果,Decode(Nvl(A.医嘱期效,0),0,'长嘱','临嘱') as 期效," & _
             " To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,A.医嘱内容,Null as 内容,A.皮试结果 as 皮试," & _
             " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位)," & _
             " '4',A.总给予量||G.计算单位,'5',Round(A.总给予量/D.住院包装,5)||D.住院单位,'6',Round(A.总给予量/D.住院包装,5)||D.住院单位,A.总给予量||B.计算单位)) as 总量," & _
             " Decode(A.首次用量,Null,'',A.首次用量 || Decode(A.诊疗类别,'4',G.计算单位,B.计算单位) || ':') || Decode(A.单次用量,NULL,NULL,decode(sign(1-A.单次用量),1,'0'||A.单次用量,A.单次用量)||Decode(A.诊疗类别,'4',G.计算单位,B.计算单位)) as 单量,A.天数," & _
             " A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('2468',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法," & _
             " A.医生嘱托,A.执行时间方案 as 执行时间,To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间," & _
             " Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室," & _
             " Decode(Instr('567E',Nvl(A.诊疗类别,'*')),0,NULL,A.执行性质) as 执行性质," & _
             " To_Char(A.上次执行时间,'YYYY-MM-DD HH24:MI') as 上次执行," & str状态SQL & " as 状态," & _
             " A.开嘱医生,To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 开嘱时间,A.校对护士,To_Char(A.校对时间,'YYYY-MM-DD HH24:MI') as 校对时间," & _
             " A.停嘱医生,To_Char(A.停嘱时间,'YYYY-MM-DD HH24:MI') as 停嘱时间,A.确认停嘱护士 as 停嘱护士," & _
             " To_Char(A.确认停嘱时间,'YYYY-MM-DD HH24:MI') as 确认停嘱时间,A.诊疗项目ID,B.试管编码,A.执行标记,A.屏蔽打印,A.前提ID,Decode(A.新开签名ID,NULL,0,1) as 签名否," & _
             " M.病历文件ID as 文件ID,Nvl(N.通用,0) as 报告项,Max(y.病历id) As 报告id, Decode(Max(NVL(y.查阅状态,0)),MiN(NVL(y.查阅状态,0)),Max(NVL(y.查阅状态,0)),2) As 查阅状态,A.收费细目ID,B.计算单位 as 单量单位,A.开嘱科室ID,A.审核状态,A.申请序号," & _
            "(Select To_Char(Max(P.打印时间), 'YYYY-MM-DD HH24:MI') From 病人医嘱打印 P Where a.Id = p.医嘱id And p.打印标记 = 0) As 打印时间,A.审核标记,Decode(D.基本药物,null,0,1) as 基本药物,d.高危药品,A.标本部位, Max(y.检查报告id) || '' As 检查报告id"
    strSQL = strSQL & _
             " From 病人医嘱记录 A,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,收费项目目录 G,病人医嘱报告 Y,病历单据应用 M,病历文件列表 N" & _
             " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
             " And A.收费细目ID=D.药品ID(+) And A.收费细目ID=G.ID(+) And A.ID=Y.医嘱ID(+)" & _
             " And (Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL) Or A.诊疗类别='E' And B.操作类型='8')" & _
             " And A.诊疗项目ID=M.诊疗项目ID(+) And M.应用场合(+)=2 And M.病历文件ID=N.ID(+) And N.种类(+)=7" & _
             " And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1" & _
             IIF(str前提IDs <> "" And bln科内, " And Nvl(A.前提ID,0) in (Select Column_Value From Table(Cast(f_Num2list([9]) As zlTools.t_Numlist)))", "") & _
             IIF(int场合 = 2, "", " And A.病人来源<>3") & strWhere & str未记帐
    strGroupBy = " Group By a.Id, a.相关id, a.序号, a.婴儿, a.医嘱状态, a.诊疗类别, b.操作类型, c.毒理分类, a.紧急标志, a.审查结果, a.医嘱期效, a.开始执行时间, a.医嘱内容, a.皮试结果," & vbNewLine & _
            "         a.总给予量, a.首次用量, g.计算单位, d.住院包装, d.住院单位, a.单次用量, a.天数, a.执行频次, a.医生嘱托, b.名称, a.执行性质, a.执行时间方案, a.执行终止时间, e.名称," & vbNewLine & _
            "         a.上次执行时间, a.开嘱时间, a.开嘱医生, a.校对护士, a.校对时间, a.停嘱医生, a.停嘱时间, a.确认停嘱护士, a.确认停嘱时间, a.诊疗项目id, b.试管编码, a.执行标记, a.屏蔽打印," & vbNewLine & _
            "         a.前提id, a.新开签名id, m.病历文件id, n.通用, a.收费细目id, b.计算单位, a.开嘱科室id, a.审核状态, a.申请序号, a.审核标记, d.基本药物, d.高危药品, a.标本部位," & vbNewLine & _
            "         a.用药目的"

    '重整显示格式处理
    If dat重整 <> CDate(0) Then
        If bln重整 Then
            '只显示最后一次重整之后的医嘱
            strSQL = strSQL & " And (Nvl(A.重整标志,1)=1 Or A.医嘱状态 IN(1,2)) " & strGroupBy & " Order by 婴儿ID,序号"
        Else
            '显示重整前后分隔
            strSQL = _
            " Select * From (" & strSQL & " And Nvl(A.重整标志,1)=0 And A.医嘱状态 Not IN(1,2) " & strGroupBy & " Order by 婴儿ID,序号)" & _
                     " Union ALL" & _
                     " Select -Null as ID,-Null as 相关ID,-Null as 序号,-Null as 婴儿ID,-Null as 医嘱状态,Null as 诊疗类别,Null as 操作类型," & _
                     " Null as 毒理分类,-Null as 标志,-Null as 审查结果,Null as 期效,Null as 开始时间,Null as 医嘱内容,Null as 内容,Null as 皮试,Null as 总量," & _
                     " Null as 单量,-Null as 天数,Null as 频率,Null as 用法,Null as 医生嘱托,Null as 执行时间,Null as 终止时间,Null as 执行科室," & _
                     " Null as 执行性质,Null as 上次执行,Null as 状态,Null as 开嘱医生,Null as 开嘱时间,Null as 校对护士,Null as 校对时间," & _
                     " Null as 停嘱医生,Null as 停嘱时间,Null as 停嘱护士,Null as 确认停嘱时间,-Null as 诊疗项目ID,Null as 试管编码,-Null as 执行标记,-Null as 屏蔽打印,-Null as 前提ID," & _
                     " -Null as 签名否,-Null as 文件ID,-Null as 报告项,-Null as 报告ID,-Null as 查阅状态, -Null as 收费细目ID, Null as 单量单位, -Null as 开嘱科室ID,-Null as 审核状态,-Null as 申请序号,Null as 打印时间,-Null as 审核标记 ,-Null as 基本药物,-Null as 高危药品,-Null as 标本部位,-NULL AS 检查报告id  From Dual" & _
                     " Union ALL" & _
                     " Select * From (" & strSQL & " And (Nvl(A.重整标志,1)=1 Or A.医嘱状态 IN(1,2)) " & strGroupBy & " Order by 婴儿ID,序号)"
        End If
    Else
        strSQL = strSQL & strGroupBy & " Order by 婴儿ID,序号"
    End If

    '访问历史空间处理
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDockInAdviceRecord", lng病人ID, lng主页ID, 0, _
                                        int婴儿, "," & str医嘱期效 & ",", "," & str医嘱状态 & ",", dat开始时间, dat结束时间, IIF(str前提IDs = "", "0", str前提IDs))

    Set rsAdvice = rsTmp
    GetDockInAdviceRecord = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAdviceExecCount(ByRef rsCount As Recordset, ByVal str单据串 As String) As Boolean
'功能：获取医嘱发送数次和已执行数次（除药品外）
'参数：rsCount/出：返回发送数次和已执行数次记录集
'      str单据串：记录性质2:No1,记帐性质2:NO2······
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "Select NO, 医嘱id, 收费细目id, 发送总数量, 已执行数量, (发送总数量 - 已执行数量) As 未执行数量" & vbNewLine & _
            "From (Select a.No, b.医嘱id, b.收费细目id, Sum(b.数量) As 发送总数量," & vbNewLine & _
            "              Sum(Decode(a.执行状态, 1, b.数量, Decode(c.医嘱id, Null, 0, b.数量))) 已执行数量" & vbNewLine & _
            "       From 病人医嘱发送 A, 医嘱执行计价 B, 病人医嘱执行 C" & vbNewLine & _
            "       Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And b.医嘱id = c.医嘱id(+) And b.发送号 = c.发送号(+) And b.要求时间 = c.要求时间(+) And" & vbNewLine & _
            "             (a.记录性质, a.No) In" & vbNewLine & _
            "             (Select To_Number(C1) As 记录性质, C2 As NO From Table(f_Str2list2([1])))" & vbNewLine & _
            "       Group By a.No, b.医嘱id, b.发送号, b.收费细目id)"


    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetAdviceExecCount", str单据串)
    
    Set rsCount = rsTmp
    GetAdviceExecCount = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAdviceRemind(ByRef rsNotify As Recordset, ByVal lng病区ID As Long, Optional ByVal str责任护士 As String) As Boolean
'功能：获取病区医嘱提醒记录集
'参数：rsNotify/出：返回当前护士站消息记录集
'      str单据串：记录性质2:No1,记帐性质2:NO2······
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim strNotifyAdvice As String
    Dim intNotifyDay As Integer
    
    strNotifyAdvice = gobjComlib.zlDatabase.GetPara("自动刷新医嘱类型", glngSys, p住院护士站, "00000000000")
    intNotifyDay = Val(gobjComlib.zlDatabase.GetPara("自动刷新医嘱天数", glngSys, p住院护士站, 1))
    
    If Mid(strNotifyAdvice, m新开, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_001"
    If Mid(strNotifyAdvice, m新停, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_002"
    If Mid(strNotifyAdvice, m新废, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CIS_003"
    If Mid(strNotifyAdvice, m安排, 1) = "1" Then strTmp = strTmp & ",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015"
    If Mid(strNotifyAdvice, m危机值, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_003,ZLHIS_PACS_005"
    If Mid(strNotifyAdvice, m输液拒绝, 1) = "1" Then strTmp = strTmp & ",ZLHIS_TRANSFUSION_001"
    If Mid(strNotifyAdvice, m销帐申请, 1) = "1" Then strTmp = strTmp & ",ZLHIS_CHARGE_001"
    If Mid(strNotifyAdvice, mRIS预约, 1) = "1" Then strTmp = strTmp & ",ZLHIS_PACS_006"
    If Mid(strNotifyAdvice, mRIS预约准备, 1) = "1" Then strTmp = strTmp & ",ZLHIS_PACS_007"
    If Mid(strNotifyAdvice, m取血通知, 1) = "1" Then strTmp = strTmp & ",ZLHIS_BLOOD_003"
    If Mid(strNotifyAdvice, m标本拒收, 1) = "1" Then strTmp = strTmp & ",ZLHIS_LIS_004"
    If Mid(strNotifyAdvice, m备血完成, 1) = "1" Then strTmp = strTmp & ",ZLHIS_BLOOD_001"
    If Mid(strNotifyAdvice, m血袋回收, 1) = "1" Then strTmp = strTmp & ",ZLHIS_BLOOD_007"
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then GetAdviceRemind = True: Exit Function
    
    If str责任护士 = "" Then
        strSQL = "Select b.病人id, b.就诊id as 主页id,a.住院号,a.姓名, a.当前床号 As 床号, Nvl(b.就诊科室id, a.当前科室id) As 就诊科室id," & _
            " Nvl(b.就诊病区id, a.当前病区id) As 就诊病区id, b.病人来源, b.消息内容, b.类型编码, b.业务标识, b.优先程度, b.登记时间,a.险类" & _
            " From 病人信息 A, 业务消息清单 B, 业务消息提醒部门 C, 业务消息提醒人员 D" & _
            " Where a.病人id = b.病人id And b.Id = c.消息id And b.Id = d.消息id(+) And b.登记时间 >=Trunc(Sysdate-" & (intNotifyDay - 1) & ") and substr(b.提醒场合,[4],1)='1'" & _
            " And Nvl(b.是否已阅, 0) = 0  And instr(','||[5]||',',','||b.类型编码||',')>0 and (c.部门id = [1] Or d.提醒人员 = [2])" & _
            " Order By b.优先程度 Desc, b.登记时间 Desc"
    Else
        strSQL = "Select b.病人id, b.就诊id as 主页id,a.住院号,a.姓名, a.当前床号 As 床号, Nvl(b.就诊科室id, a.当前科室id) As 就诊科室id," & _
            " Nvl(b.就诊病区id, a.当前病区id) As 就诊病区id, b.病人来源, b.消息内容, b.类型编码, b.业务标识, b.优先程度, b.登记时间,a.险类" & _
            " From 病人信息 A, 业务消息清单 B, 业务消息提醒部门 C, 业务消息提醒人员 D,病案主页 e" & _
            " Where a.病人id = b.病人id And b.Id = c.消息id And b.Id = d.消息id(+) And b.登记时间 >=Trunc(Sysdate-" & (intNotifyDay - 1) & ") and substr(b.提醒场合,[4],1)='1'" & _
            " And Nvl(b.是否已阅, 0) = 0 and b.病人ID=e.病人ID and b.就诊ID=e.主页ID And instr(','||[5]||',',','||b.类型编码||',')>0 and (c.部门id = [1] Or d.提醒人员 = [2])" & _
            " and e.责任护士=[6]" & _
            " Order By b.优先程度 Desc, b.登记时间 Desc"
    End If
    Screen.MousePointer = 11
    
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetAdviceRemind", lng病区ID, UserInfo.姓名, , 3, strTmp, str责任护士)
    
    Set rsNotify = rsTmp
    GetAdviceRemind = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMsgRISReady(ByVal lng病区ID As Long) As Boolean
'功能：轮询查询产生RIS预约准消息
'返回：是否有消息产生，true 有消息，false无
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    Dim str消息内容 As String
    Dim datCur As Date
    Dim datTmp As Date
    Dim strNotifyAdvice As String
    
    strNotifyAdvice = gobjComlib.zlDatabase.GetPara("自动刷新医嘱类型", glngSys, p住院护士站, "000000000")
    
    If Val(Mid(strNotifyAdvice, mRIS预约准备, 1)) = 0 Then
        Exit Function
    End If
    
    On Error GoTo errH
    datCur = gobjComlib.zlDatabase.Currentdate
    
    datTmp = DateAdd("n", 30, datCur)
    
    strSQL = "select b.病人id,b.主页id,b.医嘱内容 as 项目,a.医嘱id,c.当前病区id,c.出院科室id,a.预约开始时间" & vbNewLine & _
        "from RIS检查预约 a,病人医嘱记录 b,病案主页 c" & vbNewLine & _
        "where a.医嘱id=b.id and b.病人id=c.病人id and b.主页id=c.主页id" & vbNewLine & _
        " and not exists(select 1 from 业务消息清单 m where m.病人id=c.病人id and m.就诊id=c.主页id and m.类型编码='ZLHIS_PACS_007' and m.业务标识=to_char(a.医嘱id))" & _
        "and c.当前病区id+0=[1] and a.预约开始时间 between [2] and [3]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "RIS预约准备消息", lng病区ID, datCur, datTmp)
    arrSQL = Array()
    For i = 1 To rsTmp.RecordCount
        str消息内容 = "请【" & Format(rsTmp!预约开始时间, "yyyy-MM-dd HH:mm") & "】前去做【" & rsTmp!项目 & "】检查。"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_业务消息清单_Insert(" & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!当前病区id & "," & rsTmp!出院科室ID & _
            ",2,'" & str消息内容 & "','0010','ZLHIS_PACS_007','" & rsTmp!医嘱ID & "',1,0,null," & rsTmp!当前病区id & ",null)"
        rsTmp.MoveNext
    Next
    '提交数据
    If UBound(arrSQL) <> -1 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            gobjComlib.zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "RIS预约准备消息"
        Next
        gcnOracle.CommitTrans: blnTrans = False
        GetMsgRISReady = True
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function AdviceSendSingle(ByRef objParent As Object, ByRef arrSQL As Variant, ByVal lng医嘱ID As Long, ByVal int发送方式 As Integer, ByVal lngModule As Long, Optional ByVal bln加班加价 As Boolean, Optional ByRef lng发送号 As Long) As Boolean
'功能：发送整组临时医嘱，产生医嘱发送、计价、待执行记录、费用记录
'参数：arrSQL/出：返回执行的SQL
'      lng发送号/出：本次发送的发送号
'      objParent:父窗体
'      医嘱ID：相关ID为空的主记录
'       发送方式：0-发送收费单据，1-发送记账单据
'       lngModule=调用模块号
'       bln加班加价=表示是否加班加价
    Dim rsExec As ADODB.Recordset  '医嘱执行计价
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    Dim rsNumber As ADODB.Recordset
    Dim rsItems As ADODB.Recordset '用于医保管控的费用记录集,动态记录集
    Dim rsMoneyNow As ADODB.Recordset '当前病人本次要发送的费用,动态记录集
    Dim rsMoneyDay As ADODB.Recordset '当前病人当天已发送的费用,静态记录集
    Dim rsMoney As New ADODB.Recordset
    Dim rsBill As ADODB.Recordset
    Dim rsRXKey As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date, strCurDate As String
    Dim bln划价 As Boolean, intWarn As Integer, strNOKey As String, strNO As String, lng费用序号 As Long, lng费用父号 As Long
    Dim rsAdvice As Recordset, lng基准药房 As Long
    Dim lng材料ID As Long, lng执行科室ID As Long, dbl单价 As Double
    Dim lngTmpEOF As Long, lngNOSequence As Long
    Dim str收费项目 As String, strPrice As String
    Dim rsTmp As Recordset, rsClone As Recordset
    Dim int父序号 As Integer, lng父项目ID As Long
    Dim strHaveSub As String, strNoneSub As String
    Dim rsSeek As Recordset, strCuvetteNumber As String
    Dim lng采集科室ID As Long, int计费状态 As Integer, str分解时间 As String, str发生时间 As String
    Dim bln附加手术 As Boolean, lng病人科室ID As Long, lng费用次数 As Long
    Dim rsAudit As Recordset, strAudit As String, lng药品类别ID As Long
    Dim int付数 As Integer, dbl数量 As Double, lng卫材类别ID As Long, dbl应收 As Double, cur应收 As Currency, cur实收 As Currency
    Dim blnBool As Boolean, str实收 As String, bln保险项目否 As Boolean, lng保险大类ID As Long, cur统筹金额 As Currency, str保险编码 As String, str费用类型 As String
    Dim strTmp As String, cur合计 As Currency, str类别 As String, str类别名称 As String
    Dim str用法 As String, int药品性质 As Long, int划价 As Integer, cur记帐合计 As Currency, str登记时间 As String, str自动发料 As String
    Dim bln实时监控 As Boolean, lng发送序号 As Long, blnFirst As Boolean, dbl发送数次 As Double, str首次时间 As String, str末次时间 As String
    Dim lng组ID As Long, str医嘱IDs As String, intOutType As Integer, strAdviceInfo As String, strDiag As String
    
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.相关id,A.病人来源, Nvl(a.相关id, a.Id) As 组id, Nvl(x.序号, a.序号) As 组号, a.医嘱状态, a.诊疗类别, f.名称 As 类别名称, a.诊疗项目id," & vbNewLine & _
            "       b.名称 As 诊疗项目, a.收费细目id, a.婴儿, a.医嘱内容, a.标本部位, a.检查方法, a.执行标记, a.天数, round(a.总给予量,5) as 总量, d.住院单位, d.门诊单位,  round(a.单次用量) as 单量," & vbNewLine & _
            "       Decode(a.诊疗类别, '4', c.计算单位, b.计算单位) As 计算单位, d.剂量系数, d.住院包装, d.门诊包装, a.开始执行时间, a.执行频次, a.频率次数, a.频率间隔, a.间隔单位, a.医生嘱托," & vbNewLine & _
            "       a.执行时间方案, a.病人科室id,H.当前病区ID as 病人病区ID, a.开嘱科室id, a.开嘱医生, a.开嘱时间, a.计价特性, a.执行性质, a.执行科室id, e.名称 As 执行科室, b.操作类型, b.试管编码," & vbNewLine & _
            "       Nvl(a.可否分零, d.住院可否分零) As 可否分零, Decode(a.诊疗类别, '4', g.在用分批, d.药房分批) As 分批, g.跟踪在用, c.是否变价, c.撤档时间, c.服务对象," & vbNewLine & _
            "       a.新开签名id As 签名id, a.摘要, a.紧急标志, c.撤档时间, b.计算方式, b.执行安排,H.险类,H.费别,H.门诊号,H.住院号,H.当前床号,H.姓名,H.年龄,H.性别,Y.ID as 挂号ID,A.病人id,A.主页ID" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B, 收费项目目录 C, 药品规格 D, 病人信息 H, 部门表 E, 诊疗项目类别 F, 材料特性 G, 病人医嘱记录 X,病人挂号记录 Y" & vbNewLine & _
            "Where h.病人id = a.病人id  And  A.病人ID=Y.病人ID(+) And A.挂号单=Y.NO(+) And " & vbNewLine & _
            "      (a.id = [1] Or a.相关ID=[1]) And" & vbNewLine & _
            "      a.医嘱状态 In (1, 3, 5) And a.医嘱期效 = 1 And a.相关id = x.Id(+) And b.类别 = f.编码 And a.诊疗项目id = b.Id And" & vbNewLine & _
            "      a.收费细目id = c.Id(+) And a.收费细目id = d.药品id(+) And a.收费细目id = g.材料id(+) And a.执行科室id = e.Id(+) And" & vbNewLine & _
            "      Nvl(a.执行标记, 0) <> -1 And a.开始执行时间 Is Not Null And Nvl(a.医嘱状态, 0) <> -1" & vbNewLine & _
            "Order By a.婴儿, 组号, 组id, a.序号"
            
    Set rsAdvice = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "AdviceSendSingle", lng医嘱ID)
    If rsAdvice.RecordCount = 0 Then MsgBox "没有找到可发送的医嘱,可能是由于医嘱已被发送。", vbInformation, "AdviceSendSingle": Exit Function
    
    Call InitSendRecordset(rsExec, rsBill, rsRXKey, rsSQL, rsTotal, rsUpload, rsNumber, rsMoneyNow, rsItems) '医嘱执行计价
    lng发送号 = gobjComlib.zlDatabase.GetNextNo(10)
    curDate = gobjComlib.zlDatabase.Currentdate
    strCurDate = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    bln划价 = True '初始全部是划价
    intWarn = -1 '记帐报警时缺省要提示,与病人无关
    gstr动态费别 = Load动态费别(Val(rsAdvice!病人科室ID & ""))
    Call gobjExpense.zlGetPriceGrade(gstrNodeNo, Val(rsAdvice!病人ID & ""), Val(rsAdvice!主页ID & ""), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)

    '提取当前病人的审批项目清单
    strAudit = ""
    If Val(rsAdvice!病人来源 & "") = 2 Then
        If Not IsNull(rsAdvice!险类) Then
            bln实时监控 = gclsInsure.GetCapability(support实时监控, rsAdvice!病人ID, rsAdvice!险类)
            If Val(gobjComlib.zlDatabase.GetPara("检查医保审批", glngSys, p住院医嘱发送, "1")) = 1 Then
                Set rsAudit = GetAuditRecord(rsAdvice!病人ID, rsAdvice!主页ID)
            Else
                Set rsAudit = Nothing
            End If
        Else
            bln实时监控 = False
            Set rsAudit = Nothing '以Nothing为标志该病人不需要判断
        End If
    End If
    
    For i = 1 To rsAdvice.RecordCount
        '产生单据号分配关键字
        '-----------------------------------------------------------------------------------------
        If InStr(",5,6,", rsAdvice!诊疗类别 & "") > 0 Then
            '中西成药按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
            '一并给药的，发送到一起：包括自备药和不同药房的情况
            If lng基准药房 = 0 Then lng基准药房 = Val(rsAdvice!执行科室ID & "")
            strNOKey = "中西成药_" & _
                Val(rsAdvice!病人科室ID & "") & "_" & Val(rsAdvice!开嘱科室ID & "") & "_" & _
                rsAdvice!开嘱医生 & "_" & lng基准药房
            
        ElseIf InStr(",4,M,", rsAdvice!诊疗类别 & "") > 0 Then
            '材料按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
            strNOKey = "材料医嘱_" & Val(rsAdvice!病人ID & "") & "_" & Val(rsAdvice!主页ID & "") & "_" & _
                Val(rsAdvice!病人科室ID & "") & "_" & Val(rsAdvice!开嘱科室ID & "") & "_" & _
                rsAdvice!开嘱医生 & "_" & Val(rsAdvice!执行科室ID & "")
        ElseIf rsAdvice!诊疗类别 & "" = "7" Then
            '一个配方中的所有草药分配一个独立单据号
            strNOKey = "中药配方"
        ElseIf Val(rsAdvice!相关ID & "") <> 0 And rsAdvice!诊疗类别 & "" = "C" Then
            '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
            '同一个类检验型，同一个检验执行科室，同一采集管，同一个采集方式，同一个采集执行科室的检验分配相同的单据号
            strNOKey = "一并采集_" & rsAdvice!标本部位 & "" & "_" & _
                rsAdvice!执行科室ID & "" & "_" & rsAdvice!操作类型 & "" & "_" & rsAdvice!试管编码 & ""
        ElseIf Val(rsAdvice!相关ID & "") <> 0 And InStr(",F,D,", rsAdvice!诊疗类别 & "") > 0 Then
            '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
            strNOKey = "非药医嘱_" & Val(rsAdvice!相关ID & "")
        Else
            '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法，采集方式，麻醉方式，输血医嘱/输血途径)
            strNOKey = "非药医嘱_" & Val(rsAdvice!ID & "")
        End If
        
        '产生医嘱记帐费用:以最新价格计算
        '-----------------------------------------------------------------------------------------
        strSQL = "": str收费项目 = ""
        If InStr(",5,6,7,", rsAdvice!诊疗类别 & "") > 0 Then
            '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
            If Val(rsAdvice!执行性质 & "") <> 5 Then
                strSQL = _
                    " Select A.ID,A.类别,D.名称 as 类别名称,RTrim(A.名称||' '||A.规格) as 名称," & _
                    " A.计算单位,A.是否变价,A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,100 as 附术收费率," & _
                    " Y.住院单位,Y.住院包装,Y.门诊单位,Y.门诊包装,Y.剂量系数,Y.药房分批 as 分批,0 as 跟踪在用,B.收入项目ID," & _
                    " C.收据费目,1 as 数量,B.现价 as 单价,[2] as 执行科室ID,0 as 从项,0 as 费用性质,0 as 收费方式,I.要求审批" & _
                    " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,药品规格 Y,保险支付项目 I" & _
                    " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                    " And A.ID=Y.药品ID(+) And A.ID=[1] And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " Order by A.编码"
            End If
        Else
            '先删除原非药医嘱的计价(应该没有)
            rsSQL.AddNew
            rsSQL!类型 = 1: rsSQL!项目ID = 0: rsSQL!序号 = i
            rsSQL!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID), Val(rsAdvice!相关ID & ""))
            rsSQL!sql = "ZL_病人医嘱计价_Delete(" & Val(rsAdvice!ID) & ",1)"
            rsSQL.Update
            '不计价,手工计价；叮嘱,院外执行的医嘱不读取
            If Val(rsAdvice!计价特性 & "") = 0 And InStr(",0,5,", Val(rsAdvice!执行性质 & "")) = 0 Then
                If rsAdvice!试管编码 & "" <> "" Then
                    lng材料ID = GetTubeMaterial(rsAdvice!试管编码 & "")
                End If
                '几种对应的计价情况
                If Not IsNull(rsAdvice!标本部位) And Not IsNull(rsAdvice!检查方法) Then
                    strPrice = " And c.检查部位=[3] And c.检查方法=[4] And Nvl(c.费用性质,0)=0"
                ElseIf Nvl(rsAdvice!执行标记, 0) = 0 Then
                    strPrice = " And c.检查部位 Is Null And c.检查方法 is Null And Nvl(c.费用性质,0)=0"
                Else '目前包含床旁或术中加收的情况
                    strPrice = " And c.检查部位 Is Null And c.检查方法 is Null And Nvl(c.费用性质,0) IN(0,1)"
                End If
                
                strSQL = "Select c.类别, a.收费项目id, a.收费数量 as 数量, a.固有对照, b.收入项目id, c.加班加价, b.加班加价率, Decode(c.是否变价, 1, b.缺省价格, b.现价) As 单价, c.是否变价," & vbNewLine & _
                        "       Nvl(a.从属项目, 0) As 从项, d.跟踪在用, [2] As 执行科室id, c.屏蔽费别, Nvl(a.费用性质, 0) As 费用性质, Nvl(a.收费方式, 0) As 收费方式" & vbNewLine & _
                        "From (Select *" & vbNewLine & _
                        "       From (Select c.诊疗项目id, c.收费项目id, c.检查部位, c.检查方法, c.费用性质, c.收费数量, c.固有对照, c.从属项目, c.收费方式, c.适用科室id," & vbNewLine & _
                        "                     Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & vbNewLine & _
                        "              From 诊疗收费关系 C" & vbNewLine & _
                        "              Where c.诊疗项目id = [1] " & strPrice & _
                                                " And (c.适用科室id Is Null And c.病人来源 = 0 Or c.适用科室id = Decode([2], 0, [6], [2]) And c.病人来源 = [7]))" & vbNewLine & _
                        "       Where Nvl(适用科室id, 0) = Top) A, 收费价目 B, 收费项目目录 C, 材料特性 D" & vbNewLine & _
                        "Where a.收费项目id = b.收费细目id And a.收费项目id = c.Id And a.收费项目id = d.材料id(+) And c.服务对象 In ([7], 3) " & vbNewLine & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "B", "8", "9", "10") & _
                        "  And    (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null) And (c.站点 = " & gstrNodeNo & " Or c.站点 Is Null) And" & vbNewLine & _
                        "      ((Sysdate Between b.执行日期 And b.终止日期) Or (Sysdate >= b.执行日期 And b.终止日期 Is Null)) And" & vbNewLine & _
                        "      (Nvl(a.收费方式, 0) = 1 And c.类别 = '4' And a.收费项目id = [5] Or Not (Nvl(a.收费方式, 0) = 1 And c.类别 = '4' And [5] <> 0))" & vbNewLine & _
                        "Order By 费用性质, 从项, a.收费项目id"
                
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "AdviceSendSingle", Val(rsAdvice!诊疗项目ID), _
                        Val(Nvl(rsAdvice!执行科室ID, 0)), CStr(Nvl(rsAdvice!标本部位)), CStr(Nvl(rsAdvice!检查方法)), lng材料ID, Val(rsAdvice!病人病区ID & ""), IIF(Val(rsAdvice!病人来源 & "") = 2, 2, 1), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                strSQL = ""
                Do While Not rsTmp.EOF
                    If Val(rsAdvice!总量 & "") <> 0 Then
                        '执行科室:非药嘱药品及跟踪卫材的专门取
                        lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
                        If rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!类别) > 0 Then
                            lng执行科室ID = Get收费执行科室ID(Val(rsAdvice!病人ID & ""), 0, rsTmp!类别, rsTmp!收费项目ID, 4, Nvl(rsAdvice!病人科室ID, 0), 0, Val(rsAdvice!病人来源 & ""), lng执行科室ID, , , Val(rsAdvice!病人来源 & ""))
                        End If
                        
                        '计算单价和实收
                        If Nvl(rsTmp!是否变价, 0) = 1 And InStr(",5,6,7,", rsTmp!类别) > 0 Then
                            '非药嘱药品计价按时价计算(仅一个收入),其它变价需要由医生输入
                            dbl单价 = CalcDrugPrice(rsTmp!收费项目ID, lng执行科室ID, Format(Val(rsAdvice!总量 & ""), "0.00000") * Nvl(rsTmp!收费数量, 0), , True, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        ElseIf Nvl(rsTmp!是否变价, 0) = 1 And rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Then
                            '跟踪在用的时价卫材和药品一样计算
                            dbl单价 = CalcDrugPrice(rsTmp!收费项目ID, lng执行科室ID, Format(Val(rsAdvice!总量 & ""), "0.00000") * Nvl(rsTmp!收费数量, 0), , True, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        Else '固定价格或普通变价(只有一个收入项目)
                            dbl单价 = Nvl(rsTmp!单价, 0)
                        End If
                        
                        '药品、卫材医嘱的计价固定对应不保存；非跟踪在用的时价卫材的变价需要输入，因此要保存到计价表中
                        If InStr(",4,5,6,7,", rsAdvice!诊疗类别 & "") = 0 _
                            Or rsAdvice!诊疗类别 & "" = "4" And Nvl(rsTmp!跟踪在用, 0) = 0 And Nvl(rsTmp!是否变价, 0) = 1 Then
                            rsSQL.AddNew
                            rsSQL!类型 = 1: rsSQL!项目ID = rsTmp!收费项目ID: rsSQL!序号 = i
                            rsSQL!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID), Val(rsAdvice!相关ID & ""))
                            rsSQL!sql = "ZL_病人医嘱计价_INSERT(" & _
                                Val(rsAdvice!ID) & "," & rsTmp!收费项目ID & "," & _
                                Nvl(rsTmp!数量, 0) & "," & Format(dbl单价, gstrDecPrice) & "," & _
                                Nvl(rsTmp!从项, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                Nvl(rsTmp!费用性质, 0) & "," & Nvl(rsTmp!收费方式, 0) & ")"
                            rsSQL.Update
                        End If
                        
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select " & rsTmp!收费项目ID & " as 收费细目ID," & _
                            lng执行科室ID & " as 执行科室ID," & _
                            Nvl(rsTmp!数量, 0) & " as 数量," & Format(dbl单价, gstrDecPrice) & " as 单价," & _
                            Nvl(rsTmp!从项, 0) & " as 从项," & Nvl(rsTmp!费用性质, 0) & " as 费用性质," & _
                            Nvl(rsTmp!收费方式, 0) & " as 收费方式 From Dual"
                    End If
                    rsTmp.MoveNext
                Loop
                If strSQL <> "" Then
                    strSQL = _
                        " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,A.是否变价," & _
                        " A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,B.附术收费率,Y.住院单位,Y.住院包装,Y.门诊单位,Y.门诊包装,Y.剂量系数," & _
                        " Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批,E.跟踪在用,B.收入项目ID," & _
                        " C.收据费目,X.数量,Decode(A.是否变价,1,X.单价,B.现价) as 单价,X.执行科室ID," & _
                        " X.从项,X.费用性质,X.收费方式,I.要求审批" & _
                        " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,(" & strSQL & ") X,药品规格 Y,保险支付项目 I" & _
                        " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.ID=E.材料ID(+)" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                        " And A.类别=D.编码 And X.收费细目ID=A.ID And A.ID=Y.药品ID(+)" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                        " Order by X.费用性质,X.从项,X.收费方式 Desc,A.ID"
                        '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                End If
            End If
        End If
                        
        '汇总折扣变量初始
        int父序号 = 0: lng父项目ID = 0
        strHaveSub = "": strNoneSub = ""
        '初始化用于汇总计算折扣的临时记录集
        Set rsSeek = New ADODB.Recordset
        rsSeek.Fields.Append "费用性质", adInteger
        rsSeek.Fields.Append "主项标签", adVariant
        rsSeek.Fields.Append "主收入ID", adBigInt
        rsSeek.Fields.Append "合计", adCurrency, , adFldIsNullable
        rsSeek.CursorLocation = adUseClient
        rsSeek.LockType = adLockOptimistic
        rsSeek.CursorType = adOpenStatic
        rsSeek.Open
        
        '提前生成样本条码
        strCuvetteNumber = ""
        If Val(rsAdvice!执行性质 & "") <> 0 Then
            If gbln发送生成条形码 Then
                lngTmpEOF = rsAdvice.AbsolutePosition
                rsAdvice.Filter = "ID=" & rsAdvice!相关ID
                If rsAdvice.RecordCount > 0 Then lng采集科室ID = Val(rsAdvice!执行科室ID & "")
                rsAdvice.Filter = 0
                rsAdvice.AbsolutePosition = lngTmpEOF
                strCuvetteNumber = GetCuvetteNumber(rsNumber, rsAdvice!试管编码 & "", _
                    Val(rsAdvice!ID & ""), Val(rsAdvice!相关ID & ""), rsAdvice!诊疗类别 & "", Val(rsAdvice!操作类型 & ""), _
                    Val(rsAdvice!执行科室ID & ""), Val(rsAdvice!婴儿), Val(rsAdvice!诊疗类别 & ""), _
                    Val(rsAdvice!紧急标志 & ""), rsAdvice!标本部位 & "", lng采集科室ID)
            End If
        End If
                        
        int计费状态 = IIF(Val(rsAdvice!计价特性 & "") = 1, -1, 0) '无需计费或未计费
                
        '分解时间
        str分解时间 = ""
        If Not IsNull(rsAdvice!执行时间方案) Or Nvl(rsAdvice!间隔单位) = "分钟" Then
            str分解时间 = Calc次数分解时间(0, rsAdvice!开始执行时间, CDate("3000-01-01"), "", Nvl(rsAdvice!执行时间方案), rsAdvice!频率次数, rsAdvice!频率间隔, rsAdvice!间隔单位)
        End If
        
        If Len(str分解时间) > 4000 Then
            Screen.MousePointer = 0
            MsgBox "当前发送的医嘱时间范围太长,共需执行" & CStr(UBound(Split(str分解时间, ",")) + 1) & "次。" & vbCrLf & _
                "超过了支持的最大次数" & CStr(UBound(Split(Mid(str分解时间, 1, 4000), ",")) + 1) & "次,请调整总量或天数后重新发送！", vbInformation, "AdviceSendSingle"
            Exit Function
        End If
        
        If strSQL <> "" Then
            Set rsMoney = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "AdviceSendSingle", Val(rsAdvice!收费细目ID & ""), Val(rsAdvice!执行科室ID & ""), Val(Nvl(rsAdvice!险类, 0)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            If Not rsMoney.EOF Then
                int计费状态 = 1 '已计费
                Set rsClone = rsMoney.Clone
            End If
            
            '处理收入项目级的费用明细
            bln附加手术 = rsAdvice!诊疗类别 & "" = "F" And Val(rsAdvice!相关ID & "") <> 0
            Do While Not rsMoney.EOF
MoneyItemBegin:
                '执行科室ID
                lng执行科室ID = Nvl(rsMoney!执行科室ID, 0)
                '在原值基础上取有效的非药嘱药品及跟踪卫材的执行科室
                If InStr(",4,5,6,7", rsAdvice!诊疗类别 & "") = 0 _
                    And (rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsMoney!类别) > 0) Then
                    lng病人科室ID = Val(rsAdvice!病人科室ID & "")
                    lng执行科室ID = Get收费执行科室ID(Val(rsAdvice!病人ID & ""), 0, rsMoney!类别, rsMoney!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
                End If
                
                '----------------------------------------
                '根据收费方式，确定当前收费项目是否应收费
                If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                    If Not AdviceMoneyMake(Val(rsAdvice!病人ID & ""), Val(rsAdvice!主页ID & ""), rsMoneyNow, rsMoneyDay, _
                        IIF(Val(rsAdvice!相关ID & "") <> 0, Val(rsAdvice!相关ID & ""), Val(rsAdvice!ID & "")), _
                        Val(rsAdvice!诊疗类别 & ""), rsMoney!ID, lng执行科室ID, rsAdvice!试管编码 & "", _
                        rsMoney!类别, Nvl(rsMoney!收费方式, 0), str分解时间, IIF(Val(rsAdvice!病人来源 & "") = 2, 2, 1), lng费用次数, Val(rsAdvice!总量 & ""), _
                        Val(rsAdvice!ID & ""), lng发送号, Val(rsMoney!数量 & ""), rsExec, Val(rsAdvice!计算方式 & ""), _
                                rsAdvice!执行频次 & "", Val(rsAdvice!单量 & ""), 1, rsMoney!费用性质, rsAdvice!诊疗类别 & "") Then
                        '跳过当前收费项目(多个收入项目)
                        str收费项目 = rsMoney!费用性质 & "_" & rsMoney!ID
                        Do While rsMoney!费用性质 & "_" & rsMoney!ID = str收费项目
                            rsMoney.MoveNext
                            If rsMoney.EOF Then Exit Do
                        Loop
                        If rsMoney.EOF Then Exit Do
                        GoTo MoneyItemBegin
                    End If
                End If
                '----------------------------------------
                
                '检查是否需要和已经审批
                If Nvl(rsMoney!要求审批, 0) = 1 And Not rsAudit Is Nothing Then
                    rsAudit.Filter = "项目ID=" & rsMoney!ID
                    If rsAudit.EOF Then
                        If UBound(Split(strAudit, vbCrLf)) < 10 Then
                            If InStr(strAudit, "●" & rsMoney!名称) = 0 Then
                                strAudit = strAudit & vbCrLf & "●" & rsMoney!名称
                            End If
                        ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                            strAudit = strAudit & vbCrLf & "… …"
                        End If
                    End If
                End If
            
                If InStr(",5,6,7", rsMoney!类别) > 0 Then
                    If lng药品类别ID = 0 Then
                        MsgBox "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！", vbInformation, "AdviceSendSingle"
                        Exit Function
                    End If
                
                    If InStr(",5,6,7", rsAdvice!诊疗类别 & "") > 0 Then
                        If rsAdvice!诊疗类别 & "" = "7" Then
                            int付数 = Val(rsAdvice!总量 & "")
                            '中药药房单位按不可分零处理:每付
                            If Val(rsAdvice!可否分零 & "") = 0 Then
                                dbl数量 = Val(rsAdvice!单量 & "") / Nvl(rsMoney!剂量系数, 1)
                            Else
                                dbl数量 = IntEx(Val(rsAdvice!单量 & "") / Nvl(rsMoney!剂量系数, 1) / Nvl(IIF(Val(rsAdvice!病人来源 & "") = 2, rsMoney!住院包装, rsMoney!门诊包装), 1)) * Nvl(IIF(Val(rsAdvice!病人来源 & "") = 2, rsMoney!住院包装, rsMoney!门诊包装), 1)
                            End If
                        Else
                            int付数 = 1
                            dbl数量 = Val(rsAdvice!总量 & "") * Nvl(IIF(Val(rsAdvice!病人来源 & "") = 2, rsMoney!住院包装, rsMoney!门诊包装), 1)
                        End If
                    Else
                        int付数 = 1
                        '中药药房单位按不可分零处理:每付
                        '非药嘱药品计价:因为这里预定了售价数量,因此不作不分零处理
                        dbl数量 = Val(rsAdvice!总量 & "") * Nvl(rsMoney!数量, 0)
                    End If
                    dbl数量 = Format(dbl数量, "0.00000")
                    
                    If Nvl(rsMoney!是否变价, 0) = 1 Then
                        dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, int付数 * dbl数量, , True, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                    Else
                        dbl单价 = Format(Nvl(rsMoney!单价, 0), gstrDecPrice)
                    End If
                ElseIf rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Then
                    '检查卫生材料入出类别
                    If lng卫材类别ID = 0 Then
                        Screen.MousePointer = 0
                        MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, "AdviceSendSingle"
                        Exit Function
                    End If
                    
                    int付数 = 1
                    If InStr(",1,2,3,4,5,6,7,", Val("" & rsMoney!收费方式)) > 0 Then
                        dbl数量 = Format(lng费用次数 * Nvl(rsMoney!数量, 0), "0.00000")
                    Else
                        dbl数量 = Format(Val(rsAdvice!总量 & "") * Nvl(rsMoney!数量, 0), "0.00000")
                    End If
                    
                    '确定时价卫材价格
                    If Nvl(rsMoney!是否变价, 0) = 1 Then
                        dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, dbl数量, , True, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                    Else
                        dbl单价 = Format(Nvl(rsMoney!单价, 0), gstrDecPrice)
                    End If
                Else
                    '总量等于单次用量乘数次。一天只收一次时，有多少天要执行，就收多少次，不管单次用量（例如：每天两次）,但要管收费对照的次数
                    int付数 = 1
                    If InStr(",1,2,3,4,5,6,7,", Val("" & rsMoney!收费方式)) > 0 Then
                        dbl数量 = Format(lng费用次数 * Nvl(rsMoney!数量, 0), "0.00000")
                    Else
                        dbl数量 = Format(Val(rsAdvice!总量 & "") * Nvl(rsMoney!数量, 0), "0.00000")
                    End If
                    dbl单价 = Format(Nvl(rsMoney!单价, 0), gstrDecPrice)
                End If
                
                '非药嘱药品及跟踪卫材的库存检查
                If InStr(",4,5,6,7", rsAdvice!诊疗类别 & "") = 0 _
                    And (rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsMoney!类别) > 0) Then
                    '调用药品接口（待完成）
                End If
                    
                '发送金额
                dbl应收 = int付数 * dbl数量 * dbl单价
                If bln附加手术 Then
                    dbl应收 = dbl应收 * Nvl(rsMoney!附术收费率, 100) / 100
                End If
                
                '处理加班加价
                If bln加班加价 And Nvl(rsMoney!加班加价, 0) = 1 Then
                    dbl应收 = dbl应收 * (1 + Nvl(rsMoney!加班加价率, 0) / 100)
                End If
                
                cur应收 = Format(dbl应收, gstrDec)

                'NO,序号
                Call GetCurBillSet(rsBill, strNOKey, strNO, lng费用序号, -1, int发送方式 = 0, lngNOSequence)
                rsSQL.AddNew: blnBool = False
                If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                    lng费用父号 = lng费用序号
                    If rsMoney!从项 = 0 Then
                        '记录主项信息，主项肯定在从项前
                        '即使不汇总折扣，也要记录主从项关系
                        If InStr(strHaveSub & ",", "," & rsMoney!费用性质 & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsMoney!费用性质 & ",") = 0 Then
                            rsClone.Filter = "费用性质=" & rsMoney!费用性质 & " And 从项=1"
                            If Not rsClone.EOF Then
                                int父序号 = lng费用序号
                                lng父项目ID = rsMoney!ID
                                
                                rsSeek.AddNew
                                rsSeek!费用性质 = rsMoney!费用性质
                                rsSeek!主项标签 = rsSQL.Bookmark 'Variant(Double)
                                rsSeek!主收入ID = rsMoney!收入项目ID
                                rsSeek.Update
                                strHaveSub = strHaveSub & "," & rsMoney!费用性质
                                
                                blnBool = True
                            Else
                                strNoneSub = strNoneSub & "," & rsMoney!费用性质
                            End If
                        End If
                    End If
                End If
                
                '计算汇总折扣合计
                If gbln从项汇总折扣 And (rsMoney!从项 = 1 Or InStr(strHaveSub & ",", "," & rsMoney!费用性质 & ",") > 0) Then
                    cur实收 = cur应收
                    
                    '累计医嘱合计来计算折扣
                    rsSeek.Filter = "费用性质=" & rsMoney!费用性质
                    rsSeek!合计 = Nvl(rsSeek!合计, 0) + cur实收
                    rsSeek.Update
                ElseIf Nvl(rsMoney!屏蔽费别, 0) = 0 Then
                    cur实收 = Format(ActualMoney(Nvl(rsAdvice!费别), rsMoney!收入项目ID, cur应收, rsMoney!ID, lng执行科室ID, int付数 * dbl数量, _
                        IIF(bln加班加价 And Nvl(rsMoney!加班加价, 0) = 1, Nvl(rsMoney!加班加价率, 0) / 100, 0)), gstrDec)
                Else
                    cur实收 = cur应收
                End If
                '汇总折扣时，对主项的实收金额作特殊处理
                If gbln从项汇总折扣 And blnBool Then
                    str实收 = Chr(0) & Chr(1) & "Begin" & cur实收 & "End" & Chr(0) & Chr(1)
                Else
                    str实收 = cur实收
                End If
                
                '医保相关字段
                bln保险项目否 = False: lng保险大类ID = 0: cur统筹金额 = 0: str保险编码 = "": str费用类型 = ""
                If Not IsNull(rsAdvice!险类) Then
                    strTmp = gclsInsure.GetItemInsure(Val(rsAdvice!病人ID & ""), rsMoney!ID, cur实收, False, Val(rsAdvice!险类 & ""), Val(rsAdvice!摘要 & "") & "||" & int付数 * dbl数量)
                    If strTmp <> "" Then
                        bln保险项目否 = Val(Split(strTmp, ";")(0)) <> 0
                        lng保险大类ID = Val(Split(strTmp, ";")(1))
                        cur统筹金额 = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                        str保险编码 = CStr(Split(strTmp, ";")(3))
                        If UBound(Split(strTmp, ";")) >= 5 Then
                            If Split(strTmp, ";")(5) <> "" Then
                                str费用类型 = Split(strTmp, ";")(5)
                            End If
                        End If
                    End If
                End If
                
                '收集记帐报警类别
                cur合计 = cur合计 + cur实收
                If InStr(str类别, rsMoney!类别) = 0 Then
                    str类别 = str类别 & rsMoney!类别
                    str类别名称 = str类别名称 & "," & rsMoney!类别名称
                End If
                
                '发生时间
                If str分解时间 <> "" Then
                    str发生时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    str发生时间 = "To_Date('" & rsAdvice!开始执行时间 & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                
                '获取用法
                str用法 = ""
                If InStr(",5,6,C,", rsAdvice!诊疗类别) > 0 Then
                    '西药用法,检验采集方式
                    lngTmpEOF = rsAdvice.AbsolutePosition
                    rsAdvice.Filter = "ID=" & rsAdvice!相关ID
                    If rsAdvice.RecordCount > 0 Then str用法 = Val(rsAdvice!诊疗项目 & "")
                    rsAdvice.Filter = 0
                    rsAdvice.AbsolutePosition = lngTmpEOF
                ElseIf InStr(",7,", rsAdvice!诊疗类别) > 0 Then
                    '中药用法|中药煎法
                    lngTmpEOF = rsAdvice.AbsolutePosition
                    rsAdvice.Filter = "ID=" & rsAdvice!相关ID
                    If rsAdvice.RecordCount > 0 Then str用法 = Val(rsAdvice!诊疗项目 & "")
                    rsAdvice.Filter = "相关ID=" & rsAdvice!相关ID & " And 诊疗类别='E'"
                    If rsAdvice.RecordCount > 0 Then str用法 = str用法 & "|" & Val(rsAdvice!诊疗项目 & "")
                    rsAdvice.Filter = 0
                    rsAdvice.AbsolutePosition = lngTmpEOF
                ElseIf InStr(",K,", rsAdvice!诊疗类别) > 0 Then
                    '输血途径
                    lngTmpEOF = rsAdvice.AbsolutePosition
                    rsAdvice.Filter = "相关ID=" & rsAdvice!ID
                    If rsAdvice.RecordCount > 0 Then str用法 = Val(rsAdvice!诊疗项目 & "")
                    rsAdvice.Filter = 0
                    rsAdvice.AbsolutePosition = lngTmpEOF
                End If
                                        
                '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                rsSQL!类型 = 4: rsSQL!项目ID = rsMoney!ID: rsSQL!序号 = i
                rsSQL!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID & ""), Val(rsAdvice!相关ID & ""))
                rsSQL!NO = strNO
                If int发送方式 = 0 And Val(rsAdvice!病人来源 & "") <> 2 Then
                    '暂未取发药窗口
                    rsSQL!sql = "ZL_门诊划价记录_INSERT(" & _
                        "'" & strNO & "'," & lng费用序号 & "," & Val(rsAdvice!病人ID & "") & "," & ZVal(Val(rsAdvice!主页ID & "")) & ",'" & _
                        IIF(Val(rsAdvice!病人来源 & "") <> 2, Nvl(rsAdvice!门诊号), Nvl(rsAdvice!住院号)) & "','" & IIF(Val(rsAdvice!病人来源 & "") <> 2, "", Nvl(rsAdvice!当前床号)) & "','" & Nvl(rsAdvice!姓名) & "'," & _
                        "'" & Nvl(rsAdvice!性别) & "','" & Nvl(rsAdvice!年龄) & "'," & _
                        "'" & Nvl(rsAdvice!费别) & "',NULL," & _
                        ZVal(rsAdvice!病人科室ID & "") & "," & ZVal(rsAdvice!开嘱科室ID & "") & "," & _
                        "'" & GetAuditName(rsAdvice!开嘱医生) & "'," & IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                        rsMoney!ID & ",'" & rsMoney!类别 & "','" & Nvl(rsMoney!计算单位) & "',NULL," & _
                        int付数 & "," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                        IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                        "'" & Nvl(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                        str发生时间 & "," & strCurDate & "," & _
                        "'医嘱发送','" & UserInfo.姓名 & "'," & _
                        "'" & rsAdvice!医嘱内容 & "'," & Val(rsAdvice!ID & "") & ",'" & rsAdvice!执行频次 & "" & "'," & _
                        ZVal(rsAdvice!单量 & "") & ",'" & str用法 & "',1," & _
                        IIF(int药品性质 <> 0, int药品性质, Val(rsAdvice!计价特性 & "")) & "," & IIF(Val(rsAdvice!病人来源 & "") <> 2, "1", "2") & "," & _
                        "'" & str保险编码 & "','" & str费用类型 & "'," & IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ")"
                        
                ElseIf int发送方式 = 1 And Val(rsAdvice!病人来源 & "") <> 2 Then
                    '是否划价费用
                    If InStr(",5,6,7,", rsAdvice!诊疗类别 & "") > 0 Then
                        int划价 = IIF(InStr(gstr门诊发送划价单, "5") > 0, 1, 0)
                    Else
                        int划价 = IIF(InStr(gstr门诊发送划价单, rsAdvice!诊疗类别 & "") > 0, 1, 0)
                    End If
                    If int划价 = 0 Then int划价 = IIF(Nvl(rsMoney!费用确认, 0) = 1, 1, 0)
                    
                    If int划价 = 0 Then
                        bln划价 = False
                        If gdbl预存款消费验卡 <> 0 Then cur记帐合计 = cur记帐合计 + cur实收
                    End If
                    
                    '登记时间
                    If int划价 = 1 Then '与非划价的时间上区分开
                        str登记时间 = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        str登记时间 = strCurDate
                    End If
                
                    rsSQL!sql = "zl_门诊记帐记录_INSERT(" & _
                        "'" & strNO & "'," & lng费用序号 & "," & Val(rsAdvice!病人ID & "") & ",'" & _
                        IIF(Val(rsAdvice!病人来源 & "") <> 2, Nvl(rsAdvice!门诊号), Nvl(rsAdvice!住院号)) & "','" & rsAdvice!姓名 & "'," & _
                        "'" & rsAdvice!性别 & "','" & rsAdvice!年龄 & "'," & "'" & rsAdvice!费别 & "',0," & Val(rsAdvice!婴儿 & "") & "," & _
                        ZVal(rsAdvice!病人科室ID & "") & "," & ZVal(rsAdvice!开嘱科室ID & "") & "," & _
                        "'" & GetAuditName(rsAdvice!开嘱医生) & "'," & IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                        rsMoney!ID & ",'" & rsMoney!类别 & "','" & rsMoney!计算单位 & "'," & _
                        int付数 & "," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                        IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                        "'" & rsMoney!收据费目 & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                        str发生时间 & "," & str登记时间 & "," & _
                        "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                        "Null,'" & rsAdvice!医嘱内容 & "'," & Val(rsAdvice!ID & "") & ",'" & rsAdvice!执行频次 & "" & "'," & _
                        ZVal(rsAdvice!单量 & "") & ",'" & str用法 & "',1," & _
                        IIF(int药品性质 <> 0, int药品性质, Val(rsAdvice!计价特性 & "")) & ",1,Null,0,Null)"
                    '门诊标志要填1才是门诊记帐费用，填2就是住院记帐费用了
                Else
                    '是否划价费用
                    strTmp = rsAdvice!执行科室ID
                    If InStr(",5,6,7,", rsAdvice!诊疗类别 & "") > 0 Then
                        int划价 = IIF(InStr(gstr住院发送划价单, "5") > 0, 1, 0)
                        
                        '药品医嘱领药部门id为给药途径的执行科室id,72503
                        lngTmpEOF = rsAdvice.AbsolutePosition
                        rsAdvice.Filter = "ID=" & rsAdvice!相关ID
                        If rsAdvice.RecordCount > 0 Then strTmp = Val(rsAdvice!执行科室ID & "")
                        rsAdvice.Filter = 0
                        rsAdvice.AbsolutePosition = lngTmpEOF
                    Else
                        int划价 = IIF(InStr(gstr住院发送划价单, rsAdvice!诊疗类别 & "") > 0, 1, 0)
                    End If
                    If int划价 = 0 Then int划价 = IIF(Nvl(rsMoney!费用确认, 0) = 1, 1, 0)
                    
                    If int划价 = 0 Then
                        bln划价 = False
                    End If
                    
                    '登记时间
                    If int划价 = 1 Then '与非划价的时间上区分开
                        str登记时间 = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        str登记时间 = strCurDate
                    End If
                    
                    '收集医保上传单据号:mrsBill中的不一定产生了费用
                    If int划价 = 0 Then
                        rsUpload.Filter = "NO='" & strNO & "'"
                        If rsUpload.EOF Then
                            rsUpload.AddNew
                            rsUpload!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID & ""), Val(rsAdvice!相关ID & ""))
                            rsUpload!NO = strNO
                            rsUpload.Update
                        End If
                    End If
                    
                    rsSQL!sql = "ZL_住院记帐记录_Insert(" & _
                        "'" & strNO & "'," & lng费用序号 & "," & Val(rsAdvice!病人ID & "") & "," & ZVal(Val(rsAdvice!主页ID & "")) & "," & _
                        IIF(IsNull(rsAdvice!住院号), "NULL", "'" & rsAdvice!住院号 & "'") & ",'" & Nvl(rsAdvice!姓名) & "'," & _
                        "'" & Nvl(rsAdvice!性别) & "','" & Nvl(rsAdvice!年龄) & "'," & _
                        "'" & Nvl(rsAdvice!当前床号) & "','" & Nvl(rsAdvice!费别) & "'," & _
                        ZVal(rsAdvice!病人病区ID) & "," & ZVal(rsAdvice!病人科室ID & "") & ",0," & _
                        Val(rsAdvice!婴儿) & "," & _
                        ZVal(rsAdvice!开嘱科室ID & "") & ",'" & GetAuditName(rsAdvice!开嘱医生) & "'," & _
                        IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & rsMoney!ID & "," & _
                        "'" & rsMoney!类别 & "','" & Nvl(rsMoney!计算单位) & "'," & _
                        IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                        int付数 & "," & dbl数量 & ",NULL," & ZVal(lng执行科室ID) & "," & _
                        IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                        "'" & Nvl(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                        cur统筹金额 & "," & str发生时间 & "," & str登记时间 & "," & _
                        "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & _
                        IIF(rsMoney!类别 = "4", lng卫材类别ID, lng药品类别ID) & "," & _
                        "NULL,'" & rsAdvice!医嘱内容 & "',NULL," & Val(rsAdvice!ID & "") & "," & _
                        "'" & rsAdvice!执行频次 & "" & "'," & ZVal(rsAdvice!单量 & "") & "," & _
                        "'" & str用法 & "',1," & _
                        IIF(int药品性质 <> 0, int药品性质, Val(rsAdvice!计价特性 & "")) & "," & _
                        "Null,'" & str费用类型 & "',Null," & strTmp & ")"
                End If
                rsSQL.Update
                
                '记录自动发料的SQL
                If ((gbyt住院自动发料 = 1 Or gbyt住院自动发料 = 2 And lng执行科室ID = Val(rsAdvice!开嘱科室ID & "")) And Val(rsAdvice!病人来源 & "") = 2 And int发送方式 = 1 Or gbln门诊自动发料 And int发送方式 = 1 And Val(rsAdvice!病人来源 & "") <> 2) _
                    And int划价 = 0 And lng执行科室ID <> 0 And rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Then
                    If InStr(str自动发料 & ";", ";" & strNO & "," & lng执行科室ID & ";") = 0 Then
                        rsSQL.AddNew
                        rsSQL!类型 = 6
                        rsSQL!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID & ""), Val(rsAdvice!相关ID & ""))
                        rsSQL!项目ID = 0
                        rsSQL!序号 = i
                        rsSQL!NO = strNO
                        rsSQL!sql = "zl_材料收发记录_处方发料(" & lng执行科室ID & ",25,'" & strNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                        rsSQL.Update
                        str自动发料 = str自动发料 & ";" & strNO & "," & lng执行科室ID
                    End If
                End If
                
                '医保管控实时监测：生成费用项目记录集,以收费细目汇总
                If Not IsNull(rsAdvice!险类) And bln实时监控 Then
                    rsItems.Filter = "收费细目ID=" & rsMoney!ID
                    If rsItems.EOF Then
                        '加入收费项目对应的原始信息
                        rsItems.AddNew
                        rsItems!病人ID = Val(rsAdvice!病人ID & "")
                        rsItems!主页ID = Val(rsAdvice!主页ID & "")
                        rsItems!医嘱ID = Val(rsAdvice!ID & "")
                        rsItems!收费类别 = rsMoney!类别
                        rsItems!收费细目ID = rsMoney!ID
                        rsItems!开单人 = rsAdvice!开嘱医生
                        rsItems!开单科室 = CStr(gobjComlib.Sys.RowValue("部门表", Val(rsAdvice!开嘱科室ID & ""), "名称"))
                        
                        rsItems!数量 = int付数 * dbl数量
                        rsItems!单价 = dbl单价
                    Else
                        '基于一个医嘱(诊疗项目)的收费对照不会有重复的收费细目
                        '数量：同一收费项目的不同收入项目记录相同
                        If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                            rsItems!数量 = Nvl(rsItems!数量, 0) + int付数 * dbl数量
                        End If
                        '单价：同一收费项目的不同收入项目累加
                        If Val(rsAdvice!ID & "") = rsItems!医嘱ID Then
                            rsItems!单价 = Nvl(rsItems!单价, 0) + dbl单价
                        End If
                    End If
                    rsItems!实收金额 = Nvl(rsItems!实收金额, 0) + cur实收
                    rsItems.Update
                End If
                    
                str收费项目 = rsMoney!费用性质 & "_" & rsMoney!ID
                rsMoney.MoveNext
            Loop
        End If
        
        '对医嘱金额进行汇总折扣处理
        If gbln从项汇总折扣 And strHaveSub <> "" Then
            rsSeek.Filter = 0
            Do While Not rsSeek.EOF
                rsSQL.Bookmark = rsSeek!主项标签
                cur实收 = Format(ActualMoney(Nvl(rsAdvice!费别), rsSeek!主收入ID, rsSeek!合计), gstrDec)
                cur实收 = cur实收 - rsSeek!合计 '打折差额
                
                '医保管控实时监测：费用项目金额替换
                If Not IsNull(rsAdvice!险类) And bln实时监控 Then
                    rsItems.Filter = "收费细目ID=" & lng父项目ID
                    If Not rsItems.EOF Then
                        rsItems!实收金额 = Nvl(rsItems!实收金额, 0) + cur实收
                        rsItems.Update
                    End If
                End If
                
                '费用SQL生成替换
                cur实收 = Get实收金额(rsSQL!sql) + cur实收
                rsSQL!sql = Set实收金额(rsSQL!sql, cur实收)
                rsSQL.Update
            
                rsSeek.MoveNext
            Loop
        End If
        
        '选择要发送的医嘱自动进行校对(实际可能因为叮嘱不发送)
        If Val(rsAdvice!医嘱状态 & "") = 1 And Val(rsAdvice!相关ID & "") = 0 Then
            rsSQL.AddNew
            rsSQL!类型 = 3: rsSQL!项目ID = 0: rsSQL!序号 = i
            rsSQL!医嘱ID = Val(rsAdvice!ID & "")
            rsSQL!sql = "ZL_病人医嘱记录_校对(" & Val(rsAdvice!ID & "") & ",3," & _
                "To_Date('" & Format(rsAdvice!开嘱时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                "NULL,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If
        
        
        '产生医嘱发送记录
        '-----------------------------------------------------------------------------------------
        If Val(rsAdvice!执行性质 & "") <> 0 Then '叮嘱不发送(给药途径，配方煎法、用法,采集方法、输血途径可能为)
            '发送了出院,转院,死亡医嘱
            If rsAdvice!诊疗类别 & "" = "Z" _
                And InStr(",5,6,11,", Val(rsAdvice!操作类型 & "")) > 0 Then
            End If
            
            '一样要产生费用NO
            Call GetCurBillSet(rsBill, strNOKey, strNO, -1, lng发送序号, int发送方式 = 0, lngNOSequence)
                                                    
            '是否一组医嘱的第一医嘱行:药疗的第一药品行为第一医嘱行
            If rsAdvice.AbsolutePosition = 1 Then
                blnFirst = True
            Else
                blnFirst = False
            End If
                        
            '发送数次:药品为剂量单位的总量,其它为次数
            If rsAdvice!诊疗类别 & "" = "7" Then
                dbl发送数次 = Val(rsAdvice!总量 & "") * Val(rsAdvice!单量 & "")
            ElseIf InStr(",5,6,", rsAdvice!诊疗类别 & "") > 0 Then
                dbl发送数次 = Val(rsAdvice!总量 & "") * Val(IIF(Val(rsAdvice!病人来源 & "") = 2, rsAdvice!住院包装, rsAdvice!门诊包装) & "") * Val(rsAdvice!剂量系数 & "")
            Else
                dbl发送数次 = Val(rsAdvice!总量 & "")
            End If
            dbl发送数次 = Format(dbl发送数次, "0.00000")
                                                    
            '首末时间(不能用“str分解时间”判断，因为一次性临嘱记录的是首次时间)
            If str分解时间 <> "" Then
                str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                str末次时间 = "To_Date('" & Split(str分解时间, ",")(Val(UBound(Split(str分解时间, ",")))) & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                '无法分解或为"一次性"临嘱
                str首次时间 = "To_Date('" & Format(rsAdvice!开始执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                str末次时间 = "To_Date('" & Format(rsAdvice!开始执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                str分解时间 = Format(rsAdvice!开始执行时间, "yyyy-MM-dd HH:mm:ss")
            End If
                               
            rsSQL.AddNew
            rsSQL!类型 = 5: rsSQL!项目ID = 0: rsSQL!序号 = i
            rsSQL!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID & ""), Val(rsAdvice!相关ID & ""))
            rsSQL!NO = strNO
            rsSQL!sql = "ZL_病人医嘱发送_Insert(" & _
                Val(rsAdvice!ID & "") & "," & lng发送号 & "," & IIF(int发送方式 = 0, 1, 2) & ",'" & strNO & "'," & _
                lng发送序号 & "," & ZVal(dbl发送数次) & "," & str首次时间 & "," & str末次时间 & "," & strCurDate & "," & _
                0 & "," & ZVal(rsAdvice!执行科室ID & "") & "," & int计费状态 & "," & _
                IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.编号 & "'," & _
                "'" & UserInfo.姓名 & "','" & "'," & IIF(int发送方式 = 1 And Val(rsAdvice!病人来源 & "") = 1, 1, "Null") & ",'" & str分解时间 & "')"
            rsSQL.Update
            
            If gbln血库系统 And rsAdvice!诊疗类别 & "" = "K" Then
                rsSQL.AddNew
                rsSQL!类型 = 9 '血库配血费用
                rsSQL!项目ID = 0
                rsSQL!序号 = 0
                rsSQL!sql = "Zl_血液配血费用_Insert(" & Val(rsAdvice!ID & "") & ")"
                rsSQL.Update
            End If
            
            '医嘱执行计价
            If rsExec.RecordCount > 0 Then
                rsExec.Filter = "医嘱ID=" & Val(rsAdvice!ID & "") & " And 发送号=" & lng发送号
                If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                Do While Not rsExec.EOF
                    rsSQL.AddNew
                    rsSQL!类型 = 8
                    rsSQL!项目ID = 0
                    rsSQL!序号 = 0
                    rsSQL!医嘱ID = IIF(Val(rsAdvice!相关ID & "") = 0, Val(rsAdvice!ID & ""), Val(rsAdvice!相关ID & ""))
                    rsSQL!sql = "Zl_医嘱执行计价_Insert(" & rsExec!医嘱ID & "," & rsExec!发送号 & ",To_date('" & _
                    rsExec!要求时间 & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!收费细目ID & "")) & "," & rsExec!数量 & "," & rsExec!费用性质 & ")"
                    rsSQL.Update
                    rsExec.MoveNext
                Loop
                rsExec.Filter = 0
            End If
            
        End If
        
        rsAdvice.MoveNext
    Next
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    '提示未审核项目
    If strAudit <> "" Then
        MsgBox "病人""" & rsAdvice!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, "AdviceSendSingle"
        GoTo errH
    End If
    
    '医保管控实时监测
    If Not IsNull(rsAdvice!险类) And bln实时监控 Then
        rsItems.Filter = 0
        If Not rsItems.EOF Then
            If Not gclsInsure.CheckItem(rsAdvice!险类, 1, 2, rsItems) Then Exit Function
        End If
    End If
    
    '病人费用报警
    If int发送方式 = 1 And cur合计 > 0 Then
        str类别名称 = Mid(str类别名称, 2)
        For i = 1 To Len(str类别)
            If gobjExpense.zlBillingWarn.zlBillingWarnCheck(objParent, IIF(Val(rsAdvice!病人来源 & "") = 2, 2, 1), IIF(int发送方式 = 0, 1, 0), Val(rsAdvice!病人ID & ""), Val(rsAdvice!主页ID & ""), Val(rsAdvice!病人病区ID & ""), Split(str类别名称, ",")(i - 1), IIF(gbln报警包含划价费用, cur合计, cur记帐合计), InStr(";" & GetInsidePrivs(IIF(Val(rsAdvice!病人来源 & "") = 2, p住院医嘱下达, p门诊医嘱下达)) & ";", ";欠费强制记帐;") > 0, False, False, intOutType, False) Then
                If intOutType = 2 Or intOutType = 3 Then Exit Function
            End If
        Next
    End If
    If int发送方式 = 1 And Val(rsAdvice!病人来源 & "") = 1 And gdbl预存款消费验卡 <> 0 And cur记帐合计 > 0 Then
        If Not gobjComlib.zlDatabase.PatiIdentify(Me, glngSys, Val(rsAdvice!病人ID & ""), cur记帐合计, , , , IIF(-1 * gdbl预存款消费验卡 >= Val(cur记帐合计), False, True), , , (gdbl预存款消费验卡 <> 0), (2 = gdbl预存款消费验卡)) Then Exit Function
    End If
    Call InitObjLis(lngModule)
        
    '先调用LIS申请接口
    If Not gobjLIS Is Nothing Then
        If rsAdvice!诊疗类别 & "" = "C" Then
            rsAdvice.Filter = "ID=" & rsAdvice!相关ID
            strAdviceInfo = rsAdvice!ID
            rsAdvice.Filter = 0
            rsAdvice.MoveFirst
            strAdviceInfo = rsAdvice!ID & "," & strAdviceInfo & "," & rsAdvice!执行科室ID & "," & rsAdvice!标本部位
        End If
        If strAdviceInfo <> "" Then
            Set rsTmp = Get病人诊断记录(rsAdvice!病人ID, IIF(Val(rsAdvice!病人来源 & "") = 2, Val(rsAdvice!主页ID & ""), Val(rsAdvice!挂号ID & "")), IIF(Val(rsAdvice!病人来源 & "") = 2, "2", "1"))
            If rsTmp.RecordCount > 0 Then strDiag = rsTmp!诊断描述
        End If
    End If
            
    Call ReplaceTrueNO(rsSQL, rsUpload)
    
    '执行顺序:1-计价,2-签名,3-校对,4-费用,5-发送,6-发料
    '1.先填写费用,因为发送时可能处理费用
    '2.对费用记录按收费细目ID排序插入
    rsSQL.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
    rsSQL.Sort = "类型,项目ID,序号"
    rsUpload.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
    
    If TypeName(arrSQL) = "Empty" Then arrSQL = Array()
    Do While Not rsSQL.EOF
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = rsSQL!sql
        rsSQL.MoveNext
    Loop
    
    AdviceSendSingle = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Sub ShowArchive(ByRef objForm As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnModal As Boolean = False)
'功能：调用电子病案查阅功能
    Call gobjCISJob.ShowArchive(objForm, lng病人ID, lng主页ID, blnModal)
End Sub

Public Function GetArchiveFrom() As Object
'功能：调用电子病案查阅功能，嵌入式获取窗体对象
    Set GetArchiveFrom = gobjCISJob.GetArchiveFrom
End Function

Public Sub zlArchiveRefresh(ByVal lng病人ID As Long, ByVal lng就诊ID As Long)
'功能：刷新电子病案查阅界面
    Call gobjCISJob.zlArchiveRefresh(lng病人ID, lng就诊ID)
End Sub

Public Function zlPrintAdvice(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngBabyNum As Long, _
    ByVal int期效 As Integer, Optional ByVal strPrinterName As String, Optional ByVal intMode As Integer, Optional ByRef objReportForm As Object, _
    Optional ByRef strMsg As String) As Boolean
'功能：医嘱单打印外部调用
    zlPrintAdvice = gobjKernel.zlPrintAdvice(frmParent, lng病人ID, lng主页ID, lngBabyNum, int期效, strPrinterName, intMode, objReportForm, strMsg)
End Function

Private Sub Class_Initialize()
    lngNumPublicAdvice = lngNumPublicAdvice + 1
    If objDockInAdvices Is Nothing Then Set objDockInAdvices = New zlPublicAdvice.clsDockInAdvices
    If objDockOutAdvices Is Nothing Then Set objDockOutAdvices = New zlPublicAdvice.clsDockOutAdvices
    If objMedRecPage Is Nothing Then Set objMedRecPage = New zlPublicAdvice.clsArchiveMedRec
End Sub

Private Sub Class_Terminate()
    lngNumPublicAdvice = lngNumPublicAdvice - 1
    If lngNumPublicAdvice <= 0 Then
        lngNumPublicAdvice = 0
        If Not gobjComlib Is Nothing Then Set gobjComlib = Nothing
        If Not gobjCommFun Is Nothing Then Set gobjCommFun = Nothing
        If Not gobjControl Is Nothing Then Set gobjControl = Nothing
        If Not gobjDatabase Is Nothing Then Set gobjDatabase = Nothing
        If Not gclsInsure Is Nothing Then Set gclsInsure = Nothing
        If Not gobjLIS Is Nothing Then Set gobjLIS = Nothing
        If Not gobjKernel Is Nothing Then Set gobjKernel = Nothing
        If Not gobjCISJob Is Nothing Then Set gobjCISJob = Nothing
    End If
End Sub

Public Function GetAdviceDiag(ByVal lng医嘱ID As Long, Optional ByRef str诊断 As String) As String
'功能：获得医嘱对应的诊断信息
'参数：str诊断=关联诊断的诊断名称字符串
'返回：关联诊断的ID，逗号分隔
    Dim rsTmp As Recordset, strSQL As String
    Dim strReturn As String
    
    strSQL = "Select  A.*" & vbNewLine & _
            "From 病人诊断记录 A, 病人诊断医嘱 B" & vbNewLine & _
            "Where b.诊断id=a.id And  b.医嘱ID=[1]" & vbNewLine & _
            "Order By b.rowID"
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "获取医嘱相关诊断", lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str诊断 = str诊断 & "," & rsTmp!诊断描述 & IIF(Val(rsTmp!是否疑诊 & "") = 1, "(疑)", "")
            strReturn = strReturn & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        str诊断 = Mid(str诊断, 2)
        strReturn = Mid(strReturn, 2)
    End If
    GetAdviceDiag = strReturn
    Exit Function
errH:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckAdviceSkinResult(ByVal lng医嘱ID As Long) As String
'功能：判断医嘱是否标记了皮试结果
'返回：-1表示无需皮试或免试；0表示还未标记皮试结果或未下达皮试医嘱；1表示阴性；2表示阳性
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str皮试IDs As String
    Dim lng病人ID As Long, lng主页ID As Long, str挂号单 As String
    Dim int结果 As Integer

    On Error GoTo errH
    
    '读取医嘱对应的皮试项目
    strSQL = "select B.用法ID,A.病人ID,A.主页ID,A.挂号单 from 病人医嘱记录 A,诊疗用法用量 B where A.诊疗项目ID=B.项目ID And B.性质=0 and A.ID=[1] and A.诊疗类别 IN('5','6')"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckAdviceSkinResult", lng医嘱ID)
    Do While Not rsTmp.EOF
        If str皮试IDs = "" Then
            lng病人ID = Val(rsTmp!病人ID & "")
            lng主页ID = Val(rsTmp!主页ID & "")
            str挂号单 = rsTmp!挂号单 & ""
        End If
        str皮试IDs = str皮试IDs & "," & rsTmp!用法ID
        rsTmp.MoveNext
    Loop
    If str皮试IDs = "" Then CheckAdviceSkinResult = -1: Exit Function
    str皮试IDs = Mid(str皮试IDs, 2)
    
    '读取这些皮试相关的药品医嘱
    strSQL = "Select a.皮试结果,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位" & vbNewLine & _
            "From 病人医嘱记录 A,诊疗项目目录 B" & vbNewLine & _
            "Where A.诊疗项目ID=B.ID And a.病人id = [1] And " & IIF(str挂号单 = "", "A.主页ID=[2]", "a.挂号单 = [3]") & " And Instr([4], ',' || a.诊疗项目id || ',') > 0 And" & vbNewLine & _
            "      a.开始执行时间 >= TRUNC(Sysdate) - [5]" & vbNewLine & _
            "Order By a.开始执行时间 Desc"

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "CheckAdviceSkinResult", lng病人ID, lng主页ID, str挂号单, "," & str皮试IDs & ",", Val(gobjComlib.zlDatabase.GetPara(70, glngSys)))
    If rsTmp.RecordCount > 0 Then
        If rsTmp!皮试结果 & "" = "" Then
            CheckAdviceSkinResult = 0
        Else
            If rsTmp!皮试结果 & "" = "免试" Then
                CheckAdviceSkinResult = -1
            Else
                Call GetTestLabel(rsTmp!标本部位 & "", rsTmp!皮试结果 & "", int结果)
                If int结果 = 0 Then
                    CheckAdviceSkinResult = 1
                ElseIf int结果 = 1 Then
                    CheckAdviceSkinResult = 2
                End If
            End If
        End If
    Else
        CheckAdviceSkinResult = 0
    End If
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetDurgPrice(ByVal lng医嘱ID As Long, ByRef dbl应收单价 As Double) As Double
'功能：计算药品的应收价格
'返回：药品医嘱对应的应收金额
'参数：dbl应收单价返回药品的应收单价
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dblReturn As Double
    
    On Error GoTo errH
    strSQL = "Select 病人ID,主页ID From 病人医嘱记录 Where ID=[1]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDurgPrice", lng医嘱ID)
    
    Call gobjExpense.zlGetPriceGrade(gstrNodeNo, Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)

    
    strSQL = "Select Nvl(a.是否变价, 0) As 是否变价, Decode(Nvl(a.是否变价, 0), 1, b.缺省价格, b.现价) As 单价, c.执行科室id," & vbNewLine & _
            "       Decode(c.病人来源, 2, d.住院包装, d.门诊包装) As 药房包装,C.诊疗类别,C.收费细目ID,C.总给予量" & vbNewLine & _
            "From 病人医嘱记录 C, 药品规格 D, 收费项目目录 A, 收费价目 B" & vbNewLine & _
            "Where c.收费细目id = a.Id And c.收费细目id = d.药品id And a.Id = b.收费细目id And c.Id = [1] " & vbNewLine & _
            GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "2", "3", "4") & _
            "   And   (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      ((Sysdate Between b.执行日期 And b.终止日期) Or (Sysdate >= b.执行日期 And b.终止日期 Is Null))"

    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "GetDurgPrice", lng医嘱ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
    If rsTmp.RecordCount > 0 Then
        If Nvl(rsTmp!是否变价, 0) = 1 Then
            '求药品时价
            dbl应收单价 = CalcDrugPrice(rsTmp!收费细目ID, rsTmp!执行科室ID, Nvl(rsTmp!药房包装, 1), , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            dblReturn = FormatEx(dbl应收单价 * Nvl(rsTmp!总给予量, 0) * Nvl(rsTmp!药房包装, 1), gbytDecPrice)
        Else
            '药嘱药品为药房单价,非药药品为售价
            dbl应收单价 = FormatEx(Nvl(rsTmp!单价), gbytDecPrice)
            dblReturn = FormatEx(Nvl(rsTmp!单价) * Nvl(rsTmp!总给予量, 0) * Nvl(rsTmp!药房包装, 1), gbytDecPrice)
        End If
    End If
    GetDurgPrice = dblReturn
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function SetPhysicalAdviceNos(ByRef rsAdvice As Recordset, ByVal bln记帐 As Boolean) As Boolean
'功能：体检医嘱根据门诊的参数规则产生发送记录NO号
'参数：rsAdvice 医嘱记录集（虚拟的，还未保存的）
'      bln记帐=是否发送为记账单
    Dim intSendNo As Integer
    Dim strNOKey As String
    Dim str单据组合类别 As String
    Dim blnStartTimeDef As Boolean
    Dim lngTmp As Long
    Dim lng主诊疗项目ID As Long, lng主执行科室ID As Long
    Dim rsBill As Recordset, strNO As String

    'Set rsAdvice = gobjComlib.zlDatabase.CopyNewRec(rsTmp)
    
    If rsAdvice.RecordCount > 0 Then
        intSendNo = Val(gobjComlib.zlDatabase.GetPara("发送单据号规则", glngSys, p门诊医嘱下达)) '0-多个,1-单个,2-所有
        str单据组合类别 = gobjComlib.zlDatabase.GetPara("产生为同一单据的医嘱类别", glngSys, p门诊医嘱下达)
        blnStartTimeDef = Val(gobjComlib.zlDatabase.GetPara("开始时间不是同一天的分别产生单据", glngSys, p门诊医嘱下达, 0)) = 1
        
        Set rsBill = New ADODB.Recordset
        rsBill.Fields.Append "Key", adVarChar, 200
        rsBill.Fields.Append "NO", adVarChar, 30
        rsBill.CursorLocation = adUseClient
        rsBill.LockType = adLockOptimistic
        rsBill.CursorType = adOpenStatic
        rsBill.Open
        
        Do While Not rsAdvice.EOF
            If intSendNo = 1 And Not gbln执行前先结算 Then
                strNOKey = "只产生一个单据号"
            ElseIf intSendNo = 2 Then
                strNOKey = rsAdvice!执行科室ID & ""
            Else
                If InStr(str单据组合类别, rsAdvice!诊疗类别 & "") > 0 Then
                    strNOKey = "非药医嘱_" & rsAdvice!诊疗类别 & "_" & Val(rsAdvice!执行科室ID & "")
                    
                ElseIf Val(rsAdvice!相关ID & "") <> 0 And rsAdvice!诊疗类别 = "C" Then
                    '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
                    '同一个类检验型，同一个检验执行科室，同一采集管，同一个采集方式，同一个采集执行科室的检验分配相同的单据号
                    If rsAdvice!相关ID & "" <> "" Then
                        lngTmp = rsAdvice.AbsolutePosition
                        rsAdvice.Filter = "ID=" & rsAdvice!相关ID
                        lng主诊疗项目ID = Val(rsAdvice!诊疗项目ID & "")
                        lng主执行科室ID = Val(rsAdvice!执行科室ID & "")
                        rsAdvice.Filter = 0
                        rsAdvice.AbsolutePosition = lngTmp
                    Else
                        lng主诊疗项目ID = Val(rsAdvice!诊疗项目ID & "")
                        lng主执行科室ID = Val(rsAdvice!执行科室ID & "")
                    End If
                    strNOKey = "一并采集_" & rsAdvice!标本部位 & "_" & _
                        rsAdvice!执行科室ID & "_" & rsAdvice!操作类型 & "_" & rsAdvice!试管编码 & "_" & _
                        lng主诊疗项目ID & "_" & lng主执行科室ID
                ElseIf Val(rsAdvice!相关ID & "") <> 0 And InStr(",F,D,", rsAdvice!诊疗类别 & "") > 0 Then
                    '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
                    strNOKey = "非药医嘱_" & Val(rsAdvice!相关ID & "")
                    
                Else
                    '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法，采集方式，麻醉方式，输血医嘱/输血途径)
                    strNOKey = "非药医嘱_" & Val(rsAdvice!ID & "")
                End If
            End If
            
            '开始时间不是同一天的分别产生单据
            If blnStartTimeDef Then
                strNOKey = strNOKey & "_" & Format(rsAdvice!开始时间 & "", "YYYY-MM-DD")
            End If
            
            '开单人不同的，默认全部分别产生单据
            strNOKey = strNOKey & "_" & rsAdvice!开嘱医生
            
            strNOKey = rsAdvice!病人ID & "_" & rsAdvice!挂号单 & "_" & strNOKey
            
            rsBill.Filter = "Key='" & strNOKey & "'"
            If rsBill.EOF Then
                strNO = gobjComlib.zlDatabase.GetNextNo(IIF(bln记帐, 14, 13))
                rsAdvice!NO = strNO
                rsBill.AddNew
                rsBill!NO = strNO
                rsBill!Key = strNOKey
                rsBill.Update
            Else
                strNO = rsBill!NO
                rsAdvice!NO = strNO
            End If
            rsAdvice.Update
            rsAdvice.MoveNext
        Loop
        rsAdvice.MoveFirst
    End If
    SetPhysicalAdviceNos = True
End Function

Private Sub clsDis_PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String)
     RaiseEvent PatiTransfer(lng病人ID, str挂号No)
End Sub

Public Sub HideFrm(ByVal lngFrmID As Long)
'功能：隐藏窗体
'参数：lngFrmID,窗体标识，0－表示 frmDisRegist
    Call clsDis.HideFrm(lngFrmID)
End Sub

Private Sub clsDis_Closed(ByVal lngFunID As Long, ByVal strTag As String)
    RaiseEvent Closed(lngFunID, strTag)
End Sub

Public Function ShowDisRegist(ByRef frmParent As Object, ByVal intType As Integer, Optional ByVal lngID As Long, _
                Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号No As String, _
                Optional ByVal lng医嘱ID As Long, Optional ByVal var登记科室 As Variant, Optional ByVal dat送检时间 As Date, Optional ByVal var送检科室 As Variant, _
                Optional ByVal str送检医生 As String, Optional ByVal str标本名称 As String, Optional ByVal str反馈结果 As String, _
                Optional ByVal dat检查时间 As Date, Optional ByVal str疑似传染病 As String, Optional ByRef objMip As Object, Optional ByVal str登记人 As String) As Integer
'功能：调用传染病阳性结果反馈单
'参数： intType 0表示填写（只显示上半部分），1-表示医生处理（只有下半部分可编辑），2-查看（所有不可编辑，可查看下半部分）
'       lngID  = 疾病阳性记录 ID
'       lng病人ID = 病人ID
'       lng主页ID=住院:主页ID
'       str挂号No =门诊：挂号单NO
'       lng医嘱ID =要填写反馈单的医嘱的医嘱ID
'       var登记科室 = 当前填写反馈单的科室ID或者科室编码
'       dat送检时间 = 标本到达的时间
'       var送检科室 = 让病人前来检查的科室ID或者科室编码
'       str送检医生 = 让病人前来检查的医生
'       str标本名称 = 病人检测所用的标本
'       str反馈结果 = 当前填写的反馈结果
'       dat检查时间 = 病人的检查时间
'       str疑似传染病  病人的疑似传染病
'       objMip  消息对象
'      如果intType为0的话，lng病人ID,lng主页ID,str挂号No必填
'      如果intType为1或者2的话，lngID必填
'返回值：1：只在intType 为0表示填写时才会出现，表示已经填写好反馈单发送成功；
'        2：只在intType 为1表示医生处理时才会出现，表示已经填写处理说明，反馈单已经处理完成；
'        3：只在intType 为1表示医生处理时才会出现，表示不处理反馈单，进行了转科处理；
'        其他情况返回值为0
    ShowDisRegist = clsDis.ShowDisRegist(frmParent, intType, lngID, lng病人ID, lng主页ID, str挂号No, lng医嘱ID, var登记科室, dat送检时间, var送检科室, str送检医生, str标本名称, str反馈结果, dat检查时间, str疑似传染病, objMip, str登记人)
End Function

Public Function ShowDisQuery(ByVal var登记科室 As Variant) As Long
'功能：调用传染病阳性结果查询界面
'参数：var登记科室=疾病阳性记录 的 登记科室ID 或者 登记科室的编码
'返回值：返回查询到的记录数量
    ShowDisQuery = clsDis.ShowDisQuery(var登记科室)
End Function

Public Function ShowRegistByPati(ByRef frmParent As Object, ByVal intType As Integer, ByVal lng病人ID As Long, _
                            Optional ByVal lng主页ID As Long, Optional ByVal str挂号No As String, Optional ByVal var登记科室 As Variant) As Boolean
'功能：根据传入的病人ID和主页ID（门诊的话是挂号单NO），查询该病人的阳性结果反馈单
'参数：intType，1=医生处理；2-查看；3-修改
'      lng病人ID = 病人ID
'      lng主页ID=住院:主页ID
'      str挂号No =门诊：挂号单NO
'      var登记科室 = 填写反馈单的科室
    ShowRegistByPati = clsDis.ShowRegistByPati(frmParent, intType, lng病人ID, lng主页ID, str挂号No, var登记科室)
End Function

Public Function InitDisease(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, Optional ByRef frmMain As Object, _
        Optional ByVal lngModule As Long, Optional ByVal MainPrivs As String, Optional ByRef objMip As Object) As Boolean
'功能：初始化zl9Disease.clsDisease ，之所以单独写出来是为了LIS单独安装时好调用
    If clsDis Is Nothing Then
         Set clsDis = New zl9Disease.clsDisease
         Call clsDis.InitDisease(cnOracle, frmMain, lngSys, lngModule, MainPrivs, objMip)
    End If
End Function

Public Function GetArchiveOneFrom() As Object
'功能：调用电子病案查阅功能，单个病人查阅的窗体
    Set GetArchiveOneFrom = gobjCISJob.GetArchiveOneFrom
End Function

Public Sub ViewEPRReport(frmParent As Object, ByVal lng报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjCISJob.ViewEPRReport(frmParent, lng报告ID, CanPrint)
End Sub

Public Sub ViewPACSImage(frmParent As Object, ByVal lng医嘱ID As Long, ByVal blnMoved As Boolean)
'功能：PACS观片处理
    Call gobjCISJob.ViewPACSImage(frmParent, lng医嘱ID, blnMoved)
End Sub

Public Function OutAdviceSendDrug(frmParent As Object, ByVal str给药IDs As String, ByVal lng病人ID As Long, ByVal lng挂号id As Long, Optional ByRef lng发送号 As Long) As Boolean
'功能：处理门诊医嘱自动发送
'说明：处理处方审查自动发送医嘱，许多功能不支持，另外涉及到提示不成功的地方不提示不允许发送，业务上这类医嘱需要到医生那里去医生处理（作废重开等）；
'返回：如果成功则返回true，否则返回false
'      lng发送号；返回本次发送发送号；
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsNumber As ADODB.Recordset '用于生成条码的动态记录集
    Dim rsItems As ADODB.Recordset '用于医保管控的费用记录集,动态记录集
    Dim rsMoneyNow As ADODB.Recordset '当前病人本次要发送的费用,动态记录集
    Dim rsMoneyDay As ADODB.Recordset '当前病人当天已发送的费用,静态记录集
    Dim rsExec As ADODB.Recordset  '医嘱执行计价

    Dim rsTmp As ADODB.Recordset
    Dim rsPrice As ADODB.Recordset
    Dim rsMoney As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date
    Dim blnTran As Boolean, blnBool As Boolean
    Dim str类别 As String, lngTmp As Long

    Dim bln划价 As Boolean, int划价 As Integer, strTmp As String
    Dim int计费状态 As Integer, strNO As String
    Dim str收费项目 As String, lng费用序号 As Long, lng费用父号 As Long, lng发送序号 As Long
    Dim int付数 As Integer, dbl数量 As Double, cur合计 As Currency, cur记帐合计 As Currency
    Dim dbl单价 As Double, dbl应收 As Double, cur应收 As Currency, cur实收 As Currency
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim int配方数 As Integer, strNOKey As String, str自动发料 As String, strPre诊疗单据ID As String
    Dim str发生时间 As String, str登记时间 As String
    Dim dbl发送数次 As Double, blnFirst As Boolean '配方数及分号关键字
    Dim lng药品类别ID As Long, lng卫材类别ID As Long
    Dim lng执行科室ID As Long, int执行状态 As Integer
    Dim bln离院带药 As Boolean, bln附加手术 As Boolean, str费别 As String
    Dim str用法 As String

    Dim rsClone As ADODB.Recordset
    Dim rsSeek As ADODB.Recordset
    Dim strNoneSub As String, strHaveSub As String
    Dim int父序号 As Integer, lng父项目ID As Long, str实收 As String
    Dim bln保险项目否 As Boolean, lng保险大类ID As Long, str保险编码 As String, str费用类型 As String

    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean

    '电子签名
    Dim lng组ID As Long, str医嘱IDs As String, strSource As String
    Dim intRule As Integer, strSign As String, strTimeStamp As String, strTimeStampCode As String
    Dim lng证书ID As Long, lng签名id As Long

    Dim strCuvetteNumber As String  '生成条码
    Dim bln实时监控 As Boolean, rs医嘱诊断 As ADODB.Recordset
    Dim lng费用次数 As Long '一天只收一次时，本次发送应收取的费用次数
    Dim lng医嘱ID As Long
    Dim str医嘱诊断ids As String
    Dim lng主医嘱行 As Long
    Dim lng采集科室ID As Long
    Dim str医嘱内容 As String
    Dim str摘要 As String
    Dim rsPati As New Recordset
    Dim rsBill As ADODB.Recordset
    Dim rsRXKey As ADODB.Recordset
    Dim rsSend As New Recordset
    Dim str挂号单 As String
    Dim intSendNo As Integer
    Dim str单据组合类别 As String
    Dim blnNOCtrl As Boolean
    Dim blnStartTimeDef As Boolean
    Dim bln一并给药发送为一张 As Boolean
    Dim objfrm As New frmMain
    Dim lngNOSequence As Long
    Dim colStock1 As Collection
    Dim colStock2 As Collection
    Dim lng次数 As Long, lng最小次数 As Long, blnTmp As Boolean
    Dim str诊断 As String, rsDiag As Recordset

    On Error GoTo errH
    
    intSendNo = Val(gobjComlib.zlDatabase.GetPara("发送单据号规则", glngSys, p门诊医嘱下达)) '0-多个,1-单个,2-所有
    str单据组合类别 = gobjComlib.zlDatabase.GetPara("产生为同一单据的医嘱类别", glngSys, p门诊医嘱下达)

    '不同诊断的医嘱分别产生单据
    blnNOCtrl = Val(gobjComlib.zlDatabase.GetPara("不同诊断的医嘱分别产生单据", glngSys, p门诊医嘱下达, 0)) = 1
    
    '开始时间不是同一天的分别产生单据
    blnStartTimeDef = Val(gobjComlib.zlDatabase.GetPara("开始时间不是同一天的分别产生单据", glngSys, p门诊医嘱下达, 0)) = 1
    
    bln一并给药发送为一张 = Val(gobjComlib.zlDatabase.GetPara("一并给药发送为一张", glngSys, p门诊医嘱下达, 1)) = 1
    '检查医嘱关联诊断的填写
    str诊断 = gobjComlib.zlDatabase.GetPara("要求输入门诊诊断", glngSys, p门诊医嘱下达)
    
    If str诊断 <> "" Then
        If InStr(str诊断, "5") > 0 Then
            strSQL = "Select /*+cardinality(c,10)*/" & vbNewLine & _
                "  b.医嘱id" & vbNewLine & _
                "From Table(f_Num2list([1])) C,  病人诊断医嘱 B" & vbNewLine & _
                "Where b.医嘱id(+) = c.Column_Value And B.医嘱ID is Null"
    
            Set rsDiag = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "OutAdviceSendDrug", str给药IDs)
            If rsDiag.RecordCount > 0 Then
                GoTo FuncEnd
            End If
        End If
    End If
    
    Set colStock1 = GetStockCheck(0)
    Set colStock2 = GetStockCheck(1)
    Call gobjExpense.zlGetPriceGrade(gstrNodeNo, lng病人ID, 0, "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
    
    
    
    Set rsExec = New ADODB.Recordset

    rsExec.Fields.Append "医嘱ID", adBigInt
    rsExec.Fields.Append "发送号", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "要求时间", adDate, , adFldIsNullable
    rsExec.Fields.Append "数量", adDouble, , adFldIsNullable
    rsExec.Fields.Append "费用性质", adInteger, , adFldIsNullable

    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open

    '品时读取药品入出类别
    lng药品类别ID = ExistIOClass(8)
    If lng药品类别ID = 0 Then
        'MsgBox "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！", vbInformation, "处方自动发送"
        Exit Function
    End If
    lng卫材类别ID = ExistIOClass(40) '不能确定是否使用了卫材收费,后面再判断

    Screen.MousePointer = 11

    bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
    bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True

    Call GetOutPatiInfo(rsPati, lng病人ID, lng挂号id)
    str挂号单 = rsPati!挂号单 & ""
    If Not IsNull(rsPati!险类) Then
        bln实时监控 = gclsInsure.GetCapability(support实时监控, lng病人ID, rsPati!险类)
        If bln实时监控 Then
            strSQL = "Select A.诊断ID,A.疾病ID,B.医嘱ID From 病人诊断记录 A,病人诊断医嘱 B Where A.病人ID=[1] And 主页ID=[2] And A.ID=B.诊断ID"
            Set rs医嘱诊断 = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "SendAdvice", lng病人ID, lng挂号id)
        End If
    Else
        bln实时监控 = False
    End If

    Set rsBill = New ADODB.Recordset
    rsBill.Fields.Append "Key", adVarChar, 200
    rsBill.Fields.Append "NO", adVarChar, 30
    rsBill.Fields.Append "费用序号", adBigInt
    rsBill.Fields.Append "发送序号", adBigInt
    rsBill.CursorLocation = adUseClient
    rsBill.LockType = adLockOptimistic
    rsBill.CursorType = adOpenStatic
    rsBill.Open

    Set rsRXKey = New ADODB.Recordset
    rsRXKey.Fields.Append "Key", adVarChar, 200
    rsRXKey.Fields.Append "医嘱ID", adVarChar, 200
    rsRXKey.Fields.Append "条数", adBigInt
    rsRXKey.Fields.Append "张数", adBigInt
    rsRXKey.CursorLocation = adUseClient
    rsRXKey.LockType = adLockOptimistic
    rsRXKey.CursorType = adOpenStatic
    rsRXKey.Open

    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-计价,2-发送,3-签名,4-费用,5-发料
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '用于NO替换处理时排序
    rsSQL.Fields.Append "诊疗类别", adVarChar, 8
    rsSQL.Fields.Append "当前行医嘱ID", adInteger
    rsSQL.Fields.Append "其它", adVarChar, 38  '医嘱表格的行号 收费项目ID 执行部门ID 用 “_”分割
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open

    '计价数量累计记录集
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsTotal.Fields.Append "项目ID", adBigInt
    rsTotal.Fields.Append "库房ID", adBigInt
    rsTotal.Fields.Append "数量", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open

    '计录试管编码
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "管码", adVarChar, 18
    rsNumber.Fields.Append "相关ID", adBigInt
    rsNumber.Fields.Append "样本条码", adVarChar, 18
    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
    rsNumber.Fields.Append "婴儿", adBigInt
    rsNumber.Fields.Append "紧急标志", adBigInt
    rsNumber.Fields.Append "标本", adVarChar, 18
    rsNumber.Fields.Append "采集科室ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open

    '当前病人本次要发送的费用
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsMoneyNow.Fields.Append "诊疗项目ID", adBigInt
    rsMoneyNow.Fields.Append "收费项目ID", adBigInt
    rsMoneyNow.Fields.Append "试管编码", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "样本条码", adVarChar, 50, adFldIsNullable
    rsMoneyNow.Fields.Append "收费方式", adInteger
    rsMoneyNow.Fields.Append "收费时间", adVarChar, 10
    rsMoneyNow.Fields.Append "执行部门ID", adBigInt
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open

    '当前病人本次发送的费用项目汇总
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "病人ID", adBigInt
    rsItems.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "医嘱ID", adBigInt
    rsItems.Fields.Append "收费类别", adVarChar, 1
    rsItems.Fields.Append "收费细目ID", adBigInt
    rsItems.Fields.Append "数量", adDouble
    rsItems.Fields.Append "单价", adDouble
    rsItems.Fields.Append "实收金额", adDouble
    rsItems.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "疾病ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "诊断ID", adBigInt, , adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open


    lng发送号 = gobjComlib.zlDatabase.GetNextNo(10)
    curDate = gobjComlib.zlDatabase.Currentdate
    '避免(医嘱ID,操作时间)重复
    '本来医嘱Insert时就-2s了，这里再取不加1s也不会重复
    curDate = DateAdd("s", 1, curDate)
    bln划价 = True '初始全部是划价
    int配方数 = 1 '表示发送的第几付配方,用于分单据号
    lngNOSequence = 0 '单据号序列重新初始
    
    '读取药品医嘱
    strSQL = _
        "Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, Nvl(x.序号, a.序号) As 组号, a.诊疗类别,  a.诊疗项目id, b.名称 As 诊疗项目," & vbNewLine & _
        "       a.收费细目id,  a.婴儿, a.医嘱内容, a.标本部位, a.检查方法, a.执行标记, a.天数, a.总给予量, d.门诊单位, a.单次用量," & vbNewLine & _
        "       d.剂量系数, d.门诊包装, a.开始执行时间, a.执行频次, a.频率次数, a.频率间隔, a.间隔单位, a.医生嘱托," & vbNewLine & _
        "       a.执行时间方案, a.病人科室id, a.开嘱科室id, a.开嘱医生, a.计价特性, a.执行性质, a.执行科室id," & vbNewLine & _
        "       d.门诊可否分零 As 可否分零,  a.前提id,A.序号," & vbNewLine & _
        "       a.新开签名id As 签名id, b.试管编码, b.操作类型, b.执行分类, a.摘要, a.紧急标志, a.零费记帐, b.计算方式, a.开始执行时间, b.执行安排" & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗项目目录 B, 药品规格 D,   病人医嘱记录 X" & vbNewLine & _
        "Where a.id In (select /*+cardinality(c,10)*/ column_value from table(f_num2list([1])) C) And a.医嘱状态 = 1 And a.医嘱期效 = 1 And a.相关id = x.Id(+) And" & vbNewLine & _
        "      a.诊疗项目id = b.Id And a.收费细目id = d.药品id(+)  And" & vbNewLine & _
        "      Nvl(a.执行标记, 0) <> -1" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, Nvl(x.序号, a.序号) As 组号, a.诊疗类别,  a.诊疗项目id, b.名称 As 诊疗项目," & vbNewLine & _
        "       a.收费细目id,  a.婴儿, a.医嘱内容, a.标本部位, a.检查方法, a.执行标记, a.天数, a.总给予量, d.门诊单位, a.单次用量," & vbNewLine & _
        "       d.剂量系数, d.门诊包装, a.开始执行时间, a.执行频次, a.频率次数, a.频率间隔, a.间隔单位, a.医生嘱托," & vbNewLine & _
        "       a.执行时间方案, a.病人科室id, a.开嘱科室id, a.开嘱医生, a.计价特性, a.执行性质, a.执行科室id," & vbNewLine & _
        "       d.门诊可否分零 As 可否分零,  a.前提id,A.序号," & vbNewLine & _
        "       a.新开签名id As 签名id, b.试管编码, b.操作类型, b.执行分类, a.摘要, a.紧急标志, a.零费记帐, b.计算方式, a.开始执行时间, b.执行安排" & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗项目目录 B, 药品规格 D,   病人医嘱记录 X" & vbNewLine & _
        "Where a.相关id In  (select /*+cardinality(c,10)*/ column_value from table(f_num2list([1])) C) And a.医嘱状态 = 1 And a.医嘱期效 = 1 And a.相关id = x.Id(+) And" & vbNewLine & _
        "      a.诊疗项目id = b.Id And a.收费细目id = d.药品id(+)  And" & vbNewLine & _
        "      Nvl(a.执行标记, 0) <> -1" & vbNewLine & _
        "Order By 婴儿, 组号, 组id, 序号"


    Set rsSend = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "OutAdviceSendDrug", str给药IDs)
    
    With objfrm.vsAdvice
        '填充虚拟表格
        Call gobjComlib.Grid.BandRec(objfrm.vsAdvice, rsSend)
        
        '检查是否有处方
        If Not IsNull(rsPati!险类) Then
            If gclsInsure.GetCapability(support医生确定处方类型, lng病人ID, rsPati!险类) Then
'                str摘要 = "2"
'                strTmp = gobjComlib.ZLCommFun.ShowMsgBox("处方类型", "请确定当前医保病人本次要发送的药品处方的类型。", "!医保内(&A),医保外(&B),?取消(&C)", Me)
'                If strTmp = "" Then Exit Function
'                If strTmp = "医保内" Then str摘要 = "1"
                GoTo FuncEnd
            End If
        End If
        '处理表格数据
        For i = .FixedRows To .Rows - 1
            
            If .TextMatrix(i, .ColIndex("诊疗类别")) = "5" Or .TextMatrix(i, .ColIndex("诊疗类别")) = "6" Then
                '计算临嘱用药次数
                If Nvl(Val(.TextMatrix(i, .ColIndex("天数"))), 0) <> 0 And Not IsNull(.TextMatrix(i, .ColIndex("执行频次"))) Then
                    '一个频率周期的次数
                    If .TextMatrix(i, .ColIndex("间隔单位")) = "周" Then
                        lng次数 = IntEx(Val(.TextMatrix(i, .ColIndex("天数"))) * (.TextMatrix(i, .ColIndex("频率次数")) / 7))
                    ElseIf .TextMatrix(i, .ColIndex("间隔单位")) = "天" Then
                        lng次数 = IntEx(Val(.TextMatrix(i, .ColIndex("天数"))) * (.TextMatrix(i, .ColIndex("频率次数")) / .TextMatrix(i, .ColIndex("频率间隔"))))
                    ElseIf .TextMatrix(i, .ColIndex("间隔单位")) = "小时" Then
                        lng次数 = IntEx(Val(.TextMatrix(i, .ColIndex("天数"))) * (.TextMatrix(i, .ColIndex("频率次数")) / .TextMatrix(i, .ColIndex("频率间隔"))) * 24)
                    ElseIf .TextMatrix(i, .ColIndex("间隔单位")) = "分钟" Then
                        lng次数 = IntEx(Val(.TextMatrix(i, .ColIndex("天数"))) * (.TextMatrix(i, .ColIndex("频率次数")) / .TextMatrix(i, .ColIndex("频率间隔"))) * (24 * 60))
                    End If
                Else
                    '可分零药品时,按总量对单量的倍数计算给药途径的次数,不可分零与一次性使用药品时，按总量对（单量与剂量系数比值取整）的倍数计算给药途径的次数，
                    '否则按一个频率周期的次数计算
                    If Nvl(Val(.TextMatrix(i, .ColIndex("可否分零"))), 0) = 0 And Nvl(Val(.TextMatrix(i, .ColIndex("单次用量"))), 0) <> 0 Then
                        lng次数 = IntEx(.TextMatrix(i, .ColIndex("总给予量")) * .TextMatrix(i, .ColIndex("剂量系数")) / Val(.TextMatrix(i, .ColIndex("单次用量"))))
                    ElseIf (Nvl(Val(.TextMatrix(i, .ColIndex("可否分零"))), 0) = 1 Or Nvl(Val(.TextMatrix(i, .ColIndex("可否分零"))), 0) = 2) And Nvl(Val(.TextMatrix(i, .ColIndex("单次用量"))), 0) <> 0 Then
                        lng次数 = IntEx(.TextMatrix(i, .ColIndex("总给予量")) / IntEx(Val(.TextMatrix(i, .ColIndex("单次用量"))) / .TextMatrix(i, .ColIndex("剂量系数"))))
                    Else
                        lng次数 = Nvl(.TextMatrix(i, .ColIndex("频率次数")), 0)
                    End If
                End If
                If Not IsNull(.TextMatrix(i, .ColIndex("频率次数"))) And (Not IsNull(.TextMatrix(i, .ColIndex("执行时间方案"))) Or Nvl(.TextMatrix(i, .ColIndex("间隔单位"))) = "分钟") Then
                    .Cell(flexcpData, i, .ColIndex("ID")) = Calc次数分解时间(lng次数, .TextMatrix(i, .ColIndex("开始执行时间")), CDate("3000-01-01"), "", Nvl(.TextMatrix(i, .ColIndex("执行时间方案"))), .TextMatrix(i, .ColIndex("频率次数")), .TextMatrix(i, .ColIndex("频率间隔")), .TextMatrix(i, .ColIndex("间隔单位")))
                Else
                    '无分解时间(临嘱可能未输入执行时间而无法分解)
                    '记录费用发生时间(以医嘱开始执行时间)
                    .Cell(flexcpData, i, .ColIndex("ID")) = Format(.TextMatrix(i, .ColIndex("开始执行时间")), "yyyy-MM-dd HH:mm:ss")
                End If
                .TextMatrix(i, .ColIndex("单次用量")) = FormatEx(Val(.TextMatrix(i, .ColIndex("单次用量"))), 5)
                .TextMatrix(i, .ColIndex("总给予量")) = FormatEx(Val(.TextMatrix(i, .ColIndex("总给予量"))) / Val(.TextMatrix(i, .ColIndex("门诊包装"))), 5)
                .Cell(flexcpData, i, .ColIndex("相关ID")) = lng次数
                If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
            ElseIf .TextMatrix(i, .ColIndex("诊疗类别")) = "7" Then
                .Cell(flexcpData, i, .ColIndex("ID")) = Calc次数分解时间(.TextMatrix(i, .ColIndex("总给予量")), .TextMatrix(i, .ColIndex("开始执行时间")), CDate("3000-01-01"), "", .TextMatrix(i, .ColIndex("执行时间方案")), .TextMatrix(i, .ColIndex("频率次数")), .TextMatrix(i, .ColIndex("频率间隔")), .TextMatrix(i, .ColIndex("间隔单位")))
                .Cell(flexcpData, i, .ColIndex("相关ID")) = Val(.TextMatrix(i, .ColIndex("总给予量")))
            Else
                '一并给药的按最小次数发送(影响给药途径计费)
                If .TextMatrix(i, .ColIndex("诊疗类别")) = "E" And .TextMatrix(i, .ColIndex("操作类型")) = "2" Then '给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, .ColIndex("相关ID"))) = Val(.TextMatrix(i, .ColIndex("ID"))) Then
                            If Val(.Cell(flexcpData, j, .ColIndex("相关ID"))) > lng最小次数 Then
                                .Cell(flexcpData, j, .ColIndex("相关ID")) = lng最小次数
                                If .Cell(flexcpData, j, .ColIndex("ID")) <> "" Then
                                    .Cell(flexcpData, j, .ColIndex("ID")) = Trim分解时间(lng最小次数, .Cell(flexcpData, j, .ColIndex("ID")))
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    lng最小次数 = 0
                End If
                
                blnTmp = False
                If Val(.TextMatrix(i, .ColIndex("总给予量"))) <> 0 Then
                    If Val(.TextMatrix(i, .ColIndex("总给予量"))) < Val(.Cell(flexcpData, i - 1, .ColIndex("相关ID"))) Then
                        If .TextMatrix(i, .ColIndex("操作类型")) = "2" And (.TextMatrix(i, .ColIndex("执行分类")) = "1" Or .TextMatrix(i, .ColIndex("执行分类")) = "2") Then
                            blnTmp = True
                        End If
                    End If
                End If
                
                If blnTmp Then
                    .Cell(flexcpData, i, .ColIndex("相关ID")) = Val(.TextMatrix(i, .ColIndex("总给予量")))
                    .Cell(flexcpData, i, .ColIndex("ID")) = Trim分解时间(Val(.TextMatrix(i, .ColIndex("总给予量"))), .Cell(flexcpData, i - 1, .ColIndex("ID")))
                Else
                    .TextMatrix(i, .ColIndex("总给予量")) = .Cell(flexcpData, i - 1, .ColIndex("相关ID"))
                    .Cell(flexcpData, i, .ColIndex("相关ID")) = .Cell(flexcpData, i - 1, .ColIndex("相关ID"))
                    .Cell(flexcpData, i, .ColIndex("ID")) = .Cell(flexcpData, i - 1, .ColIndex("ID"))
                End If
            End If
        Next

        '医嘱发送处理
        For i = .FixedRows To .Rows - 1
            '产生单据号分配关键字
            '-----------------------------------------------------------------------------------------
            If intSendNo = 1 And Not gbln执行前先结算 Then
                strNOKey = "只产生一个单据号"
            ElseIf intSendNo = 2 Then
                strNOKey = Val(.TextMatrix(i, .ColIndex("执行科室ID")))
                If InStr(",5,6,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                    '再按处方条数限制进行分号
                    If gintRXCount > 0 Then
                        strTmp = ""
                        If Val(.TextMatrix(i, .ColIndex("相关ID"))) <> Val(.TextMatrix(i - 1, .ColIndex("相关ID"))) Then
                            strTmp = GetMergeIDs(objfrm.vsAdvice, i, .ColIndex("相关ID"), .ColIndex("ID")) '一并给药开始行或独立药品行才取值
                        End If
                        strNOKey = strNOKey & "_" & GetRXKey(rsRXKey, strNOKey, strTmp)
                    End If
                End If
            Else
                If InStr(",5,6,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                    '中西成药按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                    '一并给药的，发送到一起：包括自备药和不同药房的情况
                    strNOKey = "中西成药_" & lng病人ID & "_" & str挂号单 & "_" & _
                        Val(.TextMatrix(i, .ColIndex("病人科室ID"))) & "_" & Val(.TextMatrix(i, .ColIndex("开嘱科室ID"))) & "_" & _
                         GetMergeDrugStore(objfrm.vsAdvice, i)

                    If bln一并给药发送为一张 Then
                        If Val(.TextMatrix(i, .ColIndex("相关ID"))) <> Val(.TextMatrix(i - 1, .ColIndex("相关ID"))) Then
                            '再按要打印的诊疗单据分号(一并给药的，只取第一个药品的诊疗单据ID)
                            strPre诊疗单据ID = GetClinicBillID(Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))), 1)
                        End If
                        strNOKey = strNOKey & "_" & strPre诊疗单据ID
                    Else
                        strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))), 1)
                    End If

                    '再按处方条数限制进行分号
                    If gintRXCount > 0 Then
                        strTmp = ""
                        If Val(.TextMatrix(i, .ColIndex("相关ID"))) <> Val(.TextMatrix(i - 1, .ColIndex("相关ID"))) Then
                            strTmp = GetMergeIDs(objfrm.vsAdvice, i, .ColIndex("相关ID"), .ColIndex("ID")) '一并给药开始行或独立药品行才取值
                        End If
                        strNOKey = strNOKey & "_" & GetRXKey(rsRXKey, strNOKey, strTmp)
                    End If
                ElseIf .TextMatrix(i, .ColIndex("诊疗类别")) = "7" Then
                    '一个配方中的所有草药分配一个独立单据号
                    strNOKey = "中药配方_" & lng病人ID & "_" & str挂号单 & "_" & int配方数

                '非药嘱，同一类别是否按相同执行科室组合单据
                ElseIf InStr(str单据组合类别, .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                    strNOKey = "非药医嘱_" & .TextMatrix(i, .ColIndex("诊疗类别")) & "_" & Val(.TextMatrix(i, .ColIndex("执行科室ID")))
                Else
                    '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法，采集方式，麻醉方式，输血医嘱/输血途径)
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, .ColIndex("ID")))
                End If
            End If

            '不同诊断的医嘱分别产生单据,要重新加工strNOKey
            If blnNOCtrl Then
                lng医嘱ID = IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) <> 0, Val(.TextMatrix(i, .ColIndex("相关ID"))), Val(.TextMatrix(i, .ColIndex("ID"))))
                If lng医嘱ID <> lng组ID Then str医嘱诊断ids = GetAdviceDiag(lng医嘱ID)
                lng组ID = lng医嘱ID

                If str医嘱诊断ids <> "" Then strNOKey = strNOKey & "_" & str医嘱诊断ids
            End If

            '开始时间不是同一天的分别产生单据
            If blnStartTimeDef Then
                strNOKey = strNOKey & "_" & Format(.TextMatrix(i, .ColIndex("开始执行时间")), "YYYY-MM-DD")
            End If

            '开单人不同的，默认全部分别产生单据
            strNOKey = strNOKey & "_" & .TextMatrix(i, .ColIndex("开嘱医生"))

            '是否离院带药
            bln离院带药 = False
            If InStr(",5,6,7,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                If Val(.TextMatrix(i, .ColIndex("执行性质"))) = 5 Then bln离院带药 = True
            End If

            '产生医嘱记帐费用:以最新价格计算
            '-----------------------------------------------------------------------------------------
            strSQL = "": str收费项目 = ""
            If InStr(",5,6,7,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
                If Val(.TextMatrix(i, .ColIndex("执行性质"))) <> 5 Then
                    strSQL = _
                        " Select A.ID,A.类别,D.名称 as 类别名称,RTrim(A.名称||' '||A.规格) as 名称," & _
                        " A.计算单位,A.是否变价,A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,100 as 附术收费率," & _
                        " Y.门诊单位,Y.门诊包装,Y.剂量系数,Y.药房分批 as 分批,0 as 跟踪在用,B.收入项目ID," & _
                        " C.收据费目,1 as 数量,B.现价 as 单价,[2] as 执行科室ID,0 as 从项,0 as 费用性质,0 as 收费方式" & _
                        " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,药品规格 Y" & _
                        " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "3", "4", "5") & _
                        " And A.ID=Y.药品ID(+) And A.ID=[1]" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " Order by A.编码"
                End If
            Else
                '先删除原非药医嘱的计价(应该没有)
                rsSQL.AddNew
                rsSQL!类型 = 1: rsSQL!项目ID = 0: rsSQL!序号 = i
                rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) = 0, Val(.TextMatrix(i, .ColIndex("ID"))), Val(.TextMatrix(i, .ColIndex("相关ID"))))
                rsSQL!sql = "ZL_病人医嘱计价_Delete(" & Val(.TextMatrix(i, .ColIndex("ID"))) & ",1)"
                rsSQL.Update

                '不计价,手工计价；叮嘱,院外执行的医嘱不读取
                If Val(.TextMatrix(i, .ColIndex("计价特性"))) = 0 And InStr(",0,5,", Val(.TextMatrix(i, .ColIndex("执行性质")))) = 0 Then
                    strSQL = "Select * From (" & _
                            "Select C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,c.适用科室id" & _
                            " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                            " From 诊疗收费关系 C Where C.诊疗项目ID=[1]" & _
                            "      And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = [2] And C.病人来源 = 1)" & _
                            " ) Where Nvl(适用科室id, 0) = Top"
                    strSQL = _
                        " Select C.类别,A.收费项目ID,A.收费数量,A.固有对照,B.收入项目ID," & _
                        " C.加班加价,B.加班加价率,Decode(C.是否变价,1,B.缺省价格,B.现价)" & IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "") & " as 单价,C.是否变价," & _
                        " Nvl(A.从属项目,0) as 从项,D.跟踪在用,[2] as 执行科室ID,C.屏蔽费别,Nvl(A.费用性质,0) as 费用性质," & _
                        " Nvl(A.收费方式,0) as 收费方式" & _
                        " From (" & strSQL & ") A,收费价目 B,收费项目目录 C,材料特性 D" & _
                        " Where A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And A.收费项目ID=D.材料ID(+)" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "B", "3", "4", "5") & _
                        " And C.服务对象 IN(1,3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " Order by 费用性质,从项,A.收费项目ID"
                    Set rsPrice = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "OutAdviceSendDrug", Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))), _
                        Val(.TextMatrix(i, .ColIndex("执行科室ID"))), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    strSQL = ""
                    If Not rsPrice.EOF Then
                        For j = 1 To rsPrice.RecordCount
                            If Nvl(rsPrice!收费项目ID, 0) <> 0 And Nvl(rsPrice!收费数量, 0) <> 0 Then '对照数量为0的自动过滤掉
                                '普通项目的变价单价要求输入，包括非跟踪在用的时价卫材医嘱
                                If Nvl(rsPrice!单价, 0) = 0 And Nvl(rsPrice!是否变价, 0) = 1 _
                                    And Not (InStr(",5,6,7,", rsPrice!类别) > 0 Or rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1) Then
                                    Screen.MousePointer = 0
                                    'MsgBox "必须为变价的收费项目确定一个收费价格。", vbInformation, "处方自动发送"
                                    GoTo FuncEnd
                                End If

                                '计价执行科室:只保存非药品及卫材医嘱的，药品和卫材计价的执行科室
                                If InStr(",4,5,6,7,", .TextMatrix(i, .ColIndex("诊疗类别"))) = 0 _
                                    And (InStr(",5,6,7,", rsPrice!类别) > 0 Or rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1) Then
                                    lng执行科室ID = Nvl(rsPrice!执行科室ID, 0)

                                    '卫材必须设置执行科室
                                    If lng执行科室ID = 0 And rsPrice!类别 = "4" Then
                                        Screen.MousePointer = 0
                                        'MsgBox "卫材没有确定执行科室，请手工输入正确的执行科室。" & vbCrLf & _
                                            "如果不能确定正确的执行科室，请到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, "处方自动发送"
                                        GoTo FuncEnd
                                    End If
                                Else
                                    lng执行科室ID = 0
                                End If

                                '药品、卫材医嘱的计价固定对应不保存；非跟踪在用的时价卫材的变价需要输入，因此要保存到计价表中
                                If InStr(",4,5,6,7,", .TextMatrix(i, .ColIndex("诊疗类别"))) = 0 _
                                    Or .TextMatrix(i, .ColIndex("诊疗类别")) = "4" And Nvl(rsPrice!跟踪在用, 0) = 0 And Nvl(rsPrice!是否变价, 0) = 1 Then
                                    rsSQL.AddNew
                                    rsSQL!类型 = 1: rsSQL!项目ID = rsPrice!收费项目ID: rsSQL!序号 = i
                                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) = 0, Val(.TextMatrix(i, .ColIndex("ID"))), Val(.TextMatrix(i, .ColIndex("相关ID"))))
                                    rsSQL!sql = "ZL_病人医嘱计价_INSERT(" & _
                                        .TextMatrix(i, .ColIndex("ID")) & "," & rsPrice!收费项目ID & "," & _
                                        Nvl(rsPrice!收费数量, 0) & "," & Nvl(rsPrice!单价, 0) & "," & _
                                        Nvl(rsPrice!从项, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                        Nvl(rsPrice!费用性质, 0) & "," & Nvl(rsPrice!收费方式, 0) & ")"
                                    rsSQL.Update
                                End If

                                '临时病人医嘱计价表
                                If Val(.TextMatrix(i, .ColIndex("总给予量"))) <> 0 Then
                                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                        "Select " & rsPrice!收费项目ID & " as 收费细目ID," & _
                                        Nvl(rsPrice!执行科室ID, 0) & " as 执行科室ID," & _
                                        Nvl(rsPrice!收费数量, 0) & " as 数量," & Format(Nvl(rsPrice!单价, 0), gstrDecPrice) & " as 单价," & _
                                        Nvl(rsPrice!从项, 0) & " as 从项," & Nvl(rsPrice!费用性质, 0) & " as 费用性质," & _
                                        Nvl(rsPrice!收费方式, 0) & " as 收费方式 From Dual"
                                End If
                            End If

                            rsPrice.MoveNext
                        Next
                    End If
                End If

                If strSQL <> "" Then
                    strSQL = _
                        " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,A.是否变价," & _
                        " A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,B.附术收费率,Y.门诊单位,Y.门诊包装,Y.剂量系数," & _
                        " Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批,E.跟踪在用,B.收入项目ID," & _
                        " C.收据费目,X.数量,Decode(A.是否变价,1,X.单价,B.现价) as 单价,X.执行科室ID,X.从项,X.费用性质,X.收费方式" & _
                        " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,(" & strSQL & ") X,药品规格 Y" & _
                        " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.ID=E.材料ID(+)" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "3", "4", "5") & _
                        " And A.类别=D.编码 And X.收费细目ID=A.ID And A.ID=Y.药品ID(+)" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " Order by X.费用性质,X.从项,X.收费方式 Desc,A.ID"
                        '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                End If
            End If

            '汇总折扣变量初始
            int父序号 = 0: lng父项目ID = 0
            strHaveSub = "": strNoneSub = ""
            Set rsSeek = New ADODB.Recordset
            rsSeek.Fields.Append "费用性质", adInteger
            rsSeek.Fields.Append "主项标签", adVariant
            rsSeek.Fields.Append "主收入ID", adBigInt
            rsSeek.Fields.Append "合计", adCurrency, , adFldIsNullable
            rsSeek.CursorLocation = adUseClient
            rsSeek.LockType = adLockOptimistic
            rsSeek.CursorType = adOpenStatic
            rsSeek.Open
            
            '本科执行的自动执行：特殊医嘱不用处理
            int执行状态 = 0
            int计费状态 = IIF(Val(.TextMatrix(i, .ColIndex("计价特性"))) = 1, -1, 0) '无需计费或未计费
            
            str分解时间 = .Cell(flexcpData, i, .ColIndex("ID"))
            
            If strSQL <> "" Then
                Set rsMoney = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "OutAdviceSendDrug", Val(.TextMatrix(i, .ColIndex("收费细目ID"))), Val(.TextMatrix(i, .ColIndex("执行科室ID"))), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                If Not rsMoney.EOF Then
                    int计费状态 = 1 '已计费
                    Set rsClone = rsMoney.Clone
                End If

                '处理收入项目级的费用明细
                bln附加手术 = .TextMatrix(i, .ColIndex("诊疗类别")) = "F" And Val(.TextMatrix(i, .ColIndex("相关ID"))) <> 0
                Do While Not rsMoney.EOF
MoneyItemBegin:
                    '执行科室ID
                    lng执行科室ID = Nvl(rsMoney!执行科室ID, 0)
                    '在原值基础上取有效的非药嘱药品及跟踪卫材的执行科室
                    If InStr(",4,5,6,7", .TextMatrix(i, .ColIndex("诊疗类别"))) = 0 _
                        And (rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsMoney!类别) > 0) Then
                        lng执行科室ID = Get收费执行科室ID(lng病人ID, 0, rsMoney!类别, rsMoney!ID, 4, Val(.TextMatrix(i, .ColIndex("病人科室ID"))), 0, 1, lng执行科室ID)
                    End If

                    '----------------------------------------
                    '根据收费方式，确定当前收费项目是否应收费
                    If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                        If Not AdviceMoneyMake(lng病人ID, 0, rsMoneyNow, rsMoneyDay, _
                            IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) <> 0, Val(.TextMatrix(i, .ColIndex("相关ID"))), Val(.TextMatrix(i, .ColIndex("ID")))), _
                            Val(.TextMatrix(i, .ColIndex("诊疗项目ID"))), rsMoney!ID, lng执行科室ID, "", _
                            rsMoney!类别, Nvl(rsMoney!收费方式, 0), str分解时间, 1, lng费用次数, Val(.TextMatrix(i, .ColIndex("总给予量"))), _
                             Val(.TextMatrix(i, .ColIndex("ID"))), lng发送号, Val(rsMoney!数量 & ""), rsExec, Val(.TextMatrix(i, .ColIndex("计算方式"))), _
                                    .TextMatrix(i, .ColIndex("执行频次")), Val(.TextMatrix(i, .ColIndex("单次用量"))), , , .TextMatrix(i, .ColIndex("诊疗类别")), strCuvetteNumber) Then
                            '跳过当前收费项目(多个收入项目)
                            str收费项目 = rsMoney!费用性质 & "_" & rsMoney!ID
                            Do While rsMoney!费用性质 & "_" & rsMoney!ID = str收费项目
                                rsMoney.MoveNext
                                If rsMoney.EOF Then Exit Do
                            Loop
                            If rsMoney.EOF Then Exit Do
                            GoTo MoneyItemBegin
                        End If
                    End If
                    '----------------------------------------

                    If InStr(",5,6,7", rsMoney!类别) > 0 Then
                        If InStr(",5,6,7", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                            If .TextMatrix(i, .ColIndex("诊疗类别")) = "7" Then
                                int付数 = Val(.TextMatrix(i, .ColIndex("总给予量")))
                                '中药药房单位按不可分零处理:每付
                                If Val(Val(.TextMatrix(i, .ColIndex("可否分零")))) = 0 Then
                                    dbl数量 = Val(.TextMatrix(i, .ColIndex("单次用量"))) / Nvl(rsMoney!剂量系数, 1)
                                Else
                                    dbl数量 = IntEx(Val(.TextMatrix(i, .ColIndex("单次用量"))) / Nvl(rsMoney!剂量系数, 1) / Nvl(rsMoney!门诊包装, 1)) * Nvl(rsMoney!门诊包装, 1)
                                End If
                            Else
                                int付数 = 1
                                dbl数量 = Val(.TextMatrix(i, .ColIndex("总给予量"))) * Nvl(rsMoney!门诊包装, 1)
                            End If
                        Else
                            int付数 = 1
                            '中药药房单位按不可分零处理:每付
                            '非药嘱药品计价:因为这里预定了售价数量,因此不作不分零处理
                            '对于收费对照中的药品，且为当天只收取一次，数量为费用次数*对照数量
                            If InStr(",2,3,4,5,6,7,9,", Val("" & rsMoney!收费方式)) > 0 Then
                                dbl数量 = Format(lng费用次数 * Nvl(rsMoney!数量, 0), "0.00000")
                            Else
                                dbl数量 = Val(.TextMatrix(i, .ColIndex("总给予量"))) * Nvl(rsMoney!数量, 0)
                            End If
                        End If
                        dbl数量 = Format(dbl数量, "0.00000")

                        If Nvl(rsMoney!是否变价, 0) = 1 Then
                            dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, int付数 * dbl数量, , True, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        Else
                            dbl单价 = Format(Nvl(rsMoney!单价, 0), gstrDecPrice)
                        End If
                    ElseIf rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Then
                        '检查卫生材料入出类别
                        If lng卫材类别ID = 0 Then
                            Screen.MousePointer = 0
                            'MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, "处方自动发送"
                            GoTo FuncEnd
                        End If

                        int付数 = 1
                        If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!收费方式)) > 0 Then
                            dbl数量 = Format(lng费用次数 * Nvl(rsMoney!数量, 0), "0.00000")
                        Else
                            dbl数量 = Format(Val(.TextMatrix(i, .ColIndex("总给予量"))) * Nvl(rsMoney!数量, 0), "0.00000")
                        End If

                        '确定时价卫材价格
                        If Nvl(rsMoney!是否变价, 0) = 1 Then
                            dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, dbl数量, , True, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        Else
                            dbl单价 = Format(Nvl(rsMoney!单价, 0), gstrDecPrice)
                        End If
                    Else
                        '总量等于单次用量乘数次。一天只收一次时，有多少天要执行，就收多少次，不管单次用量（例如：每天两次）,但要管收费对照的次数
                        int付数 = 1
                        If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!收费方式)) > 0 Then
                            dbl数量 = Format(lng费用次数 * Nvl(rsMoney!数量, 0), "0.00000")
                        Else
                            dbl数量 = Format(Val(.TextMatrix(i, .ColIndex("总给予量"))) * Nvl(rsMoney!数量, 0), "0.00000")
                        End If
                        dbl单价 = Format(Nvl(rsMoney!单价, 0), gstrDecPrice)
                    End If
                    
                    '非药嘱药品及跟踪卫材的库存检查
                    If (rsMoney!类别 = "4" And Nvl(rsMoney!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsMoney!类别) > 0) Then
                        If TheStockCheck(colStock1, colStock2, lng执行科室ID, rsMoney!类别) <> 0 Or Nvl(rsMoney!是否变价, 0) = 1 Or Nvl(rsMoney!分批, 0) = 1 Then
                            If rsMoney!类别 = "4" Then
                                blnBool = CheckPriceStock(colStock1, colStock2, objfrm.vsAdvice, i, rsMoney, lng执行科室ID, int付数 * dbl数量, rsTotal, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                            Else
                                blnBool = CheckPriceStock(colStock1, colStock2, objfrm.vsAdvice, i, rsMoney, lng执行科室ID, int付数 * dbl数量, rsTotal, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                            End If
                            If blnBool Then
                                GoTo FuncEnd
                            End If
                        End If
                    End If

                    '发送金额
                    dbl应收 = int付数 * dbl数量 * dbl单价

                    cur应收 = Format(dbl应收, gstrDec)

                    'NO,序号
                    Call GetCurBillSet(rsBill, strNOKey, strNO, lng费用序号, -1, True, lngNOSequence)
                    rsSQL.AddNew: blnBool = False
                    If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                        lng费用父号 = lng费用序号
                        If rsMoney!从项 = 0 Then
                            '记录主项信息，主项肯定在从项前
                            '即使不汇总折扣，也要记录主从项关系
                            If InStr(strHaveSub & ",", "," & rsMoney!费用性质 & ",") = 0 _
                                And InStr(strNoneSub & ",", "," & rsMoney!费用性质 & ",") = 0 Then
                                rsClone.Filter = "费用性质=" & rsMoney!费用性质 & " And 从项=1"
                                If Not rsClone.EOF Then
                                    int父序号 = lng费用序号
                                    lng父项目ID = rsMoney!ID

                                    rsSeek.AddNew
                                    rsSeek!费用性质 = rsMoney!费用性质
                                    rsSeek!主项标签 = rsSQL.Bookmark 'Variant(Double)
                                    rsSeek!主收入ID = rsMoney!收入项目ID
                                    rsSeek.Update
                                    strHaveSub = strHaveSub & "," & rsMoney!费用性质

                                    blnBool = True
                                Else
                                    strNoneSub = strNoneSub & "," & rsMoney!费用性质
                                End If
                            End If
                        End If
                    End If

                    '计算汇总折扣合计
                    str费别 = Nvl(rsPati!费别)
                    If gbln从项汇总折扣 And (rsMoney!从项 = 1 Or InStr(strHaveSub & ",", "," & rsMoney!费用性质 & ",") > 0) Then
                        If Val(.TextMatrix(i, .ColIndex("零费记帐"))) = 1 Then
                            cur实收 = 0
                        Else
                            cur实收 = cur应收
                        End If

                        '累计医嘱合计来计算折扣
                        rsSeek.Filter = "费用性质=" & rsMoney!费用性质
                        rsSeek!合计 = Nvl(rsSeek!合计, 0) + cur实收
                        rsSeek.Update
                    ElseIf Nvl(rsMoney!屏蔽费别, 0) = 0 Then
                        str费别 = Nvl(rsPati!费别) & IIF(gstr动态费别 <> "", "," & gstr动态费别, "")

                        If Val(.TextMatrix(i, .ColIndex("零费记帐"))) = 1 Then
                            cur实收 = 0
                        Else
                            cur实收 = Format(ActualMoney(str费别, rsMoney!收入项目ID, cur应收, rsMoney!ID, lng执行科室ID, _
                                int付数 * dbl数量, 0), gstrDec)
                        End If
                        If InStr(str费别, ",") > 0 Then str费别 = Nvl(rsPati!费别)
                    Else
                        If Val(.TextMatrix(i, .ColIndex("零费记帐"))) = 1 Then
                            cur实收 = 0
                        Else
                            cur实收 = cur应收
                        End If
                    End If
                    '汇总折扣时，对主项的实收金额作特殊处理
                    If gbln从项汇总折扣 And blnBool Then
                        str费别 = Chr(0) & Chr(1) & "Begin费别" & str费别 & "End费别" & Chr(0) & Chr(1)
                        str实收 = Chr(0) & Chr(1) & "Begin实收" & cur实收 & "End实收" & Chr(0) & Chr(1)
                    Else
                        str实收 = cur实收
                    End If

                    '医保相关字段
                    bln保险项目否 = False: lng保险大类ID = 0: str保险编码 = "": str费用类型 = ""
                    If Not IsNull(rsPati!险类) Then
                        strTmp = gclsInsure.GetItemInsure(lng病人ID, rsMoney!ID, cur实收, True, rsPati!险类, .TextMatrix(i, .ColIndex("医生嘱托")) & "||" & int付数 * dbl数量)
                        If strTmp <> "" Then
                            bln保险项目否 = Val(Split(strTmp, ";")(0)) <> 0
                            lng保险大类ID = Val(Split(strTmp, ";")(1))
                            str保险编码 = CStr(Split(strTmp, ";")(3))
                            If UBound(Split(strTmp, ";")) >= 5 Then
                                If Split(strTmp, ";")(5) <> "" Then
                                    str费用类型 = Split(strTmp, ";")(5)
                                End If
                            End If
                        End If
                    End If

                    '收集记帐报警类别
                    If InStr(str类别, rsMoney!类别) = 0 Then
                        str类别 = str类别 & rsMoney!类别
                    End If

                    '发生时间
                    If str分解时间 <> "" Then
                        str发生时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        str发生时间 = "To_Date('" & .TextMatrix(i, .ColIndex("开始执行时间")) & "','YYYY-MM-DD HH24:MI:SS')"
                    End If

                    '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                    rsSQL!类型 = 2: rsSQL!项目ID = rsMoney!ID: rsSQL!序号 = i
                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) = 0, Val(.TextMatrix(i, .ColIndex("ID"))), Val(.TextMatrix(i, .ColIndex("相关ID"))))
                    rsSQL!NO = strNO
                    rsSQL!诊疗类别 = IIF(InStr(",5,6,7,", "," & .TextMatrix(i, .ColIndex("诊疗类别")) & ",") > 0, "药品", "0")
                    rsSQL!当前行医嘱ID = Val(.TextMatrix(i, .ColIndex("ID")))
                    rsSQL!其它 = i & "_" & rsMoney!ID & "_" & lng执行科室ID
                    cur应收 = Format(cur应收, gstrDec)
                    str实收 = Format(str实收, gstrDec)
                    cur合计 = cur合计 + cur实收
                    str用法 = ""
                    If InStr(",5,6,", "," & .TextMatrix(i, .ColIndex("诊疗类别")) & ",") > 0 Then
                        lngTmp = .FindRow(.TextMatrix(i, .ColIndex("相关ID")), , .ColIndex("ID"))
                        If lngTmp <> -1 Then
                            str用法 = .TextMatrix(lngTmp, .ColIndex("诊疗项目"))
                        End If
                    ElseIf InStr(",7,", "," & .TextMatrix(i, .ColIndex("诊疗类别")) & ",") > 0 Then
                        lngTmp = .FindRow(.TextMatrix(i, .ColIndex("相关ID")), , .ColIndex("ID"))
                        If lngTmp <> -1 Then
                            str用法 = .TextMatrix(lngTmp, .ColIndex("诊疗项目"))
                            If .TextMatrix(lngTmp - 1, .ColIndex("诊疗类别")) = "E" Then
                               str用法 = str用法 & "|" & .TextMatrix(lngTmp - 1, .ColIndex("诊疗项目"))
                            End If
                        End If
                    End If
                    
                    '暂未取发药窗口
                    rsSQL!sql = "ZL_门诊划价记录_INSERT(" & _
                        "'" & strNO & "'," & lng费用序号 & "," & lng病人ID & ",NULL," & _
                        IIF(IsNull(rsPati!门诊号), "NULL", "'" & rsPati!门诊号 & "'") & ",NULL,'" & rsPati!姓名 & "'," & _
                        "'" & Nvl(rsPati!性别) & "','" & Nvl(rsPati!年龄) & "'," & _
                        "'" & str费别 & "',NULL," & _
                        ZVal(.TextMatrix(i, .ColIndex("开嘱科室ID"))) & "," & ZVal(.TextMatrix(i, .ColIndex("开嘱科室ID"))) & "," & _
                        "'" & .TextMatrix(i, .ColIndex("开嘱医生")) & "'," & IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                        rsMoney!ID & ",'" & rsMoney!类别 & "','" & Nvl(rsMoney!计算单位) & "',NULL," & _
                        int付数 & "," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                        IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                        "'" & Nvl(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                        str发生时间 & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "'医嘱发送','" & UserInfo.姓名 & "'," & _
                        "'" & IIF(str摘要 = "", .TextMatrix(i, .ColIndex("医嘱内容")), str摘要) & "'," & _
                        Val(.TextMatrix(i, .ColIndex("ID"))) & ",'" & .TextMatrix(i, .ColIndex("执行频次")) & "'," & _
                        ZVal(.TextMatrix(i, .ColIndex("单次用量"))) & ",'" & str用法 & "',1," & _
                        IIF(bln离院带药, 3, Val(.TextMatrix(i, .ColIndex("计价特性")))) & ",1," & _
                        "'" & str保险编码 & "','" & str费用类型 & "'," & IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ")"
                    
                    rsSQL.Update


                    '医保管控实时监测：生成费用项目记录集,以收费细目汇总
                    If Not IsNull(rsPati!险类) And bln实时监控 Then
                        rsItems.Filter = "收费细目ID=" & rsMoney!ID
                        If rsItems.EOF Then
                            '加入收费项目对应的原始信息
                            rsItems.AddNew
                            rsItems!病人ID = lng病人ID
                            rsItems!主页ID = Null
                            rsItems!医嘱ID = Val(.TextMatrix(i, .ColIndex("ID")))
                            rsItems!收费类别 = rsMoney!类别
                            rsItems!收费细目ID = rsMoney!ID
                            rsItems!开单人 = .TextMatrix(i, .ColIndex("开嘱医生"))
                            rsItems!开单科室 = CStr(gobjComlib.Sys.RowValue("部门表", Val(.TextMatrix(i, .ColIndex("开嘱科室ID"))), "名称"))

                            rsItems!数量 = int付数 * dbl数量
                            rsItems!单价 = dbl单价

                            rs医嘱诊断.Filter = "医嘱ID=" & IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) = 0, Val(.TextMatrix(i, .ColIndex("ID"))), Val(.TextMatrix(i, .ColIndex("相关ID"))))
                            If Not rs医嘱诊断.EOF Then
                                rsItems!诊断id = rs医嘱诊断!诊断id
                                rsItems!疾病id = rs医嘱诊断!疾病id
                            End If
                        Else
                            '基于一个医嘱(诊疗项目)的收费对照不会有重复的收费细目
                            '数量：同一收费项目的不同收入项目记录相同
                            If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                                rsItems!数量 = Nvl(rsItems!数量, 0) + int付数 * dbl数量
                            End If
                            '单价：同一收费项目的不同收入项目累加
                            If Val(.TextMatrix(i, .ColIndex("ID"))) = rsItems!医嘱ID Then
                                rsItems!单价 = Nvl(rsItems!单价, 0) + dbl单价
                            End If
                        End If
                        rsItems!实收金额 = Nvl(rsItems!实收金额, 0) + cur实收
                        rsItems.Update
                    End If

                    str收费项目 = rsMoney!费用性质 & "_" & rsMoney!ID
                    rsMoney.MoveNext
                Loop
            End If

            '对医嘱金额进行汇总折扣处理
            If gbln从项汇总折扣 And strHaveSub <> "" Then
                rsSeek.Filter = 0
                Do While Not rsSeek.EOF
                    rsSQL.Bookmark = rsSeek!主项标签

                    str费别 = Nvl(rsPati!费别) & IIF(gstr动态费别 <> "", "," & gstr动态费别, "")
                    If Val(.TextMatrix(i, .ColIndex("零费记帐"))) = 1 Then
                        cur实收 = 0
                    Else
                        cur实收 = Format(ActualMoney(str费别, rsSeek!主收入ID, rsSeek!合计), gstrDec)
                    End If

                    If InStr(str费别, ",") > 0 Then str费别 = Nvl(rsPati!费别)
                    rsSQL!sql = Set动态费别(rsSQL!sql, str费别)

                    cur实收 = cur实收 - rsSeek!合计 '打折差额

                    '医保管控实时监测：费用项目金额替换
                    If Not IsNull(rsPati!险类) And bln实时监控 Then
                        rsItems.Filter = "收费细目ID=" & lng父项目ID
                        If Not rsItems.EOF Then
                            rsItems!实收金额 = Nvl(rsItems!实收金额, 0) + cur实收
                            rsItems.Update
                        End If
                    End If

                    '费用SQL生成替换
                    cur实收 = Get实收金额(rsSQL!sql) + cur实收
                    rsSQL!sql = Set实收金额(rsSQL!sql, cur实收)
                    rsSQL.Update

                    rsSeek.MoveNext
                Loop
            End If

            '产生医嘱发送记录
            '-----------------------------------------------------------------------------------------
            If Val(.TextMatrix(i, .ColIndex("执行性质"))) <> 0 Then '叮嘱不发送(给药途径，配方煎法、用法,采集方法、输血途径可能为)
                '医嘱的发送合性检查提示
                strSQL = "Select zl_AdviceSendCheck([1],[2]) as 结果 From Dual"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceSendCheck", Val(.TextMatrix(i, .ColIndex("ID"))), Replace(str摘要, "''", "'"))
                If Not rsTmp.EOF Then
                    strTmp = Nvl(rsTmp!结果)
                    If strTmp <> "" Then
                        Select Case Val(Split(strTmp, "|")(0))
                        Case 1 '提示
                            GoTo FuncEnd
                        Case 2 '禁止
                            'MsgBox Split(strTmp, "|")(1), vbInformation, "处方自动发送"
                            GoTo FuncEnd
                        End Select
                    End If
                End If

                '一样要产生费用NO
                Call GetCurBillSet(rsBill, strNOKey, strNO, -1, lng发送序号, True, lngNOSequence)

                '是否一组医嘱的第一医嘱行:药疗的第一药品行为第一医嘱行
                blnFirst = False
                If InStr(",5,6,7,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("相关ID"))) <> Val(.TextMatrix(i - 1, .ColIndex("相关ID"))) Then
                        blnFirst = True
                    End If
                End If

                '发送数次:药品为剂量单位的总量,其它为次数
                If .TextMatrix(i, .ColIndex("诊疗类别")) = "7" Then
                    dbl发送数次 = Val(.TextMatrix(i, .ColIndex("总给予量"))) * Val(.TextMatrix(i, .ColIndex("单次用量")))
                ElseIf InStr(",5,6,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0 Then
                    dbl发送数次 = Val(.TextMatrix(i, .ColIndex("总给予量"))) * Val(.TextMatrix(i, .ColIndex("门诊包装"))) * Val(.TextMatrix(i, .ColIndex("剂量系数")))
                Else
                    dbl发送数次 = Val(.TextMatrix(i, .ColIndex("总给予量")))
                End If
                dbl发送数次 = Format(dbl发送数次, "0.00000")

                '首末时间
                If str分解时间 <> "" Then
                    str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                    str末次时间 = "To_Date('" & Split(str分解时间, ",")(UBound(Split(str分解时间, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    '无法分解或为"一次性"临嘱，填为开始执行时间（74366）
                    str首次时间 = "To_Date('" & .TextMatrix(i, .ColIndex("开始执行时间")) & "','YYYY-MM-DD HH24:MI:SS')"
                    str末次时间 = "To_Date('" & .TextMatrix(i, .ColIndex("开始执行时间")) & "','YYYY-MM-DD HH24:MI:SS')"
                End If

                If Not gbln发送生成条形码 Then strCuvetteNumber = ""
                rsSQL.AddNew
                rsSQL!类型 = 3: rsSQL!项目ID = 0: rsSQL!序号 = i
                rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) = 0, Val(.TextMatrix(i, .ColIndex("ID"))), Val(.TextMatrix(i, .ColIndex("相关ID"))))
                rsSQL!NO = strNO
                rsSQL!sql = "ZL_门诊医嘱发送_Insert(" & _
                    Val(.TextMatrix(i, .ColIndex("ID"))) & "," & lng发送号 & "," & 1 & ",'" & strNO & "'," & _
                    lng发送序号 & "," & ZVal(dbl发送数次) & "," & str首次时间 & "," & str末次时间 & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    int执行状态 & "," & ZVal(.TextMatrix(i, .ColIndex("执行科室ID"))) & "," & int计费状态 & "," & _
                    IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                rsSQL.Update

                If gbln血库系统 And .TextMatrix(i, .ColIndex("诊疗类别")) = "K" Then
                    rsSQL.AddNew
                    rsSQL!类型 = 9
                    rsSQL!项目ID = 0
                    rsSQL!序号 = 0
                    rsSQL!sql = "Zl_血液配血费用_Insert(" & Val(.TextMatrix(i, .ColIndex("ID"))) & ")"
                    rsSQL.Update
                End If

                '医嘱执行计价
                If rsExec.RecordCount > 0 Then
                    rsExec.Filter = "医嘱ID=" & Val(.TextMatrix(i, .ColIndex("ID"))) & " And 发送号=" & lng发送号
                    If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                    Do While Not rsExec.EOF
                        rsSQL.AddNew
                        rsSQL!类型 = 8
                        rsSQL!项目ID = 0
                        rsSQL!序号 = 0
                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, .ColIndex("相关ID"))) = 0, Val(.TextMatrix(i, .ColIndex("ID"))), Val(.TextMatrix(i, .ColIndex("相关ID"))))
                        rsSQL!sql = "Zl_医嘱执行计价_Insert(" & rsExec!医嘱ID & "," & rsExec!发送号 & ",To_date('" & _
                        rsExec!要求时间 & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!收费细目ID & "")) & "," & rsExec!数量 & ")"
                        rsSQL.Update
                        rsExec.MoveNext
                    Loop
                    rsExec.Filter = 0
                End If
            End If

            '计算中药配方数
            If .TextMatrix(i, .ColIndex("操作类型")) = "4" And .TextMatrix(i, .ColIndex("诊疗类别")) = "E" Then '中药用法
                int配方数 = int配方数 + 1
            End If
NextAdvice:
            '----------------------------------------
        Next

        '医保管控实时监测
        If Not IsNull(rsPati!险类) And bln实时监控 Then
            rsItems.Filter = 0
            If Not rsItems.EOF Then
                If Not gclsInsure.CheckItem(rsPati!险类, 0, 2, rsItems, Replace(str摘要, "''", "'")) Then GoTo FuncEnd
            End If
        End If
        '提交病人数据
        '-----------------------------------------------------------------------------------------
        Call ReplaceTrueNO(rsSQL, Nothing)
        
        '执行顺序:计价,费用,发送,签名,发料
        '1.对费用记录按收费细目ID排序插入
        rsSQL.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        rsSQL.Sort = "类型,项目ID,序号"
        
        gcnOracle.BeginTrans: blnTran = True
            '执行HIS数据提交
            Do While Not rsSQL.EOF
                Call gobjComlib.zlDatabase.ExecuteProcedure(rsSQL!sql, "OutAdviceSendDrug")
                
                rsSQL.MoveNext
            Loop
                                
            '医保档案上传接口(事务内以限制)
            If Val(rsPati!险类 & "") <> 0 Then
                If gclsInsure.GetCapability(support上传门诊档案, lng病人ID, Val(rsPati!险类 & "")) Then
                    If Not gclsInsure.TranElecDossier(1, lng病人ID, lng挂号id, Val(rsPati!险类 & "")) Then Exit Function
                End If
            End If
        gcnOracle.CommitTrans: blnTran = False
        Screen.MousePointer = 0
    End With
    OutAdviceSendDrug = True

FuncEnd:
    '删除所有已成功发送的行
    Screen.MousePointer = 0
    Exit Function
errH:
    If blnTran Then
        gcnOracle.RollbackTrans
    End If
    If Err.Number <> 0 Then
        Screen.MousePointer = 0
        Call gobjComlib.SaveErrLog
    Else
        Screen.MousePointer = 0
    End If
End Function

Private Function GetMergeDrugStore(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long) As Long
'功能：获取一并给药的基准药房，用于生成发送NO的Key值
'说明：一并给药的药品发送到一起，包括自备药和不同药房的情况
    Dim lng药房ID As Long, lngBegin As Long, i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) <> Val(.TextMatrix(lngRow - 1, .ColIndex("相关ID"))) And Val(.TextMatrix(lngRow, .ColIndex("执行科室ID"))) <> 0 Then
            lng药房ID = Val(.TextMatrix(lngRow, .ColIndex("执行科室ID")))
        Else
            lngBegin = .FindRow(.TextMatrix(lngRow, .ColIndex("相关ID")), , .ColIndex("相关ID"))
            For i = lngBegin To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("相关ID"))) = Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) Then
                    If Val(.TextMatrix(i, .ColIndex("执行科室ID"))) <> 0 Then
                        lng药房ID = Val(.TextMatrix(i, .ColIndex("执行科室ID"))): Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetMergeDrugStore = lng药房ID
End Function

Private Function Set动态费别(ByVal strSQL As String, ByVal str费别 As String) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin费别"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End费别" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Set动态费别 = strLeft & str费别 & strRight
End Function

Private Function TheStockCheck(ByRef colStock1 As Collection, ByRef colStock2 As Collection, ByVal lng库房ID As Long, ByVal str类别 As String) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    If InStr(",5,6,7,", str类别) > 0 Then
        intStyle = colStock1("_" & lng库房ID)
    ElseIf str类别 = "4" Then
        intStyle = colStock2("_" & lng库房ID)
    End If
    Err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Function CheckPriceStock(ByRef colStock1 As Collection, ByRef colStock2 As Collection, ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    rsTotal As ADODB.Recordset, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean) As Boolean
'功能：发送过程中时，对非药嘱药品及跟踪在用的卫材计价进行库存检查(累计检查)
'参数：lngRow=医嘱行号
'      dbl数量=已计算好的计价数量(售价单位)
'      rsTotal=当前病人前面已累计发送的计价药品或卫材数量(售价单位)
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = TheStockCheck(colStock1, colStock2, lng库房ID, rsPrice!类别)
        bln分批 = Nvl(rsPrice!分批, 0) = 1
        bln时价 = Nvl(rsPrice!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = Nvl(rsPrice!门诊单位, Nvl(rsPrice!计算单位)) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品或卫材总量:门诊包装
            dbl总量 = Format(dbl数量 / Nvl(rsPrice!门诊包装, 1), "0.00000")
            
            '当前可用库存:门诊包装,减去前面相同药品医嘱要发送的库存
            If InStr(",5,6,7,", rsPrice!类别) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, .ColIndex("诊疗类别"))) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, .ColIndex("收费细目ID"))) = rsPrice!ID And Val(.TextMatrix(i, .ColIndex("执行科室ID"))) = lng库房ID
                    End If
                    If blnDo Then
                        If .TextMatrix(i, .ColIndex("诊疗类别")) = "7" Then
                            '中药药房单位按不可分零处理:每付
                            If Val(.TextMatrix(i, .ColIndex("可否分零"))) = 0 Then
                                dbl已发库存 = dbl已发库存 + _
                                    Val(.TextMatrix(i, .ColIndex("总给予量"))) * Val(.TextMatrix(i, .ColIndex("单次用量"))) _
                                    / Val(.TextMatrix(i, .ColIndex("剂量系数"))) / Val(.TextMatrix(i, .ColIndex("门诊包装")))
                            Else
                                dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, .ColIndex("总给予量"))) _
                                    * IntEx(Val(.TextMatrix(i, .ColIndex("单次用量"))) / Val(.TextMatrix(i, .ColIndex("剂量系数"))) / Val(.TextMatrix(i, .ColIndex("门诊包装"))))
                            End If
                        Else
                            dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, .ColIndex("总给予量")))
                        End If
                    End If
                Next
            End If
            '计价部份要发送的累计数量
            rsTotal.Filter = "项目ID=" & rsPrice!ID & " And 库房ID=" & lng库房ID
            Do While Not rsTotal.EOF
                dbl已发库存 = dbl已发库存 + Format(rsTotal!数量 / Nvl(rsPrice!门诊包装, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl可用库存 = Format(GetStock(rsPrice!ID, lng库房ID, 2), "0.00000")
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                CheckPriceStock = True
            End If
        End If
        
        '如果未提示或要发送,加入累计发送数量
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, .ColIndex("相关ID"))) <> 0 Then
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, .ColIndex("相关ID")))
            Else
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, .ColIndex("ID")))
            End If
            rsTotal!项目ID = rsPrice!ID
            rsTotal!库房ID = lng库房ID
            rsTotal!数量 = dbl数量
            rsTotal.Update
        End If
    End With
End Function

Public Function ShowApplyCustom(frmParent As Object, ByVal lngFileID As Long, Optional ByRef lng申请序号 As Long) As Boolean
'功能：查看自定义申请单
'参数：frmParent 父对象窗体；
    ShowApplyCustom = gobjKernel.ShowApplyCustom(frmParent, lngFileID, lng申请序号)
End Function

Public Function ShowAppCritical(frmParent As Object, ByVal blnModal As Boolean, ByVal intType As Integer, ByVal int调用类型 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal lng婴儿 As Long, ByRef lng危急值ID As Long, ByVal lng医嘱ID As Long, _
    Optional ByVal lng标本ID As Long, Optional ByVal str危急值描述 As String, Optional ByVal dat报告时间 As Date, Optional ByVal lng报告科室ID As Long, Optional ByVal str报告人 As String, Optional ByRef objMip As Object) As Boolean
'功能：危急值编辑
    ShowAppCritical = gobjKernel.ShowAppCritical(frmParent, blnModal, intType, int调用类型, lng病人ID, lng主页ID, str挂号单, lng婴儿, lng危急值ID, lng医嘱ID, lng标本ID, str危急值描述, dat报告时间, lng报告科室ID, str报告人, objMip)
End Function

Public Function ShowQueryCritical(frmParent As Object, ByVal blnModal As Boolean, ByVal int类型 As Integer, ByVal int场合 As Integer, ByVal lng科室id As Long, ByVal lng病区ID As Long, ByRef objMip As Object) As Boolean
'功能：查询危急值
    ShowQueryCritical = gobjKernel.ShowQueryCritical(frmParent, blnModal, int类型, int场合, lng科室id, lng病区ID, objMip)
End Function

Public Function ShowDealCritical(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByRef lng记录ID As Long) As Boolean
'功能：临床处理危急值
    ShowDealCritical = gobjKernel.ShowDealCritical(frmParent, lng病人ID, lng主页ID, str挂号单, lng记录ID)
End Function

Public Function AppCriticalData(ByVal intType As Integer, ByRef lngID As Long, ByVal str数据来源 As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal lng婴儿 As Long, _
    ByVal lng医嘱ID As Long, ByVal lng标本ID As Long, ByVal str危急值描述 As String, ByVal dat报告时间 As Date, ByVal lng报告科室ID As Long, ByVal str报告人 As String) As Boolean
'功能：数据库层面的接口，功能在于对  病人危急值记录 的增修删
'参数：功能参数，intType 0-新增，1-修改，2-删除，当修改或者删除时必须传入 lngID 值
'        lngID 新增数据时传0，保存后反回值
'        str数据来源  数据来源系统，空表示zlhis系统，非空则传其它系统的名称
'        lng病人ID
'        lng主页ID  如果是门诊病人传0
'        str挂号单  如果是住院病人传空
'        lng婴儿

'        lng医嘱ID
'        lng标本ID
'        dat报告时间，检测出危急值的时间
'        lng报告科室ID，检测科室ID
'        str报告人，检测人
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str姓名 As String
    Dim str性别 As String
    Dim str年龄 As String
    Dim str部门IDs As String
    Dim lng病人来源 As Long
            
    On Error GoTo errH
    
    If intType = 0 Or intType = 1 Then
        strSQL = "select a.姓名,a.性别,a.年龄,a.病人来源,a.病人id,a.主页id,a.挂号单 from 病人医嘱记录 a where a.id=[1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "clsPublicAdvice.AppCriticalData", lng医嘱ID)
        If rsTmp.EOF Then
            MsgBox "未找到相关的医嘱信息。", vbInformation, "clsPublicAdvice.AppCriticalData"
            Exit Function
        End If
        
        str姓名 = rsTmp!姓名 & ""
        str性别 = rsTmp!性别 & ""
        str年龄 = rsTmp!年龄 & ""
        lng病人来源 = Val(rsTmp!病人来源 & "") '1-门诊病人，2-住院病人，3-外来病人，4-体检病人
        
        If intType = 0 Then
            lngID = gobjComlib.zlDatabase.GetNextID("病人危急值记录")
            strSQL = "Zl_病人危急值记录_Insert(" & lngID & ","
        Else
            strSQL = "Zl_病人危急值记录_Update(" & lngID & ","
        End If
        strSQL = strSQL & IIF("" = str数据来源, "null", "'" & str数据来源 & "'") & ","
        strSQL = strSQL & lng病人ID & ","
        strSQL = strSQL & IIF(2 = lng病人来源, lng主页ID, "null") & ","
        strSQL = strSQL & IIF(1 = lng病人来源, "'" & str挂号单 & "'", "null") & ","
        strSQL = strSQL & lng婴儿 & ",'" & str姓名 & "','" & str性别 & "','" & str年龄 & "',"
        strSQL = strSQL & lng医嘱ID & "," & lng标本ID & ","
        strSQL = strSQL & "'" & str危急值描述 & "',"
        strSQL = strSQL & "to_date('" & Format(dat报告时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),"
        strSQL = strSQL & lng报告科室ID & ",'" & str报告人 & "')"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "危急值相关数据处理")
        
        '仅门诊病人或住院病人才产生消息提醒
        If intType = 0 And (lng病人来源 = 1 Or lng病人来源 = 2) Then
        
            If lng主页ID <> 0 Then
                strSQL = "select a.姓名,a.性别,a.年龄,a.主页ID as 就诊ID,a.当前病区ID as 病区ID,a.出院科室ID as 科室ID,2 as 病人来源 from 病案主页 a where a.病人id =[1] and a.主页id=[2]"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "clsPublicAdvice.AppCriticalData", lng病人ID, lng主页ID)
            Else
                strSQL = "select a.姓名,a.性别,a.年龄,a.id as 就诊ID,null as 病区ID,a.执行部门id as 科室ID,1 as 病人来源 from 病人挂号记录 a where a.no=[1]"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "clsPublicAdvice.AppCriticalData", str挂号单)
            End If
            
            If Not rsTmp.EOF Then
                str部门IDs = Val(rsTmp!科室ID & "")
                If Val(rsTmp!病区ID & "") <> 0 Then
                    If Val(rsTmp!病区ID & "") <> Val(rsTmp!科室ID & "") Then
                        str部门IDs = str部门IDs & "," & Val(rsTmp!病区ID & "")
                    End If
                End If
                '产生新开消息
                strSQL = "Zl_业务消息清单_Insert(" & lng病人ID & "," & Val(rsTmp!就诊ID & "") & "," & Val(rsTmp!科室ID) & ","
                strSQL = strSQL & IIF(Val(rsTmp!病区ID & "") = 0, "NULL", Val(rsTmp!病区ID & "")) & "," & Val(rsTmp!病人来源 & "") & ","
                strSQL = strSQL & "'" & str危急值描述 & "','1110','ZLHIS_LIS_003','" & lng医嘱ID & "',3,0,sysdate,'" & str部门IDs & "',null)"
                Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "危急值相关数据处理")
            End If
        End If
    ElseIf intType = 2 Then '删除
        strSQL = "zl_病人危急值记录_delete(" & lngID & ")"
        Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "危急值相关数据处理")
    End If
    AppCriticalData = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Sub zlExecuteCode(ByVal lngSys As Long, ByVal lngModul As Long, ByRef frmParent As Object, ByVal strPrivs As String)
'功能：危急值查询
'说明：该接口供zl9cisjob调用，通过导航台菜台进入
    Call gobjKernel.ShowMainQueryCritical(lngSys, lngModul, frmParent, strPrivs)
End Sub
