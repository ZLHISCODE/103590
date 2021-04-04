VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISKernel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'中药配方界面所需要的病人信息
Public Type TYPE_PatiInfoEx
        lng病人ID As Long
        lng主页ID As Long '主页ID
        str挂号单 As String '挂号单号
        str性别 As String  '病人性别
        int婴儿 As Integer
        int险类 As Integer '医保病人的险类
        bln医保 As Boolean '是否医保病人
        lng病人科室ID As Long '用于确定中药配方的缺省药房
End Type

Public Sub InitCISKernel(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal MainPrivs As String)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    gMainPrivs = MainPrivs
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    
    '初始化操作员信息
    Call GetUserInfo
    
    '初始化系统参数:可分散到具体的GetForm中去初始化
    Call InitSysPar
End Sub

Public Function InitCISKernelPar() As Boolean
    InitCISKernelPar = InitSysPar
End Function

Public Function ShowCommItem(frmParent As Object, ByVal strFind As String, ByRef blnCancle As Boolean, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal intType As Integer) As String
'功能：常用项目的名称
'参数：strFind -为空则查找所有，否则根据strFind查找简码，编码，名称
'      intType 1-抗菌用药理由，2-常用嘱托，3-超量说明，4－常用就诊摘要
    ShowCommItem = frmKssReasonSelect.ShowMe(frmParent, strFind, blnCancle, lngLeft, lngTop, intType)
End Function

Public Sub ShowDiagHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean)
    Call frmDiagHelp.ShowMe(bytModal, frmParent, lngItemID, blnShowInTaskBar)
End Sub

Public Sub ShowClincHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean)
    Call frmClinicHelp.ShowMe(bytModal, frmParent, lngItemID, blnShowInTaskBar)
End Sub

Public Function ShowSchemeEdit(frmParent As Object, ByVal int范围 As Long, Optional rsScheme As ADODB.Recordset, _
    Optional ByVal blnView As Boolean, Optional ByVal bln显示缺省列 As Boolean, Optional ByVal str使用科室 As String, _
    Optional ByVal byt场合 As Byte, Optional ByVal str诊疗分类 As String, Optional ByVal str操作类型 As String, _
    Optional ByVal str执行分类 As String) As ADODB.Recordset
'功能：编辑成套方案接口
'参数：int范围=1-门诊使用,2-住院使用,3-门诊和住院都可以使用
'      blnView=是否仅查看
'      bln显示缺省列=临床路径项目定义调用时"选择使用"
'返回：与表"诊疗项目组合"相同结构的动态记录集,如果取消则返回Nothing
    Set ShowSchemeEdit = frmSchemeEdit.ShowMe(frmParent, int范围, rsScheme, blnView, bln显示缺省列, str使用科室, byt场合, str诊疗分类, str操作类型, str执行分类)
End Function

Public Function ShowAdviceEdit(frmParent As Object, ByVal int场合 As Integer, ByVal bytUseType As Byte, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
      ByVal strAdivceOfPath As String, ByVal dat开始时间 As Date, arrSQL As Variant, strAdviceOfItem As String, Optional ByVal rsLastAdvice As ADODB.Recordset, _
       Optional ByVal datSendTime As Date, Optional ByRef str路径项目IDs As String, Optional ByRef objMip As Object, Optional ByRef datPathOut As Date) As Boolean
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
'      datPathOut      =bytUseType=2时,返回医嘱的开始日期,此日期用于确定该路径外项目添加到路径表单的某一天。
    ShowAdviceEdit = frmInAdviceEdit.ShowMeByPath(frmParent, int场合, gMainPrivs, bytUseType, lng病人ID, lng主页ID, strAdivceOfPath, dat开始时间, arrSQL, strAdviceOfItem, rsLastAdvice, datSendTime, str路径项目IDs, objMip, datPathOut)
End Function

Public Function ShowOutAdviceEdit(frmParent As Object, ByVal int场合 As Integer, ByVal bytUseType As Byte, ByVal lng病人ID As Long, ByVal str挂号单 As String, _
      ByVal strAdivceOfPath As String, ByVal dat开始时间 As Date, arrSQL As Variant, strAdviceOfItem As String, Optional ByVal rsLastAdvice As ADODB.Recordset, _
       Optional ByVal datSendTime As Date, Optional ByRef str路径项目IDs As String, Optional ByRef objMip As Object, Optional ByRef datPathOut As Date, Optional ByVal lng挂号科室ID As Long) As Boolean
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
'      datPathOut      =bytUseType=2时,返回医嘱的开始日期,此日期用于确定该路径外项目添加到路径表单的某一天。
'      lng挂号科室ID
    ShowOutAdviceEdit = frmOutAdviceEdit.ShowMeByPath(frmParent, int场合, gMainPrivs, bytUseType, lng病人ID, str挂号单, strAdivceOfPath, dat开始时间, arrSQL, strAdviceOfItem, rsLastAdvice, datSendTime, str路径项目IDs, objMip, datPathOut, lng挂号科室ID)
End Function

Public Function ShowAdviceOperate(frmParent As Object, ByVal int场合 As Integer, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal strAdviceOfItem As String, ByVal dateStop As Date, Optional ByRef objMip As Object) As Boolean
'功能：用于临床路径执行后调用停止长嘱的功能
'      strAdviceOfItem=路径项目相关的医嘱ID
'      dateStop=停止长嘱的时间
'      objMip 消息对象 zl9ComLib.clsMipModule
    ShowAdviceOperate = frmAdviceOperate.ShowMe(frmParent, gMainPrivs, 1, lng病人ID, lng主页ID, lng病区ID, 0, int场合 = 1, False, 1, strAdviceOfItem, dateStop, , , , , , lng病区ID, , objMip)
End Function

Public Function AdviceOperate(frmParent As Object, ByVal strMainPrivs As String, ByVal int类型 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, _
    Optional ByVal lng医嘱ID As Long, Optional ByRef objMip As Object, Optional ByRef strPatisOut As String, Optional ByVal blnOnePati As Boolean) As Boolean
'功能：双击消息提醒列表时调用校对或者确认停止功能
    AdviceOperate = frmAdviceOperate.ShowMe(frmParent, strMainPrivs, int类型, lng病人ID, lng主页ID, lng病区ID, lng医嘱ID, True, False, 0, , , blnOnePati, , , , , , , objMip, strPatisOut)
End Function

Public Function AdviceSend(frmParent As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strMainPrivs As String, Optional ByRef objMip As Object) As Boolean
'功能：双击消息提醒列表时调用医嘱发送功能
    AdviceSend = frmAdviceSendALL.ShowMe(frmParent, lng病区ID, lng病人ID, lng主页ID, strMainPrivs, False, True, , , objMip)
End Function

Public Function AdviceRisReport(frmParent As Object, ByVal lng病区ID As Long) As Boolean
'功能：批量打印RIS预约单
    AdviceRisReport = frmAdviceRisReport.ShowMe(frmParent, lng病区ID)
End Function

Public Function ShowBloodApply(frmParent As Object, ByVal lng输血医嘱ID As Long) As Boolean
'功能：外部调用输血申请单查看功能，区分门诊和住院直接由挂号单判断
    Dim rsTmp As Recordset, strSQL As String
    Dim bytBaby As Byte
    strSQL = "select 病人ID,主页ID,挂号单,婴儿 From 病人医嘱记录 Where ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "BloodApplyView", lng输血医嘱ID)
    
    If rsTmp.RecordCount > 0 Then
        If rsTmp!挂号单 & "" <> "" Then
            If gbln血库系统 Then
                Call frmApplyBloodNew.ShowMe(frmParent, Val(rsTmp!病人ID & ""), 0, 1, 2, lng输血医嘱ID, , , , , , , , 1, rsTmp!挂号单 & "", , , Val(rsTmp!婴儿 & ""))
            Else
                Call frmApplyBlood.ShowMe(frmParent, Val(rsTmp!病人ID & ""), 0, 1, 2, lng输血医嘱ID, , , , , , , , 1, rsTmp!挂号单 & "", , , Val(rsTmp!婴儿 & ""))
            End If
        Else
            bytBaby = Val(rsTmp!婴儿 & "")
            If gbln血库系统 Then
                Call frmApplyBloodNew.ShowMe(frmParent, Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), 0, 2, lng输血医嘱ID, , , , , , , , , , , , bytBaby)
            Else
                Call frmApplyBlood.ShowMe(frmParent, Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), 0, 2, lng输血医嘱ID, , , , , , , , , , , , bytBaby)
            End If
        End If
    Else
        MsgBox "没有找到您指定的输血医嘱。", vbInformation, gstrSysName
        Exit Function
    End If
    
    ShowBloodApply = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowConsultationApply(frmParent As Object, ByVal lng医嘱ID As Long, Optional ByVal intType As Integer = 2, Optional objFrmEdit As Object) As Boolean
'功能：外部调用会诊申请单查看/书写会诊意见功能，
'参数：intType 3 书写会诊意见，默认是查看会诊申请
    Dim rsTmp As Recordset, strSQL As String
    Dim blnOK As Boolean
    Dim objFrm As Object
    Dim bytBaby As Byte
    
    strSQL = "select 病人ID,主页ID,申请序号,婴儿 From 病人医嘱记录 Where ID=[1]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ConsultationApplyView", lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        bytBaby = Val(rsTmp!婴儿 & "")
        Set objFrm = New frmApplyConsultation
        Set objFrmEdit = objFrm
        blnOK = objFrm.ShowMe(frmParent, lng医嘱ID, Val(rsTmp!申请序号 & ""), intType, , Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), , , , , , , , , , bytBaby)
    Else
        MsgBox "没有找到您指定的会诊医嘱。", vbInformation, gstrSysName
        Exit Function
    End If
    ShowConsultationApply = blnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowOperationApply(frmParent As Object, ByVal lng医嘱ID As Long) As Boolean
'功能：外部调用手术申请单
    Dim rsTmp As Recordset, strSQL As String
    Dim bytBaby As Byte
    
    strSQL = "select 病人ID,主页ID,挂号单,婴儿 From 病人医嘱记录 Where ID=[1]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "OperationApplyView", lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        If rsTmp!挂号单 & "" <> "" Then
            Call frmApplyOperation.ShowMe(frmParent, 1, 2, Val(rsTmp!病人ID & ""), rsTmp!挂号单 & "", 1, lng医嘱ID)
        Else
            bytBaby = Val(rsTmp!婴儿 & "")
            Call frmApplyOperation.ShowMe(frmParent, 0, 2, Val(rsTmp!病人ID & ""), rsTmp!主页ID & "", 0, lng医嘱ID, , , , , , , , , , , , bytBaby)
        End If
    Else
        MsgBox "没有找到您指定的会诊医嘱。", vbInformation, gstrSysName
        Exit Function
    End If
    ShowOperationApply = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowPacsApplication(frmParent As Object, ByVal lng检查医嘱ID As Long, Optional ByVal lngNo As Long) As Boolean
'功能：外部调用检查申请单查看功能
    ShowPacsApplication = ShowApply检查(frmParent, lngNo, lng检查医嘱ID)
End Function

Public Function ChargeDelAudit(frmParent As Object, ByVal lng执行科室ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：双击消息提醒列表时调用销帐审核功能
    Dim strCommon As String
    Dim intAtom As Integer
    
    '调用费用部件功能
    On Error Resume Next
    
    If gobjInExse Is Nothing Then Set gobjInExse = CreateObject("zl9InExse.clsInExse")
    err.Clear: On Error GoTo 0
    
    If gobjInExse Is Nothing Then Exit Function
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    ChargeDelAudit = gobjInExse.CallReCharge(frmParent, gcnOracle, gstrDBUser, glngSys, 1, 1, lng执行科室ID, GetInsidePrivs(p医嘱附费管理), lng病人ID)
    Call GlobalDeleteAtom(intAtom)
End Function

Public Function zlPrintAdvice(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngBabyNum As Long, _
    ByVal int期效 As Integer, Optional ByVal strPrinterName As String, Optional ByVal intMode As Integer, _
    Optional ByRef objReportForm As Object, Optional ByRef strMsg As String) As Boolean
'功能：打印病人医嘱单，完整打印从第一面打到最后一页，打印之后不设置医嘱的打印标记
'参数：int期效 0－长嘱，1－临嘱；strPrinterName 打印机名称(intMode=4 传PDF文件路径)；intMode=1-预览,2-打印，默认为打印;4-PDF ;5-返回Report对象嵌入窗体
'       objReportForm intMode=5 返回 报表嵌入对象
    Dim strSQL As String
    Dim strReport As String
    Dim strWhere As String
    Dim strPrinterTmp As String
    Dim strPdfFile      As String
    
    Dim dat最近重整时间 As Date
    Dim dat日期 As Date
    
    Dim rsTmp As ADODB.Recordset
    Dim rsTable As ADODB.Recordset
    Dim rsColumn As ADODB.Recordset
    Dim rsPLR As ADODB.Recordset
    
    Dim blnHead As Boolean, blnBody As Boolean
    
    Dim i As Long
    Dim lngBodyH As Long, lngHeadH As Long
    Dim lngPrintPos As Long
    Dim lngPrintType As Long
    Dim lngPages As Long
    Dim lng报表行数 As Long
    Dim int中药分行 As Integer
    
    On Error GoTo errH
    

    strReport = IIF(int期效 = 0, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    If int期效 = 0 Then
        int中药分行 = Val(zlDatabase.GetPara("长嘱单中药医嘱单行显示字数", glngSys, p住院医嘱发送))
    Else
        int中药分行 = Val(zlDatabase.GetPara("临嘱单中药医嘱单行显示字数", glngSys, p住院医嘱发送))
    End If
    If int中药分行 > 0 Then int中药分行 = 1
     
    '1、判断有无医嘱
    strSQL = "select 1 from 病人医嘱记录 where 病人ID=[1] and 主页ID=[2] and nvl(婴儿,0)=[3] and 医嘱期效=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", lng病人ID, lng主页ID, lngBabyNum, int期效)
    If rsTmp.EOF Then strMsg = "病人没有医嘱信息！": Exit Function
    
    '2、重整时间
    strSQL = "Select 医嘱重整时间 as 时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取重整时间", lng病人ID, lng主页ID)
    dat最近重整时间 = NVL(rsTmp!时间, CDate("1900-01-01"))
    
    '3、计算报表的行数
    strSQL = "Select A.ID as 报表ID,B.ID,B.W,B.H,B.行高,B.分栏" & _
        " From zlReports A,zlRPTItems B" & _
        " Where A.ID=B.报表ID And B.类型=4 And Nvl(A.系统,0)=[1] And A.编号=[2] And B.格式号=[3]" & _
        " Order by B.W*B.H Desc"
    Set rsTable = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", glngSys, strReport, 1)
    If rsTable.EOF Then strMsg = "计算报表的行数异常退出！": Exit Function
    
    strSQL = "Select 序号,表头,内容 From zlRPTItems Where 报表ID=[1] And 格式号=[2] And 上级ID=[3] And 类型=6 Order by 序号"
    Set rsColumn = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", Val(rsTable!报表ID), 1, Val(rsTable!ID))
    If rsColumn.EOF Then strMsg = "计算报表的行数异常退出！": Exit Function
    
    '以下代码参照自定义报表中的方法编写
    '----------------------------------
    '求出表头高度:以第一列为准
    For i = 0 To UBound(Split(rsColumn!表头, "|"))
        lngHeadH = lngHeadH + Val(Split(Split(rsColumn!表头, "|")(i), "^")(1))
    Next
    
    '求出表体高度
    blnHead = False: blnBody = False
    rsColumn.MoveFirst
    Do While Not rsColumn.EOF
        i = UBound(Split(rsColumn!表头, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(rsColumn!表头, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or Not IsNull(rsColumn!内容)
        rsColumn.MoveNext
    Loop
    If Not blnHead And blnBody Then '仅有表体
        lngBodyH = rsTable!H
    Else
        If rsTable!H - lngHeadH + 15 < 0 Then
            lngBodyH = 0
        Else
            lngBodyH = rsTable!H - lngHeadH + 15
        End If
    End If
    '求出行数
    lng报表行数 = Int(lngBodyH / rsTable!行高) * NVL(rsTable!分栏, 1)
    
    '4.生成医嘱打印数据
    strSQL = "Zl_病人医嘱打印_Insert(" & lng病人ID & "," & lng主页ID & "," & lngBabyNum & "," & int期效 & "," & lng报表行数 & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "zlPrintAdvice")
    
    '5、打印医嘱
    '共有多少页
    strSQL = "select max(页号) as 页数 from 病人医嘱打印  where 病人ID=[1] And 主页ID=[2] And Nvl(婴儿,0)=[3] And 期效=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", lng病人ID, lng主页ID, lngBabyNum, int期效)
    
    lngPages = Val(rsTmp!页数 & "")
    
    '打印机设置
    If intMode = 4 Then
         Call ReportOpen(gcnOracle, glngSys, strReport, frmParent, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, "婴儿=" & lngBabyNum, "打印模式=" & lngPrintType, _
        "停嘱打印=0", "起始行号=1", "StartPageNum=1", "起始页号=1", "结束页号=" & lngPages, "中药分行=" & int中药分行, "PressWorkFirst=0", "PDF=" & strPrinterName, intMode)
    ElseIf intMode = 5 Then
        If strPrinterName <> "" Then
            SetReportPrintSet gcnOracle, glngSys, strReport, "printer", strPrinterName
        End If
        Call LoadReport(gcnOracle, glngSys, strReport, frmParent, objReportForm, Nothing, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, "婴儿=" & lngBabyNum, "打印模式=" & lngPrintType, _
                "停嘱打印=0", "起始行号=1", "StartPageNum=1", "起始页号=1", "结束页号=" & lngPages, "中药分行=" & int中药分行, "PressWorkFirst=0", 1)
    Else
        If strPrinterName <> "" Then
            SetReportPrintSet gcnOracle, glngSys, strReport, "printer", strPrinterName
        End If
        Call ReportOpen(gcnOracle, glngSys, strReport, frmParent, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, "婴儿=" & lngBabyNum, "打印模式=" & lngPrintType, _
        "停嘱打印=0", "起始行号=1", "StartPageNum=1", "起始页号=1", "结束页号=" & lngPages, "中药分行=" & int中药分行, "PressWorkFirst=0", intMode)
    End If
    zlPrintAdvice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Terminate()
'在导航台退出时激活,此时注销
    Set grs医疗付款方式 = Nothing
End Sub

Public Function ShowApplyCustom(frmParent As Object, ByVal lngFileID As Long, Optional ByRef lng申请序号 As Long) As Boolean
'功能：查看自定义申请单
'参数：frmParent 父对象窗体；
    Dim objForm As New frmApplyCustom
    
    ShowApplyCustom = objForm.ShowMe(frmParent, 0, 2, 0, 0, 0, lngFileID, lng申请序号)
End Function

Public Function ShowAppCritical(frmParent As Object, ByVal blnModal As Boolean, ByVal intType As Integer, ByVal int调用类型 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByVal lng婴儿 As Long, ByRef lng危急值ID As Long, ByVal lng医嘱ID As Long, _
    Optional ByVal lng标本id As Long, Optional ByVal str危急值描述 As String, Optional ByVal dat报告时间 As Date, Optional ByVal lng报告科室ID As Long, Optional ByVal str报告人 As String, Optional ByRef objMip As Object) As Boolean
'功能：危急值编辑
    ShowAppCritical = frmCriticalEdit.ShowMe(frmParent, blnModal, intType, int调用类型, lng病人ID, lng主页ID, str挂号单, lng婴儿, lng危急值ID, lng医嘱ID, lng标本id, str危急值描述, dat报告时间, lng报告科室ID, str报告人, objMip)
End Function

Public Function ShowQueryCritical(frmParent As Object, ByVal blnModal As Boolean, ByVal int类型 As Integer, ByVal int场合 As Integer, ByVal lng科室id As Long, ByVal lng病区ID As Long, ByRef objMip As Object) As Boolean
'功能：查询危急值
    ShowQueryCritical = frmCriticalQuery.ShowMe(frmParent, blnModal, int类型, int场合, lng科室id, lng病区ID, objMip)
End Function

Public Function ShowDealCritical(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, ByRef lng记录ID As Long) As Boolean
'功能：临床处理危急值按单个病人处理
'参数：lng记录ID 出参数，本次处理的危急值记录
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnOK As Boolean
    Dim lng危急值ID As Long
    Dim int调用类型 As Integer
    Dim lng医嘱ID As Long
 
    If lng主页ID = 0 Then
        int调用类型 = 1
        strSQL = "select a.id,a.医嘱ID,a.姓名,a.性别,a.年龄,a.报告时间,a.状态,a.是否危急值, a.危急值描述 from 病人危急值记录 a where a.挂号单=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsCIKernel", str挂号单)
    Else
        int调用类型 = 2
        strSQL = "select a.id,a.医嘱id,a.姓名,a.性别,a.年龄,a.报告时间,a.状态,a.是否危急值, a.危急值描述  from 病人危急值记录 a where a.病人ID=[1] and a.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsCIKernel", lng病人ID, lng主页ID)
    End If
    
    If rsTmp.EOF Then
        MsgBox "该患者未找到相关的危急值记录。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTmp.RecordCount > 1 Then
        lng危急值ID = frmCriticalQuery.ShowMeSelCard(frmParent, rsTmp)
        If lng危急值ID = 0 Then Exit Function
        rsTmp.Filter = "ID=" & lng危急值ID
        If rsTmp.EOF Then Exit Function
    End If
    
    lng危急值ID = Val(rsTmp!ID & "")
    lng医嘱ID = Val(rsTmp!医嘱ID & "")
    
    blnOK = frmCriticalEdit.ShowMe(frmParent, True, 3, int调用类型, lng病人ID, lng主页ID, str挂号单, 0, lng危急值ID, lng医嘱ID)
    If blnOK Then
        lng记录ID = lng危急值ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMainQueryCritical(ByVal lngSys As Long, ByVal lngModul As Long, ByRef frmParent As Object, ByVal strPrivs As String) As Boolean
'功能：独立查询功能
    ShowMainQueryCritical = frmCriticalQuery.ShowMeQuery(lngSys, lngModul, frmParent, strPrivs)
End Function
