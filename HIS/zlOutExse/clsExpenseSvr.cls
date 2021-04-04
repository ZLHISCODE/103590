VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpenseSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mlngModule  As Long, mstrPrivs As String
Private mobjExpenceSvr As zlPublicExpense.clsExpenceSvr, mobjPubService As zlPublicExpense.clsService
Private mobjDrugStuff As clsDrugStuff

Public Function InitCommon(ByVal lngModule As Long, strPrivs As String, _
    Optional ByRef objDrugStuff As clsDrugStuff, Optional ByVal blnNewBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关基础数据
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    mlngModule = lngModule: mstrPrivs = strPrivs
    If GetPubExpenceSvr(mobjExpenceSvr, lngModule) = False Then Exit Function
    
    Set mobjDrugStuff = New clsDrugStuff
    If mobjDrugStuff.InitCommon(lngModule, strPrivs, blnNewBill) = False Then Exit Function
    
    If mobjDrugStuff.GetPubService(mobjPubService, lngModule) = False Then Exit Function
    
    Set objDrugStuff = mobjDrugStuff
    InitCommon = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPubExpenceSvr(ByRef objPubExpenceSvr As Object, ByVal lngModule As Long) As Boolean
    '创建单据操作公共部件
    If gobjPubExpenceSvr Is Nothing Then
        Set gobjPubExpenceSvr = CreateObject("zlPublicExpense.clsExpenceSvr")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense.clsExpenceSvr)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    If gobjPubExpenceSvr Is Nothing Then Exit Function
    If gobjPubExpenceSvr.zlInitCommon(glngSys, lngModule, gcnOracle, gstrDBUser) = False Then Exit Function
    
    Set objPubExpenceSvr = gobjPubExpenceSvr
    GetPubExpenceSvr = True
End Function

Private Function GetRegsitIdByAdviceID(ByVal lng医嘱ID As Long, _
    ByRef lng挂号ID As Long, ByRef bln急诊 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:通过医嘱ID获取就诊ID及急诊标志
    '入参:
    '   lng医嘱ID 医嘱ID
    '出参:
    '   lng挂号ID 挂号单号ID
    '   bln急诊 是否为急诊
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSql As String
    Dim rsAdviceData As ADODB.Recordset, str挂号单 As String
    
    On Error GoTo ErrHandler
    If ZlGetAdviceInfoFromID(lng医嘱ID, rsAdviceData) = False Then Exit Function
    If rsAdviceData.EOF Then GetRegsitIdByAdviceID = True: Exit Function
    
    str挂号单 = Nvl(rsAdviceData!挂号单号)
    If str挂号单 = "" Then GetRegsitIdByAdviceID = True: Exit Function
    
    strSql = "Select ID,急诊 From 病人挂号记录 Where No=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取挂号单ID", str挂号单)
    If Not rsTmp.EOF Then
        lng挂号ID = Val(Nvl(rsTmp!ID))
        bln急诊 = Val(Nvl(rsTmp!急诊)) = 1
    End If
    GetRegsitIdByAdviceID = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BuildNewBillsData(objBill As ExpenseBill, ByRef cllRcpBillData As Collection, _
    ByVal intBillType As Byte, ByVal bln划价 As Boolean, cllElseInfo As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:构建新单据数据
    '入参:
    '   intBillType-单据类型(1-收费处方;2-记帐单处方)
    '   bln划价:是否划价单
    '   cllElseInfo 其它信息集合，成员：
    '       [病人来源,医嘱ID,险类,险类名称,医疗付款方式名称,医疗付款方式编码,是否急诊]
    '出参:
    ' cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '  intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim p As Integer, i As Integer
    Dim cllPati As Collection
    Dim cllBillLists As Collection, cllBillList As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim objBillPage As BillPage, objBillDetail As BillDetail, objBillInCome As BillInCome
    Dim int序号  As Integer, int行号 As Integer, int价格父号 As Integer
    Dim dbl数次 As Double, dbl单价 As Double, blnFirstRow As Boolean
    Dim bln自动发料 As Boolean, strPayDrugWins As String
    Dim lng病人科室ID As Long
    Dim lng医嘱ID As Long, lng就诊ID As Long, bln是否急诊 As Boolean
    
    On Error GoTo ErrHandler
    '一、单据信息---------------------------------------------------------------
    Set cllBillLists = New Collection
    
    bln自动发料 = False
    If intBillType = 2 And Not bln划价 Then
        bln自动发料 = IIf(gbln门诊自动发料, 1, 0)
    End If
    
    lng医嘱ID = cllElseInfo("医嘱ID")
    If lng医嘱ID = 0 Then
        lng就诊ID = 0
        bln是否急诊 = cllElseInfo("是否急诊") = 1
    Else
        If GetRegsitIdByAdviceID(lng医嘱ID, lng就诊ID, bln是否急诊) = False Then Exit Function
    End If
    
    blnFirstRow = True
    For p = 1 To objBill.Pages.Count
        Set objBillPage = objBill.Pages(p)
        If objBillPage.NO = "" Then '划价单不处理
            int序号 = 0: int行号 = 0
            int价格父号 = 0
            strPayDrugWins = ""
            
            Set cllBillDetails = New Collection
            For Each objBillDetail In objBillPage.Details
                If objBillDetail.数次 <> 0 Then
                    For Each objBillInCome In objBillDetail.InComes
                        int序号 = int序号 + 1
                        
                        '2.收费细目部份---------------------------------------------------------------
                        With objBillDetail
                            If .序号 <> int行号 Then   '处理从属父号
                                int行号 = .序号
                                int价格父号 = int序号
                                '重新处理从属父号
                                If objBillPage.Details(.序号).从属父号 = 0 Then
                                    For i = .序号 + 1 To objBillPage.Details.Count
                                        If objBillPage.Details(i).从属父号 = .序号 Then
                                            '当父项目有多个收入项目(多个序号)时,取第一个序号
                                            objBillPage.Details(i).从属父号 = int序号
                                        End If
                                    Next
                                End If
                            End If
                            
                            '收费、划价的药品行,处理发药窗口
                            If (intBillType = 1 Or intBillType = 2 And Not bln划价) And InStr(",5,6,7,", .收费类别) > 0 Then
                                .发药窗口 = mobjDrugStuff.GetDrugWindowFromBill(objBill, p, .执行部门ID)
                                If .发药窗口 = "" Then
                                    .发药窗口 = mobjDrugStuff.Get发药窗口(objBill.病人ID, .收费类别, .执行部门ID, blnFirstRow)
                                    If blnFirstRow Then blnFirstRow = False
                                End If
                                If InStr(strPayDrugWins, ";" & .执行部门ID & "|") = 0 And .发药窗口 <> "" Then
                                    strPayDrugWins = strPayDrugWins & ";" & .执行部门ID & "|" & .发药窗口
                                End If
                            End If
                            
                            If .付数 = 0 Then .付数 = 1
                            
                            dbl数次 = .数次
                            If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                                dbl数次 = Format(.数次 * .Detail.药房包装, "0.00000")
                            End If
                        End With
                
                        '3.收入项目部份---------------------------------------------------------------
                        With objBillInCome
                            dbl单价 = .标准单价
                            If InStr(",5,6,7,", objBillDetail.收费类别) > 0 And gbln药房单位 Then
                                dbl单价 = Format(.标准单价 / objBillDetail.Detail.药房包装, gstrFeePrecisionFmt)
                            End If
                        End With
                        
                        Set cllBillDetail = New Collection
                        '序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
                        cllBillDetail.Add int序号, "序号"
                        cllBillDetail.Add IIf(objBillDetail.从属父号 = 0, Empty, objBillDetail.从属父号), "从属父号"
                        cllBillDetail.Add objBillDetail.Detail.药名ID, "药名ID"
                        cllBillDetail.Add objBillDetail.收费细目ID, "收费细目ID"
                        cllBillDetail.Add IIf(int价格父号 = int序号, Empty, int价格父号), "价格父号"
                        cllBillDetail.Add objBillInCome.收入项目ID, "收入项目ID"
                        cllBillDetail.Add IIf(objBillDetail.费别 = "", objBill.费别, objBillDetail.费别), "费别"
                        cllBillDetail.Add IIf(objBill.婴儿费 = 0, Empty, objBill.婴儿费), "婴儿序号"
                        cllBillDetail.Add objBillDetail.收费类别, "收费类别"
                        cllBillDetail.Add objBillDetail.计算单位, "计算单位"
                        cllBillDetail.Add IIf(objBillDetail.保险项目否, 1, 0), "是否保险项目"
                        cllBillDetail.Add IIf(objBillDetail.保险大类ID = 0, Empty, objBillDetail.保险大类ID), "保险大类ID"
                        
                        '保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
                        cllBillDetail.Add objBillDetail.保险编码, "保险编码"
                        cllBillDetail.Add objBillInCome.收据费目, "收据费目"
                        cllBillDetail.Add objBillDetail.付数, "付数"
                        cllBillDetail.Add dbl数次, "数次"
                        cllBillDetail.Add dbl单价, "单价"
                        cllBillDetail.Add objBillInCome.应收金额, "应收金额"
                        cllBillDetail.Add objBillInCome.实收金额, "实收金额"
                        cllBillDetail.Add objBillInCome.统筹金额, "统筹金额"
                        cllBillDetail.Add objBillDetail.附加标志, "附加标志"
                        cllBillDetail.Add objBillDetail.执行部门ID, "执行部门ID"
                        If objBillDetail.收费类别 = "4" Then
                            cllBillDetail.Add IIf(bln自动发料, 1, 0), "是否自动发放"
                        Else
                            cllBillDetail.Add 0, "是否自动发放"
                        End If
                        cllBillDetail.Add objBillDetail.摘要, "费用摘要"
                        '医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
                        cllBillDetail.Add IIf(lng医嘱ID = 0, Empty, lng医嘱ID), "医嘱ID"
                        cllBillDetail.Add IIf(lng就诊ID = 0, Empty, lng就诊ID), "就诊ID"
                        cllBillDetail.Add objBillDetail.Detail.类型, "费用类型"
                        cllBillDetail.Add objBillDetail.Detail.中药形态, "中药形态"
                        cllBillDetail.Add objBillPage.煎法, "煎法"
                        cllBillDetail.Add Empty, "执行性质"
                        cllBillDetail.Add Empty, "是否备货卫材"
                        cllBillDetail.Add IIf(objBillDetail.Detail.批次 <= 0, Empty, objBillDetail.Detail.批次), "备货材料批次"
                        cllBillDetail.Add IIf(objBillDetail.Detail.跟踪在用, 1, 0), "是否跟踪在用"
                        cllBillDetail.Add objBillDetail.发药窗口, "发药窗口"
                        '组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
                        '医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
                        '发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
                        
                        cllBillDetails.Add cllBillDetail, "_" & int序号
                    Next '下一条收入项目
                End If
            Next '下一行收费项目
            
            If cllBillDetails.Count > 0 Then
                Set cllBillList = New Collection
                '单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
                cllBillList.Add Empty, "单据号"
                cllBillList.Add Empty, "是否多病人单"
                cllBillList.Add Empty, "医疗小组id"
                cllBillList.Add Empty, "医技补临床费用"
                cllBillList.Add Empty, "简单记帐"
                cllBillList.Add Empty, "记帐单ID"
                cllBillList.Add IIf(intBillType = 2, 2, 1), "记录性质"
                cllBillList.Add IIf(bln划价, 1, 0), "是否划价"
                cllBillList.Add IIf(bln是否急诊, 1, 0), "是否急诊"
                cllBillList.Add objBill.加班标志, "加班标志"
                cllBillList.Add IIf(objBillPage.开单部门ID = 0, Empty, objBillPage.开单部门ID), "开单科室ID"
                '领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)]
                cllBillList.Add Empty, "领药部门ID"
                cllBillList.Add objBillPage.开单人, "开单人"
                cllBillList.Add UserInfo.姓名, "划价人"
                cllBillList.Add UserInfo.姓名, "操作员姓名"
                cllBillList.Add UserInfo.编号, "操作员编号"
                cllBillList.Add objBill.发生时间, "发生时间"
                cllBillList.Add Empty, "登记时间"
                lng病人科室ID = IIf(objBill.科室ID = 0, objBillPage.开单部门ID, objBill.科室ID)
                cllBillList.Add IIf(lng病人科室ID = 0, Empty, lng病人科室ID), "病人科室ID"
                cllBillList.Add cllBillDetails, "_cllBillDetails"
                
                cllBillLists.Add cllBillList, "_" & p
                
                '保存前一张单据的药房ID,以便多张单据时确定发药窗口
                If objBill.Pages.Count > 1 Then Call mobjDrugStuff.SaveDrugID(objBill, p)
            End If
            
            objBillPage.发药窗口 = Mid(strPayDrugWins, 2) '划价收费时使用
        End If
    Next '下一张单据
    
    '二、病人信息---------------------------------------------------------------
    Set cllPati = New Collection
    '病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    cllPati.Add IIf(objBill.病人ID = 0, Empty, objBill.病人ID), "病人ID"
    cllPati.Add IIf(objBill.主页ID = 0, Empty, objBill.主页ID), "主页ID"
    cllPati.Add objBill.姓名, "姓名"
    cllPati.Add objBill.性别, "性别"
    cllPati.Add objBill.年龄, "年龄"
    cllPati.Add objBill.费别, "费别"
    cllPati.Add objBill.床号, "床号"
    lng病人科室ID = IIf(objBill.科室ID = 0, objBill.Pages(1).开单部门ID, objBill.科室ID) '取第一张单据的开单科室
    cllPati.Add IIf(lng病人科室ID = 0, Empty, lng病人科室ID), "病人科室ID"
    cllPati.Add IIf(objBill.病区ID = 0, Empty, objBill.病区ID), "病区ID"
    cllPati.Add cllElseInfo("险类"), "险类"
    '险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    cllPati.Add cllElseInfo("险类名称"), "险类名称"
    cllPati.Add cllElseInfo("病人来源"), "病人来源"
    cllPati.Add objBill.标识号, "标识号"
    '患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    cllPati.Add objBill.身份, "身份"
    cllPati.Add objBill.出生日期, "出生日期"
    cllPati.Add objBill.身份证号, "身份证号"
    cllPati.Add cllElseInfo("医疗付款方式名称"), "医疗付款方式名称"
    cllPati.Add cllElseInfo("医疗付款方式编码"), "医疗付款方式编码"
    If objBill.病人ID = 0 Then
        cllPati.Add "", "病人类型"
    Else
        cllPati.Add ZLGetPatientType(objBill.病人ID), "病人类型"
    End If
    
    Set cllRcpBillData = New Collection
    cllRcpBillData.Add cllPati, "_patiinfor"
    cllRcpBillData.Add cllBillLists, "_cllBillLists"
    
    BuildNewBillsData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function UpdateNewBillsData(ByRef cllRcpBillData As Collection, ByRef strNos As String, ByVal dtNow As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新新单据的数据
    '入参:
    ' cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '   dtNow 登记时间，格式："yyyy-MM-dd HH:mm:ss"
    '出参：
    '   strNOs 分配的单据号，与单据顺序保持一致
    '返回:
    '说明:更新内容：单据号,登记时间
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim cllBillLists As Collection
    Dim byt记录性质 As Byte, strBillNO As String
    
    On Error GoTo ErrHandler
    strNos = ""
    Set cllBillLists = cllRcpBillData("_cllBillLists")
    If cllBillLists.Count = 0 Then UpdateNewBillsData = True: Exit Function
    
    For i = 1 To cllBillLists.Count
        byt记录性质 = cllBillLists(i)("记录性质")
        
        Select Case byt记录性质
        Case 1 '收费单
            strBillNO = zlDatabase.GetNextNo(13)
        Case 2  '记帐单
            strBillNO = zlDatabase.GetNextNo(14)
        End Select
        
        '更新节点
        cllBillLists(i).Remove "单据号"
        cllBillLists(i).Add strBillNO, "单据号"
        
        cllBillLists(i).Remove "登记时间"
        cllBillLists(i).Add Format(dtNow, "yyyy-MM-dd HH:mm:ss"), "登记时间"
        
        strNos = strNos & "," & strBillNO
    Next
    strNos = Mid(strNos, 2)
    
    UpdateNewBillsData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteDrugStuffBillDel(cllBillLists As Collection, ByRef cllPro As Collection, _
    ByVal intBillType As Integer, Optional ByVal bln划价 As Boolean, _
    Optional ByVal blnExecuteDelRcp As Boolean = True, _
    Optional ByRef strJosnDelDrugRcp_Out As String, Optional ByRef strDrugSendFeeIDs_Out As String, _
    Optional ByRef strJosnDelStuffRcp_Out As String, Optional ByRef strStuffSendFeeIDs_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方单、发料单的删除
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    '   blnExecuteDelRcp 是否执行删除药品\卫材单据
    '出参:
    '   cllPro-需要执行的SQL
    '   strJosnDelDrugRcp_Out blnExecuteDelRcp=False时返回的作废处方单需要的JSON入参
    '   strDrugSendFeeIDs_Out blnExecuteDelRcp=False时返回的作废处方单需要的费用ID入参
    '   strJosnDelStuffRcp_Out blnExecuteDelRcp=False时返回的作废材料单需要的JSON入参
    '   strStuffSendFeeIDs_Out blnExecuteDelRcp=False时返回的作废材料单需要的费用ID入参
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJosnDelDrugRcp As String, strJosnDelStuffRcp As String
    Dim strDrugSendFeeIDs As String, strStuffSendFeeIDs As String
    
    On Error GoTo ErrHandler
    Set cllPro = New Collection
    strJosnDelDrugRcp_Out = "": strDrugSendFeeIDs_Out = ""
    strJosnDelStuffRcp_Out = "": strStuffSendFeeIDs_Out = ""
    
    '1.先收集费用保存数据过程及生成处方数据信息
    If mobjExpenceSvr.zlGetDelBill_SaveData(cllBillLists, cllPro, _
        strJosnDelDrugRcp, strJosnDelStuffRcp, _
        strDrugSendFeeIDs, strStuffSendFeeIDs, intBillType, bln划价, True) = False Then Exit Function
    
    If blnExecuteDelRcp Then
        '2.存在药品，需要作废处方
        If strJosnDelDrugRcp <> "" Then
            If mobjExpenceSvr.zlExcute_SaveDrugDelRecipeBill(strDrugSendFeeIDs, strJosnDelDrugRcp, True) = False Then Exit Function
        End If
        
        '3.存在卫生材料，需要作废发料单
        If strJosnDelStuffRcp <> "" Then
            If mobjExpenceSvr.zlExcute_SaveStuffDelBill(strStuffSendFeeIDs, strJosnDelStuffRcp, True) = False Then Exit Function
        End If
    Else
        '不执行时返回数据
        strJosnDelDrugRcp_Out = strJosnDelDrugRcp: strDrugSendFeeIDs_Out = strDrugSendFeeIDs
        strJosnDelStuffRcp_Out = strJosnDelStuffRcp: strStuffSendFeeIDs_Out = strStuffSendFeeIDs
    End If
    
    ExecuteDrugStuffBillDel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcute_SaveRecipeBill_Check(cllRcpBillData As Collection, ByVal intBillType As Integer, _
    Optional ByVal bln划价 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方保存前检查操作
    '入参:
    ' cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '  intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '  bln划价-是否划价单
    '出参:
    '返回:合法返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillLists As Collection
    
    On Error GoTo ErrHandler
    Set cllBillLists = cllRcpBillData("_cllBillLists")
    If cllBillLists.Count = 0 Then zlExcute_SaveRecipeBill_Check = True: Exit Function
    
    zlExcute_SaveRecipeBill_Check = mobjExpenceSvr.zlExcute_SaveRecipeBill_Check( _
        cllRcpBillData, intBillType, , bln划价, True, mlngModule)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcute_SaveRecipeBill(ByVal cllRctBillData As Collection, ByVal strFrmCaption As String, _
    ByVal intBillType As Integer, Optional ByVal bln划价 As Boolean, Optional cllPro As Collection, _
    Optional ByVal blnUpdatePatiInfo As Boolean = True, Optional ByVal str原姓名 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方数据保存操作
    '入参:
    ' cllRcpBillData(结构)
    '   |-cllPati(Collect):病人信息，成员：
    '     (病人ID,主页ID,姓名,性别,年龄,费别,床号,病人科室ID,病区ID,险类,
    '      险类名称,病人来源,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '      患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '      身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllRcpBillData(_patiinfor)
    '   |-cllBillLists(Collect):单据信息集=cllRcpBillData(_cllBillLists)
    '     |-cllBillList(Collect):单据信息，成员：
    '       (单据号,是否多病人单,医疗小组id,医技补临床费用,简单记帐,记帐单id,记录性质,是否划价,是否急诊,加班标志,开单科室ID,
    '        领药部门ID,开单人,划价人,操作员姓名,操作员编号,发生时间,登记时间,病人科室ID,[cllBillDetails(collect)])=cllBillists(_单据号)
    '       |-cllBillDetails(Collect):单据明细集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,从属父号,药名ID,收费细目ID,价格父号,收入项目ID,费别,婴儿序号,收费类别,计算单位,是否保险项目,保险大类ID,
    '            保险编码,收据费目,付数,数次,单价,应收金额,实收金额,统筹金额,附加标志,执行部门ID,是否自动发放,费用摘要(N),
    '            医嘱ID(N),就诊ID(N),费用类型,中药形态,煎法,执行性质,是否备货卫材(N),备货材料批次(N),是否跟踪在用,发药窗口(N),
    '            组内序号(N),诊疗项目ID(N),给药途径ID(N),给药途径名称(N),给药途径分类(N),给药频次ID(N)，给药频次名称（N),
    '            医嘱紧急标志(N),医嘱期效(N),计价特性(N),频次(N),单量（N),用法(N),皮试结果(N),超量说明(N),使用嘱托(N),
    '            发药方式(N),药品含量(N),门诊执行天数(N),煎法(N)
    '            【记帐表增加】:(病人ID,主页ID,姓名,性别,年龄,床号,病人科室ID,病区ID,险类,
    '              险类名称,标识号(门诊号或住院号),挂号单ID(N),代办人姓名(N),代办人身份证号(N),
    '              病人来源,患者体重(N),患者体重单位(N),诊断记录ID(N),诊断ID(N),诊断名称(N),
    '              身份,出生日期,身份证号,医疗付款方式名称,医疗付款方式编码,费用审核标志(N),住院状态(N))=cllBillDetails(_序号)
    '    以上元素中，带(N)的，表示可选节点
    '  intBillType-单据类型(1-收费单;2-记帐单;3-记帐表)
    '  cllPro 需要一起执行的SQL语句
    '  blnUpdatePatiInfo-是否更新病人信息
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillLists As Collection, cllPati As Collection
    Dim cllNewPati As Collection
    
    On Error GoTo ErrHandler
    Set cllBillLists = cllRctBillData("_cllBillLists")
    If cllBillLists.Count = 0 Then zlExcute_SaveRecipeBill = True: Exit Function
    
    '划价收费单，更新“新病人”信息
    Set cllPati = cllRctBillData("_patiinfor")
    If intBillType = 1 And bln划价 And Val(cllPati("病人ID")) <> 0 And blnUpdatePatiInfo Then
        Set cllNewPati = New Collection
        cllNewPati.Add cllPati("姓名"), "姓名"
        cllNewPati.Add cllPati("性别"), "性别"
        cllNewPati.Add cllPati("年龄"), "年龄"
        cllNewPati.Add cllPati("费别"), "费别"
        cllNewPati.Add cllPati("医疗付款方式名称"), "医疗付款方式名称"
        If zlUpdateNewPati(cllPati("病人ID"), str原姓名, cllNewPati, True) = False Then Exit Function
    End If
    
    zlExcute_SaveRecipeBill = mobjExpenceSvr.zlExcute_SaveRecipeBill( _
        cllRctBillData, strFrmCaption, intBillType, bln划价, , True, cllPro, , mlngModule)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlUpdateNewPati(ByVal lng病人ID As Long, ByVal str原姓名 As String, _
    ByVal cllNewPati As Collection, ByVal bln更新费别 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:划价收费单，更新“新病人”信息
    '入参:
    '   str原姓名-原始姓名
    '   cllNewPati(Collect)-新的病人信息，成员:(姓名,性别,年龄,费别,医疗付款方式名称)
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If lng病人ID = 0 Then zlUpdateNewPati = True: Exit Function
    
    zlUpdateNewPati = mobjPubService.ZlPatisvr_UpdateNewPati( _
        lng病人ID, str原姓名, cllNewPati("姓名"), cllNewPati("性别"), cllNewPati("年龄"), _
        cllNewPati("医疗付款方式名称"), cllNewPati("费别"), bln更新费别, mlngModule)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlVerfyBillingPriceBill(ByVal byt病人来源 As Byte, ByVal strNos As String, _
    Optional ByVal strVerfyDate As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐划价单审核为记帐单
    '入参:
    '   byt病人来源:1-门诊;2-住院
    '   strNos -单据信息, 格式：NO1:序号1,序号2,...|NO1:序号1,序号2,...|...
    '   strVerfyDate-审核日期(为空时，以当前时间为准):格式：yyyy-mm-dd HH:MM:SS
    '出参:
    '返回:执行成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlVerfyBillingPriceBill = mobjExpenceSvr.zlVerfyBillingPriceBill( _
        byt病人来源, strNos, strVerfyDate)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlDrugRecipeAffirm(ByVal strNos As String, ByVal byt单据性质 As Byte, _
    ByVal byt单据类型 As Byte, Optional ByVal blnAutoSend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费成功后药品费用确认，同时自动发药
    '入参:
    '   byt单据性质 1-门诊单据,2-住院单据
    '   byt单据类型 1-收费单;2-记帐单
    '   blnAutoSend 是否自动发药
    '出参:
    '返回:执行成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlDrugRecipeAffirm = mobjExpenceSvr.zlDrugOutRecipeAffirm(strNos, byt单据性质, byt单据类型, blnAutoSend)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlStuffBillAffirm(ByVal strNos As String, ByVal byt单据性质 As Byte, _
    ByVal byt单据类型 As Byte, Optional ByVal blnAutoSend As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费、记帐划价审核成功后卫材费用确认，同时自动发料
    '入参:
    '   byt单据性质 1-门诊单据,2-住院单据
    '   byt单据类型 1-收费单;2-记帐单
    '   blnAutoSend 是否自动发料
    '出参:
    '返回:执行成功，返回true,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    ZlStuffBillAffirm = mobjExpenceSvr.zlStuffOutBillAffirm(strNos, byt单据性质, byt单据类型, blnAutoSend)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlDelBillingBillCheck(ByVal strNos As String, ByRef cllBillLists As Collection, _
    Optional ByVal bytBillType As Byte = 1) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除记帐单前检查数据合法性
    '入参:
    '   strNos:记帐单据信息,格式：A001:序号1,序号2,...|A002:序号1,序号2,...|...
    '   bln门诊:是否门诊记帐单
    '   intBillType-单据类型(1-收费处方;2-记帐单处方;3-自动记帐单)
    '出参:
    '   cllBillLists:删除执行入参数据
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillList As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim varNos As Variant, i As Integer
    Dim varNo As Variant, j As Integer
    Dim strNo As String, strNums As String, varNums As Variant
    
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    On Error GoTo ErrHandler
    
    Set cllBillLists = New Collection
    
    varNos = Split(strNos, "|")
    For i = 0 To UBound(varNos)
        varNo = Split(varNos(i) & ":", ":")
        strNo = varNo(0): strNums = varNo(1)
        
        varNums = Split(strNums, ",")
        Set cllBillDetails = New Collection
        For j = 0 To UBound(varNums)
            Set cllBillDetail = New Collection
            cllBillDetail.Add varNums(j), "序号"
            
            cllBillDetails.Add cllBillDetail, "_" & varNums(j)
        Next
    
        Set cllBillList = New Collection
        cllBillList.Add strNo, "单据号"
        cllBillList.Add UserInfo.编号, "操作员编号"
        cllBillList.Add UserInfo.姓名, "操作员姓名"
        cllBillList.Add Empty, "已结禁止销帐"
        cllBillList.Add Empty, "禁止部分销帐"
        cllBillList.Add Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"), "登记时间"
        cllBillList.Add cllBillDetails, "_cllBillDetails"
        
        cllBillLists.Add cllBillList, "_" & strNo
    Next
    
    If mobjExpenceSvr.zlExcute_DelRecipeBill_Check(cllBillLists, bytBillType, mlngModule, False, True) = False Then Exit Function
    
    zlDelBillingBillCheck = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcute_DelRecipeBill(ByVal cllBillLists As Collection, _
    ByVal strFrmCaption As String, ByVal intBillType As Integer, _
    Optional ByVal bln划价 As Boolean, Optional cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行处方数据作废操作
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    '   cllPro 需要一起执行的SQL语句
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlExcute_DelRecipeBill = mobjExpenceSvr.zlExcute_DelRecipeBill( _
        cllBillLists, strFrmCaption, intBillType, bln划价, mlngModule, True, cllPro)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlReadOutPatientDelBillData(cllInPara As Collection, ByVal strFrmCaption As String, _
    Optional ByRef rsBillData_out As ADODB.Recordset, _
    Optional ByRef rsIncome_out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取门诊费用单据剩余/准退费用信息
    '入参:
    '   cllInPara 参数列表，集合，成员(Key)：
    '    [单据类型,单据号,显示单位(N),药品名称显示方式(N),已结禁止销帐(N),操作方式(N),医嘱ID(N),
    '     更新执行状态(N),取收入汇总(N),仅取剩余数量(N),收入汇总方式(N),取医嘱内容(N)]
    '       单据类型:1-收费单;2-记帐单;3-自动记帐单
    '       单据号:多个英文逗号分隔；格式：A001,A002,...
    '       显示单位:0-售价单位;1-门诊单位;2-住院单位 仅操作方式=0时有效
    '       药品名称显示方式:0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    '       已结禁止销帐:已结账数是否禁止销账,记帐单有效:1-是;0-否
    '       操作方式:0-按单据详细信息返回;1-仅只返回简要信息：费用ID,序号,付数,数次,执行状态
    '       更新执行状态：在获取时，发现执行状态不正确时，是否立即更新:1-是;0-否
    '       取收入汇总：是否获取收入项目统计信息:1-是;0-否
    '       仅取有剩余数量：仅取有剩余数量的数据（不减去已执行数量）:1-是;0-否
    '       收入汇总方式：0-按收据费目,1-按收入项目,2-按单据
    '出参:
    '   rsBillData_out-单据数据，字段：
    '       序号,编码,类别,名称,计算单位,商品名,规格,费用类型,付数,数次,单价,
    '       应收金额,实收金额,执行部门,附加标志,医嘱序号,跟踪在用,门诊包装,门诊单位,住院包装,住院单位,
    '       NO,费用ID,从属父号,费别,收费细目ID,项目编码,医嘱内容,开单科室,执行部门ID,
    '       操作员姓名,登记时间,挂号单,门诊标志,商品条码,内部条码,
    '       结帐ID,记录标志,原始数量,准退数量,剩余数量
    '   rsIncome_out-收入汇总集，字段：
    '       名称,应收金,实收金额
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlReadOutPatientDelBillData = mobjExpenceSvr.zlReadOutPatientDelBillData( _
        cllInPara, strFrmCaption, rsBillData_out, rsIncome_out)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlDelChargeBillCheck(cllInPara As Collection, ByRef cllBillLists As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费前调用 Zl_门诊收费记录_销帐_Check 检查数据合法性
    '入参:
    '   cllInPara 入参，成员[单据信息,退费时间,结帐ID,退费摘要,是否结算作废]
    '       单据信息:格式：A001:序号1_数量,序号2_数量,...|A002:序号1_数量,序号2_数量,...|...
    '出参:
    '   cllBillLists:删除执行入参数据
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBillList As Collection
    Dim cllBillDetails As Collection, cllBillDetail As Collection
    Dim varNos As Variant, i As Integer
    Dim varNo As Variant, j As Integer
    Dim strNo As String, strNums As String, varNums As Variant
    
    '入参:
    '     cllBillList(Collect):单据信息，成员：
    '        (单据号,[cllBillDetails(collect)],[cllAdviceUpdateDatas(collect)],
    '         记帐才有(操作状态(N),操作员编号,操作员姓名,登记时间,已结禁止销帐,禁止部分销帐)
    '         退费才有(操作员编号,操作员姓名,登记时间,结帐ID,摘要,结算作废))=cllBillLists(_单据号)
    '       |-cllBillDetails(Collect):单据信息集=cllBillList(_cllBillDetails)
    '         |-cllBillDetail(Collect):每行明细数据集，成员：
    '           (序号,销帐数量,配药IDs(N))=cllBillDetails(_序号)
    '       |-cllAdviceUpdateDatas(collect):医嘱更新数据，仅执行检查后存在=cllBillLists(_cllAdviceUpdateDatas)
    '         |-cllAdviceUpdateData(collect)每行明细数据集，成员：
    '           (医嘱ID,发送号(N),计费状态,删除附费(N))=cllAdviceUpdateDatas(i)
    '    以上元素中，带(N)的，表示可选节点。
    '   intBillType-单据类型(1-收费单;2-记帐单;3-自动记帐单)
    On Error GoTo ErrHandler
    
    Set cllBillLists = New Collection
    
    varNos = Split(cllInPara("单据信息"), "|")
    For i = 0 To UBound(varNos)
        varNo = Split(varNos(i) & ":", ":")
        strNo = varNo(0): strNums = varNo(1)
        
        varNums = Split(strNums, ",")
        Set cllBillDetails = New Collection
        For j = 0 To UBound(varNums)
            Set cllBillDetail = New Collection
            cllBillDetail.Add Split(varNums(j), "_")(0), "序号"
            cllBillDetail.Add Split(varNums(j), "_")(1), "销帐数量"
            
            cllBillDetails.Add cllBillDetail, "_" & varNums(j)
        Next
    
        Set cllBillList = New Collection
        cllBillList.Add strNo, "单据号"
        cllBillList.Add UserInfo.编号, "操作员编号"
        cllBillList.Add UserInfo.姓名, "操作员姓名"
        cllBillList.Add cllInPara("退费时间"), "登记时间"
        cllBillList.Add cllInPara("结帐ID"), "结帐ID"
        cllBillList.Add cllInPara("退费摘要"), "摘要"
        cllBillList.Add cllInPara("是否结算作废"), "结算作废"
        cllBillList.Add cllBillDetails, "_cllBillDetails"
        cllBillLists.Add cllBillList, "_" & strNo
    Next
    
    If mobjExpenceSvr.zlExcute_DelRecipeBill_Check(cllBillLists, 1, mlngModule, False, True) = False Then Exit Function
    
    zlDelChargeBillCheck = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcute_SaveDrugDelRecipeBill(ByVal strSendDrugFeeIDs As String, _
    ByVal strJosnDelDrugRcp As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废药品处方
    '入参:strSendDrugFeeIds-所涉及的发药费用IDs
    '     strJosnNewDrugRcp-作废处方的Json串
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlExcute_SaveDrugDelRecipeBill = _
        mobjExpenceSvr.zlExcute_SaveDrugDelRecipeBill(strSendDrugFeeIDs, strJosnDelDrugRcp, True)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlExcute_SaveStuffDelBill(ByVal strSendStuffFeeIDs As String, _
    ByVal strJosnDelStuffRcp As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:作废卫生材料处方
    '入参:strSendDrugFeeIds-所涉及的发药费用IDs
    '     strJosnDelStuffRcp-作废处方的Json串
    '出参:
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlExcute_SaveStuffDelBill = _
        mobjExpenceSvr.zlExcute_SaveStuffDelBill(strSendStuffFeeIDs, strJosnDelStuffRcp, True)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlPriceChargeCheck(ByVal strNo As String, ByVal lng病人ID As Long, _
    ByVal strFrmCaption As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:划价收费 检查数据合法性
    '入参:
    '   strNo:退费单据号
    '出参:
    '返回:检查通过返回True,否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlPriceChargeCheck = mobjExpenceSvr.zlPriceChargeCheck(strNo, lng病人ID, strFrmCaption)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlDelPriceBillCheck(ByVal strNos As String, Optional ByRef cllBillLists As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除划价单前调用 Zl_门诊划价记录_Delete_Check 检查数据合法性
    '入参:
    '   strNos:划价单据信息,格式：A001:序号1|序号2|...,A002:序号1|序号2|...,...
    '出参:
    '   cllBillLists:删除执行入参数据
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    zlDelPriceBillCheck = mobjExpenceSvr.zlDelPriceBillCheck(strNos, cllBillLists)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceBillingState(ByVal bytBillType As Byte, ByVal strNos As String, _
    ByVal byt操作类型 As Byte, ByRef cllAdviceData As Collection, _
    Optional ByVal cllDeleteLists As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱计费状态
    '入参:
    '   bytBillType 单据类型: 1-收费单;2-记帐单;3-自动记帐单
    '   strNos -单据信息, 格式：NO1:序号1,序号2,...|NO1:序号1,序号2,...|...
    '   byt操作类型 操作类型：0-删除划价单;1-收费或记帐;2-退费或销帐
    '   cllDeleteLists 集合,退费数据
    '       |-cllDeleteList 集合,销账数据,操作类型为0或2时必须传=cllDeleteLists("_NO")
    '           |-元素集合：序号,销帐数量(为零时，按序号直接销帐)
    '出参:
    '   cllAdviceData 集合,含计费状态的医嘱数据
    '       |元素集合：医嘱ID,费用单号,计费状态,删除附费
    '返回:获取成功返回True,获取失败返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    ZlGetAdviceBillingState = mobjExpenceSvr.ZlGetAdviceBillingState( _
        bytBillType, strNos, True, byt操作类型, cllAdviceData, mlngModule, cllDeleteLists)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlUpdateAdviceChargeTag(ByVal bytBillType As Byte, _
    ByVal cllAdviceDatas As Collection, Optional ByVal cllFeeDetails As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更改病人医嘱数据的计费状态及删除附费数据
    '入参:
    '   bytBillType 单据类型：1-收费单;2-记帐单;3-自动记帐单
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集
    '        |-成员(医嘱ID,发送号(N),费用单号(N),计费状态,删除附费(N))
    '   cllFeeDetails(collect)费用明细集，仅门诊退费时传入，用于标记 医嘱执行计价.执行状态 为 2-已退费
    '     |-cllFeeDetail(collect)每行明细数据
    '        |-成员(医嘱ID,费用单号,收费明细ID)
    '出参:
    '返回:执行成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    If cllAdviceDatas.Count = 0 Then ZlUpdateAdviceChargeTag = True: Exit Function
    ZlUpdateAdviceChargeTag = mobjPubService.ZlCisSvr_UpdAdviceChargeTag(bytBillType, cllAdviceDatas, cllFeeDetails)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetPatientType(ByVal lng病人ID As Long, Optional ByRef objPati As clsPatientInfo) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取病人类型
    '入参:
    '出参:
    '返回:返回病人类型
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    
    On Error GoTo ErrHandler
    If lng病人ID = 0 Then Exit Function
    If gobjSquare.objOneCardComLib.zlGetPatiInforFromPatiID(lng病人ID, objPati) = False Then Exit Function
    If objPati Is Nothing Then Exit Function
    ZLGetPatientType = objPati.病人类型
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceInfoFromID(ByVal str医嘱IDs As String, ByRef rsAdviceData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID获取医嘱信息
    '入参:
    '   str医嘱IDs 医嘱ID,多个英文逗号分隔
    '出参:
    '   rsAdviceData 医嘱信息：医嘱id,医嘱内容,挂号单号,诊疗项目ID,诊疗类别,操作类型,皮试结果
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As New Collection
    
    On Error GoTo ErrHandler
    cllFilter.Add Array("医嘱IDS", str医嘱IDs)
    ZlGetAdviceInfoFromID = mobjPubService.ZlCissvr_GetAdviceInfo(cllFilter, rsAdviceData)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceInfoFromRegNo(ByVal str挂号单号 As String, ByRef rsAdviceData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单获取医嘱信息
    '入参:
    '   str医嘱IDs 医嘱ID,多个英文逗号分隔
    '出参:
    '   rsAdviceData 医嘱信息：医嘱id,医嘱内容,挂号单号,诊疗项目ID,诊疗类别,操作类型,皮试结果
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    
    On Error GoTo ErrHandler
    Set cllFilter = New Collection
    cllFilter.Add Array("挂号单号", str挂号单号)
    ZlGetAdviceInfoFromRegNo = mobjPubService.ZlCissvr_GetAdviceInfo(cllFilter, rsAdviceData)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLGetAdviceSendInfo(ByVal strNos As String, ByRef rsAdviceSendData As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱发送信息
    '入参:
    '   strNos 单据信息,格式：医嘱ID:NO:记录性质,医嘱ID:NO:记录性质,...
    '出参:
    '   rsAdviceSendData 医嘱发送信息：医嘱ID,发送号,挂号单号,病人ID,主页ID,病人科室ID,开嘱科室ID,
    '                                 病人来源,诊疗类别,计价性质,相关id,门诊记帐,医嘱内容,样本条码,
    '                                 No,记录性质,首次时间,末次时间,执行状态,发送时间
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If mobjPubService.ZlCissvr_GetAdviceSendInfo(2, strNos, rsAdviceSendData, False, mlngModule) = False Then Exit Function
    ZLGetAdviceSendInfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGet诊疗收费对照(ByVal str医嘱IDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID,获取相应的诊疗收费对照项目
    '入参:
    '   str医嘱IDs-多个时，用逗号分隔
    '返回:诊疗收费对照的数据集
    '问题:33634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsAdviceData  As ADODB.Recordset
    Dim str项目IDs As String, strSql As String
     
    On Error GoTo ErrHandler
    If str医嘱IDs = "" Then Exit Function
    
    If ZlGetAdviceInfoFromID(str医嘱IDs, rsAdviceData) = False Then Exit Function
    If rsAdviceData Is Nothing Then Exit Function
    If rsAdviceData.RecordCount = 0 Then Exit Function
    
    With rsAdviceData
        Do While Not .EOF
            str项目IDs = str项目IDs & "," & Nvl(!医嘱ID) & ":" & Nvl(!诊疗项目ID)
            .MoveNext
        Loop
        str项目IDs = Mid(str项目IDs, 2)
    End With
    
    strSql = _
        " Select /*+Cardinality(j,10)*/ j.C1 As 医嘱序号," & _
        "        b.诊疗项目id, b.收费项目id As 收费细目id, b.收费数量, b.固有对照, b.从属项目" & _
        " From 诊疗收费关系 B, Table(f_Num2list2([1],',',':')) J" & _
        " Where b.诊疗项目id = j.C2 And Nvl(b.固有对照, 0) = 1"
    Set zlGet诊疗收费对照 = zlDatabase.OpenSQLRecord(strSql, "获取医嘱所对应的收费关系", str项目IDs)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlCheckIsExistAdvice(ByVal str挂号单号 As String, ByRef blnExist As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号单号检查是否存在对应医嘱数据
    '入参:
    '   str挂号单号 挂号单号,多个英文逗号分隔
    '出参:
    '   blnExist 存在医嘱信息
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    blnExist = False
    ZlCheckIsExistAdvice = mobjPubService.zlCisSvr_IsExistAdvice(str挂号单号, 0, blnExist, True, , mlngModule)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZLDelCharge_GetAdviceBillingState(ByRef cllAdviceDatas As Collection, _
    ByRef cllFeeDetails As Collection, _
    ByVal blnDataFromDB As Boolean, Optional ByVal lng结算序号 As Long, _
    Optional ByVal strDelNosDetail As String, Optional ByVal blnCancelBill As Boolean) As Boolean
    '获取医嘱执行计价
    '入参:
    '   blnDataFromDB=数据是否来源于数据库，通过 lng结算序号 获取
    '   strDelNosDetail=不来源数据库时传入，格式：A001:序号1_数量,序号2_数量,...|A002:序号1_数量,序号2_数量,...|...
    '   blnCancelBill=是否作废单据，此时 blnDataFromDB 传入 True，原样退；作废异常再作废传入False
    '出参:
    '   cllAdviceDatas(collect)-数据集，格式如下
    '     |-cllAdviceData(collect)每行明细数据集
    '        |-成员(医嘱ID,发送号(N),费用单号(N),计费状态,删除附费(N))
    '   cllFeeDetails(collect)费用明细集
    '     |-cllFeeDetail(collect)每行明细数据
    '        |-成员(医嘱ID,费用单号,记录性质,收费明细ID)
    '返回:成功返回True，失败返回False
    Dim strSql As String, rsData As ADODB.Recordset
    Dim strNos As String, strPreNo As String
    Dim varNos As Variant, var序号 As Variant
    Dim cllDeleteLists As Collection, cllDeleteList As Collection, cllTemp As Collection
    '   cllDeleteLists 集合,退费数据
    '       |-cllDeleteList 集合,销账数据,操作类型为0或2时必须传=cllDeleteLists("_NO")
    '           |-元素集合：序号,销帐数量(为零时，按序号直接销帐)
    Dim i As Long, j As Long, strNo As String
    Dim strDelNos As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    
    On Error GoTo ErrHandler
    Set cllDeleteLists = New Collection
    Set cllFeeDetails = New Collection
    
    If blnDataFromDB Then
        '从数据库中获取
        '本次退的减去重收的就是实际退的
        
        strSql = "Select Distinct NO, a.序号,Nvl(Sum(a.付数 * a.数次), 0) As 数量,a.医嘱序号,a.收费细目ID" & _
                " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & _
                " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1" & strWhere & _
                " Group By a.NO, a.序号,a.医嘱序号,a.收费细目ID" & _
                " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询退费数据", lng结算序号)
        
        rsData.Filter = "医嘱序号<>null"
        Do While Not rsData.EOF
            Set cllTemp = New Collection
            cllTemp.Add Nvl(rsData!医嘱序号), "医嘱ID"
            cllTemp.Add Nvl(rsData!NO), "费用单号"
            cllTemp.Add 1, "记录性质"
            cllTemp.Add Nvl(rsData!收费细目ID), "收费明细ID"
            cllFeeDetails.Add cllTemp
            
            rsData.MoveNext
        Loop
        
    Else
        Set rsData = New ADODB.Recordset
        With rsData.Fields
            .Append "No", adVarChar, 20, adFldIsNullable
            .Append "序号", adBigInt, , adFldIsNullable
            .Append "数量", adDouble, , adFldIsNullable
        End With
        rsData.CursorLocation = adUseClient
        rsData.LockType = adLockOptimistic
        rsData.CursorType = adOpenStatic
        rsData.Open
        
        strDelNos = ""
        'A001:序号1_数量,序号2_数量,...|A002:序号1_数量,序号2_数量,...|...
        varNos = Split(strDelNosDetail, "|")
        For i = 0 To UBound(varNos)
            strNo = Split(varNos(i), ":")(0): var序号 = Split(Split(varNos(i), ":")(1), ",")
            For j = 0 To UBound(var序号)
                rsData.AddNew
                rsData!NO = strNo
                rsData!序号 = Split(var序号(j), "_")(0)
                rsData!数量 = Split(var序号(j), "_")(1)
                rsData.Update
                
                strDelNos = strDelNos & "," & strNo & ":" & Nvl(rsData!序号)
            Next
        Next
        strDelNos = Mid(strDelNos, 2)
        
        strSql = _
            " Select Distinct a.No, a.医嘱序号, a.收费细目id" & _
            " From 门诊费用记录 A, Table(f_Str2list2([1],',',':')) B" & _
            " Where a.No=b.C1 And a.序号=b.C2 And 记录状态 In(1,3) And a.医嘱序号 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询数据", strDelNos)
        Do While Not rsTemp.EOF
            Set cllTemp = New Collection
            cllTemp.Add Nvl(rsTemp!医嘱序号), "医嘱ID"
            cllTemp.Add Nvl(rsTemp!NO), "费用单号"
            cllTemp.Add 1, "记录性质"
            cllTemp.Add Nvl(rsTemp!收费细目ID), "收费明细ID"
            cllFeeDetails.Add cllTemp
            
            rsTemp.MoveNext
        Loop
    End If
    
    rsData.Filter = ""
    rsData.Sort = "NO,序号"
    strNos = ""
    Do While Not rsData.EOF
        If strPreNo <> Nvl(rsData!NO) Then
            strNos = strNos & "|" & Nvl(rsData!NO) & ":" & Nvl(rsData!序号)
            
            If strPreNo <> "" Then cllDeleteLists.Add cllDeleteList, "_" & strPreNo
            
            strPreNo = Nvl(rsData!NO)
            Set cllDeleteList = New Collection
        Else
            strNos = strNos & "," & Nvl(rsData!序号)
        End If
        
        Set cllTemp = New Collection
        cllTemp.Add Nvl(rsData!序号), "序号"
        cllTemp.Add Nvl(rsData!数量), "销帐数量"
        cllDeleteList.Add cllTemp
        
        rsData.MoveNext
    Loop
    cllDeleteLists.Add cllDeleteList, "_" & strPreNo
    
    strNos = Mid(strNos, 2)
    
    '计算医嘱计费状态
    If blnDataFromDB And Not blnCancelBill Then '作废时需要传入退费明细
        If ZlGetAdviceBillingState(1, strNos, 2, cllAdviceDatas) = False Then Exit Function
    Else
        If ZlGetAdviceBillingState(1, strNos, 2, cllAdviceDatas, cllDeleteLists) = False Then Exit Function
    End If
    ZLDelCharge_GetAdviceBillingState = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetAdviceDiaginfo(ByVal strFeeInfo As String, _
    ByRef cllDiagnose As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱诊断信息
    '入参:
    '   strFeeInfo=费用医嘱信息,格式:NO:医嘱ID,NO:医嘱ID,...
    '出参:
    '   cllDiagnose(collect)=诊断数据集，成员:Array("诊断","NO1,NO2,...")
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varFeeInfo As Variant, varNo As Variant, i As Long
    Dim strAdviceIDs As String, rsDiagnose As ADODB.Recordset
    Dim strNos As String, strKey As String, strDiag As String
    
    On Error GoTo ErrHandler
    Set cllDiagnose = New Collection
    
    varFeeInfo = Split(strFeeInfo, ",")
    For i = 0 To UBound(varFeeInfo)
        varNo = Split(varFeeInfo(i), ":")
        If Val(varNo(1)) <> 0 And InStrEx(strAdviceIDs, varNo(1)) = False Then
            strAdviceIDs = strAdviceIDs & "," & varNo(1)
        End If
    Next
    
    If strAdviceIDs <> "" Then
        strAdviceIDs = Mid(strAdviceIDs, 2)
        If mobjPubService.ZlCissvr_GetDiaginfoFromAdviceID(strAdviceIDs, rsDiagnose, mlngModule) = False Then Exit Function
    End If
    
    For i = 0 To UBound(varFeeInfo)
        varNo = Split(varFeeInfo(i), ":")
        
        If rsDiagnose Is Nothing Then
            strDiag = ""
        Else
            rsDiagnose.Filter = "医嘱ID=" & Val(varNo(1))
            If rsDiagnose.EOF Then
                strDiag = ""
            Else
                strDiag = Nvl(rsDiagnose!诊断描述)
            End If
        End If
        
        strNos = ""
        strKey = "_" & strDiag
        If CollectionExitsValue(cllDiagnose, strKey) Then
            strNos = "," & cllDiagnose(strKey)(1)
            cllDiagnose.Remove strKey
        End If
        
        If InStrEx(strNos, varNo(0)) = False Then
            strNos = strNos & "," & varNo(0)
        End If
        cllDiagnose.Add Array(strDiag, Mid(strNos, 2)), strKey
    Next
    ZlGetAdviceDiaginfo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetPatiPageInfByRange(ByVal cllFilter As Collection, _
    ByRef rsPatiInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '   cllFilter 查询条件集:成员(Array(Key,Value),Array(Key,Value),,...)
    '       Key:病区IDS,科室IDS,病人IDS,主页IDS,入院开始时间,入院结束时间,出院开始时间,出院结束时间,费别,住院状态,病人性质
    '出参:
    '   rsPatiPageInfo 病人病案主页信息：病人ID,主页ID,姓名,性别,年龄,住院号,床号,险类,费别,病人类型,
    '                                   入院时间,出院时间,住院状态,病人性质,当前病区ID,当前病区名称,当前科室ID,当前科室名称,
    '                                   医疗付款方式名称,医疗付款方式编码,住院医师,病人备注,编目日期
    '       住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
    '       病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubService.ZlCissvr_GetPatiPageInfByRange(cllFilter, rsPatiInfo, mlngModule) = False Then Exit Function
    ZlGetPatiPageInfByRange = True
End Function

Public Function zlGetPatiPageInfo(objPati As clsPatientInfo, Optional ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从病案主页中获取病人信息
    '入参:
    '   objPati-已有病人信息
    '   lng主页ID-主页ID，为0时，取最后一次住院的
    '出参:
    '   objPati-返回病人信息对象
    '返回:成功返回True，否则返回False
    '说明:如果传入 objPati 不为Nothing，则进行信息合并
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lng主页ID = 0 Then lng主页ID = objPati.主页ID '病人信息中的主页ID为最后一次主页ID
    If mobjPubService.ZlCissvr_GetPatiPageInfo(1, objPati.病人ID & ":" & lng主页ID, rsTemp) = False Then Exit Function
    If rsTemp Is Nothing Then zlGetPatiPageInfo = True: Exit Function
    If rsTemp.EOF Then zlGetPatiPageInfo = True: Exit Function
    
    If objPati Is Nothing Then Set objPati = New clsPatientInfo
    With objPati
        .主页ID = Nvl(rsTemp!主页ID)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .年龄 = Nvl(rsTemp!年龄)
        .费别 = Nvl(rsTemp!费别)
        .医疗付款方式 = Nvl(rsTemp!医疗付款方式名称)
        .医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
        .险类 = Val(Nvl(rsTemp!险类))
        .险类名称 = GetInsureName(Val(Nvl(rsTemp!险类)))
        .病人类型 = Nvl(rsTemp!病人类型)
        .当前病区ID = Val(Nvl(rsTemp!当前病区ID))
        .当前科室ID = Val(Nvl(rsTemp!当前科室ID))
        .床号 = Nvl(rsTemp!当前床号)
        .住院号 = Nvl(rsTemp!住院号)
        .病人性质 = Val(Nvl(rsTemp!病人性质))
        .入院日期 = Nvl(rsTemp!入院时间)
        .出院日期 = Nvl(rsTemp!出院时间)
    End With
    zlGetPatiPageInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiInfo(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long) As clsPatientInfo
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息，先从病人信息中获取，再从病案主页中获取进行合并
    '入参:
    '   objPati-已有病人信息
    '   lng主页ID-主页ID，为0时，取最后一次住院的
    '出参:
    '   objPati-返回病人信息对象
    '返回:成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '读取病人信息
    If gobjSquare.objOneCardComLib.zlGetPatiInforFromPatiID(lng病人ID, objPati) = False Then Exit Function
    If objPati Is Nothing Then Exit Function
    
    '2.读取病案主页
    If lng主页ID = 0 Then lng主页ID = objPati.主页ID
    If lng主页ID > 0 Then
        If zlGetPatiPageInfo(objPati, lng主页ID) = False Then Exit Function
    End If
    
    Set zlGetPatiInfo = objPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZlGetAllInDeptInfor() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有存在在院病人的科室信息
    '入参:
    '出参:
    '返回:rsDept 科室信息,字段：ID,编码,名称,简码
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsDept As ADODB.Recordset
    
    If mobjPubService.ZlCissvr_GetInDeptInfor(0, 1, gstrNodeNo, True, rsDept) = False Then Exit Function
    Set ZlGetAllInDeptInfor = rsDept
End Function

