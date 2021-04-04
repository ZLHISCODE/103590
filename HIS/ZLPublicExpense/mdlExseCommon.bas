Attribute VB_Name = "mdlExseCommon"
Option Explicit
'*********************************************************************************************************************************************
'公共费用相关处理功能
'接口说明:
'   1.zlGetSpecialItemFee-产生工本费、就诊卡等填写住院费用记录时的必须信息(收费类别,收费细目ID,计算单位,收入项目ID,收入项目,收据费目,原价,现价,是否变价,科室标志)
'   2.GetAllAdviceIDsFromDiagnoID：根据诊断ID,获取涉及所有组医嘱ID
'出参:
'返回:成功返回true,否则返回False
'编制:刘兴洪
'日期:2019-11-23 17:33:30
'*********************************************************************************************************************************************
Public grs结算方式 As ADODB.Recordset
Public gobjService As clsService
Public gobjExpenceSvr As clsExpenceSvr
Public gobjBillPrint As Object

Public Function zlGetExpenceSvrObject(ByRef objExpenceSvr As clsExpenceSvr) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取服务对象
    '出参:objExpenceSvr-返回服务对象
    '返回:获取返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gobjExpenceSvr Is Nothing Then Set objExpenceSvr = gobjExpenceSvr: zlGetExpenceSvrObject = True: Exit Function
    Set objExpenceSvr = New clsExpenceSvr
    Call objExpenceSvr.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser)
    Set gobjExpenceSvr = objExpenceSvr
    zlGetExpenceSvrObject = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Set objExpenceSvr = Nothing
End Function
Public Function zlGetServiceObject(ByRef objService As clsService) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取服务对象
    '出参:objService-返回服务对象
    '返回:获取返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gobjService Is Nothing Then Set objService = gobjService: zlGetServiceObject = True: Exit Function
    Set objService = New clsService
    Call objService.zlInitCommon(glngSys, glngModul, gcnOracle, gstrDBUser)
    Set gobjService = objService
    zlGetServiceObject = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Set objService = Nothing
End Function



Public Function zlGetSpecialItemFee(strClass As String, Optional ByVal strPriceGrade As String, Optional ByVal lng收费细目id As Long) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:产生工本费、就诊卡等填写住院费用记录时的必须信息(收费类别,收费细目ID,计算单位,收入项目ID,收入项目,收据费目,原价,现价,是否变价,科室标志)
    '入参:
    '   strClass=工本费、就诊卡、病历费
    '   strPriceGrade 普通价格等级
    '返回:指定的数据类别的费用集
    '编制:刘兴洪
    '日期:2011-07-07 02:17:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim strWherePriceGrade As String
    Dim rsTmp As New ADODB.Recordset
  
    
    On Error GoTo errH
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [2]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    
    If lng收费细目id = 0 Then
        strSql = _
            "Select a.类别 As 收费类别, a.Id As 收费细目id, a.计算单位, c.Id As 收入项目id, Nvl(a.屏蔽费别, 0) As 屏蔽费别, c.名称 As 收入项目, c.收据费目, b.原价, b.现价," & vbNewLine & _
            "       Nvl(b.缺省价格, 0) 缺省价格, Nvl(a.是否变价, 0) As 是否变价, Nvl(a.执行科室, 0) As 科室标志" & vbNewLine & _
            "From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费特定项目 D" & vbNewLine & _
            "Where b.收费细目id = a.Id And b.收入项目id = c.Id And d.收费细目id = a.Id And d.特定项目 = [1]" & vbNewLine & _
            "      And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    Else
        strSql = _
            "Select a.类别 As 收费类别, a.Id As 收费细目id, a.计算单位, c.Id As 收入项目id, Nvl(a.屏蔽费别, 0) As 屏蔽费别, c.名称 As 收入项目, c.收据费目, b.原价, b.现价," & vbNewLine & _
            "       Nvl(b.缺省价格, 0) 缺省价格, Nvl(a.是否变价, 0) As 是否变价, Nvl(a.执行科室, 0) As 科室标志" & vbNewLine & _
            "From 收费项目目录 A, 收费价目 B, 收入项目 C " & vbNewLine & _
            "Where b.收费细目id = a.Id And b.收入项目id = c.Id And A.ID = [3]" & vbNewLine & _
            "      And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取特定项目的费用集", strClass, strPriceGrade, lng收费细目id)
    If Not rsTmp.EOF Then Set zlGetSpecialItemFee = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function zlGetUnitID(bytFlag As Byte, lngID As Long) As Long
'功能：返回收费特定项目的执行科室
'参数：bytFlag=执行科室标志,lngID=收费细目ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '无明确科室
            zlGetUnitID = UserInfo.部门ID '取操作员所在科室
        Case 4 '指定科室
            strSql = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                zlGetUnitID = rsTmp!执行科室ID '默认取第一个(如有多个)
            Else
                zlGetUnitID = UserInfo.部门ID '如没有指定，则取操作员所在科室
            End If
        Case 1, 2, 3 '病人科室,操作员科室
            zlGetUnitID = UserInfo.部门ID '都取操作员科室
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetCardFeeExcuteDeptID(ByVal lng收费细目id As Long, ByVal byt科室标志 As Byte, Optional ByVal lng病人科室ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据科室标志，获取对应的执行部门ID
    '入参:lng收费细目ID-收费细目ID
    '     byt科室标志-科室标志(0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室)
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-01-18 11:31:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngUnitID  As Long
    
    On Error GoTo errHandle
    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
     Select Case byt科室标志
         Case 4 '指定科室
             lngUnitID = zlGetUnitID(byt科室标志, lng收费细目id)
         Case 1, 2 '病人科室
             If lng病人科室ID <> 0 Then
                 lngUnitID = lng病人科室ID
             Else
                 lngUnitID = UserInfo.部门ID
             End If
         Case 0, 3, 5, 6
             lngUnitID = UserInfo.部门ID
     End Select
     zlGetCardFeeExcuteDeptID = lngUnitID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFromBalanceName(ByVal str结算方式 As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算方式获取卡对象
    '入参:str结算方式-结算方式名称
    '出参:
    '返回:返回卡对象
    '编制:刘兴洪
    '日期:2018-03-30 15:39:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHandle
    Set objCard = New Card
    Set rsTemp = zlGet结算方式
    With objCard
        .结算方式 = str结算方式
        .名称 = str结算方式
        rsTemp.Filter = "名称='" & str结算方式 & "'"
        If Not rsTemp.EOF Then
            .结算性质 = Val(Nvl(rsTemp!性质))
            .缺省标志 = Val(Nvl(rsTemp!场合缺省)) = 1
        End If
    End With
    Set zlGetCardFromBalanceName = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet结算方式() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算方式
    '返回:返回结算方式信息集
    '编制:刘兴洪
    '日期:2018-03-29 17:35:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    
    If Not grs结算方式 Is Nothing Then
        If grs结算方式.State = 1 Then
            grs结算方式.Filter = 0
            Set zlGet结算方式 = grs结算方式: Exit Function
        End If
    End If
    On Error GoTo errHandle
    strSql = "" & _
    "   Select a.编码,a.名称, a.性质,b.应用场合,nvl(a.应付款,0) as 应付款,nvl(a.应收款,0) as 应收款,nvl(a.缺省标志,0) as 缺省,nvl(b.缺省标志,0) as  场合缺省" & vbNewLine & _
    "   From 结算方式 a, 结算方式应用 b" & vbNewLine & _
    "   Where a.名称 = b.结算方式(+)    "
        
    Set grs结算方式 = zlDatabase.OpenSQLRecord(strSql, "获取结算方式")
    Set zlGet结算方式 = grs结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get结算方式(str场合 As String, Optional str性质 As String) As ADODB.Recordset
    Dim strSql As String, strIF As String
    
    On Error GoTo errH
    
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strIF = "And Instr(','||[2]||',',','||B.性质||',')>0 "
        Else
            strIF = "And B.性质 = [2]"
        End If
    End If
    strSql = _
        " Select B.编码,B.名称,Nvl(Nvl(A.缺省标志,B.缺省标志),0) as 缺省,Nvl(B.性质,1) as 性质,Nvl(B.应付款,0) as 应付款" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 " & _
        " And  B.性质<>7    " & strIF
    If InStr(1, str性质, ",9") > 0 Then
        strSql = strSql & " Union " & _
                 " Select 编码,名称,Nvl(缺省标志,0) As 缺省,Nvl(性质,1) as 性质,Nvl(应付款,0) as 应付款 " & _
                 " From 结算方式 " & _
                 " Where 性质=9 " & _
                 " Order by 性质,编码"
    Else
        strSql = strSql & " Order by 性质,lpad(编码,3,' ')"
    End If
    Set Get结算方式 = zlDatabase.OpenSQLRecord(strSql, App.ProductName, str场合, str性质)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAllAdviceIDsFromDiagnoID(ByVal str诊断IDs As String, ByRef str医嘱Ids_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据诊断ID,获取涉及所有组医嘱ID
    '入参:str诊断IDs-诊断ID,多个用逗号
    '出参:str医嘱Ids_Out-医嘱ID,多个用逗号
    '返回:获取诊断返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-01-23 18:11:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As clsService
    
    If zlGetServiceObject(objService) = False Then Exit Function
    GetAllAdviceIDsFromDiagnoID = objService.zlCisSvr_GetAdviceidsFromDiag(str诊断IDs, str医嘱Ids_Out)
End Function

 
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False, _
    Optional strLogFunName As String, Optional strLogName As String)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '     strLogName-日志类别名
    '     strLogFunName-日志功能名
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSql As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSql = cllProcs(i)
        If strLogFunName <> "" Then Call WritLog(strLogName, strLogFunName, "zlExecuteProcedureArrAy", strSql)
        Call zlDatabase.ExecuteProcedure(strSql, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Sub zlAddArray(ByRef clldata As Collection, ByVal strSql As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = clldata.Count + 1
    clldata.Add strSql, "K" & i
End Sub





Public Sub zlBillPrint_Initialize(Optional ByVal lngModul As Long)
    '功能:调用发票打印插件之初始化接口
    '入参:
    '   lngModul=模块号，门诊收费=1121、保险补充结算=1124、结帐=1137
    '问题号:140948
    Dim blnInitSuccess As Boolean
    
    '创建第三方票据打印部件
    On Error Resume Next
    If gobjBillPrint Is Nothing Then
        Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    End If
    If gobjBillPrint Is Nothing Then Exit Sub
    If lngModul = 0 Then Exit Sub
    
    blnInitSuccess = gobjBillPrint.zlInitialize(gcnOracle, glngSys, lngModul, UserInfo.编号, UserInfo.姓名)
    If blnInitSuccess = False Then
        '初始部件不成功,则作为不存在处理
        Set gobjBillPrint = Nothing: Exit Sub
    End If
End Sub

Public Function zlBillPrint_EraseBill(ByVal strNOs As String, ByVal lngBalanceID As Long) As Boolean
    '功能:调用发票打印插件之作废发票
    '入参:
    '   strNOs=门诊收费、保险补充结算：以逗号分隔的带引号的多个单据号:'F0000001','F0000002',...
    '   lngBalanceId=结帐：结帐单ID
    '问题号:140948
    
    On Error GoTo ErrHandler
    If gobjBillPrint Is Nothing Then zlBillPrint_EraseBill = True: Exit Function
    If gobjBillPrint.zlEraseBill(strNOs, lngBalanceID) = False Then Exit Function
    
    zlBillPrint_EraseBill = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlBillPrint_Terminate() As Boolean
    '功能:调用发票打印插件之终止接口
    '问题号:140948
    
    On Error GoTo ErrHandler
    If gobjBillPrint Is Nothing Then zlBillPrint_Terminate = True: Exit Function
    If gobjBillPrint.zlTerminate() = False Then Exit Function
    Set gobjBillPrint = Nothing
    
    zlBillPrint_Terminate = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlGetMedicalGroupID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng开单科室ID As Long, ByVal str开单人 As String, ByVal dt发生时间 As Date) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件获取对应的医疗小组ID
    '入参:
    '   dt发生时间=费用发生时间
    '出参:
    '返回:获取到的医疗小组ID
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    Dim lng组id As Long
    
    On Error GoTo ErrHandler
    If zlGetServiceObject(objService) = False Then Exit Function
    If objService.ZlCissvr_GetMedicalGroupID(lng病人ID, lng主页ID, _
        lng开单科室ID, str开单人, dt发生时间, lng组id) = False Then Exit Function
        
    ZlGetMedicalGroupID = lng组id
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

