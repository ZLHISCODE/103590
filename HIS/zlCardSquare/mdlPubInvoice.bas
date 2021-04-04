Attribute VB_Name = "mdlPubInvoice"
Option Explicit
Public Type Ty_FactProperty
    lngShareUseID As Long   '共享领用批次ID
    strUseType As String ' 使用类别
    intInvoiceFormat As Integer '打印的发票格式,发票格式序号
    intInvoicePrint As Integer     '打印方式:0-不打印;1-自动打印;2-提示打印
    bln严格控制 As Boolean
End Type

Public Function GetShareInvoiceGroupID(ByVal bytKind As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定票种的共用票据批次
    '编制:刘兴洪
    '日期:2011-04-29 10:24:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, cllBillInfor As Collection
    On Error GoTo errH
    If zl_ExseSvr_GetReceiveInvoice(bytKind, "", cllBillInfor, False, "", , , , , , 1, True, rsTemp) = False Then Exit Function
    Set GetShareInvoiceGroupID = rsTemp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'功能：获取张数够用并且指定票据在其可用范围内的领用ID
'参数：bytKind      =   票种
'      intNum       =   要打印的票据张数
'      lngLastUseID =   上次使用的领用ID
'      lngShareUseID=   本地参数指定的共用ID
'      strBill      =   当前票据号，用于检查领用批次的票据范围
'      strUseType-使用类别
'返回：
'      >0   =   成功，可用的领用ID
'      =0   =   失败
'      -1   =   没有自用(用完或不够，或未领用),未设置共用
'      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
'      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
'      -4   =   指定批次的票据不够用
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    Dim cllBillInfo As Collection, cllTemp As Collection
 
    On Error GoTo errH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
    
        If zl_ExseSvr_GetReceiveInvoice(bytKind, lngLastUseID, cllBillInfo, False, strUseType, "", intNum, True, , glngModul) = False Then Exit Function
        '        item_list C
        '            recv_id N   1   领用ID
        '            use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
        '            use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
        '            prefix_text C   1   前缀文本
        '            start_no    C   1   开始号码
        '            end_no  C   1   终止号码
        '            inv_no_cur  C   1   当前号码
        '            surplus_num C   1   剩余数量
        '            create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
        '            use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
        '            recvtr  C   1   领用人
        If Not cllBillInfo Is Nothing Then
            If cllBillInfo.Count <> 0 Then
                 If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '可能没有当前票据号
                
                Set cllTemp = cllBillInfo(1)
                blnTmp = False
                strPre = nvl(cllTemp("_prefix_text"))
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(nvl(cllTemp("_end_no"))) And Len(strBill) = Len(nvl(cllTemp("_start_no")))) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                GetInvoiceGroupID = -4: Exit Function
            End If
        ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
            GetInvoiceGroupID = -4: Exit Function
        End If
    End If
    
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    If zl_ExseSvr_GetReceiveInvoice(bytKind, 0, cllBillInfo, True, strUseType, UserInfo.姓名, intNum, True, , glngModul) Then
        '        item_list C
        '            recv_id N   1   领用ID
        '            use_mode    N   1   使用方式:1-自用：该票据仅供领用者自己使用；2-共用：该票据由多个人员共同使用
        '            use_type    C   1   票据使用类别:1,4: 票据使用类别.名称;2预见:1-门诊预交;2-住院预交;5:存储的是医疗卡类别.ID
        '            prefix_text C   1   前缀文本
        '            start_no    C   1   开始号码
        '            end_no  C   1   终止号码
        '            inv_no_cur  C   1   当前号码
        '            surplus_num C   1   剩余数量
        '            create_time C   1   登记时间:yyyy-mm-dd hh24:mi:ss
        '            use_time    C   1   使用时间:yyyy-mm-dd hh24:mi:ss
        '            recvtr  C   1   领用人
        For i = 1 To cllBillInfo.Count
            Set cllTemp = cllBillInfo(i)
            
            If strBill = "" Then GetInvoiceGroupID = Val(nvl(cllTemp("_recv_id"))): Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = nvl(cllTemp("_prefix_text"))
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(nvl(cllTemp("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllTemp("_start_no"))))) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = Val(nvl(cllTemp("_recv_id"))): Exit Function
        Next
        lngReturn = IIf(cllBillInfo.Count > 0, -3, -1)
    Else
        lngReturn = -1
    End If

    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        If zl_ExseSvr_GetReceiveInvoice(bytKind, lngLastUseID, cllBillInfo, False, strUseType, "", intNum, True, , glngModul) = False Then
            GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
            Exit Function
        End If
        
        If cllBillInfo.Count = 0 Then
            lngReturn = -2
            GetInvoiceGroupID = lngReturn
            Exit Function
        End If
        
        Set cllTemp = cllBillInfo(1)
        If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '第一次使用时没有当前票据号
        blnTmp = False
        strPre = nvl(cllTemp("_prefix_text"))
        If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
            blnTmp = True
        ElseIf Not (UCase(strBill) >= UCase(nvl(cllTemp("_start_no"))) And UCase(strBill) <= UCase(nvl(cllTemp("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllTemp("_start_no"))))) Then
            blnTmp = True
        End If
        If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
        lngReturn = -3
    End If
    GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
    '参数：bytKind=票种
    '      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
    '      strBill=要检查范围的票据号
    '说明：
    '    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
    '    2.在检查范围时,长度也在检查范围之内。
    '    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
    '返回：
    '      正常：票据领用ID>0
    '      0=失败
    '      -1:没有自用(用完或未领用)、也没有共用(未设置)
    '      -2:设置的共用已用完
    '      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    Dim cllBillInfos As Collection, cllBillItem As Collection
    Dim cllSharess  As Collection, cllSharessItem As Collection
    Dim i As Long
    
    On Error GoTo errH
    
      '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    If Not zl_ExseSvr_GetReceiveInvoice(bytKind, "", cllBillInfos, True, strUseType, UserInfo.姓名, 1, True, , glngModul) Then
        Set cllBillInfos = New Collection
    End If
    If cllBillInfos Is Nothing Then Set cllBillInfos = New Collection


    If lng领用ID = 0 Then
         '程序中第一次检查,且没有设置本地共用
         If cllBillInfos.Count = 0 Then CheckUsedBill = -1: Exit Function  '也没有自用票据
         '有自用票据 , 按优先原则返回
         Set cllBillItem = cllBillInfos(1)
         lngReturn = Val(nvl(cllBillItem("_recv_id")))
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        If Not zl_ExseSvr_GetReceiveInvoice(bytKind, lng领用ID, cllSharess, True, strUseType, "", 0, True, , glngModul) Then
            Set cllSharess = New Collection
        End If
        
        If cllSharess.Count = 0 Then CheckUsedBill = -2: Exit Function
        Set cllSharessItem = cllSharess(1)
        
        
        If Val(nvl(cllSharessItem("_use_mode"))) = 2 Then '共用,要先看有没有自用
            If cllBillInfos.Count <> 0 Then
                '有自用的，优先
                Set cllBillItem = cllBillInfos(1)
                lngReturn = Val(cllBillItem("_recv_id"))
            Else
                '没有自用取共用
                If Val(nvl(cllSharessItem("_surplus_num"))) = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                lngReturn = Val(cllSharessItem("_recv_id"))
                blnTmp = True
            End If
        Else
            '自用票据
            If Val(nvl(cllSharessItem("_surplus_num"))) > 0 Then
                '有剩余
                lngReturn = Val(cllSharessItem("_recv_id"))
            Else
                '其它有剩余的自用
                If cllBillInfos.Count = 0 Then CheckUsedBill = -1: Exit Function      '其它自用也没有剩余
                Set cllBillItem = cllBillInfos(1)
                lngReturn = Val(cllBillItem("_recv_id"))
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(nvl(cllSharessItem("_prefix_text"))))) <> UCase(nvl(cllSharessItem("_prefix_text"))) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(nvl(cllSharessItem("_start_no"))) And UCase(strBill) <= UCase(nvl(cllSharessItem("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllSharessItem("_start_no"))))) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            Set cllBillItem = mdlPubJson.zlGetNodeObjectFromCollect(cllBillInfos, "_" & lngReturn)
           ' rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(nvl(cllBillItem("_prefix_text"))))) <> UCase(nvl(cllBillItem("_prefix_text"))) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(nvl(cllBillItem("_start_no"))) And UCase(strBill) <= UCase(nvl(cllBillItem("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllBillItem("_start_no"))))) Then
                blnTmp = True
            End If
            
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                For i = 1 To cllBillInfos.Count
                    Set cllBillItem = cllBillInfos(i)
                    
                    If nvl(cllBillItem("_recv_id")) <> lngReturn Then
                        blnTmp = False
                        
                        If UCase(Left(strBill, Len(nvl(cllBillItem("_prefix_text"))))) <> UCase(nvl(cllBillItem("_prefix_text"))) Then
                            blnTmp = True
                        ElseIf Not (UCase(strBill) >= UCase(nvl(cllBillItem("_start_no"))) And UCase(strBill) <= UCase(nvl(cllBillItem("_end_no"))) And Len(strBill) = Len(UCase(nvl(cllBillItem("_start_no"))))) Then
                            blnTmp = True
                        End If
                        If Not blnTmp Then lngReturn = Val(cllBillItem("_recv_id")): Exit For
                    End If
                Next
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng领用ID As Long) As String
'功能：根据领用批次ID,获取下一个实际票据号
'说明：1.当取不到范围内的有效票据时,返回空由用户输入
'      2.排开已报损的号码
    Dim strSQL As String, strBill As String
    Dim cllBillInfos As Collection, cllTemp As Collection
    Dim str下一张发票 As String
    
    On Error GoTo errH
    
    If Not zl_ExseSvr_GetReceiveInvoice(0, lng领用ID, cllBillInfos, True, , UserInfo.姓名, 1, True, , glngModul) Then
        Set cllBillInfos = New Collection
    End If
    
    If cllBillInfos.Count = 0 Then Exit Function
    
    Set cllTemp = cllBillInfos(1)
    '取下一个号码
    If nvl(cllTemp("_inv_no_cur")) = "" Then
        strBill = UCase(nvl(cllTemp("_start_no")))
    Else
        strBill = UCase(zlCommFun.IncStr(nvl(cllTemp("_inv_no_cur"))))
    End If
    
    '检查使用明细是否使用该票据
    If Not zl_ExseSvr_GetNextInvoice(lng领用ID, strBill, str下一张发票) Then Exit Function
    If str下一张发票 = "" Then Exit Function
    strBill = str下一张发票
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function zl_GetInvoicePreperty(ByVal lngModule As Long, _
    ByVal int票据 As Integer, Optional str使用类别 As String) As Ty_FactProperty
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票格式
    '入参:int票据:1- 收费收据, 2 - 预交收据, 3 - 结帐收据, 4 - 挂号收据, 5 - 就诊卡
    '返回:发票的相关数据
    '编制:刘兴洪
    '日期:2011-07-19 16:43:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Ty_Fact As Ty_FactProperty, strFactType As String, varData As Variant, varTemp As Variant
    Dim strShareTypeUseID As String, lng共用票据 As Long, lng使用票据 As Long
    Dim strFactTypeFormat As String, strFacePrintMode As String
    Dim intPrintMode As Long, intPrintMode1 As Long, lng领用ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset, strValue As String
    Dim i As Long, lngFormat As Long, lngFormat1 As Long
    
    strFactType = Switch(int票据 = 1, "共用收费票据批次", int票据 = 2, "共用预交票据批次", int票据 = 3, "共用结帐票据批次", int票据 = 4, "共用挂号票据批次", int票据 = 5, "共用医疗卡批次", True, "")
    strFactTypeFormat = Switch(int票据 = 1, "收费发票格式", int票据 = 2, "预交发票格式", int票据 = 3, "结帐发票格式", int票据 = 4, "挂号发票格式", int票据 = 5, "医疗卡发票格式", True, "")
    strFacePrintMode = Switch(int票据 = 1, "收费发票打印方式", int票据 = 2, "预交发票打印方式", int票据 = 3, "病人结帐打印", int票据 = 4, "挂号发票打印方式", int票据 = 5, "医疗卡发票打印方式", True, "")
    
    If strFactType = "" Then Exit Function
    
    
    '票号严格控制
    If int票据 >= 1 And int票据 <= 4 Then
        strValue = zlDatabase.GetPara("票号严格控制", glngSys, , "00000")
        Ty_Fact.bln严格控制 = Mid(strValue, int票据, 1) = "1"
    End If
    
    '用四位数表示, 每一位代表不同的业务类型:
    '第一位:         收费
    '第二位:         预交
    '第三位:         结帐
    '第四位:         挂号
    '每位用1或0表示,1表示严格控制;0-表示非严格控制

    
    '78751:李南春,2014/10/20,增加预交票据打印格式
    Ty_Fact.strUseType = str使用类别
 
    strFactTypeFormat = Trim(zlDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strFactTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = str使用类别 And lngFormat <> 0 Then
            Ty_Fact.intInvoiceFormat = lngFormat: Exit For
        End If
    Next
    If Ty_Fact.intInvoiceFormat = 0 And lngFormat1 <> 0 Then Ty_Fact.intInvoiceFormat = lngFormat
 
    '打印方式(0-不打印;1-自动打印;2-提示打印)
    '问题50656
'    If int票据 = 2 Then
'        '预交暂为自动打印
'        Ty_Fact.intInvoicePrint = 1
'    Else
        '因为Getpara就缓存了的,所以不用先用变量进行记录
        strFacePrintMode = Trim(zlDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
        Ty_Fact.intInvoicePrint = -1
        '格式:使用类别1,打印方式1|使用类别2,打印方式2...
        varData = Split(strFacePrintMode, "|")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,", ",")
            intPrintMode = Val(varTemp(1))
            If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
            If Trim(varTemp(0)) = str使用类别 Then
                Ty_Fact.intInvoicePrint = intPrintMode: Exit For
            End If
        Next
        If Ty_Fact.intInvoicePrint < 0 Then Ty_Fact.intInvoicePrint = intPrintMode1
'    End If
    '共享批次
    
    '格式:领用ID1,使用类别1|....
    strShareTypeUseID = Trim(zlDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng领用ID = Val(varTemp(0))
        If int票据 = 2 Or int票据 = 5 Then
            If Val(varTemp(1)) = 0 Then lng共用票据 = lng领用ID    '共用的.
            If Val(varTemp(1)) = Val(str使用类别) And lng领用ID <> 0 Then
                lng使用票据 = lng领用ID
            End If
        Else
            If Trim(varTemp(1)) = "" Then lng共用票据 = lng领用ID    '共用的.
            If Trim(varTemp(1)) = str使用类别 And lng领用ID <> 0 Then
                lng使用票据 = lng领用ID
            End If
        End If
    Next
    
    On Error GoTo errHandle
    '优先顺序
    '1.先使用
    '2.使用类别不区分的
    '3.具体使用类别的
    Dim cllBillInfo As Collection
    If zl_ExseSvr_GetReceiveInvoice(0, lng共用票据 & "," & lng使用票据, cllBillInfo, False) = False Then
        zl_GetInvoicePreperty = Ty_Fact
        Exit Function
    End If
    If cllBillInfo.Count <> 0 Then
         Ty_Fact.lngShareUseID = Val(nvl(cllBillInfo(1)("_recv_id"))) ' '共用的领用ID
    End If
    zl_GetInvoicePreperty = Ty_Fact
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_GetInvoiceUserType(ByVal lng病人ID As Long, ByVal lng主页id As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的使用类别
    '返回:发票的使用类别
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str使用类别 As String
    On Error GoTo errHandle
    If zl_ExseSvr_GetPatiInvoiceClass(lng病人ID, lng主页id, intInsure, str使用类别) = False Then Exit Function
    zl_GetInvoiceUserType = str使用类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal int票种 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln是否启用 As Boolean
    
    On Error GoTo errHandle
    If zl_ExseSvr_InvoiceClassUsed(int票种, bln是否启用, True, , glngModul) = False Then Exit Function
    zlStartFactUseType = bln是否启用
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


