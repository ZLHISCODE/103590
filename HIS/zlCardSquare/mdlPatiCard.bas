Attribute VB_Name = "mdlPatiCard"
Option Explicit
Public grs医疗卡类别  As ADODB.Recordset
Public gObjYLCards As clsCards
Public gObjYLCardObjs As clsCardObjects   '当前启用有效的医疗卡
Public gfrmCardMgr As Object
Public gblnNotCloseWindows  As Boolean '不关闭窗口
Public grsSystem As ADODB.Recordset
Public Function zlInitPatiCards(Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡集
    '编制:刘兴洪
    '日期:2011-05-23 17:54:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int自动读取 As Integer, bln启用 As Boolean, str部件 As String, objCard As clsCard
    Dim int自动读取间隔 As Integer, str读卡性质 As String
    Dim objBrushCards As Object
    
    Err = 0: On Error GoTo Errhand:
    
    
    Set gObjYLCards = New clsCards
    Set gObjYLCardObjs = New clsCardObjects
    
    Set grs医疗卡类别 = Nothing: Set grsStatic.rs消费卡接口 = Nothing
    
    Set rsTemp = zlGet医疗卡类别(cnOracle)
    With rsTemp
        '自制卡(即消费卡)
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & nvl(!编码), "自动读取", "0"))
            int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & nvl(!编码), "自动读取间隔", "300"))
            bln启用 = Val(nvl(rsTemp!是否启用)) = 1
            
            '90875,李南春,2016/1/22:证件类型都启用
            If bln启用 Then
                If Val(nvl(rsTemp!是否自制)) = 1 Or Val(nvl(rsTemp!是否证件)) = 1 Then   '自制卡,都启用
                    bln启用 = True
                Else
                    '问题号:54098
                    If (nvl(rsTemp!名称) Like "*身份证*" Or nvl(rsTemp!名称) Like "*IC卡*") And Val(nvl(rsTemp!是否固定)) = 1 And nvl(rsTemp!部件) = "" Then
                        bln启用 = True
                    Else
                        bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\医疗卡\" & nvl(!编码), "启用", "0")) = 1
                    End If
                End If
            End If
            str部件 = Trim(nvl(rsTemp!部件))
            'ID,编码,名称,短名,前缀文本,卡号长度,缺省标志,是否固定,是否严格控制,是否刷卡,是否自制,是否存在帐户,是否全退,部件,备注,特定项目,结算方式,是否启用
            '77872,李南春,2014/10/28:是否支持转帐及代扣
            '85565,李南春,2015/7/10:读卡性质
            '90875,李南春,2016/1/22:是否证件
            '103310,李南春,2016/12/7:卡号后增加回车符位
            Set objCard = New clsCard
            With objCard
                .卡种类 = EM_CardType_Square
                .接口序号 = nvl(rsTemp!id)
                .接口编码 = nvl(rsTemp!编码)
                .短名 = nvl(rsTemp!短名)
                .名称 = nvl(rsTemp!名称)
                .前缀文本 = nvl(rsTemp!前缀文本)
                .卡号长度 = Val(nvl(rsTemp!卡号长度)) + Val(nvl(rsTemp!设备是否启用回车))
                .缺省标志 = Val(nvl(rsTemp!缺省标志)) = 1
                .系统 = Val(nvl(rsTemp!是否固定)) = 1
                .是否严格控制 = Val(nvl(rsTemp!是否严格控制)) = 1
                .是否自动读取 = int自动读取
                .自动读取间隔 = int自动读取间隔
                .自制卡 = Val(nvl(rsTemp!是否自制)) = 1
                .是否存在帐户 = Val(nvl(rsTemp!是否存在帐户)) = 1
                .是否全退 = Val(nvl(rsTemp!是否全退)) = 1
                .卡号重复使用 = Val(nvl(rsTemp!是否重复使用)) = 1
                .结算方式 = nvl(rsTemp!结算方式)
                .接口程序名 = nvl(rsTemp!部件)
                .特定项目 = nvl(rsTemp!特定项目)
                .启用 = bln启用
                .备注 = nvl(rsTemp!备注)
                .卡号密文规则 = nvl(rsTemp!卡号密文)
                .是否退现 = Val(nvl(rsTemp!是否退现)) = 1
                .密码长度 = Val(nvl(rsTemp!密码长度))
                .密码长度限制 = Val(nvl(rsTemp!密码长度限制))
                .密码规则 = Val(nvl(rsTemp!密码规则))
                .密码输入限制 = Val(nvl(rsTemp!密码输入限制))
                .是否缺省密码 = Val(nvl(rsTemp!是否缺省密码)) = 1
                .是否制卡 = Val(nvl(rsTemp!是否制卡)) = 1   '56615
                .是否发卡 = Val(nvl(rsTemp!是否发卡)) = 1 Or .自制卡
                .是否写卡 = Val(nvl(rsTemp!是否写卡)) = 1
                .是否模糊查找 = Val(nvl(rsTemp!是否模糊查找)) = 1
                .是否转帐及代扣 = Val(nvl(rsTemp!是否转帐及代扣)) = 1
                str读卡性质 = nvl(rsTemp!读卡性质, "1000")
                .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
                .是否扫描 = Mid(str读卡性质, 2, 1) = 1
                .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
                .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
                .是否证件 = Val(nvl(rsTemp!是否证件)) = 1
                .是否持卡消费 = Val(nvl(rsTemp!是否持卡消费)) = 1
                .发送调用接口 = Val(nvl(rsTemp!发送调用接口)) = 1
                .是否退款验卡 = Val(nvl(rsTemp!是否退款验卡)) = 1
                .设备是否启用回车 = Val(nvl(rsTemp!设备是否启用回车)) = 1
                .是否缺省退现 = Val(nvl(rsTemp!是否缺省退现)) = 1
            End With
            gObjYLCards.Add objCard, "K" & objCard.接口序号
            If zlCreatePatiCardObject(objCard, objBrushCards) Then
                gObjYLCardObjs.Add objBrushCards, objCard.自制卡, objCard.接口序号, objCard, False, "K" & objCard.接口序号
            End If
            .MoveNext
        Loop
    End With
    
    Set rsTemp = zlGet消费卡接口(cnOracle)
    With rsTemp
        '自制卡(即消费卡)
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & nvl(!编号), "自动读取", "0"))
            int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & nvl(!编号), "自动读取间隔", "300"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & nvl(!编号), "启用", "0")) = 1
            
            '编号,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度,部件,系统,是否密文
            str部件 = Trim(nvl(rsTemp!部件))
            Set objCard = New clsCard
            With objCard
                .卡种类 = EM_CardType_Consume
                .接口序号 = nvl(rsTemp!编号)
                .接口编码 = nvl(rsTemp!编号)
                .短名 = Left(nvl(rsTemp!名称), 1)   '默认取第一个
                .名称 = nvl(rsTemp!名称)
                .前缀文本 = nvl(rsTemp!前缀文本)
                .卡号长度 = Val(nvl(rsTemp!卡号长度))
                .系统 = Val(nvl(rsTemp!系统)) = 1
                .是否严格控制 = False
                .是否自动读取 = int自动读取
                .自动读取间隔 = int自动读取间隔
                .自制卡 = Val(nvl(rsTemp!自制卡)) = 1
                .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
                .是否全退 = Val(nvl(rsTemp!是否全退)) = 1
                .结算方式 = nvl(rsTemp!结算方式)
                .接口程序名 = nvl(rsTemp!部件)
                .特定项目 = ""
                .启用 = bln启用
                .卡号重复使用 = True
                .备注 = ""
                .卡号密文规则 = nvl(rsTemp!是否密文)
                .消费卡 = True
                .是否退现 = Val(nvl(rsTemp!是否退现)) = 1
                .密码长度 = Val(nvl(rsTemp!密码长度))
                .密码长度限制 = Val(nvl(rsTemp!密码长度限制))
                .密码规则 = Val(nvl(rsTemp!密码规则))
                .密码输入限制 = Val(nvl(rsTemp!密码输入限制))
                .是否缺省密码 = Val(nvl(rsTemp!是否缺省密码)) = 1
                .是否制卡 = Val(nvl(rsTemp!是否制卡)) = 1   '56615
                .是否发卡 = Val(nvl(rsTemp!是否发卡)) = 1 Or .自制卡
                .是否写卡 = Val(nvl(rsTemp!是否写卡)) = 1
                
                str读卡性质 = nvl(rsTemp!读卡性质, "1000")
                .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
                .是否扫描 = Mid(str读卡性质, 2, 1) = 1
                .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
                .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            End With
            gObjYLCards.Add objCard, "X" & objCard.接口序号
            If zlCreatePatiCardObject(objCard, objBrushCards) Then
                gObjYLCardObjs.Add objBrushCards, objCard.自制卡, objCard.接口序号, objCard, True, "X" & objCard.接口序号
            End If
            .MoveNext
        Loop
    End With
    zlInitPatiCards = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlGetYLCardObjs(ByRef objYlCardObjects As clsCardObjects) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡对象
    '出参:objYlCardObjects-返回卡对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-23 13:59:24
    '说明:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gObjYLCardObjs Is Nothing Then
        Set objYlCardObjects = gObjYLCardObjs
        zlGetYLCardObjs = True
        Exit Function
    End If
    If gcnOracle.State <> 1 Then Exit Function
    If zlInitPatiCards = False Then Exit Function
    Set objYlCardObjects = gObjYLCardObjs
    zlGetYLCardObjs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetCards_YL(ByRef objCards As clsCards) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别
    '出参:objCards-医疗卡类别对象
    '返回:返回成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-23 12:03:26
    '说明:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not gObjYLCards Is Nothing Then
        Set objCards = gObjYLCards: zlGetCards_YL = True: Exit Function
    End If
    If zlInitPatiCards = False Then Exit Function
    Set objCards = gObjYLCards
    zlGetCards_YL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSystemNo(ByVal lngSys As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的系统号(如:体检2100,则共享的HIS号为100或101)
    '返回:返回共享的号码(没有共享的,直接返回当前系统号)
    '编制:刘兴洪
    '日期:2011-09-21 16:13:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Int(lngSys / 100) = 1 Then zlGetSystemNo = lngSys: Exit Function
    
    On Error GoTo errHandle
    
   gstrSQL = "Select 编号,名称,共享号,所有者  From zltools.zlsystems "
    If grsSystem Is Nothing Then
        Set grsSystem = zlDatabase.OpenSQLRecord(gstrSQL, "获取共享系统号")
    ElseIf grsSystem.State <> 1 Then
        Set grsSystem = zlDatabase.OpenSQLRecord(gstrSQL, "获取共享系统号")
    End If
    grsSystem.Filter = "编号=" & lngSys
    If grsSystem.EOF = False Then
        If Val(nvl(grsSystem!共享号)) <> 0 Then zlGetSystemNo = grsSystem!共享号
    End If
    grsSystem.Filter = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetPatiDayMoney(lng病人ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人当天发生的费用总额
    '返回:获取病人当天发生的费用总额
    '编制:刘兴洪
    '日期:2011-06-23 10:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(intTYPE As Integer, lng病人ID As Long) As Double
'功能:获取指定病人的划价单金额合计
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnAllFee As Boolean, strWhere As String
        
    On Error GoTo errH
    
    '记帐报警包含所有住院划价费用
    If intTYPE = 1 Then
        blnAllFee = Val(zlDatabase.GetPara("记帐报警包含所有住院划价费用", glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(主页ID,0) = (Select Nvl(主页ID,0) From 病人信息 Where 病人ID = [1])"
        End If
    Else
        strWhere = ""
    End If
    
    If intTYPE = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志=2" & strWhere
    Else
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计 " & _
        "   From 门诊费用记录  " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]  and 门诊标志<>2" & _
        "   Union ALL   " & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志<>2 "
        strSQL = "" & _
        "   Select Sum(nvl(划价费用合计,0)) as 划价费用合计  " & _
        "   From ( " & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取指定病人的划价总额", lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGet医疗卡类别(Optional cnOracle As ADODB.Connection) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别
    '返回:返回医疗卡类别的记录集
    '编制:刘兴洪
    '日期:2011-05-23 17:25:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDatabase As Object, objTemp As clsDataBase
    
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    
    '问题号:51072,56615:是否制卡,是否发卡,是否写卡
    '先缓存到本地
    '77872,李南春,2014/10/28:是否支持转帐及代扣
    '90875,李南春,2016/1/22:是否证件类型
    '104238:李南春，2017/2/15，医疗卡类别增加发卡卡号控制
    gstrSQL = "" & _
    "   Select A.Id, A.编码, A.名称, A.短名, A.前缀文本, A.卡号长度, A.缺省标志, A.是否固定, A.是否严格控制, " & _
    "           nvl(A.是否自制,0) as 是否自制, nvl(A.是否存在帐户,0) as 是否存在帐户, " & _
    "           nvl(A.是否全退,0) as 是否全退,nvl(A.是否重复使用,0) as 是否重复使用 , nvl(A.发卡性质,0) as 发卡性质, " & _
    "           nvl(A.密码长度,10) as 密码长度,nvl(A.密码长度限制,0) as 密码长度限制,nvl(A.密码规则,0) as 密码规则," & _
    "           nvl(A.是否退现,0) as 是否退现,A.部件, A.备注, A.特定项目, A.结算方式, A.是否启用, A.卡号密文,Nvl(A.密码输入限制,0) as 密码输入限制,Nvl(A.是否缺省密码,0) as 是否缺省密码," & _
    "           nvl(A.是否模糊查找,0) as 是否模糊查找,nvl(A.是否制卡,0) as 是否制卡, decode(nvl(A.是否自制,0),1,1,nvl(A.是否发卡,0)) as 是否发卡, nvl(A.是否写卡,0) as 是否写卡," & _
    "           B.性质  as 结算性质, nvl(A.是否启用,0) as 是否启用,nvl(是否证件,0) as 是否证件, " & _
    "           nvl(A.是否转帐及代扣,0) as 是否转帐及代扣, nvl(A.读卡性质,'1000') as 读卡性质, " & _
    "           nvl(A.是否持卡消费,0) as 是否持卡消费,nvl(A.发送调用接口,0) as 发送调用接口," & _
    "           Nvl(a.是否退款验卡,0) As 是否退款验卡, A.设备是否启用回车, nvl(A.发卡控制,0) as 发卡控制," & _
    "           Nvl(A.是否缺省退现,0) as 是否缺省退现 " & _
    "    From 医疗卡类别 A,结算方式 B" & _
    "    Where A.结算方式=B.名称(+)" & _
    "    Order by 编码"
    
    If grs医疗卡类别 Is Nothing Then
        Set grs医疗卡类别 = objDatabase.OpenSQLRecord(gstrSQL, "获取消费卡接口 ")
    ElseIf grs医疗卡类别.State <> 1 Then
        Set grs医疗卡类别 = objDatabase.OpenSQLRecord(gstrSQL, "获取消费卡接口 ")
    End If
    grs医疗卡类别.Filter = 0
    Set zlGet医疗卡类别 = grs医疗卡类别
    Set objDatabase = Nothing
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        Set objTemp = Nothing: Set objDatabase = Nothing
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objTemp = Nothing: Set objDatabase = Nothing
End Function

Public Function zlCreatePatiCardObject(ByVal objCard As clsCard, ByRef objCardObject As Object, Optional blnAdviceSend As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定卡的对象
    '出参:objCardObject-被创建的对象
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-25 10:47:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String, strHead As String
    If Not objCard.启用 And Not blnAdviceSend Then
        Set objCardObject = Nothing: Exit Function
    End If
    
    '检查设备是否启用
    strHead = IIf(objCard.消费卡, "", "zl9Card_")
    If objCard.接口程序名 = "" Then
        '99858:李南春,2016/9/2,三方账户医疗卡、非自制消费卡必须有接口部件
        If objCard.消费卡 And objCard.自制卡 Then
            Set objCardObject = New clsSimulateSquareCard: zlCreatePatiCardObject = True: Exit Function
        ElseIf Not objCard.是否存在帐户 Then
            Set objCardObject = New clsOwnerCardObject: zlCreatePatiCardObject = True: Exit Function
        End If
        MsgBox objCard.名称 & "未设置接口部件，请在" & IIf(objCard.消费卡, "【消费卡管理】", "【医疗卡类别管理】") & "中设置部件名!"
        Exit Function
    End If
    strCommpentName = GetCardComponentsStr(objCard.接口程序名, strHead)
    Err = 0: On Error Resume Next
    Set objCardObject = CreateObject(strCommpentName)
    If Err <> 0 Then
        ShowMsgbox "部件:" & objCard.接口编码 & "-" & objCard.名称 & "( " & strCommpentName & ")创建失败,请与系统管理员联系!" & vbCrLf & "详细的信息为:" & Err.Description
        Call WritLog("mdlCardSquare.zlCreatePatiCardObject", "", "部件:" & objCard.接口编码 & "-" & objCard.名称 & "创建失败!详细的信息为:" & Err.Description)
        Exit Function
    End If
    zlCreatePatiCardObject = True
End Function

Public Function zlGetComponentObject(ByVal lng卡类别ID As Long, _
     Optional bln消费卡 As Boolean = False) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的卡对象
    '入参:lng卡类别ID-卡类别ID
    '        bln消费卡
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-06-25 23:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim objYlCardObjs As clsCardObjects
    strKey = IIf(bln消费卡, "X", "K") & lng卡类别ID
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Err = 0: On Error Resume Next
    Set zlGetComponentObject = objYlCardObjs(strKey).CardObject
    If Err <> 0 Then
        Err.Clear: On Error GoTo 0
    End If
End Function
Public Function zlGetClsCardObject(ByVal lng卡类别ID As Long, _
     Optional bln消费卡 As Boolean = False) As clsCardObject
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的卡特性
    '入参:lng卡类别ID-卡类别ID
    '        bln消费卡
    '       blnChkExeObject-检查执行对象
    '       blnChkInitComents-检查初始部件
    '出参:
    '返回:合法,返回clsCardObject对象
    '编制:刘兴洪
    '日期:2011-06-25 23:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    strKey = IIf(bln消费卡, "X", "K") & lng卡类别ID
    Err = 0: On Error Resume Next
    Set zlGetClsCardObject = objYlCardObjs(strKey)
    If Err <> 0 Then
        Err.Clear: On Error GoTo 0
    End If
End Function
   
Public Function zlItemsMoney(ByVal strIDs As String, Optional ByVal strPriceGrade As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定项目的数据
    '入参:传入ID集，用逗号分离
    '返回:相关价格集
    '编制:刘兴洪
    '日期:2011-05-31 15:24:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, j As Long
    Dim strSubTable As String, varData() As Variant
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    Call zlGetSubTable(0, strIDs, strSubTable, varData)
    '价格等级
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [" & UBound(varData) + 2 & "]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [" & UBound(varData) + 2 & "]" & vbNewLine & _
            "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    
    strSubTable = " With 挂号项目 as ( " & strSubTable & ") "
    
    strSQL = strSubTable & _
        "   Select  /*+ rule */  1 as 性质,A.类别,A.ID as 主项ID,0 as 从项目ID, " & _
        "               A.编码 as 项目编码,A.名称 as 项目名称, A.计算单位,A.屏蔽费别," & _
        "               1 as 数次,B.现价 as 单价, " & _
        "               C.ID as 收入项目ID,C.名称 as 收入项目, " & _
        "               C.编码 as 收入编码,C.收据费目" & _
        "   From 收费项目目录 A,收费价目 B,收入项目 C,挂号项目 M" & _
        "   Where A.ID =M.ID and B.收费细目ID=A.ID And B.收入项目ID=C.ID  " & _
        "               And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
                strWherePriceGrade & vbNewLine & _
        "   Union ALL " & _
        "   Select 2 as 性质,A.类别,D.主项ID,A.ID as 项目ID, " & _
        "               A.编码 as 项目编码,A.名称 as 项目名称,A.计算单位,A.屏蔽费别," & _
        "               D.从项数次 as 数次,B.现价 as 单价, " & _
        "               C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目" & _
        "   From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D,挂号项目 M" & _
        "   Where A.ID=D.从项ID And D.主项ID =M.ID And A.ID=B.收费细目ID And B.收入项目ID=C.ID   " & _
        "           And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
                strWherePriceGrade
    strSQL = "Select /*+ RULE */  * From (" & strSQL & ")        "
    
    ReDim Preserve varData(UBound(varData) + 1)
    varData(UBound(varData)) = strPriceGrade
    Set zlItemsMoney = zlDatabase.OpenSQLRecordByArray(strSQL, "获取挂号价格", varData)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlGetRecodersFieldIns(ByVal rsTemp As ADODB.Recordset, ByVal strFieldNames As String, ByRef cllData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定记录集中的字段集
    '入参:rsTemp-指定的记录集
    '        strFields-指定的字段,比如:项目ID,科室ID
    '出参:cllData-返回指定的集(以:项目ID,科室ID为主键建立)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-31 17:38:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varFields As Variant, i As Long
    Dim strFields() As String, strTemp As String
    varFields = Split(strFieldNames, ",")
    ReDim Preserve strFields(0 To UBound(varFields)) As String
    With rsTemp
        Do While Not .EOF
            For i = 0 To UBound(varFields)
                If InStr(1, strFields(i) & ",", "," & .Fields(varFields(i)).value & ",") = 0 Then
                    strFields(i) = strFields(i) & "," & .Fields(varFields(i)).value
                End If
            Next
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        Set cllData = New Collection
        For i = 0 To UBound(varFields)
            strTemp = strFields(i)
            If Trim(strTemp) <> "" Then strTemp = Mid(strTemp, 2)
            cllData.Add strTemp, varFields(i)
        Next
    End With
End Function
Public Function zlGetSubTable(ByVal bytType As Byte, ByVal strValues_IN As String, _
    strSubTable As String, varPara() As Variant, Optional intParaStep As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将字符串分解成子表查询(用Num2list),超过的20个的
    '入参:bytType: 0-Num2List;1-Str2List;2-Num2List2;3Str2List2
    '       strValues_IN:bytType=0,1时,之间用逗号分离
    '                            bytType=2,3时,列之间用:分离行之间用,分离:如:张三:22,李四:22
    '       varParaStep_in:参数的启始步长
    '出参:varPara-返回重0-20的数组
    '返回:成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-06-01 10:50:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, j As Long
    Dim strTemp As String, strSplit As String
    Dim varValue As Variant
    Dim strTable As String
    i = intParaStep: strSplit = ","
    If bytType = 0 Or bytType = 2 Then
        Do While strValues_IN <> ""
            ReDim Preserve varPara(0 To i) As Variant
            If Len(strValues_IN) > 4000 Then
                j = InStr(IIf(bytType = 0, 3982, 3958), strValues_IN, strSplit)
                strTemp = Mid(strValues_IN, 1, j - 1): strValues_IN = Mid(strValues_IN, j + 1)
                varPara(i) = strTemp
            Else
                strTemp = strValues_IN
                varPara(i) = strTemp
                strValues_IN = ""
            End If
            i = i + 1
        Loop
    Else
        varValue = Split(strValues_IN, strSplit)
        strTemp = ""
        For j = 0 To UBound(varValue)
              If zlCommFun.ActualLen(strTemp & "," & varValue(j)) > 4000 Then
                  ReDim Preserve varPara(0 To i) As Variant
                  varPara(i) = Mid(strTemp, 2): i = i + 1
                  strTemp = ""
              End If
              strTemp = strTemp & "," & varValue(j)
        Next
        If strTemp <> "" Then
            ReDim Preserve varPara(0 To i) As Variant
            varPara(i) = Mid(strTemp, 2)
        End If
    End If
    For i = intParaStep To UBound(varPara)
        If varPara(i) <> "" Then
            j = i + 1
            If bytType = 0 Then
                strTable = strTable & " Union All Select Column_Value as ID From Table( f_Num2list([" & j & "])) "
            ElseIf bytType = 1 Then
                strTable = strTable & " Union All Select Column_Value From Table( f_str2list([" & j & "])) "
            ElseIf bytType = 2 Then
                strTable = strTable & " Union All Select  C1,C2 From Table( f_Num2list2([" & j & "])) "
            Else
                strTable = strTable & " Union All Select  C1,C2 From Table( f_Str2list2([" & j & "])) "
            End If
        End If
    Next
    If strTable = "" Then Exit Function
    strSubTable = "Select distinct  * From ( " & Mid(strTable, 11) & ")"
    zlGetSubTable = True
End Function

Public Function zlGetActualMoney(ByVal str费别 As String, ByVal lng收入ID As Long, ByVal dbl应收 As Double, ByVal lng收费细目ID As Long) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的费别和收入项目或收费项目,计算指定金额的实际收款金额
    '入参:str费别-费别
    '        lng收入ID-收入项目ID
    '        dbl应收-应收金额值
    '出参:
    '返回:实际应收的金额
    '编制:刘兴洪
    '日期:2011-06-02 11:50:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4])  as 实收金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str费别, lng收费细目ID, lng收入ID, dbl应收)
    If rsTmp.EOF Then
        zlGetActualMoney = dbl应收
    Else
        zlGetActualMoney = Round(Val(Split(nvl(rsTmp!实收金额) & ":", ":")(1)), 5)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetIDKindStr(Optional strIDKindStr As String = "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;就|就诊卡|0", Optional blnOnlyAccouct As Boolean = False, Optional objSquare As clsCardSquare) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效医疗卡字符串
    '入参:strIDKindStr    String  IN
    '有两种格式:
    '一种是缺省的:短名1|全名1|读卡标志1;…. ;短名n|全名n|读卡标志n
    '另一种是该接口返回的格式:短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志|是否存在帐户;…
    '出参:
    '返回: 短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)|是否扫描|是否接触式读卡|是否非接触式读卡;…
    '        其中:卡类别ID|长度是本次增加的,由调用者根据情况来确认.
    '       比如:身|身份证号|0|0|18|0;IC|IC卡号|1|0|8|0;门|门诊号|0|0|0|0;就|就诊卡|0|0|0|1;建|建行卡|0|0|10|0
    '      出现错误时,返回空
    '编制:刘兴洪
    '日期:2011-06-14 14:43:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim strNewIdKindStr As String, strTemp As String
    Dim lngMaxLen As Long, blnPassText As Boolean
    Dim blnExists As Boolean '是否存在模糊查找
    '77076,冉俊明,2014-8-25,同时打开医疗卡发放管理和病人信息管理（或病人入院管理）,在病人信息管理中打开登记,医疗卡发放管理窗体自动关闭
    Dim objCard As New Card, objCards As New Cards, objCardSquare As clsCardSquare
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandle
        
    If objSquare Is Nothing Then
        Set objCardSquare = New clsCardSquare
    Else
        Set objCardSquare = objSquare
    End If
    strNewIdKindStr = ""
    varData = Split(strIDKindStr, ";")
    blnExists = False
    '76187,冉俊明,2014-8-4
    objCardSquare.mblnYLMgr = True
    Set objCards = objCardSquare.zlGetCards(1)
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|||||||", "|")
            blnFind = False
            For Each objCard In objCards
                If objCard.名称 = varTemp(1) Then blnFind = True: Exit For
            Next
            If blnFind Then
                '85565,李南春,2015/7/10:读卡性质
                strNewIdKindStr = strNewIdKindStr & ";" & objCard.短名 & "|" & objCard.名称 & "|" & IIf(objCard.是否刷卡, 0, 1) & _
                                "|" & objCard.接口序号 & "|" & objCard.卡号长度 & "|" & IIf(objCard.缺省标志, 1, 0) & _
                                "|" & IIf(objCard.是否存在帐户, 1, 0) & "|" & objCard.卡号密文规则 & _
                                "|" & IIf(objCard.是否扫描, 1, 0) & "|" & IIf(objCard.是否接触式读卡, 1, 0) & "|" & IIf(objCard.是否非接触式读卡, 1, 0)
                strTemp = strTemp & "," & objCard.接口序号
            Else
                '短名|全名|读卡标志|卡类别ID(-1代表模糊查找)|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)|是否扫描|是否接触式读卡|是否非接触式读卡;…
                strNewIdKindStr = strNewIdKindStr & ";" & varTemp(0) & "|" & varTemp(1) & "|" & Val(varTemp(2)) & "|" & varTemp(3) & "|" & varTemp(4) & "|||"
            End If
            If varTemp(1) = "模糊查找" And Val(varTemp(3)) < 0 Then blnExists = True
        End If
    Next

    For Each objCard In objCards
        If InStr(1, strTemp & ",", "," & objCard.接口序号 & ",") = 0 Then
            If Not blnOnlyAccouct Or (blnOnlyAccouct And objCard.是否存在帐户) Then
                '85565,李南春,2015/7/10:读卡性质
                strNewIdKindStr = strNewIdKindStr & ";" & objCard.短名 & "|" & objCard.名称 & "|" & IIf(objCard.是否刷卡, 0, 1) & _
                                "|" & objCard.接口序号 & "|" & objCard.卡号长度 & "|" & IIf(objCard.缺省标志, 1, 0) & _
                                "|" & IIf(objCard.是否存在帐户, 1, 0) & "|" & objCard.卡号密文规则 & _
                                "|" & IIf(objCard.是否扫描, 1, 0) & "|" & IIf(objCard.是否接触式读卡, 1, 0) & "|" & IIf(objCard.是否非接触式读卡, 1, 0)
            End If
        End If
    Next
    
    If strNewIdKindStr <> "" Then strNewIdKindStr = Mid(strNewIdKindStr, 2)
    GetIDKindStr = strNewIdKindStr
    Exit Function
errHandle:
    GetIDKindStr = ""
End Function

Private Function IsCreateObject(ByVal str部件 As String, Optional strHead As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部件是否创建成功
    '入参:strHead-部件名的文件头:比如:医疗卡以部件名:zl9Card_打头.
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 15:43:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Object
    If str部件 = "" Then IsCreateObject = True: Exit Function
    str部件 = GetCardComponentsStr(str部件)
    Err = 0: On Error Resume Next
    Set objTemp = CreateObject(str部件)
    If Err <> 0 Then
       Err.Clear: On Error GoTo 0
       Set objTemp = Nothing
       IsCreateObject = False: Exit Function
    End If
    Set objTemp = Nothing
    IsCreateObject = True
End Function
Private Function GetCardComponentsStr(ByVal str部件 As String, Optional strHead As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取部件名称
    '返回:
    '编制:刘兴洪
    '日期:2011-06-22 13:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If str部件 = "" Then GetCardComponentsStr = "": Exit Function
    If strHead <> "" Then
        '有开头部分,需要检查传入是否包含这部分
        If str部件 Like strHead & "*" Then
            str部件 = str部件 & "." & "cls" & Mid(str部件, Len(strHead) + 1)
        Else
            str部件 = strHead & str部件 & "." & "cls" & Replace(Replace(UCase(str部件), "ZL9", ""), "ZL", "")
        End If
    Else
        str部件 = str部件 & "." & "cls" & Replace(Replace(UCase(str部件), "ZL9", ""), "ZL", "")
    End If
    GetCardComponentsStr = str部件
End Function
Public Function zlGetPrivFuns(ByVal lngModule As Long, Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定模块的权限
    '返回:返回权限串
    '编制:刘兴洪
    '日期:2015-06-03 09:46:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String, objDatabase As clsDataBase
    On Error GoTo errHandle
    If cnOracle Is Nothing Then
        strPrivs = ";" & GetPrivFunc(glngSys, lngModule) & ";"
    Else
        Set objDatabase = New clsDataBase
        Call objDatabase.InitCommon(cnOracle)
        strPrivs = ";" & objDatabase.GetPrivFunc(glngSys, lngModule) & ";"
        Set objDatabase = Nothing
    End If
    zlGetPrivFuns = strPrivs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetAvailabilityCardType(Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前工作站有效的支付卡,比如银行卡;消费卡等
    '返回:格式:短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡|是否退现|是否全退|是否扫描|是否接触式读卡|是否非接触式读卡;…
    '编制:刘兴洪
    '日期:2011-06-14 15:16:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int自动读取 As Integer, str医疗卡 As String
    Dim bln启用 As Boolean, strCardStr As String
    Dim objCard As clsCard, i As Long, blnAdd As Boolean
    Dim strPrivs As String, bln支持三方接口 As Boolean  '-False,表示只支持消费卡;True:支持消费卡;银行卡等
    Dim objYlCardObjs As clsCardObjects
    Dim objDatabase As clsDataBase
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    On Error GoTo errHandle
    
    '是否有一卡通消费操作的虚拟模块
    strPrivs = zlGetPrivFuns(1151, cnOracle)
    
    bln支持三方接口 = InStr(1, strPrivs, ";三方接口消费;") > 0
    
    For i = 1 To objYlCardObjs.count
        If objYlCardObjs(i).CardPreporty.启用 And objYlCardObjs(i).CardPreporty.是否存在帐户 Then
            If Not objYlCardObjs(i).CardObject Is Nothing Then
                blnAdd = True
                If bln支持三方接口 = False Then
                    blnAdd = False
                    If objYlCardObjs(i).CardPreporty.消费卡 And _
                        objYlCardObjs(i).CardPreporty.自制卡 Then
                        blnAdd = True
                    End If
                End If
                If blnAdd Then
                    '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡|是否退现|是否全退|是否扫描|是否接触式读卡|是否非接触式读卡;…
                    '85565,李南春,2015/7/10:读卡性质
                    Set objCard = objYlCardObjs(i).CardPreporty
                    strCardStr = strCardStr & ";" & objCard.短名 & "|" & objCard.名称 & "|" & IIf(objCard.是否刷卡, 0, 1)
                    strCardStr = strCardStr & "|" & objCard.接口序号 & "|" & objCard.卡号长度
                    strCardStr = strCardStr & "|" & IIf(objCard.消费卡, 1, 0) & "|" & objCard.结算方式
                    strCardStr = strCardStr & "|" & objCard.卡号密文规则 & "|" & IIf(objCard.自制卡, 1, 0)
                    strCardStr = strCardStr & "|" & IIf(objCard.是否退现, 1, 0) & "|" & IIf(objCard.是否全退, 1, 0)
                    strCardStr = strCardStr & "|" & IIf(objCard.是否扫描, 1, 0) & "|" & IIf(objCard.是否接触式读卡, 1, 0)
                    strCardStr = strCardStr & "|" & IIf(objCard.是否非接触式读卡, 1, 0) & "|" & IIf(objCard.是否退款验卡, 1, 0)
                    '问题:50120
                End If
            End If
        End If
    Next
    If strCardStr <> "" Then strCardStr = Mid(strCardStr, 2)
     GetAvailabilityCardType = strCardStr
    Exit Function
errHandle:
     GetAvailabilityCardType = ""
End Function

Public Function GetIDKindCardTypeID(ByVal strIDKindStr As String, ByVal strIDKind As String, _
    ByRef lngCardTypeID As Long, ByRef lngCardLen As Long, _
    Optional ByRef strKindName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:主要是根据上IDKind部件中的IDKind来获取相应卡类别ID,以便查找相关的病人信息:
    '入参:strIDKindStr  -缺省的StrIDKindStr:格式:短名|全名|读卡标志|卡类别ID|卡号长度;…如
    '       身|身份证号|0|0|18;IC|IC卡号|1|0|8;门|门诊号|0|0|0;就|就诊卡|0|0|0;建|建行卡|0|0|10
    '       strIDKind-可以为缺省的名称;也可以为索引(索引从0…N):用名称,重复时,指向第一个.
    '出参:lngCardTypeID- 卡类别ID
    '       lngCardLen- 卡号长度
    '       strKindName-名称
    '返回:Boolean 返回    成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-06-14 16:11:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, blnIndex As Boolean, i As Long
    
    On Error GoTo errHandle
    varData = Split(strIDKindStr, ";")
    lngCardTypeID = -1: lngCardLen = -1
    blnIndex = IsNumeric(strIDKind)
    For i = 0 To UBound(varData)
            If blnIndex Then
                If i = Val(strIDKind) Then
                    '格式:短名|全名|读卡标志|卡类别ID|卡号长度;…
                    varTemp = Split(varData(i) & "|||||", "|")
                    lngCardTypeID = Val(varTemp(3))
                    lngCardLen = Val(varTemp(4))
                    strKindName = Trim(varTemp(1))
                    Exit For
                End If
            Else
                    varTemp = Split(varData(i) & "|||||", "|")
                    If varTemp(1) = strIDKind Then
                        lngCardTypeID = Val(varTemp(3))
                        lngCardLen = Val(varTemp(4))
                        strKindName = Trim(varTemp(1))
                        Exit For
                    End If
            End If
    Next
    If lngCardTypeID >= 0 Then
         GetIDKindCardTypeID = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetCardFindPati( _
    ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, _
    Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, _
    Optional cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡号，模糊查找病人
    '        strCardNo-卡号
    '        blnNotShowErrMsg-不显示错误的提示信息
    '出参:strErrMsg-返回的错误信息
    '        lng病人ID-返回的病人ID
    '        strCardPass-返回卡号的密码
    '        lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str卡类别ID As String, str病人ID As String, lngTemp As Long
    Dim objDatabase As Object, objTemp As clsDataBase
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    
    
    '模糊查找
    '76020,冉俊明,2014-7-30,仍正常提取出了卡类别是支持模糊查找、已停用的持卡人信息
    '114161:李南春,2017/11/7,挂失有效天数为Null或0时，表示挂失一直生效
    strSQL = "" & _
            " Select a.卡类别id, a.病人id, a.密码," & _
            "       Case" & _
            "         When Nvl(a.状态, 0) = 1" & _
            "           And (Nvl(b.有效天数, 0)  = 0 Or Nvl(a.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.有效天数, 0) > Sysdate) Then 1" & _
            "         When Nvl(a.状态, 0) = 2 Then 2" & _
            "         Else 0" & _
            "       End As 状态" & _
            " From 病人医疗卡信息 A, 医疗卡挂失方式 B, 医疗卡类别 C" & _
            " Where a.卡类别id = c.Id And Nvl(c.是否模糊查找, 0) = 1" & _
            "      And a.卡号 = [2] And a.挂失方式 = b.名称(+)" & _
            "      And Nvl(c.是否启用, 0) = 1" & _
            " Order By 状态"
    Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "获取病人ID", lngCardTypeID, strCardNo)
    
    Set objDatabase = Nothing: Set objTemp = Nothing
    If rsTemp.EOF Then Exit Function
    rsTemp.Filter = "状态=0"
    '0-正常有效卡;1-已挂失; 2-补卡停用
    If rsTemp.RecordCount = 1 Then
        '有一条，直接返回
        lng病人ID = Val(nvl(rsTemp!病人ID))
        strCardPassWord = nvl(rsTemp!密码)
        rsTemp.Close: Set rsTemp = Nothing
        GetCardFindPati = True: Exit Function
    End If
    If rsTemp.RecordCount = 0 Then Exit Function
    '多条
    rsTemp.MoveFirst
    With rsTemp
        str病人ID = ""
        Do While Not .EOF
            lngTemp = Val(nvl(!卡类别id))
            If lngTemp <> 0 Then
                If InStr(1, str卡类别ID & ",", "," & lngTemp & ",") = 0 Then str卡类别ID = str卡类别ID & "," & lngTemp
                If InStr(1, str病人ID & ",", "," & Val(nvl(!病人ID)) & ",") = 0 Then str病人ID = str病人ID & "," & Val(nvl(!病人ID))
            End If
            .MoveNext
        Loop
        If str病人ID <> "" Then str病人ID = Mid(str病人ID, 2)
        If str卡类别ID <> "" Then str卡类别ID = Mid(str卡类别ID, 2)
        If InStr(1, str病人ID, ",") = 0 Then
            .MoveFirst
            lng病人ID = Val(nvl(rsTemp!病人ID)): lngCardTypeID = Val(nvl(rsTemp!卡类别id))
            strCardPassWord = nvl(rsTemp!密码)
            rsTemp.Close: Set rsTemp = Nothing
            GetCardFindPati = True: Exit Function
        End If
        If frmSelectType.zlSelect(Nothing, str卡类别ID, lngCardTypeID) = False Then lngCardTypeID = 0: Exit Function
        rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " And 状态=0"
        If rsTemp.EOF Then lngCardTypeID = 0: Exit Function
        lng病人ID = Val(nvl(rsTemp!病人ID))
        strCardPassWord = nvl(rsTemp!密码)
        rsTemp.Close: Set rsTemp = Nothing: GetCardFindPati = True: Exit Function
    End With
    '肯定有误，按第一条提示
     rsTemp.Filter = 0: rsTemp.MoveFirst
     If Val(nvl(rsTemp!状态)) = 1 Then
        strErrMsg = "卡号为" & strCardNo & "已经被挂失!"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Val(nvl(rsTemp!状态)) = 2 Then
        strErrMsg = "卡号为" & strCardNo & "已经被停用!"
         If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiID(ByVal strCardType As String, _
    ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, _
    Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, _
    Optional objCtl As Object = Nothing, _
    Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, _
    Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional cnOracle As ADODB.Connection, _
    Optional ByRef blnCertificate As Boolean = False, _
    Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       frmMain-调用的主窗体
    '       objCtl-调用的控件
    '       blnShowMergePati-当出现多个满足条件的病人时,是否显示合并功能按钮
    '       blnOnlyContractPati-签约病人
    '       blnUserCancel-选择器中，用户选择了取消
    '       lngShowCardNoTypeID-过滤出多条病信息时，弹出选择器中显示的卡号的卡类别ID,0-表示不显示卡号；>0表示显示指定卡号类别的ID
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str卡类别ID As String, str病人ID As Long, lngTemp As Long
    Dim strWhere As String, blnCard As Boolean '
    Dim str类别名称 As String, str标识号 As String
    Dim objDatabase  As Object, objTemp As clsDataBase
    
    On Error GoTo errHandle
    Set objDatabase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDatabase = objTemp
    End If
    
    strCardPassWord = "": strErrMsg = ""
    lng病人ID = 0
    If strCardType = "" Then Exit Function
    If Val(strCardType) = -1 Then
       GetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID, cnOracle)
       Exit Function
    End If
    
    str类别名称 = ""
    If strCardType Like "*身份证*" Or strCardType Like "*IC卡*" Then
        Set rsTemp = zlGet医疗卡类别
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Val(nvl(rsTemp!是否固定)) = 1 And Val(nvl(rsTemp!是否启用)) = 1 Then
                If strCardType Like "*身份证*" And nvl(rsTemp!名称) Like "*身份证*" And strCardType <> "联系人身份证" Then
                    str类别名称 = nvl(rsTemp!名称)
                     strCardType = Val(nvl(rsTemp!id)): Exit Do
                ElseIf strCardType Like "*IC卡*" And nvl(rsTemp!名称) Like "*IC卡*" Then
                     str类别名称 = nvl(rsTemp!名称)
                     strCardType = Val(nvl(rsTemp!id))
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If IsNumeric(strCardType) Then  '以卡类别ID为准
        strSQL = "" & _
        "   Select  A.卡类别ID, A.病人ID, 密码, A.状态, " & _
        "               nvl(挂失时间,to_date('3000-01-01','yyyy-mm-dd'))+nvl(B.有效天数,0) as 挂失时间," & _
        "               sysdate as 当前时间  " & _
        "   From 病人医疗卡信息 A,医疗卡挂失方式 B" & _
        "   Where  A.卡类别ID=[1] and A.卡号=[2] And A.挂失方式=B.名称(+)"
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "获取病人ID", Val(strCardType), strCardNo)
        If Not rsTemp.EOF Then
            lng病人ID = Val(nvl(rsTemp!病人ID))
            strCardPassWord = nvl(rsTemp!密码)
            If Val(nvl(rsTemp!状态)) = 1 Then
                If Format(rsTemp!挂失时间, "yyyy-mm-dd") <= Format(rsTemp!当前时间, "yyyy-mm-dd") Then
                    strErrMsg = "卡号为" & strCardNo & "已经被挂失!"
                    If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    rsTemp.Close: Set rsTemp = Nothing
                    Exit Function
                End If
            End If
            If Val(nvl(rsTemp!状态)) = 2 Then
                strErrMsg = "卡号为" & strCardNo & "已经被停用!"
                 If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 rsTemp.Close: Set rsTemp = Nothing
                Exit Function
            End If
            GetPatiID = True
            Exit Function
        End If
        If blnOnlyContractPati Then Exit Function
        
        If str类别名称 = "" Then
              Set rsTemp = zlGet医疗卡类别(cnOracle)
              rsTemp.Filter = "ID=" & Val(strCardType) & " And 是否启用=1"
              If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
              If Not rsTemp.EOF Then
                  If Val(nvl(rsTemp!是否固定)) = 1 Then
                      str类别名称 = rsTemp!名称
                  End If
              End If
              rsTemp.Filter = 0
          End If
          
        If str类别名称 Like "*身份证*" Then
            strCardType = "身份证"
        ElseIf UCase(str类别名称) Like "*IC卡*" Then
            strCardType = "IC卡"
        Else
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
    End If
    
    '90875:李南春,2015/12/16,医疗卡证件类型
    If blnCertificate Then
        strSQL = "" & _
        "   Select  A.卡类别ID, A.病人ID, 密码, A.状态 " & _
        "   From 病人医疗卡信息 A,医疗卡类别 B" & _
        "   Where A.卡类别ID=B.ID And A.状态=0 And B.是否启用=1 And B.名称=[1] and A.卡号=[2] And Nvl(B.是否证件,0)=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人ID", strCardType, strCardNo)
        If rsTemp.EOF Then Exit Function
        lng病人ID = Val(nvl(rsTemp!病人ID))
        strCardPassWord = nvl(rsTemp!密码)
        If Val(nvl(rsTemp!状态)) = 1 Then
            strErrMsg = "卡号为" & strCardNo & "已经被挂失!"
            If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
        If Val(nvl(rsTemp!状态)) = 2 Then
            strErrMsg = "卡号为" & strCardNo & "已经被停用!"
             If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
             rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
        GetPatiID = True
        Exit Function
    End If
    
     blnCard = True
    '问题:47939
    Select Case UCase(strCardType)
    Case "IC卡", "IC卡号"
        strWhere = "IC卡号=[2] "
    Case "身份证", "身份证号"
        strWhere = "身份证号=[2] "
    Case "联系人身份证" '问题号:51071
        strWhere = "联系人身份证号=[2]"
    Case "医保号", "医保证号"
        strWhere = "医保号=[2] "
    Case "手机号"
        strWhere = "手机号=[2] "
    Case "门诊号"
        strWhere = "门诊号=[3] "
        str标识号 = strCardNo
    '84247:李南春,2015/4/24,住院号查找病人
    Case "住院号"
        strWhere = "a.病人ID = (Select Nvl(Max(病人ID),0) As 病人ID From 病案主页 Where 住院号 = [3]) "
        str标识号 = strCardNo
    Case Else
        strWhere = "" & strCardType & "=[2] "
        blnCard = False
    End Select
    strSQL = "" & _
    "Select Rownum As ID, a.*" & vbNewLine & _
    "From (Select Decode(Nvl(max(a.在院), 0), 1, '√', '') As 在院, a.病人id, max(a.姓名) As 名称, max(a.性别)as 性别, max(a.年龄) as 年龄, max(a.身份证号) as 身份证号," & vbNewLine & _
    "             max(a.Ic卡号) as IC卡号,max( a.门诊号)as 门诊号, max(a.住院号)as 住院号,max(a.手机号)as 手机号,max( a.出生日期) as 出生日期, max(a.出生地点) as 出生地点," & vbNewLine & _
    "             max(a.费别) as 费别, max(a.医疗付款方式)as 医疗付款方式, max(a.民族) as 民族,max(a.家庭地址) as 家庭地址, max(a.家庭电话) as 家庭电话," & vbNewLine & _
    "             max(a.联系人姓名) as 联系人姓名, max(a.联系人关系)as 联系人关系, max(a.联系人电话) as 联系人电话,max(a.联系人身份证号) as 联系人身份证号," & vbNewLine & _
    "             max(a.住院次数)as 住院次数,max(a.卡验证码) As 密码id," & vbNewLine & _
    "             LTrim(To_Char(max(Decode(类型, 1, Nvl(b.预交余额, 0), 0)), '99999999990.00')) As 门诊预交余额," & vbNewLine & _
    "             LTrim(To_Char(max(Decode(类型, 1, 0, Nvl(b.预交余额, 0))), '99999999990.00')) As 住院预交余额" & IIf(lngShowCardNoTypeID <> 0, ",max(c.卡号) as 卡号", "") & vbNewLine & _
    "       From 病人信息 A, 病人余额 B" & IIf(lngShowCardNoTypeID <> 0, ",病人医疗卡信息 C", "") & vbNewLine & _
    "       Where a.停用时间 Is Null And a.病人id = b.病人id(+) And b.性质(+) = 1 And " & IIf(lngShowCardNoTypeID <> 0, "a.病人ID=c.病人ID(+) And c.卡类别ID(+)=[5] And ", "") & strWhere & vbNewLine & _
    "       group by a.病人ID" & vbNewLine & _
    "       Order By 病人id" & IIf(lngShowCardNoTypeID <> 0, ", 卡号", "") & " ) A"
    Dim frmSel As New frmPatiSelect
    '52913
    '80886,冉俊明,2014-12-18,当卡号为"310D664700068D9E",使用Val(strCardNo)会报"溢出"错误
    If Not frmSel.ShowSelect(frmMain, cnOracle, glngSys, glngModul, objCtl, strSQL, "病人选择", "当前搜索出多个病人信息,请选择指定的病人", True, blnShowMergePati, _
                             IIf(objCtl Is Nothing, False, True), "", "密码ID,ID", rsTemp, blnUserCancel, Val(strCardType), strCardNo, Val(str标识号), lng病人ID, lngShowCardNoTypeID) Then
        If Not frmSel Is Nothing Then Unload frmSel
        Set frmSel = Nothing: Exit Function
    End If
    If Not frmSel Is Nothing Then Unload frmSel
    Set frmSel = Nothing
    
    If rsTemp Is Nothing Then GoTo GoClsObject:
    If rsTemp.State <> 1 Then GoTo GoClsObject:
    If rsTemp.EOF Then
        If blnCard Then
            'IC卡,可能是模糊查找
            GetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID)
        End If
        rsTemp.Close
        GoTo GoClsObject:
        Exit Function
    End If
    
    lng病人ID = Val(nvl(rsTemp!病人ID))
    strCardPassWord = nvl(rsTemp!密码ID)
    GetPatiID = True
    Set objTemp = Nothing: Set objDatabase = Nothing
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        GoTo GoClsObject: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
GoClsObject:
    Set objTemp = Nothing: Set objDatabase = Nothing
    Set rsTemp = Nothing
End Function
 
Private Function GetCardNODencodeRule(ByVal lng卡类别ID As Long, _
    Optional bln消费卡 As Boolean = False, Optional cnOracle As ADODB.Connection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡类别ID的规则
    '入参:lng卡类别ID-卡类别ID
    '        bln消费卡-是否消费卡
    '返回:卡类别的卡号编码规则
    '编制:刘兴洪
    '日期:2011-06-22 11:01:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    If bln消费卡 Then
        Set rsTemp = zlGet消费卡接口(cnOracle)
        rsTemp.Filter = "编号=" & lng卡类别ID
        If rsTemp.EOF Then GoTo GoEnd:
        GetCardNODencodeRule = nvl(rsTemp!是否密文)
        GoTo GoEnd:
    End If
    Set rsTemp = zlGet医疗卡类别(cnOracle)
    rsTemp.Filter = "ID=" & lng卡类别ID
    If rsTemp.EOF Then GoTo GoEnd:
    GetCardNODencodeRule = nvl(rsTemp!卡号密文)
GoEnd:
    rsTemp.Filter = 0
End Function
Public Function GetCardNODencode(ByVal strCardNo As String, _
    Optional lng卡类别ID As Long = 0, _
    Optional strRule As String = "", Optional bln消费卡 As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡编码
    '入参:lng卡类别ID-卡类别ID或消费卡序号,如果传入,将以医疗卡类别或消费卡别中的"卡号密文"或是否密文进行加密
    '       strRule-规则:2-4表示从2位到4位用*代替,如无-号,则表示从最后几位显示为*
    '       strCardNo-卡号
    '出参:
    '返回:带**的卡号,如果错误,返回空
    '编制:刘兴洪
    '日期:2011-06-21 14:21:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPass As Variant
    Dim strCardPassText As String, i As Long, j As Long
    If bln消费卡 Then
        If Val(strRule) = 1 Then GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        If lng卡类别ID = 0 Then GetCardNODencode = strCardNo: Exit Function
        If Val(GetCardNODencodeRule(lng卡类别ID, True)) = 1 Then
            GetCardNODencode = String(Len(strCardNo), "*"): Exit Function
        Else
            GetCardNODencode = strCardNo: Exit Function
        End If
    End If
    If lng卡类别ID <> 0 And strRule = "" Then
        strCardPassText = GetCardNODencodeRule(lng卡类别ID)
    Else
        '取号规则
        strCardPassText = strRule
    End If
    If strCardPassText = "" Then
       GetCardNODencode = strCardNo
    End If
    varPass = Split(strCardPassText & "-", "-")
    If Val(varPass(0)) = 0 Or Val(varPass(1)) = 0 Then
        '最后几位显示*
        i = IIf(Val(varPass(0)) = 0, Val(varPass(1)), Val(varPass(0)))
        If i = 0 Then GetCardNODencode = strCardNo: Exit Function
        j = Len(strCardNo) - i: j = IIf(j < 0, 0, j)
        GetCardNODencode = Mid(strCardNo, 1, j) & String(i, "*")
        Exit Function
    End If
    i = Val(varPass(0)): j = Val(varPass(1))
    If i > Len(strCardNo) Then GetCardNODencode = strCardNo: Exit Function
    If j > Len(strCardNo) Then j = Len(strCardNo)
    If j < i Then j = i
   GetCardNODencode = Mid(strCardNo, 1, i - 1) & String(j - i + 1, "*") & Mid(strCardNo, j + 1)
End Function
Public Function InitInterFacel(ByVal frmMain As Object, ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional bln消费卡 As Boolean = False, Optional ByRef objPatiCard As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化指定卡接口
    '入参:lngCardTypeID-指定卡类别
    '       bln消费卡-是否消费卡
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, objCardObject As Object, strKey As String, strExpand As String
    Dim blnOnlyNotObject As Boolean
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    If Not objPatiCard Is Nothing Then
        If objPatiCard.接口序号 = lngCardTypeID And objPatiCard.消费卡 = bln消费卡 Then
            If Not objPatiCard.InitCompents Then
                If objPatiCard.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not objPatiCard.CardObject.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then                     '初始化部件
                    Exit Function
                End If
                objPatiCard.InitCompents = True
            End If
            InitInterFacel = True
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln消费卡, "X", "K") & lngCardTypeID
    '59760
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Set objPatiCard = objYlCardObjs(strKey)
    If Err <> 0 Then
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!"
                Call WritLog("zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!")
                Exit Function
            End If
            '重新创建
            If zlCreatePatiCardObject(objCard, objCardObject) = False Then Exit Function
            '增加对应
           Set objPatiCard = objYlCardObjs.Add(objCardObject, objCard.自制卡, objCard.接口序号, objCard, bln消费卡, strKey)
    End If
    
    If objPatiCard Is Nothing Then
        If Not objCard Is Nothing Then
                MsgBox "注意:" & vbCrLf & "调用接口(" & objCard.接口编码 & "-" & objCard.名称 & ")调用失败,请检查!", vbInformation, gstrSysName
        Else
                MsgBox "注意:" & vbCrLf & "调用接口(" & lngCardTypeID & ")调用失败,请检查!", vbInformation, gstrSysName
        End If
        Exit Function
    End If

    Err = 0: On Error Resume Next
    Set objCard = objPatiCard.CardPreporty
    If Err <> 0 Then
        ShowMsgbox "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description
        Call WritLog("clsPatiCard.zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description)
        Exit Function
    End If
    If Not objPatiCard.InitCompents Then
        If Not objPatiCard.CardObject.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then Exit Function
         objPatiCard.InitCompents = True
    End If
    InitInterFacel = True
End Function

Public Function zlOnlyBrushCard(ByVal objEdit As Object, KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡操作(目前只支持有卡进行刷卡)
    '返回:是否刷卡结束后,返回true
    '编制:刘兴洪
    '日期:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean
    Dim strText As String
    '刷卡时含有特殊符号的由调用方取消输入
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then Exit Function
    strText = objEdit.Text
    If objEdit.SelLength = Len(objEdit.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    sngNow = timer
    If objEdit.Text = "" Or strText = "" Then
        sngInputBegin = sngNow
    Else
        If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '用一台笔记本测试，一般在0.014左右
    End If
    If Not blnCard Then
        blnCard = KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1
    End If
    
    If Not blnCard Then
        If gblnTestCardNo Then  '不限制为刷卡
            If KeyAscii = 13 And Trim(objEdit.Text) <> "" Then
                 zlOnlyBrushCard = True
            End If
            Exit Function
        End If
        If KeyAscii <> 8 And KeyAscii <> 13 Then
            objEdit.Text = Chr(KeyAscii): objEdit.SelStart = Len(objEdit)
        Else
            objEdit.Text = ""
        End If
        If KeyAscii <> 13 Then
            KeyAscii = 0:
        End If
        Exit Function
    End If
    If KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1 Then
        zlOnlyBrushCard = True
    End If
End Function
Public Function zlGetCardObj(ByVal frmMain As Object, ByVal lngCardTypeID As Long, _
    Optional bln消费卡 As Boolean = False, _
    Optional ByRef objPatiCardObj As clsCardObject, _
    Optional ByRef blnNotParaCreateObject As Boolean = False, _
    Optional ByVal blnNotStartCreateObject As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡结算对象
    '入参:lngCardTypeID-指定卡类别
    '       bln消费卡-是否消费卡
    '       blnNotParaCreateObject-不根据参数创建对象
    '       blnNotStartCreateObject-为true时，未设置启用的也要创建接口对象, _
    '                                                  为False时, 只有设置了启用才创建对象
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, objCardObject As Object, strKey As String, strExpand As String
    Dim blnOnlyNotObject As Boolean
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    
    If Not objPatiCardObj Is Nothing Then
        If objPatiCardObj.接口序号 = lngCardTypeID And objPatiCardObj.消费卡 = bln消费卡 Then
            If Not objPatiCardObj.InitCompents Then
                If objPatiCardObj.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not objPatiCardObj.CardObject.zlInitComponents(frmMain, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then                    '初始化部件
                    Exit Function
                End If
                objPatiCardObj.InitCompents = True
            End If
            zlGetCardObj = True
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln消费卡, "X", "K") & lngCardTypeID
    '59760
    '检查设备是否启用
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Set objPatiCardObj = objYlCardObjs(strKey)
    If Err <> 0 Or _
        blnNotParaCreateObject And objPatiCardObj.CardObject Is Nothing _
        Or blnNotStartCreateObject Then
          '错误或不根据参数创建对象时,才需要重新创建对象
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!"
                Call WritLog("zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!")
                Exit Function
            End If
            '未启用也要创建对象
            If blnNotStartCreateObject Then objCard.启用 = True
            '重新创建
            If zlCreatePatiCardObject(objCard, objCardObject) = False Then Exit Function
            '增加对应
           Set objPatiCardObj = objYlCardObjs.Add(objCardObject, objCard.自制卡, objCard.接口序号, objCard, bln消费卡, strKey)
    End If
    
    If objPatiCardObj Is Nothing Then
        If Not objCard Is Nothing Then
                MsgBox "注意:" & vbCrLf & "调用接口(" & objCard.接口编码 & "-" & objCard.名称 & ")调用失败,请检查!", vbInformation, gstrSysName
        Else
                MsgBox "注意:" & vbCrLf & "调用接口(" & lngCardTypeID & ")调用失败,请检查!", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    Err = 0: On Error Resume Next
    Set objCard = objPatiCardObj.CardPreporty
    If Err <> 0 Then
        ShowMsgbox "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description
        Call WritLog("clsPatiCard.zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description)
        Exit Function
    End If
    If Not objPatiCardObj.InitCompents Then
        If Not objPatiCardObj.CardObject.zlInitComponents(frmMain, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then Exit Function
         objPatiCardObj.InitCompents = True
    End If
    zlGetCardObj = True
End Function
Public Function zlSelectPayType(ByVal frmMain As Object, ByRef lngCardTypeID As Long, Optional blnNotTheeInterface As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择支付类型
    '出参:lngCardTypeID-卡类别ID
    '       blnNotTheeInterface-不存在三方接口
    '编制:刘兴洪
    '日期:2012-06-11 14:11:20
    '返回:选择成功,返回true,否则返回False
    ''问题:50120
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTypes As String, varData As Variant, strCardTypeIDs As String
    Dim i As Long, varTemp As Variant
    
    strTypes = GetAvailabilityCardType
    '短|全名|读卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    On Error GoTo errHandle
    strCardTypeIDs = ""
    varData = Split(strTypes, ";")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & "|||||||", "|")
        '暂不包含消费卡
        If Val(varTemp(3)) <> 0 And Val(varTemp(5)) = 0 Then strCardTypeIDs = strCardTypeIDs & "," & Val(varTemp(3))
    Next
    If strCardTypeIDs = "" Then blnNotTheeInterface = True: Exit Function
    strCardTypeIDs = Mid(strCardTypeIDs, 2)
    If InStr(strCardTypeIDs, ",") = 0 Then
        '只有一种类别时
        lngCardTypeID = Val(strCardTypeIDs): zlSelectPayType = True: Exit Function
    End If
    '多种类别，选择一种
    If Not frmSelectType.zlSelect(frmMain, strCardTypeIDs, lngCardTypeID, "支付方式选择") Then
      lngCardTypeID = 0: Exit Function
    End If
    If lngCardTypeID = 0 Then Exit Function
    zlSelectPayType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlSelectWriteCardType(ByVal frmMain As Object, ByRef lngCardTypeID As Long, _
    Optional cnOracle As ADODB.Connection, Optional ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择(门诊/住院)写卡类别
    '出参:lngCardTypeID-卡类别ID
    '编制:刘兴洪
    '日期:2012-2-12 15:21:22
    '返回:选择成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardTypeIDs As String
    Dim i As Long, objCard As clsCard
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    For i = 1 To objYlCardObjs.count
        Set objCard = objYlCardObjs(i).CardPreporty
        If objCard.启用 And objCard.是否写卡 And objCard.消费卡 = False Then
            strCardTypeIDs = strCardTypeIDs & "," & objCard.接口序号
        End If
    Next
    If strCardTypeIDs <> "" Then strCardTypeIDs = Mid(strCardTypeIDs, 2)
    strCardTypeIDs = ZLGetPatiCardFromCards(strCardTypeIDs, lng病人ID)
    If strCardTypeIDs = "" Then Exit Function
    If InStr(strCardTypeIDs, ",") = 0 Then
        '只有一种类别时
        lngCardTypeID = Val(strCardTypeIDs): zlSelectWriteCardType = True: Exit Function
    End If
    '多种类别，选择一种
    If Not frmSelectType.zlSelect(frmMain, strCardTypeIDs, lngCardTypeID, "选择写卡类别", cnOracle) Then
      lngCardTypeID = 0: Exit Function
    End If
    If lngCardTypeID = 0 Then Exit Function
    zlSelectWriteCardType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ZLGetPatiCardFromCards(ByVal strCardTypeIDs As String, ByVal lng病人ID As Long) As String
    '从给定卡类别中检索指定病人持有有效卡的卡类别
    '入参:
    '   strCardTypeIDs 给定卡类别，多个用逗号分隔
    '   lng病人ID
    '返回：返回病人持有有效卡的卡类别，多个用逗号分隔
    '问题号：113121
    Dim strCards As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If strCardTypeIDs = "" Or lng病人ID = 0 Then Exit Function
    strSQL = _
        "Select Distinct a.卡类别id" & vbNewLine & _
        "From 病人医疗卡信息 A" & vbNewLine & _
        "Where a.病人id = [1] And Nvl(a.状态, 0) = 0" & vbNewLine & _
        "      And 卡类别id In (Select /*+cardinality(j,10)*/  j.Column_Value From Table(f_Num2list([2])) J)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatiCard", lng病人ID, strCardTypeIDs)
    Do While Not rsTemp.EOF
        strCards = strCards & "," & Val(nvl(rsTemp!卡类别id))
        rsTemp.MoveNext
    Loop
    If strCards <> "" Then strCards = Mid(strCards, 2)
    ZLGetPatiCardFromCards = strCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardProperty(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As clsCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡片属性
    '入参:lngCardTypeID-卡类别ID
    '       bln消费卡-是否消费卡
    '出参:objCard-卡对象
    '返回:存在指定卡类别ID的，返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-01-17 16:07:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As clsCards, objTemp As clsCard
    On Error GoTo errHandle
    If zlGetCards_YL(objCards) = False Then Exit Function
    If objCards Is Nothing Then Exit Function
    For Each objTemp In objCards
        If objTemp.接口序号 = lngCardTypeID And objTemp.消费卡 = bln消费卡 Then
            Set objCard = objTemp
            zlGetCardProperty = True: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

