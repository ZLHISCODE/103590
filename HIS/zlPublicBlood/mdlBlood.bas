Attribute VB_Name = "mdlBlood"
Option Explicit
Public gobjCardSquare As Object '一卡通对象
Public gobjPublicExpense As Object '费用公共对象
Public gobjRegister As Object          '注册授权部件zlRegister

Private mstrSQL As String
Public Enum COLOR
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    
    原始单据 = 0
    冲销记录 = &HFF
    停用项目 = &H8000000C
    启用项目 = 0
    
    公共模块色 = &HC00000
    
    
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    
End Enum

Public Enum Enum_Inside_Program
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p医嘱附费管理 = 1257
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    P新版护士站 = 1265
    p输血审核管理 = 1268
    p输血反应管理 = 1938
    p血液接收登记 = 1910
End Enum

Public Function GetObjectRegister() As Boolean
'创建注册授权部件zlRegister
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
    
        If gobjRegister Is Nothing Then
            Set gobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If gobjRegister Is Nothing Then
                MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    GetObjectRegister = True
End Function

Public Function InitObjPublicExpense(ByVal lngSys As Long) As Boolean
    If gobjPublicExpense Is Nothing Then
        On Error Resume Next
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Not gobjPublicExpense Is Nothing Then
            Call gobjPublicExpense.zlInitCommon(lngSys, gcnOracle, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    InitObjPublicExpense = Not gobjPublicExpense Is Nothing
End Function

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjCardSquare Is Nothing Then
        Set gobjCardSquare = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjCardSquare.zlInitComponents(frmMain, lngModule, lngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjCardSquare Is Nothing Then Exit Sub
    If Not gobjCardSquare Is Nothing Then
         Set gobjCardSquare = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
End Sub



Public Function GetDeptList(ByVal str工作性质 As String, Optional ByVal int服务对象 As Integer = -1, Optional ByVal blnShowAll As Boolean = True) As ADODB.Recordset
        '******************************************************************************************************************
    '功能：读取部门选择
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim bytService(3) As Byte
    
    Select Case int服务对象
    Case -1         '不判断
        bytService(0) = 0
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    Case 0          '不服务于病人
        bytService(0) = 0
        bytService(1) = 0
        bytService(2) = 0
        bytService(3) = 0
    Case 1          '服务于门诊病人
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 9
        bytService(3) = 3
    Case 2          '服务于住院病人
        bytService(0) = 9
        bytService(1) = 9
        bytService(2) = 2
        bytService(3) = 3
    Case 3          '服务于门诊和住院病人
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    End Select
        
    If blnShowAll Then '所有部门
        mstrSQL = "SELECT Distinct A.简码,A.名称, A.ID,A.编码 FROM 部门表 A,部门性质说明 B WHERE B.服务对象 In ([2],[3],[4],[5]) And (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质=[1] ORDER BY A.编码"
        Set GetDeptList = gobjDatabase.OpenSQLRecord(mstrSQL, "获取部门列表", str工作性质, bytService(0), bytService(1), bytService(2), bytService(3))
    Else
        mstrSQL = _
            " Select Distinct a.简码, a.名称, a.Id, a.编码, c.缺省" & vbNewLine & _
            " From 部门表 a, 部门性质说明 b, 部门人员 c" & vbNewLine & _
            " Where (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And a.Id = b.部门id And b.工作性质 = [1] And" & vbNewLine & _
            "      b.服务对象 In ([3],[4],[5],[6]) And a.Id = c.部门id And c.人员id = [2]" & vbNewLine & _
            " Order By a.简码 || '-' || a.名称"
        Set GetDeptList = gobjDatabase.OpenSQLRecord(mstrSQL, "获取部门列表", str工作性质, UserInfo.id, bytService(0), bytService(1), bytService(2), bytService(3))
    End If

End Function

Public Function GetPatientOtherInfo(ByVal lng病人ID As Long, ByVal str信息名 As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：获取病人信息从表内容
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************

    mstrSQL = "Select 病人id,信息名,信息值 From 病人信息从表 Where 病人id=[1] And 信息名=[2]"
    Set GetPatientOtherInfo = gobjDatabase.OpenSQLRecord(mstrSQL, "病人信息从表", lng病人ID, str信息名)
End Function


Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：工具栏中的一些功能：大图标，标准按钮，状态栏等
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
    
    Select Case Control.id
        Case conMenu_View_ToolBar_Button '工具栏
        
            For lngLoop = 2 To frmMain.cbsMain.Count
                frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
            Next
            frmMain.cbsMain.RecalcLayout
            
        Case conMenu_View_ToolBar_Text '按钮文字
        
            For lngLoop = 2 To frmMain.cbsMain.Count
                For Each objControl In frmMain.cbsMain(lngLoop).Controls
                    If objControl.Type = xtpControlButton Then
                        objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                    End If
                Next
            Next
            frmMain.cbsMain.RecalcLayout
            
        Case conMenu_View_ToolBar_Size '大图标
        
            frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
            frmMain.cbsMain.RecalcLayout
            
        Case conMenu_View_StatusBar '状态栏
        
            frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
            frmMain.cbsMain.RecalcLayout
                
    End Select
    CommandBarExecutePublic = True
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '******************************************************************************************************************
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(gobjDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(gobjDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case Else
        If strMode = Val(strMode) Then
            If bytFlag = 1 Then
                GetDateTime = Format(DateAdd("d", -Val(strMode), CDate(Format(gobjDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
            Else
                GetDateTime = Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
            End If
        End If
    End Select
    
End Function

Public Function GetDepartPeople(ByVal str工作性质 As String, Optional ByVal int服务对象 As Integer = -1, Optional ByVal blnShowAll As Boolean = True)
Dim bytService(3) As Byte
    
    Select Case int服务对象
    Case -1         '不判断
        bytService(0) = 0
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    Case 0          '不服务于病人
        bytService(0) = 0
        bytService(1) = 0
        bytService(2) = 0
        bytService(3) = 0
    Case 1          '服务于门诊病人
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 9
        bytService(3) = 9
    Case 2          '服务于住院病人
        bytService(0) = 9
        bytService(1) = 9
        bytService(2) = 2
        bytService(3) = 9
    Case 3          '服务于门诊和住院病人
        bytService(0) = 9
        bytService(1) = 1
        bytService(2) = 2
        bytService(3) = 3
    End Select
    
    If blnShowAll Then
        mstrSQL = " Select Distinct d.姓名 " & _
                  " From 部门性质说明 a, 部门表 b, 部门人员 c, 人员表 d " & _
                  " Where a.工作性质 = [1] And a.部门id = b.Id And c.部门id = b.Id And c.人员id = d.Id and a.服务对象 in([2],[3],[4],[5])"
        Set GetDepartPeople = gobjDatabase.OpenSQLRecord(mstrSQL, "获取部门人员信息", str工作性质, bytService(0), bytService(1), bytService(2), bytService(3))
    Else
        mstrSQL = "select 姓名 from 人员表 where id=[1]"
        Set GetDepartPeople = gobjDatabase.OpenSQLRecord(mstrSQL, "获取部门人员信息", UserInfo.id)
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'医嘱操作相关
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function zlBloodInstantRptPrint(ByVal objFrm As Object, ByVal lngActiveID As Long) As Boolean
'功能：供护士站或医技工作站调用(输血执行单打印)
'参数： objFrm--调用主窗体
'           lngActiveID--医嘱ID
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = _
        " Select  c.操作类型, c.执行分类,b.医嘱状态" & vbNewLine & _
        " From 诊疗项目目录 c, 病人医嘱记录 a, 病人医嘱记录 b" & vbNewLine & _
        " Where c.Id = a.诊疗项目id And a.相关id = b.Id And a.诊疗类别 = 'E' And b.Id = [1] And b.诊疗类别 = 'K'"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "zlBloodInstantRptPrint", lngActiveID)
    If rsTmp.EOF Then
        MsgBox "选中的医嘱并非输血医嘱，请选择输血医嘱！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not (Val("" & rsTmp!操作类型) = 8 And Val("" & rsTmp!执行分类) = 1) Then
        MsgBox "选中的医嘱并非输血类的用血医嘱，请选择用血医嘱！", vbInformation, gstrSysName
        Exit Function
    End If
    zlBloodInstantRptPrint = frmBloodInstantRptPrint.ShowMe(objFrm, lngActiveID)
End Function

Public Function zlAdviceOperation(ByVal lngMoudle As Long, ByVal lng医嘱ID As Long, ByVal intOperation As Enum_Advice, Optional ByVal blnMoved As Boolean = False, _
        Optional ByRef strErrInfo As String = "") As Boolean
'功能：医嘱操作调用接口（新开、删除、发送、回退时此方法的调用请放在医嘱操作事物中调用，修改、校对、作废为操作校验检查，放在事物之前）
'入参:
'       lngMoudle:调用模块号
'       lng医嘱ID:血液医嘱主医嘱ID
'       intOperation:医嘱操作类型(枚举),含：新开、修改、删除、校对、作废、发送、回退
'       blnMoved:病人历史数据是否转出
'出参：
'       strErrInfo：接口返回FALSE时的信息
'返回：成功=TRUE，失败=False
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    '医嘱内容变量
    Dim int病人来源 As Integer, lng病人ID As Long, lng主页id As Long, lng执行科室ID As Long, lng相关ID As Long
    Dim int审核状态 As Long, str检查方法 As String, int操作类型 As Integer, int执行分类 As Integer, int医嘱状态 As Integer
    Dim bln用血 As Boolean
    
    On Error GoTo ErrHand
    If blnMoved = True Then
        strErrInfo = "病人的数据已经转出到后备数据库，不允许操作。" & vbCrLf & "您可以与系统管理员联系，将相应数据抽选返回。"
        Exit Function
    End If
    strSQL = _
        " Select a.id 相关ID,b.病人来源, b.病人id, b.主页id, b.执行科室id, b.审核状态, b.检查方法, c.操作类型, c.执行分类,B.医嘱状态" & vbNewLine & _
        " From 诊疗项目目录 c, 病人医嘱记录 a, 病人医嘱记录 b" & vbNewLine & _
        " Where c.Id = a.诊疗项目id And a.相关id = b.Id And a.诊疗类别 = 'E' And b.Id = [1] And b.诊疗类别 = 'K'"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "zlAdviceOperation", lng医嘱ID)
    '血液医嘱肯定查得到数据，查不到数据则退出
    If rsTmp.EOF Then
        zlAdviceOperation = True
        Exit Function
    End If
    
    lng相关ID = Val("" & rsTmp!相关ID)
    int病人来源 = Val("" & rsTmp!病人来源)
    lng病人ID = Val("" & rsTmp!病人id)
    lng主页id = Val("" & rsTmp!主页id)
    lng执行科室ID = Val("" & rsTmp!执行科室ID)
    int审核状态 = Val("" & rsTmp!审核状态)
    str检查方法 = "" & rsTmp!检查方法
    int操作类型 = Val("" & rsTmp!操作类型)
    int执行分类 = Val("" & rsTmp!执行分类)
    int医嘱状态 = Val("" & rsTmp!医嘱状态)
    If str检查方法 = "" Then
        If int操作类型 = "8" And int执行分类 = 1 Then
            bln用血 = True
        End If
    Else
        bln用血 = Val(str检查方法) = 1
    End If
    Select Case intOperation
        Case Advice_新开
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If bln用血 = True And gbln医嘱发送后发血 = False Then
                strSQL = "Zl_血液医嘱记录_Insert(" & lng医嘱ID & "," & lng病人ID & "," & IIf(lng主页id = 0, "NULL", lng主页id) & "," & int病人来源 & "," & lng执行科室ID & "," & 2 & ")"
                Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Insert")
            End If
        Case Advice_修改
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            
            If int审核状态 = 5 Or int审核状态 = 2 Then
                strErrInfo = "该医嘱目前输血科已经接收，不允许对医嘱进行操作！"
                Exit Function
            End If
            
            If bln用血 = True And int医嘱状态 = 1 Then  '医嘱还是新开状态（则修改血液配血记录信息）
                '检查是否存在配血记录
                strSQL = "select id From  血液配血记录 where 申请ID=[1]"
                Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "zlAdviceOperation", lng医嘱ID)
                If Not rsData.EOF Then
                    strSQL = "Zl_血液医嘱记录_Insert(" & lng医嘱ID & "," & lng病人ID & "," & IIf(lng主页id = 0, "NULL", lng主页id) & "," & int病人来源 & "," & lng执行科室ID & "," & 2 & ")"
                    Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Insert")
                End If
            End If
        Case Advice_删除
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int审核状态 = 5 Or int审核状态 = 2 Then
                strErrInfo = "该医嘱目前输血科已经接收，不允许对医嘱进行操作！"
                Exit Function
            End If
            strSQL = "Zl_血液医嘱记录_Delete(" & lng医嘱ID & "," & IIf(bln用血 = False, 1, 2) & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Delete")
        Case Advice_校对
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int审核状态 = 5 Or int审核状态 = 2 Then
                zlAdviceOperation = True
                Exit Function
            End If
            '医嘱校对直接发送的情况
            If int医嘱状态 = 8 Then
                If bln用血 = False Or (bln用血 = True And gbln医嘱发送后发血 = True) Then
                    strSQL = "Zl_血液医嘱记录_Insert(" & lng医嘱ID & "," & lng病人ID & "," & IIf(lng主页id = 0, "NULL", lng主页id) & "," & int病人来源 & "," & lng执行科室ID & "," & 2 & ")"
                    Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Insert")
                End If
            End If
        Case Advice_作废
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int审核状态 = 5 Or int审核状态 = 2 Then
                strErrInfo = "该医嘱目前输血科已经接收，不允许对医嘱进行操作！"
                Exit Function
            End If
            '门诊病人作废删除配血信息，住院病人备血医嘱回退删除、用血医嘱存在作废删除和回退删除的情况(根据参数：gbln医嘱发送后发血决定)
            strSQL = "Zl_血液医嘱记录_Delete(" & lng医嘱ID & "," & IIf(bln用血 = False, 1, 2) & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Delete")
        Case Advice_回退作废 '住院病人医嘱作废可以回退
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            '主要针住院病人的用血医嘱，因为备血医嘱的发送后才产生配血信息
            If int病人来源 = 2 Then
                If bln用血 = True And gbln医嘱发送后发血 = False Then
                    strSQL = "Zl_血液医嘱记录_Insert(" & lng医嘱ID & "," & lng病人ID & "," & IIf(lng主页id = 0, "NULL", lng主页id) & "," & int病人来源 & "," & lng执行科室ID & "," & 2 & ")"
                    Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Insert")
                End If
            End If
        Case Advice_发送
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            strSQL = "Zl_血液医嘱记录_Insert(" & lng医嘱ID & "," & lng病人ID & "," & IIf(lng主页id = 0, "NULL", lng主页id) & "," & int病人来源 & "," & lng执行科室ID & "," & IIf(bln用血 = True, 2, 1) & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Insert")
        Case Advice_回退
            '老的用血医嘱不进行任何处理
            If bln用血 = True And str检查方法 = "" Then
                zlAdviceOperation = True
                Exit Function
            End If
            If int审核状态 = 5 Or int审核状态 = 2 Then
                strErrInfo = "该医嘱目前输血科已经接收，不允许对医嘱进行操作！"
                Exit Function
            End If
            If bln用血 = False Or (bln用血 = True And gbln医嘱发送后发血 = True) Then
                strSQL = "Zl_血液医嘱记录_Delete(" & lng医嘱ID & "," & IIf(bln用血 = False, 1, 2) & ")"
                Call gobjDatabase.ExecuteProcedure(strSQL, "Zl_血液医嘱记录_Delete")
            End If
    End Select
    zlAdviceOperation = True
    Exit Function
ErrHand:
    If gcnOracle.Errors.Count <> 0 Then
        strErrInfo = gcnOracle.Errors(0).Description
        If InStr(UCase(strErrInfo), "[ZLSOFT]") > 0 Then
            strErrInfo = Split(strErrInfo, "[ZLSOFT]")(1)
        End If
    Else
        strErrInfo = Err.Description
    End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'医嘱执行相关
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ItemCanCancel(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal lng组ID As Long, str诊疗类别 As String, _
    ByVal bln单独执行 As Boolean, ByVal blnMove As Boolean, ByVal byt来源 As Byte) As Boolean
'功能：判断指定项目是否可以取消执行
'参数：byt来源=1:门诊，2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If gbytBillOpt = 0 Then ItemCanCancel = True: Exit Function
    
    On Error GoTo errH
    
    If bln单独执行 Then
        strSQL = _
            " Select Distinct NO From 病人医嘱发送 Where 记录性质=2 And 医嘱ID=[1] And 发送号=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From 病人医嘱附费 Where 记录性质=2 And 医嘱ID=[1] And 发送号=[2]"
    Else
        strSQL = _
            " Select Distinct NO From 病人医嘱发送 Where 记录性质=2 And 医嘱ID=[1] And 发送号=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From 病人医嘱附费 Where 记录性质=2 And 发送号=[2]" & _
            " And 医嘱ID IN(Select ID From 病人医嘱记录 Where (ID=[3] Or 相关ID=[3]) And 诊疗类别=[4])"
    End If
    If blnMove Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "ItemCanCancel", lng医嘱ID, lng发送号, lng组ID, str诊疗类别)
    
    Do While Not rsTmp.EOF
        '处理中排开了结帐金额为0的，即零耗费用登记
        If HaveBilling(rsTmp!NO, True, "", IIf(bln单独执行, lng医嘱ID, 0), byt来源) <> 0 Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("该项目包含已经结帐的费用,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    MsgBox "该项目包含已经结帐的费用,操作不能继续。", vbExclamation, gstrSysName
                    Exit Function
            End Select
        End If
        rsTmp.MoveNext
    Loop
    ItemCanCancel = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function HaveBilling(ByVal strNO As String, ByVal blnALL As Boolean, _
     ByVal strTime As String, ByVal lng医嘱ID As Long, ByVal byt来源 As Byte) As Integer
'功能：判断一张记帐单/表是否已经结帐
'参数：strNO=记帐单据号,不分门诊及住院
'      blnALL=是否对整张单据内容进行判断,否则只对未销帐部分进行判断(销帐时)
'      byt来源=1:门诊，2-住院
'返回：0-未结帐,1=已全部结帐,2-已部分结帐
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    Dim strTab As String
    
    On Error GoTo errH
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
        
    '求未作废的费用行
    strSQL = _
        " Select 序号 From (" & _
        " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        " Avg(Nvl(付数, 1) * 数次) As 数量" & _
        " From " & strTab & "" & _
        " Where NO=[1] And 记录性质=2" & _
        " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    
    '求每行的结帐情况
    strSQL = _
        "Select Nvl(价格父号,序号) as 序号,Sum(Nvl(结帐金额,0)) as 结帐金额" & _
        " From " & strTab & "" & _
        " Where NO=[1] And 记录性质 IN(2,12)" & _
        IIf(Not blnALL, " And Nvl(价格父号,序号) IN(" & strSQL & ")", "") & _
        IIf(strTime <> "", " And 登记时间=[2]", "") & _
        IIf(lng医嘱ID <> 0, " And 医嘱序号+0=[3]", "") & _
        " Group by Nvl(价格父号,序号)"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "HaveBilling", strNO, CDate(IIf(strTime = "", "1990-01-01", strTime)), lng医嘱ID)
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '单据行数
        rsTmp.Filter = "结帐金额<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '无结帐行
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 '全部行已结帐
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '部分行已结帐
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function ItemHaveCash(ByVal int病人来源 As Integer, ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, _
    ByVal lng发送号 As Long, ByVal str类别 As String, ByVal str单据号 As String, ByVal int记录性质 As Integer, ByVal int门诊记帐 As Integer, ByVal int方式 As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByVal dat发送时间 As Date, Optional ByRef str医嘱IDs As String, Optional ByRef strNOs As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'功能：判断当前的执行医嘱是否已收费或记帐划价单是否已审核
'参数：int病人来源=1-门诊,2-住院
'      str类别=诊疗类别，用于从一组医嘱中区分分开执行的内容
'      int方式=0-检查是否存在未收费记录
'              1-检查是否存在已收费记录
'      int门诊记帐=1=住院发送到门诊记帐
'      返回：str医嘱IDs=该医嘱及相关的医嘱ID,NOs=医嘱发送的单据号和补的附费中的单据号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTab As String
    
    If int病人来源 = 2 And int记录性质 = 2 And int门诊记帐 = 0 Then
        strTab = "住院费用记录"
    Else
        strTab = "门诊费用记录"
    End If
    ItemHaveCash = True
    str医嘱IDs = ""
    strNOs = ""
    
    '对应的费用中是否存在未收费[或已作废]的内容
    '和清单只显示已收费不同：
    '1.检查了医嘱附费(不加记录性质的条件，因为可能补收费单或记帐单)
    '2.记帐划价也显示为未收(清单需要先显出来执行后审核)
    '3.按NO对应到相关医嘱的费用检查(清单是按显示的医嘱ID)
    strSQL = _
        " Select A.记录状态,Nvl(B.相关ID,B.ID) as 医嘱ID,B.诊疗类别,A.执行状态,A.NO" & IIf(strTab = "住院费用记录", ",0 as 费用状态", ",NVL(A.费用状态,0) as 费用状态") & _
        " From " & strTab & " A,病人医嘱记录 B" & _
        " Where A.NO=[4] And A.记录状态 IN(0,1,3) And A.医嘱序号+0=B.ID And A.记录性质=[5]" & IIf(bln单独执行, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select B.记录状态,Nvl(C.相关ID,C.ID) as 医嘱ID,C.诊疗类别,B.执行状态,A.NO" & IIf(strTab = "住院费用记录", ",0 as 费用状态", ",NVL(b.费用状态,0) as 费用状态") & _
        " From 病人医嘱记录 C," & strTab & " B,病人医嘱附费 A" & _
        " Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱ID=B.医嘱序号+0" & IIf(bln单独执行, " And A.医嘱ID=[2]", _
            " And A.医嘱ID IN (Select ID From 病人医嘱记录 Where (ID=[1] Or 相关ID=[1]) And 诊疗类别=[6])") & _
        " And A.发送号=[3] And B.记录状态 IN(0,1,3) And A.医嘱ID=C.ID And A.记录性质=[5]"
    If blnMove Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
        strSQL = Replace(strSQL, strTab, "H" & strTab)
    ElseIf gobjDatabase.DateMoved(dat发送时间) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID), lng医嘱ID, lng发送号, str单据号, int记录性质, str类别)
    If Not rsTmp.EOF Then
        If int方式 = 0 Then
            rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 费用状态=1"
            If Not rsTmp.EOF Then
                blnIsAbnormal = True
                ItemHaveCash = False
            Else
                rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态=0"
                If Not rsTmp.EOF Then ItemHaveCash = False
            End If
            
            While Not rsTmp.EOF
                If InStr("," & str医嘱IDs & ",", "," & rsTmp!医嘱ID & ",") = 0 Then
                    str医嘱IDs = str医嘱IDs & "," & rsTmp!医嘱ID
                End If
                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNOs = strNOs & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Wend
            strNOs = Mid(strNOs, 2)
            str医嘱IDs = Mid(str医嘱IDs, 2)
        ElseIf int方式 = 1 Then
            rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态<>1 And 费用状态<>1"
            If Not rsTmp.EOF Then ItemHaveCash = False
        End If
    ElseIf int方式 = 1 Then
        ItemHaveCash = False
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAdviceMoney(ByVal str组ID As String, ByVal str医嘱ID As String, ByVal str发送号 As String, _
    str类别 As String, str类别名 As String, ByVal bln单独执行 As Boolean, ByVal byt来源 As Byte) As Currency
'功能：根据指定的医嘱ID串，获取医嘱对应未审核的记帐费用合计
'参数：str组ID,str医嘱ID,str发送号="ID1,ID2,..."
'      bln单独执行=检验项目单独执行，这时只有一个医嘱ID
'      byt来源，1:门诊，2-住院
'返回：str类别,str类别名=用于报警提示
'说明：当系统参数为执行后审核费用时才返回。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, curMoney As Currency
    Dim strTab As String
    
    str类别 = "": str类别名 = ""
    
    On Error GoTo errH
     
    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
    
    If bln单独执行 Then
        strSQL = _
            " Select B.编码,B.名称,Sum(A.实收金额) as 金额" & _
            " From " & strTab & " A,收费项目类别 B" & _
            " Where A.医嘱序号 + 0 = [2] And (A.记录性质, A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费 Where 医嘱id = [2] And 发送号 + 0 = [3]" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = [2] And 发送号 + 0 = [3])" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
            " Group by B.编码,B.名称"
    Else
        strSQL = _
            " Select /*+ RULE */ B.编码,B.名称,Sum(A.实收金额) as 金额" & _
            " From " & strTab & " A,收费项目类别 B" & _
            " Where A.医嘱序号 + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From 病人医嘱记录" & _
            "       Where 相关id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.记录性质, A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费" & _
            "       Where 医嘱id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From 病人医嘱记录" & _
                "       Where 相关id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And 发送号 + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送" & _
            "       Where 医嘱id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And 发送号 + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
            " Group by B.编码,B.名称"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str组ID, str医嘱ID, str发送号)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!金额, 0)
        str类别 = str类别 & rsTmp!编码
        str类别名 = str类别名 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    str类别名 = Mid(str类别名, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function PatiCanBilling(ByVal lng病人ID As Long, ByVal lng主页id As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
'功能：检查指定病人是否具有相关权限
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "出院未结强制记帐") > 0 And InStr(strPrivs, "出院结清强制记帐") > 0 Then
        Exit Function
    End If
    On Error GoTo errH
    strSQL = "Select NVL(B.姓名,A.姓名) 姓名,B.出院日期,B.状态,X.费用余额" & _
        " From 病人信息 A,病案主页 B,病人余额 X" & _
        " Where A.病人ID=B.病人ID And A.病人ID=X.病人ID(+) And X.类型(+) = 2" & _
        " And A.病人ID=[1] And B.主页ID=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页id)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!出院日期) And Nvl(rsTmp!状态, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "出院未结强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) <> 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用未结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If InStr(strPrivs, "出院结清强制记帐") = 0 Then
            If Nvl(rsTmp!费用余额, 0) = 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用已结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If lngModual = p医嘱附费管理 Or lngModual = p住院医嘱发送 Or lngModual = p住院医嘱下达 Then
            '68081不允许出院病人处理医嘱费用
            strMsg = """" & rsTmp!姓名 & """已经出院(或预出院)，不能对该病人的医嘱进行发送、超期收回、执行、回退。"
        End If
        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function IsBloodMessageDone(ByVal intMode As Integer, ByVal lng病人ID As Long, ByVal lng就诊id As Long, _
                                    ByVal int阅读场合 As Integer, ByVal lng阅读部门id As Long) As Boolean
'参数：intMode  1-血袋是否已经回收；2-血袋是否填写输血反应
'功能：查询医护站的血库相关消息是否进行了后续操作
    Dim rsTmp As New ADODB.Recordset, rsMsg As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lng收发id As Long, str收发ids As String
    Dim arr() As String
    Dim blnTrans As Boolean, arrSQL As Variant
    On Error GoTo errH
    arrSQL = Array()
    Select Case intMode
        Case 1
            strSQL = "select id,业务标识 from 业务消息清单 where 类型编码 = [1] and 病人ID = [2] and 就诊id = [3] and 是否已阅 = 0 "
            Set rsMsg = gobjDatabase.OpenSQLRecord(strSQL, "消息状态", "ZLHIS_BLOOD_007", lng病人ID, lng就诊id)
            Do While Not rsMsg.EOF
                arr = Split(rsMsg!业务标识, ":")
                If UBound(arr) > 0 Then
                    str收发ids = str收发ids & ":" & Val(arr(2))
                End If
                rsMsg.MoveNext
            Loop
            strSQL = "select /*+ CARDINALITY(c,10) */" & vbNewLine & _
                    "       a.id, b.记录状态 from 血液收发记录 a ,血液配血回收 b,table(f_str2list([1],':')) c" & vbNewLine & _
                    "       where a.id = b.收发id and b.血袋编号 = a.血袋编号 and a.id = c.column_value"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "回收状态", Mid(str收发ids, 2, Len(str收发ids)))
            
                rsTmp.Filter = "记录状态 = '1'"
            If rsTmp.RecordCount = 0 Then
                strSQL = "zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_007'," _
                                                & int阅读场合 & ",'" & UserInfo.姓名 & "'," & lng阅读部门id & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                IsBloodMessageDone = True
                Else
                '将记录状态为1的消息的收发id从串中剔除
                rsTmp.MoveFirst
                str收发ids = str收发ids & ":"
                Do While Not rsTmp.EOF
                    str收发ids = Replace(str收发ids, ":" & rsTmp!id & ":", ":")
                    rsTmp.MoveNext
                Loop
                '遍历存在血液回收的消息，收发id串中存放的为不为记录状态不为1的相关血袋的收发id，将这部分挨个设为已读
                rsMsg.MoveFirst
                Do While Not rsMsg.EOF
                    rsTmp.MoveFirst
                        If InStr(str收发ids, Mid(rsMsg("业务标识"), InStr(rsMsg("业务标识"), ":"))) > 0 Then
                            strSQL = "zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_007'," _
                                                    & int阅读场合 & ",'" & UserInfo.姓名 & "'," & lng阅读部门id & ",null," & rsMsg("ID") & " ,null)"
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = strSQL
                End If
                    rsMsg.MoveNext
                Loop
                IsBloodMessageDone = False
            End If
        Case 2
            strSQL = "select id,业务标识 from 业务消息清单 where 类型编码 = [1] and 病人ID = [2] and 就诊id = [3] and 是否已阅 = 0 "
            Set rsMsg = gobjDatabase.OpenSQLRecord(strSQL, "消息状态", "ZLHIS_BLOOD_006", lng病人ID, lng就诊id)
            Do While Not rsMsg.EOF
                arr = Split(rsMsg!业务标识, ":")
                If UBound(arr) > 0 Then
                    str收发ids = str收发ids & ":" & Val(arr(1))
                End If
                rsMsg.MoveNext
            Loop
            If str收发ids = "" Then IsBloodMessageDone = False: Exit Function
            strSQL = "SELECT /*+ CARDINALITY(b,10) */ a.收发id, a.有无输血反应 FROM 输血反应记录 a, TABLE(f_Str2list([1], ':')) b " _
                    & "WHERE a.收发id = b.Column_Value"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "输血反应记录", Mid(str收发ids, 2, Len(str收发ids)))
            
            If rsTmp.RecordCount = rsMsg.RecordCount Then           '每条未读消息都有对应的输血反应记录，表示都已填写，则将所有未读消息设为已读
                strSQL = "zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_006'," _
                                                & int阅读场合 & ",'" & UserInfo.姓名 & "'," & lng阅读部门id & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                IsBloodMessageDone = True
            Else            '遍历存在输血反应记录的血袋，对应的消息设为已读
                If rsTmp.RecordCount <> 0 Then
                    rsMsg.MoveFirst
                    Do While Not rsMsg.EOF
                        rsTmp.MoveFirst
                        Do While Not rsTmp.EOF
                            If InStr(rsMsg("业务标识") & ":", ":" & rsTmp("收发id") & ":") > 0 Then
                                strSQL = "zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_007'," _
                                                    & int阅读场合 & ",'" & UserInfo.姓名 & "'," & lng阅读部门id & ",null," & rsMsg("ID") & " ,null)"
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = strSQL
                                Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsMsg.MoveNext
                    Loop
                End If
                IsBloodMessageDone = False
            End If
    End Select
    If UBound(arrSQL) < 0 Then Exit Function
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), "血液相关消息处理")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Function
errH:
    IsBloodMessageDone = False
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Public Function GetReactionTips(lng库房id As Long) As Recordset
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "select x.消息内容,x.业务标识 收发id,x.病人id,x.就诊id,x.病人来源 from 业务消息清单 x, 业务消息提醒部门 b" & vbNewLine & _
                "where x.类型编码 = 'ZLHIS_BLOOD_008' and x.id = b.消息id and x.是否已阅 = 0 "
    If lng库房id > 0 Then
        strSQL = strSQL & " and b.部门id = [1] "
    Else
        strSQL = strSQL & "and b.部门id in (SELECT Distinct A.ID FROM 部门表 A,部门性质说明 B" & vbNewLine & _
                        "WHERE B.服务对象 In (9,1,2,3) And (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID  AND B.工作性质='血库')"
    End If
    strSQL = strSQL & " Order by x.登记时间 desc "
    Set rs = gobjDatabase.OpenSQLRecord(strSQL, "输血反应消息", lng库房id)
    Set GetReactionTips = rs
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
