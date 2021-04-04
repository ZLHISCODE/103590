VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************************************************************
'功能：药品及卫生材料相关处理
'   *CreateDrugAndStuffSvr:创建药品卫材服务公共部件
'   *CreateDrugPacker:创建自助发药机(自动化药房)
'   *GetDefaultWindow:获取缺省的药房窗口设置
'   DrugMachine_Charge:收费或记帐后发药机数据上传
'   DrugMachine_Del:退费或销账后发药机数据上传
'   Get发药窗口:分配发药窗口
'   InitCommon:初始化相关基础数据
'*************************************************************************************************************************************************

Private mlngModule As Long, mstrPrivs As String

Private mobjService As Object 'zlPublicExpense.clsService

Private mobjDrugPacker  As Object ' 自动发药机(更新发药窗口)
Private mobjDrugMachine As Object '自动发药机(新）

Private mcllWindows As Collection 'array(药房ID,窗口)
Private mstr西窗 As String, mstr成窗 As String, mstr中窗 As String '记录门诊病人连续收费的窗口分配
Private mlng西药房 As Long, mlng成药房 As Long, mlng中药房 As Long '记录门诊病人连续收费的药房分配

Public Function InitCommon(ByVal lngModule As Long, strPrivs As String, ByVal blnCliniqueRoomPay As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关基础数据
    '入参:
    '   blnCliniqueRoomPay 是否诊间无窗口支付
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    mlngModule = lngModule: mstrPrivs = strPrivs
    If CreateDrugAndStuffSvr(lngModule) = False Then Exit Function
    
    If blnCliniqueRoomPay Then
        '药房、窗口分配方式
        mstr中窗 = zlDatabase.GetPara(49, glngSys, lngModule)
        mstr西窗 = zlDatabase.GetPara(50, glngSys, lngModule)
        mstr成窗 = zlDatabase.GetPara(51, glngSys, lngModule)
        
        mlng西药房 = Val(zlDatabase.GetPara(18, glngSys, lngModule))
        mlng成药房 = Val(zlDatabase.GetPara(19, glngSys, lngModule))
        mlng中药房 = Val(zlDatabase.GetPara(20, glngSys, lngModule))
    Else
        mstr西窗 = "": mstr中窗 = "": mstr成窗 = ""
        mlng中药房 = 0: mlng西药房 = 0: mlng成药房 = 0
    End If
    
    InitCommon = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateDrugAndStuffSvr(ByVal lngMoudle As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建药品卫材服务公共部件
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjService Is Nothing Then
        Set mobjService = CreateObject("zlPublicExpense.clsService")
        If mobjService Is Nothing Then
            MsgBox "费用公共部件(zlPublicExpense.clsService)创建失败，请与系统管理员联系！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    'zlInitCommon(ByVal lngSys As Long, ByVal lngMoudle As Long, ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String, _
        Optional objService As Object, Optional objStuffService As Object)
    If mobjService.zlInitCommon(glngSys, lngMoudle, gcnOracle, gstrDBUser) = False Then
        MsgBox "费用公共部件(zlPublicExpense.clsService)初始化失败，请与系统管理员联系！", vbInformation, gstrSysName
        Set mobjService = Nothing: Exit Function
    End If
    
    CreateDrugAndStuffSvr = True
End Function

Private Function CreateDrugPacker() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建自助发药机(自动化药房)
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String, strMessage As String
    Dim blnSuccess As Boolean

    If Not mobjDrugMachine Is Nothing Or Not mobjDrugPacker Is Nothing Then CreateDrugPacker = True: Exit Function

    On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err <> 0 Then Set mobjDrugMachine = Nothing
        
        On Error GoTo 0
        If Not mobjDrugMachine Is Nothing Then
            '权限检查
            strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
            If zlstr.IsHavePrivs(strPrivs, "基本") = False Then Set mobjDrugMachine = Nothing: Exit Function
            blnSuccess = mobjDrugMachine.Init(1, objComLib, strMessage)
            If blnSuccess = False Then Set mobjDrugMachine = Nothing: Exit Function
            CreateDrugPacker = True: Exit Function
        End If
    End If
    
    '旧部件
    On Error Resume Next
    Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
    If Err <> 0 Then Set mobjDrugMachine = Nothing
    If mobjDrugPacker Is Nothing Then Exit Function
    
    On Error GoTo 0
    blnSuccess = mobjDrugPacker.DYEY_MZ_IniSoap()
    If blnSuccess = False Then Set mobjDrugPacker = Nothing: Exit Function
    
    CreateDrugPacker = True
End Function

Public Sub DrugMachine_Charge(ByVal bytBillType As Byte, ByVal strNOs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费或记帐后发药机数据上传
    '入参:
    '   bytBillType 单据类型：1-收费单,2-记帐单
    '   strNos 费用单据，格式：A001,A002
    '出参:
    '返回:
    '说明:110319,95366
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strData As String, strReturn As String
    
    On Error GoTo ErrHandler
    If CreateDrugPacker() = False Then Exit Sub
    
    Select Case bytBillType
    Case 0: '收费单
        If Not mobjDrugMachine Is Nothing Then
            '门诊格式：1|单据1,处方号1;单据2,处方号2
            strData = "1|" & "8," & Replace(strNOs, ",", ";8,")
            Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
            Exit Sub
        End If
        
        If mobjDrugPacker Is Nothing Then Exit Sub
        '格式：单据1,处方号1|单据2,处方号2
        strData = "8," & Replace(strNOs, ",", "|8,")
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.编号, UserInfo.姓名, 0, strData, strReturn)
    Case 2: '记帐单
        If mobjDrugMachine Is Nothing Then Exit Sub
        
        '门诊格式：1|单据1,处方号1;单据2,处方号2
        strData = "1|" & "9," & Replace(strNOs, ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-配药[门诊和住院处方明细上传]"), strData, strReturn)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DrugMachine_Del(ByVal bytBillType As Byte, ByVal varValue As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退费或销账后发药机数据上传
    '入参:
    '   bytBillType 单据类型：1-收费单,2-记帐单
    '   varValue bytBillType=0,结算序号；bytBillType=2,单据号
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsData As ADODB.Recordset
    Dim strReturn As String, strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
    Dim lng结算序号 As Long, strNo As String
    
    On Error GoTo ErrHandler
    If CreateDrugPacker() = False Then Exit Sub
    
    Select Case bytBillType
    Case 0: '收费单
        lng结算序号 = varValue
        If Not mobjDrugMachine Is Nothing Then
            '本次退的减去重收的就是实际退的
            strSQL = "Select Max(Decode(a.记录状态, 2, a.Id, 0)) As 费用id, -1 * Nvl(Sum(a.付数 * a.数次), 0) As 退药数量" & _
                    " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & _
                    " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1 And a.收费类别 In('5','6','7')" & _
                    " Group By NO, Nvl(价格父号, 序号)" & _
                    " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询退费数据", lng结算序号)
            If rsData.RecordCount = 0 Then Exit Sub
            
            Do While Not rsData.EOF
                strData = strData & ";" & nvl(rsData!费用id) & "," & nvl(rsData!退药数量)
                rsData.MoveNext
            Loop
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
            Exit Sub
        End If
        
        If mobjDrugPacker Is Nothing Then Exit Sub
        strSQL = "Select a.No, a.执行部门id" & _
                " From 门诊费用记录 A, 病人预交记录 B" & _
                " Where a.结帐id = b.结帐id And a.记录状态=2  And a.收费类别 In('5','6','7') And b.结算序号 = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询退费数据", lng结算序号)
        If rsData.RecordCount = 0 Then Exit Sub
        
        Do While Not rsData.EOF
            If InStr(strData & "|", "|" & nvl(rsData!NO) & "," & nvl(rsData!执行部门ID) & "|") = 0 Then
                strData = strData & "|" & nvl(rsData!NO) & "," & nvl(rsData!执行部门ID)
            End If
            rsData.MoveNext
        Loop
        strData = Mid(strData, 2)
        Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strData, strReturn)
        
    Case 2: '记帐单
        If mobjDrugMachine Is Nothing Then Exit Sub
        
        strNo = varValue
        strSQL = "Select Id As 费用id, -1 * Nvl(付数, 1) * 数次 As 退药数量" & _
                " From 门诊费用记录" & _
                " Where 记录性质 = 2 And 记录状态 = 2 And NO = [1] And 收费类别 In('5','6','7')" & _
                "       And 登记时间 + 0 = (Select Max(登记时间)" & _
                "                       From 门诊费用记录" & _
                "                       Where 记录性质 = 2 And 记录状态 = 2 And NO = [1])"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询销账数据", strNo)
        If rsData.RecordCount = 0 Then Exit Sub
        
        Do While Not rsData.EOF
            strData = strData & ";" & nvl(rsData!费用id) & "," & nvl(rsData!退药数量)
            rsData.MoveNext
        Loop
        strData = Mid(strData, 2)
        Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
        
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDefaultWindow(ByVal str类别 As String, ByVal lng药房ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的药房窗口设置
    '入参:
    '出参:
    '返回:
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str类别
        Case "5"
            If InStr(mstr西窗, ":") > 0 Then '旧数据没有存药房ID
                strTmp = mstr西窗
            ElseIf mlng西药房 > 0 And mstr西窗 <> "" Then
                strTmp = mlng西药房 & ":" & mstr西窗
            End If
        Case "6"
            If InStr(mstr中窗, ":") > 0 Then
                strTmp = mstr中窗
            ElseIf mlng中药房 > 0 And mstr中窗 <> "" Then
                strTmp = mlng中药房 & ":" & mstr中窗
            End If
        Case "7"
            If InStr(mstr中窗, ":") > 0 Then
                strTmp = mstr成窗
            ElseIf mlng成药房 > 0 And mstr成窗 <> "" Then
                strTmp = mlng成药房 & ":" & mstr成窗
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str类别
                Case "5"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function

Public Function Get发药窗口(ByVal lng病人ID As Long, ByVal str类别 As String, ByVal lng药房ID As Long, _
    Optional ByVal blnFirst As Boolean, Optional ByVal bytBillType As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:分配发药窗口
    '入参:
    '   bytBillType 单据类型：1-收费单,2-记帐单
    '出参:
    '返回:分配成功，返回true,否则返回False
    '说明:
    '   根据药房ID来确定,相同的药房ID分配相同的发药窗口
    '   修改时保持原有发药窗口，不会进来
    '规则说明：
    '    1.前面单据和前面行已分配发药窗口，则以前面的为准（保证相同药房有同一个窗口）
    '    2.判断指定病人在指定药房的未发药品记录中是否存在正在上班的发药窗口
    '      a.发药窗口存在，返回填制日期最近的发药窗口
    '      b.发药窗口不存在:
    '        i:如果存在缺省的发药窗口，且正在上班，则返回缺省的发药窗口，如果窗口未上班则返回null
    '        ii:如果不存在缺省的发药窗口，则根据动态分配规则（0-闲忙;1-平均）获取非专家的发药窗口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str发药窗口 As String
    Dim rsData As ADODB.Recordset, str缺省 As String
    
    On Error GoTo ErrHandler
    If blnFirst Then Set mcllWindows = New Collection
     
    If InStr(",5,6,7,", str类别) = 0 Then Exit Function
        
    '分配窗口时，如果发现药房与上张单据不同，则清除缺省窗口,以免药房不同分配到相同窗口
    If str类别 = "5" Then
        If lng药房ID <> mlng西药房 And mlng西药房 <> 0 Then mstr西窗 = ""
        mlng西药房 = lng药房ID '记录该病人使用的药房(划价已定)
    ElseIf str类别 = "6" Then
        If lng药房ID <> mlng成药房 And mlng成药房 <> 0 Then mstr成窗 = ""
        mlng成药房 = lng药房ID
    ElseIf str类别 = "7" Then
        If lng药房ID <> mlng中药房 And mlng中药房 <> 0 Then mstr中窗 = ""
        mlng中药房 = lng药房ID
    End If
    
    '2.从已分配中取:检查是否存在该窗口，保证相同药房有同一个窗口
    For i = 1 To mcllWindows.Count
        If mcllWindows(i)(0) = lng药房ID Then
            Get发药窗口 = mcllWindows(i)(1): Exit Function
        End If
    Next
    
    '3.指定时固定分配(指定是指没有对应药房上班时指定)
    Select Case str类别
    Case "5"
        If mstr西窗 <> "" Then
            str缺省 = mstr西窗
        ElseIf mlng西药房 > 0 Then
            str缺省 = GetDefaultWindow(str类别, lng药房ID)
        End If
    Case "6"
        If mstr成窗 <> "" Then
            str缺省 = mstr成窗
        ElseIf mlng成药房 > 0 Then
            str缺省 = GetDefaultWindow(str类别, lng药房ID)
        End If
    Case "7"
        If mstr中窗 <> "" Then
            str缺省 = mstr中窗
        ElseIf mlng中药房 > 0 Then
            str缺省 = GetDefaultWindow(str类别, lng药房ID)
        End If
    End Select
    
    '4.由药品系统那边分配
    If mobjService.zlDrugSvr_GetSendWindows(1, lng病人ID, 0, _
         lng药房ID & "," & str缺省, rsData, mlngModule) = False Then Exit Function
    If Not rsData.EOF Then str发药窗口 = nvl(rsData!发药窗口)
    
    If str发药窗口 <> "" Then
        Select Case str类别
            Case "5"
                mstr西窗 = str发药窗口
            Case "6"
                mstr成窗 = str发药窗口
            Case "7"
                mstr中窗 = str发药窗口
        End Select
        
        mcllWindows.Add Array(lng药房ID, str发药窗口), "K" & lng药房ID
    End If
    Get发药窗口 = str发药窗口
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


 


Private Sub Class_Terminate()
    '释放对象
    Set mobjService = Nothing
    Set mobjDrugPacker = Nothing
    Set mobjDrugMachine = Nothing
End Sub
