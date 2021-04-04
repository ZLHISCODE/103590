VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
'    On Error Resume Next
'    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
'    If Err <> 0 Or gobjComLib Is Nothing Then
'        If gblnShowMsg Then
'            MsgBox "引用zl9ComLib部件失败！", vbInformation, GSTR_MESSAGE
'        End If
'        Exit Sub
'    End If
    gintCount = gintCount + 1
End Sub

Private Sub Class_Terminate()
    gintCount = gintCount - 1
    If gintCount <= 0 Then
        Set gobjComLib = Nothing
        If gcnOutside Is Nothing Then Exit Sub
        If gcnOutside.State = adStateOpen Then gcnOutside.Close
    End If
End Sub

Public Function DBConnect() As Boolean
'连接中间数据库
    Dim strServer As String, strDBName As String, strUser As String, strPassword As String
    Dim blnConnectFinish As Boolean
    Dim strStart As String
    
    strStart = GetSetting("ZLSOFT", "公共模块\自动发药机", "启用住院药房")
    
    If gobjComLib Is Nothing Then
        On Error Resume Next
        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
        If Err.Number = 0 Then
            gstrUnit = gobjComLib.GetUnitName
        End If
        On Error GoTo 0
    End If
    
    '如果存在注册表并且为0表示未启用住院药房接口
    If strStart = "0" Then Exit Function
    
    '查询注册表有无连接服务器的信息
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    If Trim(strUser) = "" Then
        '无：默认信息
        DBConnect = MSSQLServerOpen(MSTR_SERVER, MSTR_DBNAME, MSTR_USER, MSTR_PASSWORD)
    Else
        '有：注册表信息
        strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
        strDBName = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
        strPassword = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
        strPassword = StringEnDeCodecn(strPassword, 68)     '解密
        DBConnect = MSSQLServerOpen(strServer, strDBName, strUser, strPassword)
    End If
    
    Call OutputLog("授权单位：" & gstrUnit)
    If DBConnect Then
        Call OutputLog("DBConnect成功：" & Now)
    Else
        Call OutputLog("DBConnect失败：" & Now)
    End If
End Function

Public Function SetDBConnect() As Boolean
    '设置数据库服务器连接参数
    Dim strStart As String
    
    strStart = GetSetting("ZLSOFT", "公共模块\自动发药机", "启用住院药房")
    
    '如果存在注册表并且为0表示未启用住院药房接口
    If strStart = "0" Then Exit Function
    
    frmOutsideLinkSet.Show vbModal
    '返回值
    SetDBConnect = frmOutsideLinkSet.gblnSetupFinish
End Function

Public Function TranDrugPacker(ByVal strData As String) As String
'功能： 传送药品自动分包数据
'参数： 分包数据字符串
'格式： 病区编码;库房组号;分包设备编号;发药号|收发ID1;病例号;...|收发ID2;病例号;...|收发ID3;病例号;...
'       山西阳煤医院需要传入发药号
'规则： 0-收发ID,1-病例号,2-病人ID,3-姓名,4-病区编码,5-病区名称,6-药师姓名,7-床号,8-服用方法,9-服药时间说明,
'       10-药品编码,11-药品名称,12-规格,13-剂量,14-剂量单位,15-服用数量,16-服用时间,17-分包设备编号,18-优先标记,19-医嘱类型,20-审核人
'       ,21-执行频次,22-执行性质,23-执行标记,24-毒理分类,25-批号(未用),26-效期
'       21-24山西阳煤医院需要额外传这几个数据，由接口进行判断是否上传
'       25-山西阳煤医院补充上传的信息，其他用户可不传 by:李业庆,128515,2018/7/9
'返回值：未成功传送的收发ID字符串
    Dim arrPrimary As Variant, arrSecondly As Variant, arrSecondlyVals As Variant
    Dim strInsert As String, strTmp As String, strID As String, strPageNO As String
    Dim i As Integer, j As Integer, intPageNO As Integer
    Dim rsInsert As New ADODB.Recordset
    Dim blnRollback As Boolean, blnInsert As Boolean, blnInserted As Boolean
    Dim blnContinue As Boolean
    Dim str毒理分类 As String
    Dim intCount As Integer
    
    Call OutputLog("" & Now)
    Call OutputLog("进入传送药品数据的过程")
    
    '如果不存在注册表或者为0表示不接受毒理分类为毒麻精神类的药品
    str毒理分类 = GetSetting("ZLSOFT", "公共模块\自动发药机\住院数据规则", "毒理分类", "0")
    
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        MsgBox "你未连接数据库，请先执行DBConnect()函数！", vbCritical, GSTR_MESSAGE
        TranDrugPacker = "NOT"
        Exit Function
    End If
    
    strTmp = Trim(strData)
    If strTmp = "" Then Exit Function
    
    If InStr(strTmp, "|") > 0 Then
        Call OutputLog("参数：" & Left(strTmp, InStr(1, strTmp, "|") - 1))
    End If
    
    arrPrimary = Split(Mid(strTmp, 1, InStr(1, strTmp, "|") - 1), ";")
    strTmp = Mid(strTmp, InStr(1, strTmp, "|") + 1)
    arrSecondly = Split(strTmp, "|")
    
    Call OutputLog("授权单位：" & gstrUnit)
    
    '取PageNO号
    If gstrUnit = GCST_UNIT_YQMY And UBound(arrPrimary) >= 3 Then
        '阳泉媒业：按“发药号”方式
        strPageNO = arrPrimary(3)
    Else
        strTmp = "select convert(char(6),getdate(),12) + right('000000'+cast(isnull(max(substring(page_no,7,len(page_no))),0)+1 as varchar(4)),4) max_no " _
               & "from dbo.atf_ypxx where convert(char(6),getdate(),12)=left(page_no,6)"
        rsInsert.Open strTmp, gcnOutside
        strPageNO = rsInsert!max_no
        rsInsert.Close
    End If
    
    '先传送表数据(从)
'    intPageNO = 1   '计数
'    intAbate = 0    '回滚数
    If gstrUnit = GCST_UNIT_YQMY Then
        strInsert = "insert into dbo.atf_ypxx " _
                  & "(DETAIL_SN,inpatient_no,p_id,name,ward_sn,ward_name,doctor,bed_no,comment,comm2,drug_code,drugname" _
                  & ",specification,dosage,dos_unit,amount,occ_time,atf_no,pri_flag,Mz_flag,dept_name,page_no,lot) " & Chr(13)
    Else
        strInsert = "insert into dbo.atf_ypxx " _
                  & "(DETAIL_SN,inpatient_no,p_id,name,ward_sn,ward_name,doctor,bed_no,comment,comm2,drug_code,drugname" _
                  & ",specification,dosage,dos_unit,amount,occ_time,atf_no,pri_flag,Mz_flag,dept_name,page_no) " & Chr(13)
    End If
    
    strTmp = ""
    For i = LBound(arrSecondly) To UBound(arrSecondly)
        '得到元素
        arrSecondlyVals = Split(arrSecondly(i), ";")
        
        blnContinue = True
        
        '阳煤医院规则
        '当医嘱频率为一次性、自备药、不取药的时候数据不传到包药机
        '毒麻精神类药品根据注册表项确定是否上传
        If gstrUnit = GCST_UNIT_YQMY Then
            '医嘱数据规则：
            '执行频次:一次性
            '执行性质:5-院外执行；药品医嘱界面表现为正常、外购药、出院带药。
            '执行标记:对药品，根据医嘱录入为：0-常规,1-自取药，2-不取药
            If UBound(arrSecondlyVals) >= 24 Then
                If arrSecondlyVals(21) = "一次性" Or arrSecondlyVals(22) = 5 Or arrSecondlyVals(23) = 2 Then
                    blnContinue = False
                End If
                
                '默认不传毒麻精神类药品
                If str毒理分类 = "0" And arrSecondlyVals(24) <> "普通药" Then
                    blnContinue = False
                End If
            End If
        End If
        
        If blnContinue = True Then
            '组织字符串
            strTmp = strTmp & "select "
            For j = LBound(arrSecondlyVals) To UBound(arrSecondlyVals)
                Select Case j
                Case 0
                    strTmp = strTmp & "'" & arrSecondlyVals(j) & "'"
                Case 1 To 12, 14, 20, 16 To 18
                    strTmp = strTmp & ",'" & arrSecondlyVals(j) & "'"
                Case 13, 15, 19
                    strTmp = strTmp & "," & arrSecondlyVals(j)
                End Select
            Next
            strTmp = strTmp & ",'" & strPageNO & "'"
            If gstrUnit = GCST_UNIT_YQMY Then
                '阳煤需要上次的额外数据：效期
                strTmp = strTmp & ",'" & arrSecondlyVals(26) & "'"
            End If
            strTmp = strTmp & " union all " & Chr(13)
            '判断下条记录是否为同一收发ID
            strID = arrSecondlyVals(0)
            If i = UBound(arrSecondly) Then
                blnInsert = True
            Else
                If Mid(arrSecondly(i + 1), 1, InStr(1, arrSecondly(i + 1), ";") - 1) = strID Then
                    blnInsert = False
                Else
                    blnInsert = True
                End If
            End If
            '是否执行Insert语句
            If blnInsert = True Then
                blnRollback = False
                strTmp = Left(strTmp, Len(strTmp) - 11)
                
                Call OutputLog(strInsert & strTmp)
                
                gcnOutside.BeginTrans
                On Error GoTo errRollback
                rsInsert.Open strInsert & strTmp, gcnOutside
                On Error GoTo 0
                intCount = 0
                If blnRollback = False Then
                    gcnOutside.CommitTrans
                    blnInserted = True
                Else
    '                intPageNO = intPageNO - intAbate - 1
                    '记录未提交的收发ID
                    TranDrugPacker = TranDrugPacker & strID & ";"
                End If
                If rsInsert.State = adStateOpen Then rsInsert.Close
                strTmp = ""
    '            intAbate = 0
            Else
                strTmp = strTmp & Chr(13)
                '记录多少条相同的
    '            intAbate = intAbate + 1
            End If
    '        intPageNO = intPageNO + 1
    
        Else
            Call OutputLog("  不提交的药品：" & arrSecondly(i))
        End If
    Next
    
    If rsInsert.State = adStateOpen Then rsInsert.Close
    '先传送表数据(主)
    If blnInserted Then
        blnRollback = False
        strTmp = "insert into dbo.atf_yp_page_no (ward_sn,group_no,atf_no,submit_time,page_no) " & Chr(13)
        strTmp = strTmp & "select "
        For i = LBound(arrPrimary) To UBound(arrPrimary)
            Select Case i
            Case 0 To 2     '病区编码;库房组号;分包设备编号
                strTmp = strTmp & "'" & arrPrimary(i) & "',"
            End Select
        Next
        
        strTmp = strTmp & "getdate(),'" & strPageNO & "'"
        
        Call OutputLog(strTmp)
        
        '提交数据
        gcnOutside.BeginTrans
        On Error GoTo errRollback
        rsInsert.Open strTmp, gcnOutside
        On Error GoTo 0
        intCount = 0
        If blnRollback = False Then
            gcnOutside.CommitTrans
        Else
            '如果主表数据失败，同时删除从表对应数据
            strTmp = "delete dbo.atf_ypxx where page_no='" & strPageNO & "'"

            On Error Resume Next
            If rsInsert.State = adStateOpen Then rsInsert.Close
            rsInsert.Open strTmp, gcnOutside
            If rsInsert.State = adStateOpen Then rsInsert.Close
            '返回所有收发ID字符串
            strID = "": TranDrugPacker = ""
            For i = LBound(arrSecondly) To UBound(arrSecondly)
                If Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1) <> strID Then
                    strID = Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1)
                    TranDrugPacker = TranDrugPacker & strID & ";"
                End If
            Next
        End If
    End If
    'If gcnOutside.State = adStateOpen Then gcnOutside.Close
    If Trim(TranDrugPacker) <> "" Then
        '返回收发ID字符串
        TranDrugPacker = Left(TranDrugPacker, Len(TranDrugPacker) - 1)
    End If
    
    Call OutputLog("传送药品数据的过程成功")
    
    Exit Function

errRollback:
    '支持3次提交尝试，尝试3次后就放弃当前insert的提交
    intCount = intCount + 1
    If intCount <= 3 Then
        Sleep 300       '等待0.3秒
        Resume          '重试
    Else
        gcnOutside.RollbackTrans
        Call OutputLog("异常：" & Err.Description)
        blnRollback = True
        Resume Next     '放弃
    End If
End Function

Public Function DrugPackerSet(ByVal cnZLHIS As ADODB.Connection, ByVal blnOutsideConnected As Boolean) As String
    Dim frmTmp As New frmDrugPackerSet
    frmTmp.ConnectHIS = cnZLHIS
    frmTmp.OutsideConnected = blnOutsideConnected
    frmTmp.Show vbModal
    '返回值
    DrugPackerSet = frmTmp.ResultString
    Set frmTmp = Nothing
End Function

Public Function TranDrugSingle(ByVal cnHIS As ADODB.Connection, ByVal lngID As Long) As Boolean
    Dim cmdInsert As New ADODB.Command, rsTmp As New ADODB.Recordset
    Dim strTmp As String, strInsert As String
    Dim lngExec As Long

    If cnHIS Is Nothing Or cnHIS.State = adStateClosed Then
        'MsgBox "ZLHIS数据库未连接！", vbCritical, GSTR_MESSAGE
        Exit Function
    End If
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        'MsgBox "你未连接数据库，请先执行DBConnect()函数！", vbCritical, GSTR_MESSAGE
        Exit Function
    End If

    strTmp = "Select A.*, B.编码 厂商编码 " _
           & "From (Select A.药品id, B.编码, B.名称, B.规格, D.药品剂型 药品类型, A.剂量系数 剂量, E.计算单位 剂量单位," _
           & "        A.住院包装 包装数量, A.住院单位 包装单位, Nvl(A.上次产地, B.产地) 生产厂商," _
           & "        zlTools.zlSpellCode(B.名称) 拼音码" _
           & "      From 药品规格 A, 收费项目目录 B, 药品特性 D, 诊疗项目目录 E" _
           & "      Where A.药品id = B.ID And A.药名id = D.药名id And A.药名id = E.ID And B.类别 In ('5', '6', '7') And" _
           & "        Nvl(B.撤档时间, To_Date('3000-1-1', 'yyyy-mm-dd')) = To_Date('3000-1-1', 'yyyy-mm-dd') And" _
           & "        A.药品ID=" & lngID & ") A, 药品生产商 B " _
           & "Where A.生产厂商 = B.名称(+) order by cast(a.编码 as int) "
    rsTmp.Open strTmp, cnHIS
    If rsTmp.RecordCount <> 1 Then
        rsTmp.Close
        Exit Function
    End If
    gcnOutside.BeginTrans
    On Error GoTo errHand
    strInsert = "delete dbo.atf_his_druginfo where drug_code='" & rsTmp!编码 & "' and drugname='" & rsTmp!名称 & "' " & Chr(13) _
              & "insert into dbo.atf_his_druginfo (drug_code,drugname,specification,drug_type," _
              & "dosage,dos_unit,pack_amount,pack_name,manufactory,py_code,manu_no) " & Chr(13)
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        strTmp = "select '" & rsTmp!编码 & "'," _
               & "'" & rsTmp!名称 & "'," _
               & "'" & rsTmp!规格 & "'," _
               & "'" & rsTmp!药品类型 & "'," _
               & CDbl(rsTmp!剂量) & "," _
               & "'" & rsTmp!剂量单位 & "'," _
               & CDbl(rsTmp!包装数量) & "," _
               & "'" & rsTmp!包装单位 & "'," _
               & "'" & IIf(IsNull(rsTmp!生产厂商), "", rsTmp!生产厂商) & "'," _
               & "'" & rsTmp!拼音码 & "'," _
               & "'" & IIf(IsNull(rsTmp!厂商编码), "", rsTmp!厂商编码) & "' union all "
        strInsert = strInsert & strTmp & Chr(13)
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    strInsert = Left(strInsert, Len(strInsert) - 11)
    '更新数据
    With cmdInsert
        .ActiveConnection = gcnOutside
        .CommandText = strInsert
        .Execute 'lngExec, , adCmdText
    End With
    'If lngExec <= 0 Then GoTo errHand

    gcnOutside.CommitTrans
    TranDrugSingle = True
    Exit Function

errHand:
    gcnOutside.RollbackTrans
End Function

Public Function DYEY_MZ_IniSoap(Optional ByVal blnShowMsg As Boolean = True, Optional ByRef strMsg As String, _
    Optional ByVal strUser As String, _
    Optional ByVal cnOracle As ADODB.Connection, _
    Optional ByVal intMode As Integer = 0) As Boolean
    '功能：初始化接口
    '参数：
    '  blnShowMsg：显示消息提示对话框，默认显示
    '  strMsg：返回消息
    '  intMode：0-ZLHIS业务调用；1-无线业务调用
    '返回值：true-成功，false-失败
    
    Dim strUrl As String '获取路径
    Dim strOutput As String
    
    strOutput = "" & vbCrLf & CStr(Now)
    strOutput = strOutput & vbCrLf & "调用接口：DYEY_MZ_IniSoap"
    gintMode = intMode      '无线业务不用ZLHIS登录，所有zl9ComLib部件不调用

    If gintMode = 0 Then
        If gobjComLib Is Nothing Then
            strOutput = strOutput & vbCrLf & "开始创建公共部件：zl9ComLib.clsComLib"
            Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
            strOutput = strOutput & vbCrLf & "创建公共部件成功"
        End If
    End If
    
'    If cnOracle Is Nothing Then
'        call OutPutLog(CStr(Now) & " " & "建立数据库连接")
'        gobjComLib.InitCommon cnOracle
'    End If
    
    If intMode = 1 Then
        If cnOracle Is Nothing Then
            strOutput = strOutput & vbCrLf & "初始化本部件失败：未传入ZLHIS的Oracle连接对象"
            Call OutputLog(strOutput)
            Exit Function
        Else
            If cnOracle.State <> adStateOpen Then
                strOutput = strOutput & vbCrLf & "初始化本部件失败：传入ZLHIS的Oracle连接对象未打开"
                Call OutputLog(strOutput)
                Exit Function
            End If
            Set gcnWireless = cnOracle
        End If
        gstrUnit = mdlDrugPacker.zlRegInfo("单位名称", , -1)
        If gstrUnit = "" Then
            '如果非ZLHIS登录，就取注册表
            gstrUnit = GetSetting("ZLSOFT", "公共模块\自动发药机", "单位名称")
            If gstrUnit = "" Then
                '如果不能访问注册表，就取配置文件zlDrugPacker.ini
                gstrUnit = GetINIInfo("单位名称")
            End If
        End If
    Else
        Set gcnWireless = Nothing
        gstrUnit = gobjComLib.GetUnitName
        If gstrUnit = "" Then
            gstrUnit = GetSetting("ZLSOFT", "公共模块\自动发药机", "单位名称")
        End If
    End If
    
    strOutput = strOutput & vbCrLf & "获取用户名：" & gstrUnit
    
    '该接口采取WebService访问方式，通过用户名来判断
    If gstrUnit <> GCST_UNIT_DYEY And gstrUnit <> GCST_UNIT_YZSZYY And gstrUnit <> GCST_UNIT_JLSZXYY And gstrUnit <> GCST_UNIT_CQFLQZYY _
            And gstrUnit <> GCST_UNIT_YNYXRMYY And gstrUnit <> GCST_UNIT_BTSZXYY Then
        strOutput = strOutput & vbCrLf & "当前用户名与接口支持的用户名不符合"
        strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_IniSoap"
        Call OutputLog(strOutput)
        strMsg = "该单位未授权使用"
        
        Exit Function
    End If
    
    '检查“是否上传”
    gblnUpdateFlag = HaveUpdateFlag()
    
    On Error Resume Next
    
    strOutput = strOutput & vbCrLf & "创建MSSoap组件30"
    
    Set gobjSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo errSoap
        
        strOutput = strOutput & vbCrLf & "创建MSSoap组件20"
        
        Set gobjSOAP = CreateObject("MSSOAP.SoapClient")
    End If
    
    gblnShowMsg = blnShowMsg
    
    GoTo RRow
    
errSoap:
    If blnShowMsg Then
        MsgBox "系统初始化失败！" & vbCrLf & vbCrLf & "客户端未安装SOAP！" & vbCrLf & vbCrLf & "错误信息如下：" & vbCrLf & vbCrLf & Err.Description, vbCritical, GSTR_MESSAGE
    Else
        strMsg = "客户端未安装SOAP！"
    End If
    
    strOutput = strOutput & vbCrLf & "客户端未安装SOAP，初始化失败"
    Call OutputLog(strOutput)
    
    Exit Function
    
RRow:
    On Error GoTo errHandle
    
    strUrl = GetSetting("ZLSOFT", "公共模块\WebService路径", "WebUrl")
    If Trim(strUrl) = "" Then
        strUrl = GetINIInfo("WebService路径")
    End If
    
    strOutput = strOutput & vbCrLf & "获取WebService地址：" & strUrl
    
    If strUrl = "" Then
        If blnShowMsg Then
            MsgBox "未设置WebServic服务地址，请在程序中进行设置！", vbInformation + vbOKOnly, GSTR_MESSAGE
        Else
            strMsg = "未设置WebServic服务地址！"
        End If
        
        strOutput = strOutput & vbCrLf & "未设置WebServic服务地址，初始化失败"
        strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_IniSoap"
        
        DYEY_MZ_IniSoap = False
    Else
        Call gobjSOAP.MSSoapInit(strUrl & IIf(UCase(Right(strUrl, 5)) = "?WSDL", "", "?WSDL"))
        
        strOutput = strOutput & vbCrLf & "已连接WebServic，初始化成功"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_IniSoap"
        
        DYEY_MZ_IniSoap = True
    End If
    
    Call OutputLog(strOutput)
    
    Exit Function
    
errHandle:
    DYEY_MZ_IniSoap = False
    
    If blnShowMsg Then
        MsgBox "初始化失败，请检测服务器地址。", vbCritical, GSTR_MESSAGE
    Else
        strMsg = "初始化失败，请检测服务器地址。"
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_IniSoap"
    
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_SetServer() As Boolean
    '设置WebService服务地址
    '返回值：true-成功，false-失败
    Dim strUrl As String
    Dim strOutput As String
    
    On Error GoTo errHandle
    
    strOutput = "" & vbCrLf & CStr(Now)
    strOutput = strOutput & vbCrLf & "调用接口：DYEY_MZ_SetServer"
        
    Call frmDYEY_MZ_SetServer.ShowMe(strUrl)
            
    If strUrl = "" Then
        MsgBox "未设置WebServic服务地址，请在程序中进行设置！", vbInformation + vbOKOnly, GSTR_MESSAGE
        
        strOutput = strOutput & vbCrLf & "未设置WebServic服务地址"
        strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_SetServer"
    Else
        If strUrl <> GetSetting("ZLSOFT", "公共模块\WebService路径", "WebUrl") Then
            Call gobjSOAP.MSSoapInit(strUrl & IIf(UCase(Right(strUrl, 5)) = "?WSDL", "", "?WSDL"))
            Call SaveSetting("zlsoft", "公共模块\WebService路径", "WebUrl", strUrl)
        End If
        
        DYEY_MZ_SetServer = True
        
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_SetServer"
    End If
    
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    MsgBox "初始化失败，请检测服务器地址。", vbCritical, GSTR_MESSAGE
    
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_SetServer"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_TransDrug(ByVal intOprID As Integer, ByVal strUserCode As String, ByVal strUserName As String, ByRef strReturn As String) As Boolean
'HIS端向接口传递药品基础信息
'返回值：true-成功，false-失败
'适用接口：韦乐海茨CONSIS系统，注意有版本区别
    Dim arrXML As Variant
    Dim strOutput As String
    
    On Error GoTo errHandle
    
    strOutput = " " & vbCrLf & CStr(Now)
    strOutput = strOutput & vbCrLf & "调用接口：DYEY_MZ_TransDrug"
    strOutput = strOutput & vbCrLf & "intOprId=" & intOprID
    strOutput = strOutput & vbCrLf & "strUserCode=" & strUserCode
    strOutput = strOutput & vbCrLf & "strUserName=" & strUserName
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        arrXML = GetXML_Drug_CQFLQZYY(strOutput)
    Else
        arrXML = GetXML_Drug
    End If
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransDrug = True
        
        strOutput = strOutput & vbCrLf & "无数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransDrug"
        Call OutputLog(strOutput)
        
        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        If DYEY_MZ_TransData_CQFLQZYY(gType.IntDrug, intOprID, strUserCode, strUserName, arrXML, strReturn, strOutput) Then
            DYEY_MZ_TransDrug = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransDrug"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransDrug"
        End If
    Else
        If DYEY_MZ_TransData(gType.IntDrug, intOprID, strUserCode, strUserName, arrXML, strReturn) Then
            DYEY_MZ_TransDrug = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransDrug"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransDrug"
        End If
    End If
    
    Call OutputLog(strOutput)
    Exit Function
errHandle:
    If gintMode <> 1 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransDrug"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_TransStockDevice(ByVal strUserCode As String, ByVal strUserName As String, _
    ByRef strStore As String, _
    ByRef strReturn As String) As Boolean
'功能：向摆药机读取设备的药品库存信息
'参数：
'  strUserCode：用户编码
'  strUserName：用户姓名
'  strStore：设备库存信息
'  strReturn：返回信息
'返回：

'<ROOT>
'    <OPSYSTEM>HIS</OPSYSTEM>
'    <OPWINID>DeviceQty</OPWINID>
'    <OPTYPE>601</OPTYPE>
'    <OPIP></OPIP>
'    <OPMANNO></OPMANNO>
'    <OPMANNAME></OPMANNAME>
'</ROOT>

    Dim strInput As String, strOut As String, strOutput As String
    
    strInput = "<ROOT>" & _
                   "<OPSYSTEM>HIS</OPSYSTEM>" & _
                   "<OPWINID>DeviceQty</OPWINID>" & _
                   "<OPTYPE>601</OPTYPE>" & _
                   "<OPIP></OPIP>" & _
                   "<OPMANNO></OPMANNO>" & _
                   "<OPMANNAME></OPMANNAME>" & _
               "</ROOT>"
    
    On Error GoTo hErr
    strOut = gobjSOAP.HisTransData(strInput)
    On Error GoTo 0
    
    strStore = strOut
    Call OutputLog(strStore)
    DYEY_MZ_TransStockDevice = True
    Exit Function
    
hErr:
    strReturn = Err.Description
    strOutput = strOutput & vbCrLf & "发生异常错误" & vbCrLf & Err.Description
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransStockDevice"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_TransStock(ByVal intOprID As Integer, ByVal strUserCode As String, ByVal strUserName As String, _
    ByVal lngStockID As Long, ByRef strReturn As String) As Boolean
'HIS端向接口传递药品库存信息
'返回值：true-成功，false-失败
'适用接口：韦乐海茨CONSIS系统，注意有版本区别
    Dim arrXML As Variant
    Dim strOutput As String
    Dim str注册药房 As String
    
    On Error GoTo errHandle
    
    strOutput = "" & vbCrLf & CStr(Now)
    strOutput = strOutput & "调用接口：DYEY_MZ_TransStock" & vbCrLf
    strOutput = strOutput & "intOprId=" & intOprID & vbCrLf
    strOutput = strOutput & "strUserCode=" & strUserCode & vbCrLf
    strOutput = strOutput & "strUserName=" & strUserName & vbCrLf
    strOutput = strOutput & "LngStockID=" & lngStockID & vbCrLf
    
    '如果门诊药房只有一个或只有一个启用自动发药系统，可以固定药房ID''''''''''''''''''''''''''''''''''''''''''''''''
    str注册药房 = GetSetting("ZLSOFT", "公共模块\自动发药机", "门诊药房")
    If Val(str注册药房) = 0 Then
        str注册药房 = GetINIInfo("门诊药房")
    End If

    If Val(str注册药房) = 0 Then
        strOutput = strOutput & vbCrLf & "未注册药房"
        strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeDetail"
        Call OutputLog(strOutput)

        Exit Function
    End If
    
    If IsRegisterStock(lngStockID, str注册药房) = False Then
        DYEY_MZ_TransStock = True

        strOutput = strOutput & vbCrLf & "当前药房不是自动发药机注册药房，不进行数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeDetail"
        Call OutputLog(strOutput)

        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        arrXML = GetXML_Stock_CQFLQZYY(lngStockID, strOutput)
    Else
        arrXML = GetXML_Stock(lngStockID)
    End If
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransStock = True
        
        strOutput = strOutput & vbCrLf & "无数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransStock"
        Call OutputLog(strOutput)
        
        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        If DYEY_MZ_TransData_CQFLQZYY(gType.IntStore, intOprID, strUserCode, strUserName, arrXML, strReturn, strOutput) Then
            DYEY_MZ_TransStock = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransStock"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransStock"
        End If
    Else
        If DYEY_MZ_TransData(gType.IntStore, intOprID, strUserCode, strUserName, arrXML, strReturn) Then
            DYEY_MZ_TransStock = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransStock"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransStock"
        End If
    End If
    
    Call OutputLog(strOutput)
    
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransStock"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_TransDept(ByVal strProperty As String, ByVal strUserCode As String, ByVal strUserName As String, _
    ByRef strReturn As String) As Boolean
    
'功能：提交部门数据给设备接口（韦乐海茨）
'参数：
'  strProperty：部门性质字符串；格式：性质1,性质2,...
'  strReturn：设备接口的返回信息
'返回：True成功；False失败

    Dim strLog As String
    Dim arrXML As Variant
    
strLog = strLog & Now() & vbCrLf
strLog = strLog & "准备提交部门数据！" & vbCrLf & "部门工作性质：" & strProperty & vbCrLf
    
    arrXML = GetXML_Dept(strProperty, strLog)
    
    If UBound(arrXML) = -1 Then
strLog = strLog & "无部门数据可提交！" & vbCrLf
        DYEY_MZ_TransDept = True
        Exit Function
    End If
    
    If DYEY_MZ_TransData_CQFLQZYY(gType.IntDept, 0, strUserCode, strUserName, arrXML, strReturn, strLog) Then
strLog = strLog & "部门数据提交完成！" & vbCrLf
        DYEY_MZ_TransDept = True
    Else
strLog = strLog & "部门数据提交失败！" & vbCrLf
    End If
    
    Call OutputLog(strLog)
    
    Exit Function

hErr:
strLog = strLog & "部门数据提交异常！" & vbCrLf & Err.Description
End Function

Public Function DYEY_MZ_TransRecipeDetail(ByVal intOprID As Integer, ByVal strUserCode As String, _
    ByVal strUserName As String, ByVal lngStockID As Long, _
    ByVal strNO As String, ByRef strReturn As String, _
    Optional ByRef strWinInfo As String) As Boolean
'HIS端向接口传递处方明细信息
'返回值：true-成功，false-失败
'适用接口：韦乐海茨CONSIS系统，注意有版本区别

    Dim arrXML As Variant
    Dim strOutput As String
    Dim str注册药房 As String
    Dim arrStock As Variant
    Dim lngID As Long
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strOutput = "" & vbCrLf & CStr(Now)
    strOutput = strOutput & vbCrLf & "调用接口：DYEY_MZ_TransRecipeDetail"
    strOutput = strOutput & vbCrLf & "intOprId=" & intOprID
    strOutput = strOutput & vbCrLf & "strUserCode=" & strUserCode
    strOutput = strOutput & vbCrLf & "strUserName=" & strUserName
    strOutput = strOutput & vbCrLf & "LngStockID=" & lngStockID
    strOutput = strOutput & vbCrLf & "strNO=" & strNO
    
    strWinInfo = ""
    
    '如果门诊药房只有一个或只有一个启用自动发药系统，可以固定药房ID''''''''''''''''''''''''''''''''''''''''''''''''
    '门诊药房格式：药房ID[;药房ID[...]]
    str注册药房 = Trim(GetSetting("ZLSOFT", "公共模块\自动发药机", "门诊药房"))
    If Val(str注册药房) = 0 Then
        str注册药房 = GetINIInfo("门诊药房")
        
        If Val(str注册药房) = 0 Then
            strOutput = strOutput & vbCrLf & "未注册药房"
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeDetail"
            Call OutputLog(strOutput)
    
            Exit Function
        End If
    End If
    
    arrStock = Split(str注册药房, ";")
    
    If gstrUnit = GCST_UNIT_YNYXRMYY Then
        '云南玉溪医院，碟和系统
        For i = LBound(arrStock) To UBound(arrStock)
            lngID = Val(arrStock(i))
            If lngID <> 0 Then
                Call HisTransData_DIH(gType.IntDetail, lngID, strNO, strReturn)
                '更新上传标志
                Call UpdateFlag(lngID, strNO)
            End If
        Next
        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        arrXML = GetXML_RecipeDetail_CQFLQZYY(str注册药房, strNO, strOutput)
    Else
        arrXML = GetXML_RecipeDetail(str注册药房, strNO)
    End If
    
    If UBound(arrXML) = -1 Then
        If gintMode = 1 And gstrUnit <> GCST_UNIT_BTSZXYY Then
            '移动业务：即使没有包药机处理的数据，也要向终端返回病人ID、药房、窗口、药品信息
            strWinInfo = PackingWindow_DYEY(strNO, strOutput)
            Call OutputLog("无数据上传包药机，接口回写移动XML信息：" & strWinInfo)
        End If
        
        DYEY_MZ_TransRecipeDetail = True
        
        strOutput = strOutput & vbCrLf & "无数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeDetail"
        Call OutputLog(strOutput)
        
        '更新上传标志
        For i = LBound(arrStock) To UBound(arrStock)
            lngID = Val(arrStock(i))
            If lngID <> 0 Then Call UpdateFlag(lngID, strNO)
        Next
        
        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        If DYEY_MZ_TransData_CQFLQZYY(gType.IntDetail, intOprID, strUserCode, strUserName, arrXML, strReturn, strOutput, strNO, lngStockID) Then
            DYEY_MZ_TransRecipeDetail = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeDetail"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeDetail"
        End If
    Else
        If DYEY_MZ_TransData(gType.IntDetail, intOprID, strUserCode, strUserName, arrXML, strReturn, strNO, lngStockID) Then
            Call OutputLog("上传包药机，包药机接口返回信息：" & strReturn)
            If gintMode = 1 Then
                '移动业务：向终端返回病人ID、药房、窗口、药品信息
                strWinInfo = PackingWindow_DYEY(strNO, strOutput)
                Call OutputLog("上传包药机，DYEY_MZ_TransRecipeDetail 回写移动XML信息：" & strWinInfo)
            End If
            
            DYEY_MZ_TransRecipeDetail = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeDetail"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeDetail"
        End If
    End If
    
    '更新上传标志
    For i = LBound(arrStock) To UBound(arrStock)
        lngID = Val(arrStock(i))
        If lngID <> 0 Then Call UpdateFlag(lngID, strNO)
    Next
    
    Call OutputLog(strOutput)
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeDetail"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_TransRecipeList(ByVal strOprId As String, ByVal strUserCode As String, ByVal strUserName As String, ByVal lngStockID As Long, ByVal strNO As String, ByRef strReturn As String, Optional ByVal intSendType As Integer) As Boolean
'HIS端向接口传递处方主单
'intSendType:0-开始发药,1-结束发药
'返回值：true-成功，false-失败
'适用接口：韦乐海茨CONSIS系统，注意有版本区别
    Dim arrXML As Variant
    Dim strOutput As String
    Dim str注册药房 As String
    Dim intOprID As Integer
    
    On Error GoTo errHandle
    
    intOprID = Val(strOprId)
    
    strOutput = "" & vbCrLf & CStr(Now)
    strOutput = strOutput & vbCrLf & "调用接口：DYEY_MZ_TransRecipeList"
    strOutput = strOutput & vbCrLf & "GINT_SEND_TYPE=" & GINT_SEND_TYPE
    strOutput = strOutput & vbCrLf & "intOprId=" & intOprID
    strOutput = strOutput & vbCrLf & "strUserCode=" & strUserCode
    strOutput = strOutput & vbCrLf & "strUserName=" & strUserName
    strOutput = strOutput & vbCrLf & "LngStockID=" & lngStockID
    strOutput = strOutput & vbCrLf & "strNO=" & strNO
    
    '如果门诊药房只有一个或只有一个启用自动发药系统，可以固定药房ID''''''''''''''''''''''''''''''''''''''''''''''''
    str注册药房 = GetSetting("ZLSOFT", "公共模块\自动发药机", "门诊药房")
    If Val(str注册药房) = 0 Then
        Call OutputLog("未注册药房1")
        str注册药房 = GetINIInfo("门诊药房")
    End If

    If Val(str注册药房) = 0 Then
        strOutput = strOutput & vbCrLf & "未注册药房2"
        strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeList"
        Call OutputLog(strOutput)

        Exit Function
    End If
    
    If IsRegisterStock(lngStockID, str注册药房) = False Then
        DYEY_MZ_TransRecipeList = True

        strOutput = strOutput & vbCrLf & "当前药房不是自动发药机注册药房，不进行数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeList"
        Call OutputLog(strOutput)

        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_YNYXRMYY Then
        '云南玉溪医院，碟和系统
        Call OutputLog("执行开始：DYEY_MZ_TransRecipeList")
        DYEY_MZ_TransRecipeList = HisTransData_DIH(IIf(intSendType = 0, gType.IntStartList, gType.IntEndList), lngStockID, strNO, strReturn)
        
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeList"
        Call OutputLog(strOutput)
        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        arrXML = GetXML_RecipeList_CQFLQZYY(lngStockID, strNO, strOutput)
    Else
        Call OutputLog("开始组装XML")
        arrXML = GetXML_RecipeList(lngStockID, strNO)
    End If
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransRecipeList = True
        
        strOutput = strOutput & vbCrLf & "无数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeList"
        Call OutputLog(strOutput)
        
        Exit Function
    End If
    
    If gstrUnit = GCST_UNIT_CQFLQZYY Or gstrUnit = GCST_UNIT_BTSZXYY Then
        If DYEY_MZ_TransData_CQFLQZYY(IIf(intSendType = 0, gType.IntStartList, gType.IntEndList), intOprID, strUserCode, strUserName, arrXML, strReturn, strOutput) Then
            DYEY_MZ_TransRecipeList = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeList"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeList"
        End If
    Else
        If DYEY_MZ_TransData(IIf(intSendType = 0, gType.IntStartList, gType.IntEndList), intOprID, strUserCode, strUserName, arrXML, strReturn) Then
            DYEY_MZ_TransRecipeList = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeList"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeList"
        End If
    End If
    
    Call OutputLog(strOutput)
    Exit Function
    
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误：" & Err.Description
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeList"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_TransRecipeReturn(ByVal intOprID As Integer, ByVal strUserCode As String, ByVal strUserName As String, _
    ByVal strReturnRecipt As String, ByRef strReturn As String) As Boolean
    
'HIS端向接口传递处方退费(全退模式)
'返回值：true-成功，false-失败
'适用接口：韦乐海茨CONSIS系统v4.3
'intOprId：操作终端号
'strUserCode：操作员编码
'strUserName：操作员名称
'strReturnRecipt：退费处方信息，格式：NO,药房id|NO,药房id
'strReturn：接口错误提示信息

    Dim arrXML As Variant
    Dim strOutput As String
    Dim strStockIDs As String
    
    On Error GoTo errHandle
    
    strOutput = "" & vbCrLf & CStr(Now)
    strOutput = strOutput & vbCrLf & "调用接口：DYEY_MZ_TransRecipeReturn"
    strOutput = strOutput & vbCrLf & "intOprId=" & intOprID
    strOutput = strOutput & vbCrLf & "strUserCode=" & strUserCode
    strOutput = strOutput & vbCrLf & "strUserName=" & strUserName
    strOutput = strOutput & vbCrLf & "strReturnRecipt=" & strReturnRecipt
    
    '如果门诊药房只有一个或只有一个启用自动发药系统，可以固定药房ID''''''''''''''''''''''''''''''''''''''''''''''''
    strStockIDs = Trim(GetSetting("ZLSOFT", "公共模块\自动发药机", "门诊药房"))
    If Val(strStockIDs) = 0 Then
        strStockIDs = GetINIInfo("门诊药房")
    End If
    If Val(strStockIDs) = 0 Then
        strOutput = strOutput & vbCrLf & "未注册药房"
        strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeReturn"
        Call OutputLog(strOutput)
        Exit Function
    End If
    
    If gstrUnit <> GCST_UNIT_CQFLQZYY And gstrUnit <> GCST_UNIT_BTSZXYY Then
        strOutput = strOutput & "无需执行：DYEY_MZ_TransRecipeReturn"
        DYEY_MZ_TransRecipeReturn = True
        Exit Function
    End If
    
    '以下代码为涪陵区中医院专用，其他医院不适用
    arrXML = GetXML_RecipeReturn_CQFLQZYY(strReturnRecipt, strStockIDs, strOutput)
    
    If UBound(arrXML) = -1 Then
        DYEY_MZ_TransRecipeReturn = True
        
        strOutput = strOutput & vbCrLf & "无数据上传"
        strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeReturn"
    Else
        If DYEY_MZ_TransData_CQFLQZYY(gType.IntReturnAll, intOprID, strUserCode, strUserName, arrXML, strReturn, strOutput) Then
            DYEY_MZ_TransRecipeReturn = True
            strOutput = strOutput & vbCrLf & "执行成功：DYEY_MZ_TransRecipeReturn"
        Else
            strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeReturn"
        End If
    End If
    
    Call OutputLog(strOutput)
    Exit Function
errHandle:
    If gintMode = 0 Then
        If gobjComLib.ErrCenter = 1 Then Resume
        Call gobjComLib.SaveErrLog
    End If
    
    strOutput = strOutput & vbCrLf & "发生异常错误"
    strOutput = strOutput & vbCrLf & "执行失败：DYEY_MZ_TransRecipeReturn"
    Call OutputLog(strOutput)
End Function

Public Function DYEY_MZ_GetSendType() As Integer
'向HIS端返回发药流程方式
    DYEY_MZ_GetSendType = GINT_SEND_TYPE
End Function

Public Function DYEY_MZ_GetStartSendType() As Integer
'向HIS端返回开始发药操作方式
    DYEY_MZ_GetStartSendType = GINT_STARTSEND_TYPE
End Function



