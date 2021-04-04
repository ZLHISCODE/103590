VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_YUYAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'说明：第三方接口（YUYAMA优雅玛）类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mcnThird As ADODB.Connection

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mcnThird = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'说明：开启接口为活动状态
    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mcnThird Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "连接YUYAMA服务器失败", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "开始连接YUYAMA服务器", 1
        mcnThird.Open mobjBase.URL
        mobjBase.mobjLog.Add "成功连接YUYAMA服务器", 1
    Else
        Set mcnThird = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Add "连接YUYAMA服务器异常", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Private Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)
End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim objDB As Object
    
    'Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    
    '如果不使用基类的记录集对象，请将前面代码屏蔽，再以下面编写新的代码
    'TODO：
    
    On Error GoTo hErr
    
    strInfo = Split(strData, "|")(1)
    
    Select Case intType
    Case 21
        If TypeName(mobjBase.mobjComLib) = "clsComLib" Then
            'ZLHIS、自助机、药品自动化设备管理工具
            Set objDB = mobjBase.mobjComLib.zlDatabase
        ElseIf TypeName(mobjBase.mobjComLib) = "clsPublic" Then
            '支付宝
            Set objDB = mobjBase.mobjComLib
        Else
            mobjBase.mobjLog.Add "对象类型不正确，终止获取标准数据", 1
            mobjBase.mobjLog.Add "" & intType, 1, 1
            mobjBase.mobjLog.Save
            Exit Function
        End If
        
        strSQL = "Select /*+cardinality(z,10)*/ Distinct Zl_Fun_Drug_Machine(c.库房id, g.药品剂型, c.Id) 接口编号, c.库房id, I1.名称 As 发药药房," & vbNewLine & _
                "    Decode(Nvl(o.标志, 0), 11, 1, 12, 1, 0, 1, 0) 是否传送, b.标识号 住院号," & vbNewLine & _
                "    c.汇总发药号 As 发药号, Sysdate As 传送时间, c.对方部门id As 领药科室id, I2.名称 As 领药科室, j.病人id, j.姓名 As 病人姓名," & vbNewLine & _
                "    j.性别 As 病人性别, j.年龄 As 病人年龄, j.就诊卡号, b.病人病区id As 病区id, I3.名称 As 病区, j.当前床号 As 床号," & vbNewLine & _
                "    k.优先级, a.开嘱医生, a.开嘱时间, a.Id As 医嘱id, g.药品剂型, c.药品id, e.名称 As 药品名称, f.名称 As 商品名称, c.用法," & vbNewLine & _
                "    c.单量, h.计算单位 As 单量单位, d.住院单位, a.医嘱期效, m.要求时间 As 执行时间, a.医生嘱托, g.毒理分类, a.紧急标志," & vbNewLine & _
                "    Decode(m.要求时间, Null, Round(c.实际数量 / d.住院包装, 2), Round((a.单次用量 / d.剂量系数) / d.住院包装, 2)) As 住院单位数量," & vbNewLine & _
                "    d.剂量系数, a.校对时间 护士校对时间, c.填制人 收发填制人, c.审核人 收发审核人, c.ID 收发ID ,m.医嘱id," & vbNewLine & _
                "    replace(to_char(m.要求时间,'hh24-mi'),'-','') As 时间编码," & vbNewLine & _
                "    Case When replace(to_char(m.要求时间,'hh24-mi'),'-','')<1200 And replace(to_char(m.要求时间,'hh24-mi'),'-','')>=0000 Then   '早上'" & vbNewLine & _
                "         When replace(to_char(m.要求时间,'hh24-mi'),'-','')<1800 And replace(to_char(m.要求时间,'hh24-mi'),'-','')>=1200 Then '中午'" & vbNewLine & _
                "         When replace(to_char(m.要求时间,'hh24-mi'),'-','')<0000 And replace(to_char(m.要求时间,'hh24-mi'),'-','')>=1800 Then '晚上'" & vbNewLine & _
                "    End  As  服用时间,to_char(m.要求时间,'YYYY-MM-DD') As 用药时间 " & vbNewLine & _
                "From 病人医嘱记录 A, 住院费用记录 B, 药品收发记录 C, 药品规格 D, 收费项目目录 E, 收费项目别名 F, 药品特性 G, 诊疗项目目录 H," & vbNewLine & _
                "    部门表 I1, 部门表 I2, 部门表 I3, 病人信息 J, 身份 K, 病人医嘱发送 L, 医嘱执行时间 M, 病人医嘱记录 N," & vbNewLine & _
                "    药品收发住院标志 O, Table(f_Num2list([1], ';')) Z " & vbNewLine & _
                "Where a.Id = b.医嘱序号 And b.Id = c.费用id And c.药品id = d.药品id And c.药品id = e.Id And e.Id = f.收费细目id(+) " & vbNewLine & _
                "    And d.药名id = g.药名id And g.药名id = h.Id And c.库房id = I1.Id And c.对方部门id = I2.Id And b.病人病区id = I3.Id " & vbNewLine & _
                "    And b.病人id = j.病人id And j.身份 = k.名称(+) And a.相关id = n.Id And n.Id = m.医嘱id " & vbNewLine & _
                "    And b.医嘱序号 = l.医嘱id(+) And b.No = l.No And b.执行部门id = l.执行部门id " & vbNewLine & _
                "    And l.发送号 = m.发送号(+) And c.Id = o.收发id(+) And f.性质(+) = 3 And f.码类(+) = 1 " & vbNewLine & _
                "    And c.Id = z.Column_Value " & vbNewLine & _
                "Order By c.库房id, c.对方部门id, j.病人id"
                
        mobjBase.mobjLog.Add strSQL, 1, 1
        Set rsResult = objDB.OpenSQLRecord(strSQL, "获取开始发药数据", strInfo)
        mobjBase.mobjLog.Add "获取开始发药数据完成", 1
        
        Set clsINF_GetData = rsResult
    Case Else
        '使用基类的数据
        Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    End Select
    Exit Function
    
hErr:
    blnFinish = False
    mobjBase.mobjLog.Add "获取标准数据失败", 1
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_KEY(ByVal RHS As String)
    mobjBase.KEY = RHS
End Property

Private Property Get clsINF_KEY() As String
    clsINF_KEY = mobjBase.KEY
End Property

Private Property Set clsINF_mobjComLib(ByVal RHS As Object)
    Set mobjBase.mobjComLib = RHS
End Property

Private Property Get clsINF_mobjComLib() As Object
    Set clsINF_mobjComLib = mobjBase.mobjComLib
End Property

Private Property Set clsINF_mobjLog(ByVal RHS As clsLog)
    Set mobjBase.mobjLog = RHS
End Property

Private Property Get clsINF_mobjLog() As clsLog
    Set clsINF_mobjLog = mobjBase.mobjLog
End Property

Private Property Let clsINF_mstrNOs(ByVal RHS As String)
    mobjBase.mstrNOs = RHS
End Property

Private Property Get clsINF_mstrNOs() As String
    clsINF_mstrNOs = mobjBase.mstrNOs
End Property

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, Optional strWinInfo As String) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的SendData方法

    Dim strSQL As String
    Dim colProc As New Collection
    Dim intSN As Integer, intMacNo As Integer
    Dim lng病人id As Long

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "开始传送数据", 1
    
    Select Case intType
    Case 21
        rsData.Filter = "是否传送=1"
    Case Else
        mobjBase.mobjLog.Add "功能不支持", 1
        GoTo makExit
    End Select
    
    If rsData.RecordCount <= 0 Then
        mobjBase.mobjLog.Add "无数据可传送", 1
makExit:
        mobjBase.mobjLog.Save
        clsINF_SendData = True
        Exit Function
    End If
    
    '个性化处理（选择机器号）
    If mdlDefine.gtypYUYAMA.BottomLevel Then
        Call frmINF_YUYAMA_MacNo.ShowMe(intMacNo)
    Else
        intMacNo = mdlDefine.gtypYUYAMA.MacNO
    End If
    
    With rsData
        intSN = 1
        Do While .EOF = False
            '该存储过程的名称固定
            Select Case intType
            Case Val("21-配药")
                '序号处理
                If lng病人id <> mdlDrugMachine.NVL(!病人id) Then
                    intSN = 1
                End If
                
                '存储过程
                strSQL = "ZL_药品设备接口_YYM_INS("
                '处方号
                strSQL = strSQL & "To_Char(Sysdate, 'MMDDHH') || Replace(To_Char(" & mdlDrugMachine.NVL(!病人id) & ", '9999999999999'), ' ', '0')"
                '序号
                strSQL = strSQL & ", " & intSN
                '小组编号
                strSQL = strSQL & ", 1"
                '机器编号（0：所有；1-n：指定机器编号包药）
                strSQL = strSQL & ", " & intMacNo
                '处理状态（0：未处理；1：正常；2：异常；3：取消处理；-1：处理中））
                strSQL = strSQL & ", 0"
                '病人ID
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!病人id, 0) & "'"
                '病人姓名
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!病人姓名) & "'"
                '门诊与住院区分（1：门诊；2：住院长嘱；3：住院临嘱）
                strSQL = strSQL & ", " & IIf(mdlDrugMachine.NVL(!医嘱期效, 0) = 0, "'2'", "'3'")
                '门诊（科室编码）；住院（病区编码）
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!病区id, 0) & "'"
                '门诊（科室名称）；住院（病区名称）
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!病区, 0) & "'"
                '输入处方日期
                strSQL = strSQL & ", To_Date('" & Format(mdlDrugMachine.NVL(!开嘱时间, "2000/1/1"), "yyyy-mm-dd hh:mm:ss") & "', 'yyyy-mm-dd hh24:mi:ss')"
                '紧急类别（0：不紧急；1-9：紧急，由缓到急）
                strSQL = strSQL & ", " & mdlDrugMachine.NVL(!紧急标志, 0)
                '药品编码
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!药品id, 0) & "'"
                '药品名称
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!药品名称) & "'"
                '摆药单位
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!住院单位) & "'"
                '天数或次数
                strSQL = strSQL & ", 1"     '1次1条记录
                '用法
                strSQL = strSQL & ", '" & mdlDrugMachine.NVL(!用法) & "'"
                '服用时间
                strSQL = strSQL & ", To_Date('" & Format(mdlDrugMachine.NVL(!执行时间, "2000/1/1"), "yyyy-mm-dd hh:mm:ss") & "', 'yyyy-mm-dd hh24:mi:ss')"
                'HIS的发药时间
                strSQL = strSQL & ", Sysdate"
                strSQL = strSQL & ")"
                
            Case Else
                strSQL = ""
            End Select
            
            If strSQL <> "" Then
                mdlDrugMachine.AddArray colProc, strSQL
            End If
            
            lng病人id = mdlDrugMachine.NVL(!病人id)
            
            .MoveNext
            intSN = intSN + 1
        Loop
    End With
    
    '传送数据至存储过程
    On Error GoTo makFail
    Call mdlDrugMachine.ExecuteProcedureBeach(colProc, "", mcnThird, mobjBase.mobjLog)
    On Error GoTo hErr
    
    '标记数据传送成功
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, True)
    
    mobjBase.mobjLog.Add "完成传送数据", 1
    mobjBase.mobjLog.Save
    
    '执行基类的代码
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
makFail:
    '标记数据传送失败
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, False)
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "传送数据失败", 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
    '
End Property

Private Property Get clsINF_Token() As String
    '
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
    
    Set mcnThird = New ADODB.Connection
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property

