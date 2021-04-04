VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'说明：第三方接口基类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Public mobjComLib As Object
Public mobjLog As clsLog
Public mstrNOs As String        '格式：单据1,处方号1[[;单据2,处方号2]...]

Private mstrURL As String
Private mblnAction As Boolean
Private mstrKey As String
Private mstrToken As String
Private mstrUser As String
Private mstrInfCode As String

'ZLHIS数据库帐户名
Public Property Get UserName() As String
    UserName = mstrUser
End Property
Public Property Let UserName(ByVal strUser As String)
    mstrUser = strUser
End Property

'接口连接状态；True连接；False未连接
Public Property Get Action() As Boolean
    Action = mblnAction
End Property
Public Property Let Action(ByVal blnValue As Boolean)
    mblnAction = blnValue
End Property

'OLEDB连接串；或是WebService服务地址；目前只支持这两类
Public Property Get URL() As String
    URL = mstrURL
End Property
Public Property Let URL(ByVal strValue As String)
    mstrURL = strValue
End Property

'信息交互平台需要的KEY
Public Property Get KEY() As String
    KEY = mstrKey
End Property
Public Property Let KEY(ByVal strKey As String)
    mstrKey = strKey
End Property

'信息交互平台需要的Token
Public Property Get Token() As String
    Token = mstrToken
End Property
Public Property Let Token(ByVal strToken As String)
    mstrToken = strToken
End Property

Public Function GetData(ByVal intType As Integer, ByVal strData As String, ByRef blnFinish As Boolean) As ADODB.Recordset
'功能：获取ZLHIS的标准数据
'参数：
'  intType：业务类别
'           1-部门信息；
'           2-人员信息；
'           3-药品目录；
'           4-药品库存与库位；
'           5-发药窗口；
'           21-配药[门诊和住院处方明细上传]；
'           22-开始发药；
'           23-完成发药；
'           24-处方(完整/部分)退药；
'           25-处方完整退药；
'  strData：业务数据
'           业务类别=1：接口编号|部门性质1;部门性质2;…
'           业务类别=2：接口编号|人员性质1;人员性质2;…
'           业务类别=3：接口编号|剂型编码1;剂型编码2;…
'           业务类别=4：接口编号|库房id1;库房id2;…
'           业务类别=5：接口编号|库房id1;库房id2;…
'           业务类别=21 to 23：
'               门诊：1|单据1,处方号1;单据2,处方号2;…
'               住院：2|收发ID1;收发ID2;…
'           业务类别=24：收费ID1,退药数量1;收费ID2,退药数量2;…   退药数量的单位为ZLHIS药品最小单位
'           业务类别=25：单据1,处方号1,库房id1;单据2,处方号2,库房id2;…
'  blnFinish：True正常完成；False异常中断
'返回：记录集对象

    Dim objDB As Object
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strINF As String, strInfo As String, strDetail As String
    Dim strTmp As String
    Dim intIO As Integer
    Dim arrTmp As Variant, arrInfo As Variant, arrDetail As Variant, arrNew As Variant
    Dim l As Long
    
    mobjLog.Add "准备获取标准数据", 1

    If TypeName(mobjComLib) = "clsComLib" Then
        'ZLHIS、自助机、药品自动化设备管理工具
        Set objDB = mobjComLib.zlDatabase
    ElseIf TypeName(mobjComLib) = "clsPublic" Then
        '支付宝
        Set objDB = mobjComLib
    Else
        mobjLog.Add "对象类型不正确，终止获取标准数据", 1
        mobjLog.Add "" & intType, 1, 1
        mobjLog.Save
        Exit Function
    End If
    
    '检查strData参数
    mobjLog.Add strData, 1, 1
    
    arrTmp = Split(strData, "|")
    Select Case intType
    Case 1 To 5
        If UBound(arrTmp) = 1 Then
            strINF = arrTmp(0)          '接口编码
            strInfo = Trim(arrTmp(1))   '业务数据
        ElseIf UBound(arrTmp) = 0 Then
            strINF = arrTmp(0)
            strInfo = ""
        Else
            mobjLog.Add "“基础数据”参数无内容，终止获取标准数据", 1
            mobjLog.Save
            Exit Function
        End If
    
    Case 21 To 23
        If UBound(arrTmp) = 1 Then
            If Trim(arrTmp(1)) = "" Then
errExit:
                mobjLog.Add "“业务数据”参数无单据信息，终止获取标准数据", 1
                mobjLog.Save
                Exit Function
            Else
                intIO = Val(arrTmp(0))  '门诊、住院
                strInfo = arrTmp(1)     '业务数据
                If intIO < 1 Or intIO > 2 Then
                    mobjLog.Add "“业务数据”参数无法识别门诊与住院，终止获取标准数据", 1
                    mobjLog.Save
                    Exit Function
                End If
            End If
        Else
            GoTo errExit
        End If
    
    Case 24
        If UBound(arrTmp) = 0 Then
            If Trim(arrTmp(0)) = "" Then
                GoTo errExit
            Else
                If Not arrTmp(0) Like "*,*" Then
                    mobjLog.Add "“业务数据”参数无法识别退药数据，终止获取标准数据", 1
                    mobjLog.Save
                    Exit Function
                End If
            End If
            strInfo = arrTmp(0)
            intIO = 1
        Else
            GoTo errExit
        End If
        
    Case 25
        If UBound(arrTmp) = 0 Then
            If Trim(arrTmp(0)) = "" Then
                GoTo errExit
            Else
                If Not arrTmp(0) Like "*,*,*" Then
                    mobjLog.Add "“业务数据”参数无法识别完整退药数据，终止获取标准数据", 1
                    mobjLog.Save
                    Exit Function
                End If
            End If
            
            '将“单据,处方号,库房id;...”转成“处方号,库房id;...”格式
            arrNew = Split(arrTmp(0), ";")
            strTmp = ""
            For l = LBound(arrNew) To UBound(arrNew)
                If arrNew(l) Like "*,*" Then
                    strTmp = strTmp & ";" & Mid(arrNew(l), InStr(arrNew(l), ",") + 1)
                End If
            Next
            If Left(strTmp, 1) = ";" Then strTmp = Mid(strTmp, 2)
            Erase arrNew
            
            strInfo = strTmp
            intIO = 1
        Else
            GoTo errExit
        End If
        
    Case Else
        mobjLog.Add "“业务类别”参数不正确，终止获取标准数据", 1
        mobjLog.Save
        Exit Function
        
    End Select
    
    On Error GoTo hErr
    
    mobjLog.Add "strInfo：" & strInfo, 1
    
    '按业务类别获取标准数据
    Select Case intType
    Case Val("1-部门信息")
        strSQL = _
            "Select b.Id, b.上级id, b.编码, b.名称, b.简码," & vbNewLine & _
            "    f_List2str(Cast(Collect(Cast(a.服务对象 As Varchar2(20))) As t_Strlist)) 服务对象," & vbNewLine & _
            "    f_List2str(Cast(Collect(a.工作性质) As t_Strlist)) 工作性质, " & vbNewLine & _
            "    Decode(To_Char(Nvl(b.撤档时间, To_Date('3000-1-1', 'yyyy-mm-dd')), 'yyyy'), '3000', 0, 1) 停用 " & vbNewLine & _
            "From 部门性质说明 A, 部门表 B " & _
            IIf(strInfo = "", "", ", Table(f_Str2list([1], ';')) C ") & vbNewLine & _
            "Where a.部门id(+) = b.Id " & _
            IIf(strInfo = "", "", " And a.工作性质 = c.Column_Value ") & vbNewLine & _
            "Group By b.Id, b.上级id, b.编码, b.名称, b.简码, b.撤档时间 "
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取部门信息", strInfo)
        mobjLog.Add "获取部门信息完成", 1
        
    Case Val("2-人员信息")
        strSQL = _
            "Select b.Id, b.编号, b.姓名, b.简码, b.出生日期, b.性别, b.民族," & vbNewLine & _
            "    f_List2str(Cast(Collect(a.人员性质) As t_Strlist)) 人员性质," & vbNewLine & _
            "    f_List2str(Cast(Collect(Cast(c.部门id As Varchar2(20))) As t_Strlist)) 部门id," & vbNewLine & _
            "    f_List2str(Cast(Collect(d.编码) As t_Strlist)) 部门编码, " & vbNewLine & _
            "    Decode(To_Char(Nvl(b.撤档时间, To_Date('3000-1-1', 'yyyy-mm-dd')), 'yyyy'), '3000', 0, 1) 停用 " & vbNewLine & _
            "From 人员性质说明 A, 人员表 B, 部门人员 C, 部门表 D " & _
            IIf(strInfo = "", "", ", Table(f_Str2list([1], ';')) E") & vbNewLine & _
            "Where a.人员id(+) = b.Id And b.Id = c.人员id(+) And c.部门id = d.Id " & _
            IIf(strInfo = "", "", " And a.人员性质 = e.Column_Value ") & vbNewLine & _
            "Group By b.Id, b.编号, b.姓名, b.简码, b.出生日期, b.性别, b.民族, b.撤档时间 "
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取人员信息", strInfo)
        mobjLog.Add "获取人员信息完成", 1
        
    Case Val("3-药品目录")
        strSQL = _
            "Select Decode(Nvl(To_Char(a.撤档时间, 'yyyy'), '3000'), '3000', 0, 1) 停用, a.Id, a.编码, a.名称 通用名," & vbNewLine & _
            "    a.规格, a.计算单位, c.药库单位, c.住院单位, c.门诊单位, c.药库包装, c.住院包装, c.门诊包装, c.剂量系数," & vbNewLine & _
            "    c.批准文号, d.药品剂型, d.毒理分类, e.计算单位 剂量单位, f.编码 生产商编码, f.名称 生产商," & vbNewLine & _
            "     Decode(a.类别, '5', '西成药', '6', '中成药', '中草药') 药品类别," & vbNewLine & _
            "     Max(Decode(b.性质, 1, b.简码, Null)) 通用名简码, Max(Decode(b.性质, 3, b.名称, Null)) 商品名," & vbNewLine & _
            "     g.名称 英文名 " & vbNewLine & _
            "From 收费项目目录 A, 收费项目别名 B, 药品规格 C, 药品特性 D, 诊疗项目目录 E, 药品生产商 F, 诊疗项目别名 G " & vbNewLine & _
            IIf(strInfo = "", "", ", Table(f_Str2list([1], ';')) H ") & _
            "Where a.Id = b.收费细目id(+) And a.Id = c.药品id And c.药名id = d.药名id And c.药名id = e.Id " & vbNewLine & _
            "    And a.产地 = f.名称(+) And e.Id = g.诊疗项目id(+) " & _
            IIf(strInfo = "", "", " And d.药品剂型 = h.Column_Value ") & vbNewLine & _
            "    And g.性质(+) = 2 And g.码类(+) = 0 And a.类别 In ('5', '6', '7') And b.码类(+) = 1" & vbNewLine & _
            "Group By a.撤档时间, a.Id, a.编码, a.名称, a.规格, a.计算单位, a.产地, c.药库单位, c.住院单位, c.门诊单位," & vbNewLine & _
            "    c.药库包装, c.住院包装, c.门诊包装, c.剂量系数, c.批准文号, d.药品剂型, d.毒理分类, e.计算单位, a.类别," & vbNewLine & _
            "    f.编码, f.名称, g.名称 "

        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取药品目录", strInfo)
        mobjLog.Add "获取药品目录完成", 1
    
    Case Val("4-药品库存与库位")
        strSQL = _
            "Select a.库房id, a.药品id, a.批次, a.上次批号, a.上次生产日期, a.效期, a.实际数量, b.编码 药品编码, b.计算单位," & vbNewLine & _
            "    c.药库单位, c.住院单位, c.门诊单位, c.药库包装, c.住院包装, c.门诊包装, d.库房货位, e.编码 生产商编码," & vbNewLine & _
            "    e.名称 生产商, f.编码 供应商编码, f.名称 供应商 " & vbNewLine & _
            "From 药品库存 A, 收费项目目录 B, 药品规格 C, 药品储备限额 D, 药品生产商 E, 供应商 F " & _
            IIf(strInfo = "", "", " , Table(f_Num2list([1], ';')) G") & vbNewLine & _
            "Where a.药品id = b.Id And a.药品id = c.药品id And a.药品id = d.药品id(+) And a.库房id = d.库房id(+) And " & vbNewLine & _
            "    a.上次产地 = e.名称(+) And a.上次供应商id = f.Id(+) " & _
            IIf(strInfo = "", "", " And a.库房id = g.Column_Value ")
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取药品库存信息", strInfo)
        mobjLog.Add "获取药品库存信息完成", 1
    
    Case Val("5-发药窗口")
        strSQL = _
            "Select a.编码, a.名称, a.上班否 上班, a.药房id " & vbNewLine & _
            "From 发药窗口 A " & _
            IIf(strInfo = "", "", ", Table(f_Num2list([1], ';')) B") & vbNewLine & _
            "Where 1 = 1 " & _
            IIf(strInfo = "", "", " And a.药房id = b.Column_Value ")
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取药房发药窗口", strInfo)
        mobjLog.Add "获取药房发药窗口完成", 1
    
    Case Val("21-配药")
        If intIO = 1 Then
            strTmp = "门诊配（摆）药"
            strSQL = _
                "Select /*+cardinality(z,10)*/ " & vbNewLine & _
                "    a.No 处方号, a.病人id, a.登记时间 处方时间, a.是否急诊, b.ID 收发ID, b.单据, b.序号, b.库房id, b.药品id, Y1.名称 发药药房," & vbNewLine & _
                "    c.出生日期, a.姓名, a.性别, a.年龄, c.就诊卡号, c.医疗付款方式 医保类别, c.费别, d.优先级, a.发生时间 缴费时间, E1.开嘱科室id," & vbNewLine & _
                "    Y2.名称 开嘱科室, E1.开嘱医生, E1.Id 医嘱id, f.编码, f.规格, f.名称 通用名, g.名称 商品名, f.产地 生产商," & vbNewLine & _
                "    h.门诊单位, Cast(b.实际数量 / h.门诊包装 As Number(16, 5)) 门诊单位数量, b.用法, b.单量 单次用量, i.计算单位 剂量单位," & vbNewLine & _
                "    E1.执行频次, j.药品剂型, j.毒理分类, E1.医生嘱托, l.诊断描述, c.身份, " & vbNewLine & _
                "    Decode(Nvl(m.标志, 0), 11, 1, 12, 1, 0, 1, 0) 是否传送, " & vbNewLine & _
                "    Zl_Fun_Drug_Machine(b.库房id, j.药品剂型, b.ID) 接口编号 " & vbNewLine & _
                "From 门诊费用记录 A, 药品收发记录 B, 病人信息 C, 身份 D, 病人医嘱记录 E1, 收费项目目录 F," & vbNewLine & _
                "    收费项目别名 G, 药品规格 H, 诊疗项目目录 I, 药品特性 J, 病人诊断医嘱 K, 病人诊断记录 L, " & vbNewLine & _
                "    药品收发门诊标志 M, 部门表 Y1, 部门表 Y2, " & vbNewLine & _
                "    Table(f_Str2list2([1], ';', ',')) Z " & vbNewLine & _
                "Where a.Id = b.费用id And a.病人id = c.病人id(+) And c.身份 = d.名称(+) " & vbNewLine & _
                "    And a.医嘱序号 = E1.Id(+) And b.库房id = Y1.Id And E1.开嘱科室id = Y2.Id(+) And a.收费细目id = f.Id(+) " & vbNewLine & _
                "    And f.Id = g.收费细目id(+) And a.收费细目id = h.药品id(+) And E1.相关id = k.医嘱id(+) And h.药名id = i.Id(+) " & vbNewLine & _
                "    And h.药名id = j.药名id(+) And k.诊断id = l.Id(+) " & vbNewLine & _
                "    And b.no = m.处方号(+) And b.单据 = m.单据(+) And b.库房id = m.库房id(+) And m.业务分类(+) = 1 " & vbNewLine & _
                "    And b.单据 = z.C1 And b.No = z.C2 " & vbNewLine & _
                "    And (Mod(b.记录状态, 3) = 0 Or b.记录状态 = 1) And a.收费类别 In ('5', '6', '7') " & vbNewLine & _
                "    And E1.相关id(+) Is Not Null And g.性质(+) = 3 And g.码类(+) = 1 " & vbNewLine & _
                "Order By 接口编号, a.病人id, a.NO, b.单据, b.库房id "
            
        Else
            strTmp = "住院配（包）药"
            strSQL = _
                "Select /*+cardinality(z,10)*/ Distinct Zl_Fun_Drug_Machine(c.库房id, g.药品剂型, c.Id) 接口编号, c.库房id, I1.名称 As 发药药房," & vbNewLine & _
                "    Decode(Nvl(o.标志, 0), 11, 1, 12, 1, 0, 1, 0) 是否传送, b.标识号 住院号, " & vbNewLine & _
                "    c.汇总发药号 As 发药号, Sysdate As 传送时间, c.对方部门id As 领药科室id, I2.名称 As 领药科室, j.病人id, j.姓名 As 病人姓名," & vbNewLine & _
                "    j.性别 As 病人性别, j.年龄 As 病人年龄, j.就诊卡号, b.病人病区id As 病区id, I3.名称 As 病区, j.当前床号 As 床号," & vbNewLine & _
                "    k.优先级, a.开嘱医生, a.开嘱时间, a.Id As 医嘱id, g.药品剂型, c.药品id, e.名称 As 药品名称, f.名称 As 商品名称, c.用法," & vbNewLine & _
                "    c.单量, h.计算单位 As 单量单位, d.住院单位, a.医嘱期效, m.要求时间 As 执行时间, a.医生嘱托, g.毒理分类, a.紧急标志," & vbNewLine & _
                "    Decode(m.要求时间, Null, Round(c.实际数量 / d.住院包装, 2), Round((a.单次用量 / d.剂量系数) / d.住院包装, 2)) As 住院单位数量," & vbNewLine & _
                "    d.剂量系数, a.校对时间 护士校对时间, c.填制人 收发填制人, c.审核人 收发审核人, c.ID 收发ID " & vbNewLine & _
                "From 病人医嘱记录 A, 住院费用记录 B, 药品收发记录 C, 药品规格 D, 收费项目目录 E, 收费项目别名 F, 药品特性 G, 诊疗项目目录 H," & vbNewLine & _
                "    部门表 I1, 部门表 I2, 部门表 I3, 病人信息 J, 身份 K, 病人医嘱发送 L, 医嘱执行时间 M, 病人医嘱记录 N," & vbNewLine & _
                "    药品收发住院标志 O, Table(f_Num2list([1], ';')) Z" & vbNewLine & _
                "Where a.Id = b.医嘱序号 And b.Id = c.费用id And c.药品id = d.药品id And c.药品id = e.Id And e.Id = f.收费细目id(+) " & vbNewLine & _
                "    And d.药名id = g.药名id And g.药名id = h.Id And c.库房id = I1.Id And c.对方部门id = I2.Id And b.病人病区id = I3.Id " & vbNewLine & _
                "    And b.病人id = j.病人id And j.身份 = k.名称(+) And a.相关id = n.Id And n.Id = m.医嘱id " & vbNewLine & _
                "    And b.医嘱序号 = l.医嘱id(+) And b.No = l.No And b.执行部门id = l.执行部门id " & vbNewLine & _
                "    And l.发送号 = m.发送号(+) And c.Id = o.收发id(+) And f.性质(+) = 3 And f.码类(+) = 1 " & vbNewLine & _
                "    And c.Id = z.Column_Value " & vbNewLine & _
                "Order By c.库房id, c.对方部门id, j.病人id"

        End If
        
        mobjLog.Add strSQL, 1, 1
        Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取" & strTmp & "数据", strInfo)
        mobjLog.Add "获取" & strTmp & "数据完成", 1
        
    Case Val("22-开始发药"), Val("23-完成发药")
        If intIO = 1 Then
            If intType = 22 Then
                strTmp = "门诊开始发药"
            Else
                strTmp = "门诊完成发药"
            End If
            
            strSQL = _
                "Select Distinct a.单据, a.No 处方号, b.登记时间 处方时间, a.库房id, a.药品id, b.病人id, d.药品剂型," & _
                "    Decode(Nvl(e.标志, 0), 11, 1, 12, 1, 0, 1, 0) 是否传送, f.编码 窗口编码, " & _
                "    Zl_Fun_Drug_Machine(a.库房id, d.药品剂型) 接口编号 " & vbNewLine & _
                "From 药品收发记录 A, 门诊费用记录 B, 药品规格 C, 药品特性 D, 药品收发门诊标志 E, 发药窗口 F " & vbNewLine & _
                "   , Table(f_Str2list2([1], ';', ',')) Z " & vbNewLine & _
                "Where a.费用id = b.Id And a.药品id = c.药品id And c.药名id = d.药名id(+) " & vbNewLine & _
                "    And a.发药窗口 = f.名称(+) And a.库房id = f.药房id(+) And a.单据 = z.C1 " & vbNewLine & _
                "    And a.NO = e.处方号(+) And a.单据 = e.单据(+) And a.库房id = e.库房id(+) And e.业务分类(+) = [2] " & vbNewLine & _
                "    And a.No = z.C2 And (Mod(a.记录状态, 3) = 0 Or a.记录状态 = 1) " & vbNewLine & _
                "Order By 接口编号, a.NO, a.单据, a.库房id "

            mobjLog.Add strSQL, 1, 1
            Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取" & strTmp & "数据", strInfo, intType - 20)
            mobjLog.Add "获取" & strTmp & "数据完成", 1
        Else
            strTmp = "住院开始发药不支持"
            
            Set rsTmp = Nothing
            mobjLog.Add strTmp, 1
        End If
    
    Case Val("24-处方退药")
        If intIO = 1 Then
            strTmp = "门诊处方退药"
            strSQL = _
                "Select /*+cardinality(h,10)*/ " & vbNewLine & _
                "    Decode(Mod(B.记录性质, 10), 1, 8, 9) 单据, b.No 处方号, b.执行部门id 库房id, b.登记时间 处方时间, " & vbNewLine & _
                "    b.医嘱序号 医嘱id, e.编码 药品编码, " & vbNewLine & _
                "    Cast((b1.数次 * b1.付数 - Decode(b1.数次, 0, 0, 1) * h.C2) / c.门诊包装 As Number(16, 5)) 门诊单位数量,  " & vbNewLine & _
                "    Decode(Nvl(f.标志, 0), 11, 1, 12, 1, 0, 1, 0) 是否传送, " & vbNewLine & _
                "    Zl_Fun_Drug_Machine(b.执行部门id, d.药品剂型, Null) 接口编号 " & vbNewLine & _
                "From 门诊费用记录 B, 门诊费用记录 B1, 药品规格 C, 药品特性 D, 收费项目目录 E, 药品收发门诊标志 F " & vbNewLine & _
                "    , Table(f_Str2list2([1], ';', ',')) H " & vbNewLine & _
                "Where b.NO = b1.NO(+) And b.记录性质 = b1.记录性质(+) And b.医嘱序号 = b1.医嘱序号(+) " & vbNewLine & _
                "    And b.收费细目id = c.药品id And c.药名id = d.药名id(+) And b.收费细目id = e.Id And b.No = f.处方号(+) " & vbNewLine & _
                "    And Decode(Mod(B.记录性质, 10), 1, 8, 9) = f.单据 And b.执行部门id = f.库房id(+) And b.Id = h.C1 " & vbNewLine & _
                "    And b1.记录状态(+) = 3 And f.业务分类(+) = 4 " & vbNewLine & _
                "Order By 接口编号, 处方号, 单据, 库房id "
                
            mobjLog.Add strSQL, 1, 1
            Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取" & strTmp & "数据", strInfo)
            mobjLog.Add "获取" & strTmp & "数据完成", 1
        Else
            strTmp = "住院退药不支持"

            Set rsTmp = Nothing
            mobjLog.Add strTmp, 1
        End If
        
    Case Val("25-处方完整退药")
        If intIO = 1 Then
            strTmp = "处方完整退药"
            strSQL = _
                "Select /*+cardinality(h,10)*/ Distinct " & vbNewLine & _
                "    Decode(Mod(B.记录性质, 10), 1, 8, 9) 单据, b.No 处方号, b.执行部门id 库房id, b.登记时间 处方时间, " & vbNewLine & _
                "    Decode(Nvl(f.标志, 0), 11, 1, 12, 1, 0, 1, 0) 是否传送, " & vbNewLine & _
                "    Zl_Fun_Drug_Machine(b.执行部门id, d.药品剂型, Null) 接口编号 " & vbNewLine & _
                "From 门诊费用记录 B, 药品规格 C, 药品特性 D, 药品收发门诊标志 F, " & vbNewLine & _
                "    Table(f_Str2list2([1], ';', ',')) H " & vbNewLine & _
                "Where b.收费细目id = c.药品id And c.药名id = d.药名id(+) And " & vbNewLine & _
                "    b.No = f.处方号(+) And Decode(Mod(B.记录性质, 10), 1, 8, 9) = f.单据(+) And " & vbNewLine & _
                "    b.执行部门id = f.库房id(+) And b.No = h.C1 And b.执行部门id = h.C2 " & vbNewLine & _
                "Order By 接口编号, 处方号, 单据, 库房id "
            
            mobjLog.Add strSQL, 1, 1
            Set rsTmp = objDB.OpenSQLRecord(strSQL, "获取" & strTmp & "数据", strInfo)
            mobjLog.Add "获取" & strTmp & "数据完成", 1
        Else
            strTmp = "住院退药不支持"

            Set rsTmp = Nothing
            mobjLog.Add strTmp, 1
        End If
    
    End Select
    
    mobjLog.Save
    If Not rsTmp Is Nothing Then
        Set GetData = rsTmp.Clone
    End If
    
    blnFinish = True
    
    Exit Function
    
hErr:
    blnFinish = False
    mobjLog.Add "获取标准数据失败", 1
    mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjLog.Save
End Function

Private Sub Class_Initialize()
    Set mobjLog = New clsLog
End Sub

Private Sub Class_Terminate()
    Set mobjLog = Nothing
End Sub

Public Function FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, ByRef rsFinish As ADODB.Recordset) As Boolean
'功能：过滤出药品自动化接口的数据
'参数：
'  intType：业务类别
'           1-部门信息；
'           2-人员信息；
'           3-药品目录；
'           4-药品库存与库位；
'           5-发药窗口；
'           21-配药[门诊和住院处方明细上传]；
'           22-开始发药；
'           23-完成发药；
'           24-处方(完整/部分)退药；
'  rsData：标准的业务记录集对象
'  rsFinish：已过滤完成的记录集对象
'返回：True完成；False失败
'说明：具体代码请在对应的接口类模块编写

    Dim strTemp As String

    On Error GoTo hErr
    
    mobjLog.Add "开始过滤数据", 1
    
    Set rsFinish = mdlDrugMachine.CopyStructure(rsData.Fields)
    mobjLog.Add "完成记录集结构复制", 1
    
    If intType > 20 Then
        rsData.Filter = "接口编号='" & strCode & "'"
        mobjLog.Add rsData.Filter, 1, 1
    End If
    Do While rsData.EOF = False
        strTemp = CopyRecord(rsData.Fields, rsFinish)
    
        rsData.MoveNext
    Loop
    
    FilterData = True
    
    mobjLog.Add "完成过滤数据", 1
    mobjLog.Save

    Exit Function
    
hErr:
    mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjLog.Save
End Function

Public Function SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, ByRef strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'功能：向药品自动化接口传送数据
'参数：
'  intAppType：应用类别。1-ZLHIS；2-接口管理工具（定时向第三方接口传送数据）；3-移动（支付宝）；4-中联自助平台
'  intType：业务类型；详见 GetData 方法的参数
'  intIO：1-门诊；2-住院
'  strCode：接口编号
'  rsData：已过滤完成的记录集对象
'  strMessage：返回消息
'  strWinInfo：移动（支付宝）需求的
'返回：True完成；False失败
'说明：具体代码请在对应的接口类模块编写

    Dim strMsg As String

    On Error GoTo hErr
    
    mstrInfCode = strCode
    If intAppType = Val("3-移动") And intType = Val("21-收费") And intIO = Val("1-门诊") Then
        '移动（支付宝）缴费，接口返回窗口信息给移动业务
        mobjLog.Add "接口返回窗口信息给移动业务", 1
        strWinInfo = MobilePayment(mobjLog, mstrNOs, strMsg)
        strMessage = IIf(strMessage <> "", strMessage & vbNewLine, "") & strMsg
    End If
    
    mobjLog.Save
    SendData = True
    Exit Function
    
hErr:
    strWinInfo = ""
    mobjLog.Add Err.Description, 1
    mobjLog.Save
End Function

Private Function MobilePayment(ByVal objLog As clsLog, ByVal strNO As String, _
    Optional ByRef strMessage As String) As String
    
'功能：获取传入单据号的病人ID、发药药房、发药窗口、药品名称信息。主要是移动业务在使用
'参数：
'  objLog：日志类对象
'  strNO：单据信息，格式详见调用层说明
'  strMessage（实参）：异常信息
'返回：病人ID、发药药房、发药窗口、药品名称信息的XML结构字符串
'XML格式：
'<OUTPUT>
'  <BRID>病人ID</BRID>
'  <ITEM>
'    <YFMC>药房名称</YFMC>
'    <YFCK>发药窗口</YFCK>
'    <YFMX>
'      <ITEM>
'        <MC>药品名称1</MC>
'      </ITEM>
'      <ITEM>
'        <MC>药品名称2</MC>
'      </ITEM>
'      <ITEM>
'        <MC>药品名称...</MC>
'      </ITEM>
'    </YFMX>
'  </ITEM>
'  <ITEM>
'    ...
'  </ITEM>
'</OUTPUT>

    Const STR_OUT As String = "OUTPUT"
    Const STR_BRID As String = "BRID"
    Const STR_ITEM As String = "ITEM"
    Const STR_YFMC As String = "YFMC"
    Const STR_YFCK As String = "YFCK"
    Const STR_YFMX As String = "YFMX"
    Const STR_MC As String = "MC"

    Dim objXML As New clsXML
    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String, strWin As String
    Dim lngStoreID As Long
    
    On Error GoTo hErr
    
    strSQL = "Select Distinct b.病人id, a.库房id, d.名称 As 药房名称, a.发药窗口, c.名称 As 药品名称 " & vbCr & _
             "From 药品收发记录 A, 门诊费用记录 B, 收费项目目录 C, 部门表 D, Table(f_Str2list2([1], ';', ',')) E " & vbCr & _
             "Where a.费用id = b.Id And a.药品id = c.Id And a.库房id = d.Id And a.单据 = e.C1 And a.No = e.C2 " & vbCr & _
             "Order By Nvl(b.病人id, 0) Desc, a.库房id, c.名称 "
    Set rsSQL = mobjComLib.OpenSQLRecord(strSQL, "获取药房发药信息", strNO)
    With rsSQL
        If .RecordCount > 0 Then
            objXML.AppendNode STR_OUT, False
            objXML.AppendData STR_BRID, mdlDrugMachine.NVL(!病人id)
            
            lngStoreID = -9999
            Do While .EOF = False
                strWin = mdlDrugMachine.NVL(!发药窗口)
                If lngStoreID <> Val(mdlDrugMachine.NVL(!库房id)) Then
                    objXML.AppendNode STR_ITEM, False
                    objXML.AppendData STR_YFMC, mdlDrugMachine.NVL(!药房名称)
                    objXML.AppendData STR_YFCK, strWin
                    
                    objXML.AppendNode STR_YFMX, False
                    GoSub makAdd
                Else
makAdd:
                    objXML.AppendNode STR_ITEM, False
                    objXML.AppendData STR_MC, mdlDrugMachine.NVL(!药品名称)
                    objXML.AppendNode STR_ITEM, True
                End If
                
                lngStoreID = Val(mdlDrugMachine.NVL(!库房id))
                .MoveNext
                
                If .EOF Then
                    GoTo makEnd
                Else
                    If lngStoreID <> Val(mdlDrugMachine.NVL(!库房id)) Then
makEnd:
                        objXML.AppendNode STR_YFMX, True
                        objXML.AppendNode STR_ITEM, True
                    End If
                End If
            Loop
            objXML.AppendNode STR_OUT, True
        End If
        .Close
    End With
    
    MobilePayment = objXML.XmlText
    
    objLog.Add MobilePayment, 1
    objLog.Add "返回“病人ID、发药药房、发药窗口、药品名称”信息成功", 1
    objLog.Save
    
    Set objXML = Nothing
    Exit Function
    
hErr:
    strMessage = "返回“病人ID、发药药房、发药窗口、药品名称”信息失败"
    objLog.Add strNO, 2, 1
    objLog.Add strSQL, 2, 1
    objLog.Add strMessage, 1
    objLog.Save
    MobilePayment = ""
End Function


