VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPatient"
Attribute VB_Ext_KEY = "Member0" ,"cPatient"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'局部变量，保存集合
Private mCol As Collection
Public mSeatings As Seatings
Public Enum PATI_COLUM
    col_排队状态 = 0
    col_图标 = 1
    col_calling = 2
    col_姓名 = 3
    
    col_接受时间 = 4
    col_顺序号 = 5
    col_穿刺台 = 6
    col_挂号单 = 7
    col_主页ID = 8
    col_座位号 = 9
    col_就诊卡号 = 10
    
    col_加权号 = 11
    col_性别 = 12
    col_年龄 = 13
    
    col_病人ID = 14
    col_门诊号 = 15
    col_病人科室 = 16
    col_费别 = 17
    col_门诊诊断 = 18
    col_医生 = 19
    col_身份证号 = 20
    col_IC卡号 = 21
    col_挂号时间 = 22
    col_开始时间 = 23
    col_开始操作员 = 24
    col_结束时间 = 25
    col_结束操作员 = 26
    col_key = 27
End Enum

Private blnRptInit As Boolean

Public Function ClearPatients(ByRef Rpt As Object) As Boolean
'清除集合及指定控件中的数据
    Dim lngIndex As Long
    
    For lngIndex = 1 To Me.Count
        Remove 1
    Next
    
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            Rpt.Records.DeleteAll
        End If
    End If
End Function

Public Function FetchPatients(ByVal deptID As Long, ByVal BeginTime As Date, ByVal EndTime As Date, _
        Optional ByVal blnIsReady As Boolean = False, Optional ByVal strInfo As String, Optional ByRef Rpt As Object, _
        Optional ByRef img As Object, _
        Optional ByRef objSquareCard As Object)
'从数据库中取得数据加入集合，并将数据填到指定控件中
'目前只支持ReportControl
'写入前会调用ClearPatients过程。
'参数：
'  strInfo：提取指定病人信息；格式： 类型(1-就诊卡;2-门诊号;3-单据号;4-姓名;5-身份证;6-IC卡)或一卡能|卡号|一卡通卡类别ID

    Dim blnToRpt As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, objPati As cPatient
    Dim str诊断 As String, rsZD As ADODB.Recordset, strPar As String, strType As String, i As Integer
    Dim strSeatKey As String, strTmp As String, strTmpA As String, strTmpB As String
    Dim strVal As String
    Dim intType As Integer
    Dim lng病人ID As Long, lngCardType As Long
    Dim dateTmp As Date
    Dim arrVal As Variant
    Dim strWatchSQL As String

    On Error GoTo errHandle
    
    strPar = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
    For i = 0 To 3
        strType = strType & IIf(Val(Split(strPar, ",")(i)) = 1, "," & i, "")
    Next
    
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            blnToRpt = True
        End If
    End If
    If EndTime < BeginTime Then Exit Function
    
    If Trim(strInfo) = "" And blnIsReady = False Then
        '刷新数据
        dateTmp = DateAdd("d", -3, BeginTime)
    Else
        '接单或搜索指定单号
        arrVal = Split(strInfo, "|")
        intType = Val(arrVal(0))        '所有卡的分类
        strVal = arrVal(1)              '卡号
        Select Case intType
            Case 2
                strTmpA = " And c.门诊号=[5] "
            Case 3
                strTmpA = " And a.挂号单=[5] "
            Case 4
                strTmpA = " And c.姓名=[5] "
            Case 5
                strTmpA = " And c.身份证号=[5] "
            Case 1, 6 To 99
                '1-就诊卡；6-99医疗卡；病人信息表不一定能通过IC卡找到病人，此处通过公共函数得到病人ID
                lngCardType = Val(arrVal(2))    '一卡通类别ID
                If objSquareCard Is Nothing Then
                    strTmpA = " And c.IC卡号=[5] "
                Else
                    strTmpA = " And c.病人id=[6] "
                    
                    '取卡类别
                    If objSquareCard.zlGetPatiID(lngCardType, UCase(strVal), , lng病人ID) = False Then lng病人ID = 0
                End If
            Case Else
                strTmpA = " And a.挂号单 is null "
        End Select
    End If
    
    
    '初始化座位类
    mSeatings.getSeatings (deptID)
    
    Call ClearPatients(Rpt)
           
'    '未接单-门诊留观
'    strWatchSQL = "union all " & vbCr & _
'                  "select b.病人id, c.门诊号, c.就诊卡号, c.费别, c.姓名, c.性别, c.年龄, b.开始执行时间, b.开嘱时间, c.就诊诊室, b.开嘱医生," & vbCr & _
'                  "    b.主页id, null 挂号id, null 挂号单, e.名称 || '[' || e.编码 || ']' As 病人科室, c.身份证号, c.Ic卡号, 1 病人来源 " & vbCr & _
'                  "from 诊疗项目目录 F, 部门表 E, 病案主页 d, 病人信息 c, 病人医嘱记录 B, 病人医嘱发送 A " & vbCr & _
'                  "where b.id=a.医嘱id and b.诊疗项目id=f.id and b.执行科室id = e.id and b.病人id = d.病人id and b.主页id= d.主页id and b.病人id = c.病人id " & vbCr & _
'                  "    and (f.执行分类 is null or f.执行分类 in (0,1,2,3)) And d.病人性质 = 1 and d.出院日期 is null  " & vbCr & _
'                  "    and a.发送时间 Between [2] and [3] And a.执行部门id = [1] " & strTmpC1 & vbCr
'
'    '未接单。指定医嘱发送日期范围的病人信息；如果“接单”，医嘱发送日期范围为一年内
'    strSQL = "Select a.*, f_List2str(Cast(Collect(b.诊断描述) As t_Strlist), '，') 诊断描述 " & _
'             "From ( " & _
'             "Select Distinct *" & vbNewLine & _
'             "From (Select b.病人id, Nvl(b.门诊号, 0) As 门诊号, Nvl(b.就诊卡号, '无') 就诊卡号, b.身份证号, b.Ic卡号, Nvl(b.费别, '无') 费别, b.姓名," & vbNewLine & _
'             "              Nvl(b.性别, '未知') 性别, Nvl(b.年龄, '未知') 年龄, To_Char(Nvl(d.日期,B.执行时间), 'yyyy-MM-dd HH24:MI:SS') As 接受时间,b.登记时间 as 挂号时间, " & vbNewLine & _
'             "              Nvl(b.执行人, '无') 医生, Nvl(b.挂号单, '无') 挂号单, Nvl(c.类别, 0) 座位类别, Nvl(c.编号, '无') As 座位号, LPad(To_Char(Nvl(d.顺序, 0)), 3, '0') 顺序," & vbNewLine & _
'             "              Decode(Nvl(d.加权号, 0), 0, Replace(Nvl(顺序, 0), '_', ''), d.加权号) As 加权号," & vbNewLine & _
'             "              Decode(Nvl(d.状态, -1), 1, '1-待配液', 2, '2-弃号', 3, '3-退号', 4, '4-结束', 5, '5-待穿刺', 6, '6-待执行', 7, '7-执行中','0-未接单') As 排队状态, " & vbNewLine & _
'             "              Nvl(b.病人科室, '无') 病人科室, b.主页id, b.挂号id, d.呼叫标志, d.开始时间, d.开始操作员, d.结束时间, d.结束操作员,d.穿刺台,b.病人来源 " & vbNewLine & _
'             "       From (Select 病人id, 日期, To_Char(日期, 'hhmmss') || '_' || LPad(To_Char(Nvl(顺序号, 0)), 3, '0') As 顺序, 加权号, 状态, 呼叫标志, 开始时间," & vbNewLine & _
'             "                     开始操作员, 挂号单, 结束操作员, 结束时间,穿刺台 " & vbNewLine & _
'             "              From 排队记录" & vbNewLine & _
'             "              Where 日期 Between [2] And [3] And 科室id = [1]) D," & vbNewLine & _
'             "            (Select 病人id, 编号, 类别 From 座位状况记录 Where 科室id = [1]) C," & vbNewLine & _
'             "            (Select b.病人id, c.门诊号, c.就诊卡号, c.费别, c.姓名, c.性别, c.年龄, d.执行时间, d.登记时间, c.就诊诊室, d.执行人, " & vbNewLine & _
'             "                  null 主页id, d.Id As 挂号id, b.挂号单, e.名称 || '[' || e.编码 || ']' As 病人科室, c.身份证号, c.Ic卡号, 0 病人来源 " & vbNewLine & _
'             "              From 诊疗项目目录 F, 部门表 E, 病人挂号记录 D, 病人信息 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
'             "              Where Instr([4], Nvl(f.执行分类, 0)) > 0 And b.诊疗项目id = f.Id And d.执行部门id = e.Id And b.病人id = c.病人id And" & vbNewLine & _
'             "                    a.医嘱id = b.Id And b.挂号单 = d.No And b.病人来源 = 1 And d.记录性质 = 1 And d.记录状态 = 1 And" & vbNewLine & _
'             "                    a.发送时间 Between [2] And [3] And a.执行部门id = [1]" & vbCr & _
'             strTmpA & vbCr & strWatchSQL & _
'             "      ) B " & vbNewLine & _
'             "Where b.病人id = c.病人id(+) And b.病人id = d.病人id(+) And b.挂号单 = d.挂号单(+) And d.挂号单(+) Is Null "

    '未接单（门诊、门诊留观）。指定医嘱发送日期范围的病人信息；如果“接单”，医嘱发送日期范围为一年内
    ''说明：Nvl(D2.病人性质, -1) 表示如果值为Null，就是门诊病人
    strSQL = "Select a.*, f_List2str(Cast(Collect(b.诊断描述) As t_Strlist), '，') 诊断描述 " & _
             "From ( " & _
             "Select Distinct *" & vbNewLine & _
             "From (Select b.病人id, Nvl(b.门诊号, 0) As 门诊号, Nvl(b.就诊卡号, '无') 就诊卡号, b.身份证号, b.Ic卡号, Nvl(b.费别, '无') 费别, b.姓名," & vbNewLine & _
             "          Nvl(b.性别, '未知') 性别, Nvl(b.年龄, '未知') 年龄, To_Char(Nvl(d.日期,B.执行时间), 'yyyy-MM-dd HH24:MI:SS') As 接受时间,b.登记时间 as 挂号时间, " & vbNewLine & _
             "           Nvl(b.执行人, '无') 医生, Nvl(b.挂号单, '无') 挂号单, Nvl(c.类别, 0) 座位类别, Nvl(c.编号, '无') As 座位号, LPad(To_Char(Nvl(d.顺序, 0)), 3, '0') 顺序," & vbNewLine & _
             "           Decode(Nvl(d.加权号, 0), 0, Replace(Nvl(顺序, 0), '_', ''), d.加权号) As 加权号," & vbNewLine & _
             "           Decode(Nvl(d.状态, -1), 1, '1-待配液', 2, '2-弃号', 3, '3-退号', 4, '4-结束', 5, '5-待穿刺', 6, '6-待执行', 7, '7-执行中','0-未接单') As 排队状态, " & vbNewLine & _
             "           Nvl(b.病人科室, '无') 病人科室, b.主页id, b.挂号id, d.呼叫标志, d.开始时间, d.开始操作员, d.结束时间, d.结束操作员,d.穿刺台,b.病人来源 " & vbNewLine & _
             "      From (Select 病人id, 主页id, 日期, To_Char(日期, 'hhmmss') || '_' || LPad(To_Char(Nvl(顺序号, 0)), 3, '0') As 顺序, " & vbNewLine & _
             "                加权号, 状态, 呼叫标志, 开始时间, 开始操作员, 挂号单, 结束操作员, 结束时间,穿刺台 " & vbNewLine & _
             "            From 排队记录 " & vbNewLine & _
             "            Where 日期 Between [2] And [3] And 科室id = [1]) D," & vbNewLine & _
             "      (Select 病人id, 编号, 类别 From 座位状况记录 Where 科室id = [1]) C," & vbNewLine & _
             "      (Select b.医嘱id,a.病人id, c.门诊号, c.就诊卡号, c.费别, c.姓名, c.性别, c.年龄, Decode(D1.执行时间, Null, a.开始执行时间, D1.执行时间) 执行时间," & vbNewLine & _
             "           Decode(D1.登记时间, Null, a.开嘱时间, D1.登记时间) 登记时间, c.就诊诊室, decode(D1.执行人, null, a.开嘱医生, D1.执行人) 执行人," & vbNewLine & _
             "           D2.主页id, D1.Id 挂号id, a.挂号单, f.名称 || '[' || f.编码 || ']' As 病人科室, c.身份证号, c.Ic卡号, " & vbNewLine & _
             "           Decode(Nvl(D2.主页id, 0), 0, 0, 1) 病人来源, Nvl(D2.病人性质, -1) 病人性质 " & vbNewLine & _
             "       From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 病人挂号记录 D1, 病案主页 D2, 诊疗项目目录 E, 部门表 F " & vbNewLine & _
             "       Where a.Id = b.医嘱id And a.病人id = c.病人id And a.挂号单 = D1.No(+) And a.病人id = D2.病人id(+) And a.主页id = D2.主页id(+) " & vbNewLine & _
             "           And a.诊疗项目id = e.Id And a.执行科室id = f.Id And a.病人来源 In (1, 2) " & vbNewLine & _
             "           And b.执行部门id = [1] And b.发送时间 between [2] and [3] " & vbNewLine & _
             "           And D1.记录性质(+) = 1 And D1.记录状态(+) = 1 " & vbNewLine & _
             "           And Instr([4], Nvl(e.执行分类, 0)) > 0 "
    strSQL = strSQL & strTmpA & vbCr & ") B " & vbCr & _
             "Where b.病人id = c.病人id(+) And b.病人id = d.病人id(+) And b.挂号单 = d.挂号单(+) And nvl(b.主页id,0)=nvl(d.主页id(+),0) And b.病人性质 in (-1,1) "  '即门诊和门诊留观
    

'    '已接单-门诊留观
'    strWatchSQL = "union all " & vbCr & _
'                  "select c.病人id, b.门诊号, Nvl(b.就诊卡号, '无') 就诊卡号, b.身份证号, b.ic卡号, nvl(b.费别, '无'), b.姓名,Nvl(b.性别, '未知') ,Nvl(b.年龄, '未知')," & vbCr & _
'                  "    To_Char(c.日期, 'yyyy-MM-dd HH24:MI:SS'),a.登记时间, a.门诊医师, Nvl(c.挂号单, '无'), Nvl(e.类别, 0), Nvl(e.编号, '无')," & vbCr & _
'                  "    LPad(To_Char(Nvl(c.顺序号, 0)), 3, '0'), To_Char(Nvl(c.加权号, 0))," & vbCr & _
'                  "    Decode(Nvl(c.状态, -1), 1, '1-待配液', 2, '2-弃号', 3, '3-退号', 4, '4-结束', 5, '5-待穿刺', 6, '6-待执行', 7, '7-执行中', '0-未接单')," & vbCr & _
'                  "    Nvl(d.名称 || '[' || d.编码 || ']', '无'), a.主页id, null 挂号id, c.呼叫标志, c.开始时间, c.开始操作员, c.结束时间, c.结束操作员, c.穿刺台, 1 病人来源 " & vbCr & _
'                  "from 病案主页 a, 病人信息 b, 排队记录 c, 部门表 d, 座位状况记录 e " & vbCr & _
'                  "where a.病人id = b.病人id And a.主页id = b.主页id and b.病人id = c.病人id and c.病人id= e.病人id(+) And c.科室id = e.科室id(+) and a.入院科室id=d.id " & vbCr & _
'                  IIf(Trim(strInfo) = "" And blnIsReady = False, _
'                        " And (c.日期 Between [5] And [2] and Instr('1,5,6,7', c.状态) > 0 Or c.日期 between [2] and [3]) ", _
'                        " And c.日期 Between [2] And [3] ") & vbCr & _
'                  "    and c.科室id = [1] " & strTmpC2
'
'    '已接单。三天内的病人排队记录（1,5,6,7)和指定科室、指定日期的病人排队记录
'    strSQL = strSQL & " Union all " & vbNewLine & _
'             "Select c.病人id, nvl(b.门诊号,0) 门诊号, nvl(b.就诊卡号, '无') 就诊卡号, b.身份证号, b.ic卡号, nvl(b.费别, '无') 费别, " & _
'             "    b.姓名, nvl(b.性别, '未知') 性别, nvl(b.年龄, '未知') 年龄, To_Char(c.日期, 'yyyy-MM-dd HH24:MI:SS') As 接受时间, " & _
'             "    a.登记时间 挂号时间, nvl(a.执行人, '无') 医生, nvl(c.挂号单, '无') 挂号单, Nvl(e.类别, 0) 座位类别, Nvl(e.编号, '无') As 座位号, " & _
'             "    LPad(To_Char(nvl(c.顺序号, 0)), 3, '0') 顺序, To_Char(nvl(c.加权号, 0)) 加权号, " & _
'             "    Decode(Nvl(c.状态, -1), 1, '1-待配液', 2, '2-弃号', 3, '3-退号', 4, '4-结束', 5, '5-待穿刺', 6, '6-待执行', 7, '7-执行中', '0-未接单') As 排队状态, " & _
'             "    Nvl(d.名称 || '[' || d.编码 || ']', '无') 病人科室, null 主页id, a.id 挂号id, c.呼叫标志, c.开始时间, c.开始操作员, c.结束时间, c.结束操作员, c.穿刺台, 0 病人来源 " & vbNewLine & _
'             "From 病人挂号记录 A, 病人信息 B, 排队记录 C, 部门表 D, 座位状况记录 E " & vbNewLine & _
'             "Where a.No = c.挂号单 And a.执行部门id = d.Id And b.病人id = c.病人id And c.病人id = e.病人id(+) And c.科室id = e.科室id(+) " & _
'             "    And a.执行状态 In (1, 2) And a.记录性质 = 1 And a.记录状态 = 1 " & _
'             IIf(Trim(strInfo) = "" And blnIsReady = False, _
'                " And ( c.日期 Between [5] And [2] And instr('1,5,6,7',c.状态)>0 Or c.日期 Between [2] And [3] ) ", _
'                " And c.日期 Between [2] And [3] ") & _
'             "    And c.科室id = [1] " & strTmpB & vbNewLine & _
'             strWatchSQL

    '已接单（门诊、门诊留观）。三天内的病人排队记录（1,5,6,7)和指定科室、指定日期的病人排队记录
    strSQL = strSQL & vbNewLine & _
             "Union all " & vbNewLine & _
             "Select a.病人id, nvl(c.门诊号,0) 门诊号, nvl(c.就诊卡号, '无') 就诊卡号, c.身份证号, c.Ic卡号, nvl(c.费别, '无') 费别, c.姓名, nvl(c.性别, '未知') 性别, nvl(c.年龄, '未知') 年龄," & vbNewLine & _
             "    To_Char(a.日期, 'yyyy-MM-dd HH24:MI:SS') 接受时间, Decode(b.登记时间, Null, d.登记时间, b.登记时间) 挂号时间," & vbNewLine & _
             "    nvl(b.执行人, d.门诊医师) 医生, nvl(a.挂号单, '无') 挂号单, Nvl(e.类别, 0) 座位类别, nvl(e.编号, '无') 座位号," & vbNewLine & _
             "    LPad(To_Char(Nvl(a.顺序号, 0)), 3, '0') 顺序, To_Char(Nvl(a.加权号, 0)) 加权号," & vbNewLine & _
             "    Decode(Nvl(a.状态, -1), 1, '1-待配液', 2, '2-弃号', 3, '3-退号', 4, '4-结束', 5, '5-待穿刺', 6, '6-待执行', 7, '7-执行中', '0-未接单') As 排队状态," & vbNewLine & _
             "    Decode(F1.名称, Null, F2.名称 || '[' || F2.编码 || ']', F1.名称 || '[' || F1.编码 || ']') 病人科室, a.主页id, b.id 挂号id, a.呼叫标志," & vbNewLine & _
             "    a.开始时间, a.开始操作员, a.结束时间, a.结束操作员, a.穿刺台, Decode(Nvl(a.主页id, 0), 0, 0, 1) 病人来源 " & vbNewLine & _
             "From 排队记录 A, 病人挂号记录 B, 病人信息 C, 病案主页 D, 座位状况记录 E, 部门表 F1, 部门表 F2 " & vbNewLine & _
             "Where a.病人id = b.病人id(+) And a.挂号单 = b.NO(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.病人id = c.病人id " & vbNewLine & _
             "    And a.科室id = e.科室id(+) And a.病人id = e.病人id(+) And b.执行部门id = F1.Id(+) And d.入院科室id = F2.Id(+) " & vbNewLine & _
             "    And b.记录状态(+) = 1 And b.记录性质(+) = 1 And Instr(',1,2', b.执行状态(+)) > 0 " & _
             IIf(Trim(strInfo) = "" And blnIsReady = False, _
                 " And (a.日期 between [5] And [2] And instr('1,5,6,7', a.状态) > 0 or a.日期 between [2] and [3])", _
                 " And a.日期 Between [2] And [3] ") & _
             "    And a.科室id = [1] " & strTmpA & vbNewLine

    strSQL = strSQL & _
             ") ) A, 病人诊断记录 B " & vbNewLine & _
             "Where a.病人id=b.病人id(+) and a.挂号id = b.主页id(+) " & _
             "Group by a.病人id, a.门诊号, a.就诊卡号, a.身份证号, a.Ic卡号, a.费别, a.姓名, a.性别, a.年龄, a.接受时间, a.挂号时间, a.医生, a.挂号单, a.座位类别" & _
             "  , a.座位号, a.顺序, a.加权号, a.排队状态, a.病人科室, a.主页id, a.挂号id, a.病人来源, a.呼叫标志, a.开始时间,a.开始操作员, a.结束时间, a.结束操作员, a.穿刺台 " & _
             "Order by a.排队状态, a.接受时间 "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FetchPatients", deptID, BeginTime, EndTime, _
                            strType, IIf(Trim(strInfo) = "" And blnIsReady = False, dateTmp, strVal), lng病人ID)
    
    If blnToRpt And Not blnRptInit Then
        Call initObjRpt(Rpt, img)
    End If
    
    Do Until rsTmp.EOF
        Set objPati = New cPatient
        With rsTmp

            str诊断 = zlCommFun.NVL(.Fields("诊断描述"), "(无)")
            
            objPati.病人ID = .Fields("病人ID")
            objPati.病人科室 = .Fields("病人科室")
            objPati.费别 = .Fields("费别")
            objPati.挂号单 = zlCommFun.NVL(.Fields("挂号单"))
            objPati.就诊卡号 = .Fields("就诊卡号")
            
            objPati.病人来源 = .Fields("病人来源")
            If objPati.病人来源 = 1 Then
                objPati.单据ID = zlCommFun.NVL(.Fields("主页ID"), 0)
            Else
                objPati.单据ID = zlCommFun.NVL(.Fields("挂号ID"), 0)
            End If
            
            objPati.接受时间 = CDate(zlCommFun.NVL(.Fields("接受时间"), 0))
            objPati.结束操作员 = "" & zlCommFun.NVL(.Fields("开始操作员"))
            objPati.结束时间 = CDate(IIf("" & .Fields("开始时间") = "", 0, "" & .Fields("开始时间")))
            objPati.开始操作员 = "" & zlCommFun.NVL(.Fields("结束操作员"))
            objPati.开始时间 = CDate(IIf("" & .Fields("结束时间") = "", 0, "" & .Fields("结束时间")))
            objPati.穿刺台 = "" & zlCommFun.NVL(.Fields("穿刺台"))

            objPati.门诊号 = CStr(zlCommFun.NVL(.Fields("门诊号"), "无"))
            objPati.门诊诊断 = str诊断
            objPati.年龄 = .Fields("年龄")
            objPati.顺序号 = Format(.Fields("接受时间"), "yyMMddHHmmss") & "_" & .Fields("顺序")
            
            objPati.姓名 = "" & zlCommFun.NVL(.Fields("姓名"))
            objPati.性别 = "" & zlCommFun.NVL(.Fields("性别"))
            objPati.医生 = "" & zlCommFun.NVL(.Fields("医生"))
            objPati.排队状态 = .Fields("排队状态")
            objPati.座位号 = .Fields("座位号")
            objPati.加权号 = .Fields("加权号")
            
            objPati.呼叫标志 = Val("" & .Fields("呼叫标志"))
            
            objPati.身份证号 = "" & zlCommFun.NVL(.Fields("身份证号"))
            objPati.IC卡号 = "" & zlCommFun.NVL(.Fields("IC卡号"))
            
            objPati.挂号时间 = CDate(zlCommFun.NVL(.Fields("挂号时间"), 0))
            
            If "" & .Fields("座位号") <> "" And "" & .Fields("座位号") <> "无" Then
                strSeatKey = "" & .Fields("座位类别") & "_" & .Fields("座位号")
                strTmp = Trim$("" & .Fields("开始操作员"))
                If strTmp <> "" Then strTmp = strTmp & " " & Format("" & .Fields("开始时间"), "HH:mm")
                    
                mSeatings(strSeatKey).开始时间 = strTmp
                mSeatings(strSeatKey).诊断 = str诊断
                mSeatings(strSeatKey).挂号单 = .Fields("挂号单")
            End If
            
            '载入病人信息对象的集合
            If objPati.病人来源 = 1 Then
                Call Add(objPati, zlStr.FormatString("[1]_[2]", objPati.病人ID, objPati.单据ID))
            Else
                Call Add(objPati, objPati.挂号单)
            End If
            Call AddToRpt(objPati, Rpt)
            
            .MoveNext
        End With
        
    Loop
    If blnToRpt Then Rpt.Populate

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Add(ByVal objPati As cPatient, Optional sKey As String) As cPatient
    '增加对象到集合和rpt控件中
    '创建新对象
    Dim objNewMember As cPatient
    Dim blnAdd As Boolean

    On Error Resume Next
    
    blnAdd = False
    Set objNewMember = mCol(sKey)
    If objNewMember Is Nothing Then
        Set objNewMember = New cPatient
        blnAdd = True
    End If
    
    On Error GoTo 0
    
    '设置传入方法的属性
    objNewMember.病人ID = objPati.病人ID
    objNewMember.病人科室 = objPati.病人科室
    objNewMember.费别 = objPati.费别
    objNewMember.挂号单 = objPati.挂号单
    objNewMember.就诊卡号 = objPati.就诊卡号

    objNewMember.接受时间 = objPati.接受时间
    objNewMember.挂号时间 = objPati.挂号时间
    objNewMember.门诊号 = objPati.门诊号
    objNewMember.门诊诊断 = objPati.门诊诊断
    objNewMember.年龄 = objPati.年龄
    objNewMember.顺序号 = objPati.顺序号

    objNewMember.姓名 = objPati.姓名
    objNewMember.性别 = objPati.性别
    objNewMember.医生 = objPati.医生
    objNewMember.排队状态 = objPati.排队状态
    objNewMember.座位号 = objPati.座位号
    objNewMember.加权号 = objPati.加权号
    objNewMember.呼叫标志 = objPati.呼叫标志
    objNewMember.身份证号 = objPati.身份证号
    objNewMember.IC卡号 = objPati.IC卡号
    objNewMember.穿刺台 = objPati.穿刺台
    objNewMember.开始操作员 = objPati.开始操作员
    objNewMember.结束操作员 = objPati.结束操作员
    objNewMember.开始时间 = objPati.开始时间
    objNewMember.结束时间 = objPati.结束时间
    
    objNewMember.病人来源 = objPati.病人来源
    objNewMember.单据ID = objPati.单据ID

    objNewMember.Key = sKey
    
    If blnAdd Then
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
    End If
    
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Private Sub AddToRpt(ByVal objPati As cPatient, ByRef Rpt As Object)
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim intIcon As Integer
    Dim str顺序号 As String
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            Set objRecord = Rpt.Records.Add
            Call Add_rptItem(objRecord, objPati.排队状态)
            
            
            Set objItem = Add_rptItem(objRecord, "")
            Select Case objPati.座位号
            Case "无"
                intIcon = 3
            Case Else
                intIcon = 0
            End Select
            objItem.Icon = intIcon
            'calling
            Set objItem = Add_rptItem(objRecord, "")
            'objItem.Icon = intIcon
            If objPati.呼叫标志 = 1 Then
                objItem.Icon = 5
            Else
                objItem.Icon = 6
            End If
            
            Call Add_rptItem(objRecord, objPati.姓名)
            
            Call Add_rptItem(objRecord, Format(objPati.接受时间, "yy-MM-dd hh:mm"))
            'If InStr(objPati.顺序号, "_") > 0 Then
            '    str顺序号 = Mid(objPati.顺序号, InStr(objPati.顺序号, "_") + 1)
            'Else
                str顺序号 = objPati.顺序号
            'End If
            Call Add_rptItem(objRecord, IIf(Val(str顺序号) = "0", "", str顺序号))
            Call Add_rptItem(objRecord, objPati.穿刺台)
            
            '2015-05-06，门诊为“挂号单号”；门诊留观为“病人ID_主页ID”
            If objPati.病人来源 = 1 Then
                Call Add_rptItem(objRecord, "")
                Call Add_rptItem(objRecord, objPati.单据ID)
            Else
                Call Add_rptItem(objRecord, objPati.挂号单)
                Call Add_rptItem(objRecord, "")
            End If
            
            Call Add_rptItem(objRecord, objPati.座位号)
            Call Add_rptItem(objRecord, objPati.就诊卡号)
            Call Add_rptItem(objRecord, objPati.加权号)
            Call Add_rptItem(objRecord, objPati.性别)
            Call Add_rptItem(objRecord, objPati.年龄)
            Call Add_rptItem(objRecord, objPati.病人ID)
            Call Add_rptItem(objRecord, objPati.门诊号)
            Call Add_rptItem(objRecord, objPati.病人科室)
            Call Add_rptItem(objRecord, objPati.费别)
            Call Add_rptItem(objRecord, objPati.门诊诊断)
            Call Add_rptItem(objRecord, objPati.医生)
            
            Call Add_rptItem(objRecord, objPati.身份证号)
            Call Add_rptItem(objRecord, objPati.IC卡号)
            Call Add_rptItem(objRecord, objPati.挂号时间)
            
            Call Add_rptItem(objRecord, objPati.开始时间)
            Call Add_rptItem(objRecord, objPati.开始操作员)
            Call Add_rptItem(objRecord, objPati.结束时间)
            Call Add_rptItem(objRecord, objPati.结束操作员)
            Call Add_rptItem(objRecord, objPati.Key)
        End If
        '在此可以添加对其他控件的支持
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As cPatient
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    If Err.Number <> 0 Then
          If TypeName(vntIndexKey) = "String" Then
              SaveOperLog 0, vntIndexKey, QUEUE, "病人类的Item方法"
          End If
    End If
End Property

Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
    Set mSeatings = New Seatings
End Sub

Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
    Set mSeatings = Nothing
End Sub

Private Function Add_rptItem(ByRef objRecord As ReportRecord, ByVal strValues As String) As ReportRecordItem
    
    Set Add_rptItem = objRecord.AddItem(strValues)
    Add_rptItem.Caption = strValues
    
End Function

Public Sub initObjRpt(ByRef rptPati As Object, ByRef img16 As Object)
    
    Dim objCol As ReportColumn

    With rptPati
        .Records.DeleteAll
        .Columns.DeleteAll
        .GroupsOrder.DeleteAll
        .SortOrder.DeleteAll
        
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(col_排队状态, "状态", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_图标, "", 18, False): objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_calling, "", 18, False): objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_姓名, "姓名", 55, True)
        
        Set objCol = .Columns.Add(col_接受时间, "就诊时间", 95, True)
        Set objCol = .Columns.Add(col_顺序号, "顺序号", 60, True)
        Set objCol = .Columns.Add(col_穿刺台, "穿刺台", 35, True)
        Set objCol = .Columns.Add(col_挂号单, "挂号单", 68, True)
        Set objCol = .Columns.Add(col_主页ID, "主页ID", 45, True)
        Set objCol = .Columns.Add(col_座位号, "座位号", 45, True)
        Set objCol = .Columns.Add(col_就诊卡号, "就诊卡号", 60, True)
        
        '隐藏数据列
        Set objCol = .Columns.Add(col_加权号, "加权号", 0, False)
        Set objCol = .Columns.Add(col_性别, "性别", 0, False)
        Set objCol = .Columns.Add(col_年龄, "年龄", 0, False)
        Set objCol = .Columns.Add(col_病人ID, "病人ID", 0, False)
        Set objCol = .Columns.Add(col_门诊号, "门诊号", 0, False)
        Set objCol = .Columns.Add(col_病人科室, "病人科室", 0, False)
        Set objCol = .Columns.Add(col_费别, "费别", 0, False)
        
        Set objCol = .Columns.Add(col_门诊诊断, "门诊诊断", 0, False)
        Set objCol = .Columns.Add(col_医生, "医生", 0, False)
        
        Set objCol = .Columns.Add(col_身份证号, "身份证号", 0, False)
        Set objCol = .Columns.Add(col_IC卡号, "IC卡号", 0, False)
        
        Set objCol = .Columns.Add(col_挂号时间, "挂号时间", 0, False)
        Set objCol = .Columns.Add(col_开始时间, "开始时间", 0, False)
        Set objCol = .Columns.Add(col_开始操作员, "开始操作员", 0, False)
        Set objCol = .Columns.Add(col_结束时间, "结束时间", 0, False)
        Set objCol = .Columns.Add(col_结束操作员, "结束操作员", 0, False)
        Set objCol = .Columns.Add(col_key, "Key", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = col_排队状态
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
            
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList img16
        
        .GroupsOrder.Add .Columns(col_排队状态)
        '.GroupsOrder.Add .Columns(col_接受时间)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_加权号)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_顺序号)
        .SortOrder(1).SortAscending = True
    End With
    blnRptInit = True
    
End Sub
Public Function DeptToCbo(ByRef cboDept As Object, ByVal strPrivs As String) As Boolean

'功能：初始化临床科室

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String
    
    On Error GoTo errH
    
    '包含门诊医技科室
    If Not cboDept Is Nothing Then
        If TypeName(cboDept) <> "ComboBox" Then
            Call Err.Raise(ccTypeMismatch, "cPatients.DeptToCbo", "DeptToCbo函数要求传入ComboBox型控件！")
        End If
    End If
    
    If InStr(strPrivs, "所有科室") > 0 Then
        strSQL = _
            " Select /*+Rule */ Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(1,3) And B.工作性质 IN('治疗','临床')" & _
            " And (A.站点=[2] Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        strSQL = _
            " Select /*+Rule */ Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And B.服务对象 IN(1,3) And B.工作性质 IN('治疗','临床')" & _
            " And (A.站点=[2] Or A.站点 is Null)" & _
            " Order by A.编码"
    End If
    
    cboDept.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "DeptToCbo", UserInfo.ID, zl9ComLib.gstrNodeNo)
    str科室IDs = GetUser科室IDs
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.部门ID Then
            Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex) '直接所属优先
        End If
        If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then
            Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
        End If
        
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    
    
    DeptToCbo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Private Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    If gbln病区科室独立 Then
        GetUser科室IDs = UserInfo.部门ID
    Else
        strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1]"
        If bln病区 Then
            strSQL = strSQL & " Union" & _
                " Select Distinct B.科室ID From 部门人员 A,床位状况记录 B" & _
                " Where A.部门ID=B.病区ID And A.人员ID=[1]"
        End If
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlTransfusion", UserInfo.ID)
        For i = 1 To rsTmp.RecordCount
            GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
            rsTmp.MoveNext
        Next
        GetUser科室IDs = Mid(GetUser科室IDs, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub PatiListRefresh(ByRef Rpt As Object, Optional ByVal strStat As String = "")
    '根据当前数据，填充RPT控件
    'rpt 表格控件
    
    'strStat : 填弃哪一种状态的数据到控件  以，分隔的状态标志
    
    Dim objPati As cPatient, blnToRpt As Boolean
    Dim intLoop As Integer
    
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            blnToRpt = True
        End If
    End If
    If Not blnToRpt Then Exit Sub
    
    If strStat <> "" Then
        If Right(strStat, 1) <> "," Then strStat = strStat & ","
        If Left(strStat, 1) <> "," Then strStat = "," & strStat
    End If
    
    Call Rpt.Records.DeleteAll
    Rpt.Populate
    For intLoop = 1 To Me.Count
        blnToRpt = True
        Set objPati = Me.Item(intLoop)
        If strStat <> "" Then
            If InStr(strStat, "," & Val(objPati.排队状态) & ",") <= 0 Then blnToRpt = False
        End If
        If blnToRpt Then Call AddToRpt(objPati, Rpt)
    Next
    Rpt.Populate
End Sub

