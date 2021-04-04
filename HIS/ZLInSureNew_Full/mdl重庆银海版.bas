Attribute VB_Name = "mdl重庆银海版"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

Private Const strFolder As String = "C:\CQYB_YH"        '交换目录
Private Const strRecipe As String = "Recipe.txt"        '处方明细
Private Const strBalance As String = "Balance.txt"      '结算信息
Private Const strDeal As String = "Deal.txt"            '待遇信息
Private Const str门诊处方明细 As String = "Upload.txt"  '用于门诊结算
Private mobjFileSystem As New FileSystemObject
Private mobjStream As TextStream
Public gcn重庆银海版 As New ADODB.Connection
Private mstrBusiness As String
Private mstrInput As String
Private mstrAppMsg As String
Public gstrReturn_重庆银海版 As String                            '全局使用
Public Const gstrSplit_Row_重庆银海版 As String = "$"
Public Const gstrSplit_Col_重庆银海版 As String = "|"
        
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Private gobjYH As Object   '定义存放引用对象的变量。
'Private gobjYH As New clsT_CQYHYB                       '调试用
Private mblnInit As Boolean
Private mstrOwner As String                             '中间库用户名

Private Type ComInfo_重庆银海版
    医院编码 As String
    业务类型 As String
    个人编号 As String
    就诊流水号 As String
    结算流水号 As String
    疾病编码 As String                      '保存身份验证后返回的疾病编码
    并发症 As String
    统筹区号 As String
    帐户余额 As Currency
    总费用 As Currency                      'HIS
    总费用_中心 As Currency                 '中心的费用总额
    就诊时间 As String
    冲销ID As Long
End Type
Public gComInfo_重庆银海版 As ComInfo_重庆银海版

Enum 操作类型_重庆银海版
    待遇信息
    处方明细
    结算信息
    门诊处方明细
End Enum
Private rsRecipe As New ADODB.Recordset                 '用来保存门诊处方

'以下结构体用来纪录虚拟结算结果，用于结算时核对
Private Type typBalance
    cur医保基金 As Double
    cur公务员补助 As Double
    cur个人帐户 As Double
    cur大病基金 As Double
End Type
Private pre_Balance As typBalance

Private Function MakeFile_Recipe(ByVal rsDetail As ADODB.Recordset, Optional ByVal bln门诊 As Boolean = True, _
    Optional ByVal bln预结算 As Boolean = True, Optional ByRef str处方流水号_UP As String) As Boolean
    'str处方流水号_Up:用来记录本次产生了处方明细计算部分的流水号，以","分隔
    Dim intDO As Integer
    Dim lng病人ID As Long, lng主页ID As Long
    Dim bln急诊 As Boolean, bln药品 As Boolean, bln上传 As Boolean, bln血液白蛋白 As Boolean, bln高收费项目 As Boolean
    Dim str业务 As String, str个人编号 As String, str流水号 As String, str结算流水号 As String, str统筹区号 As String
    Dim str处方流水号 As String, str退单处方流水号 As String
    Dim str项目流水号 As String, str项目类别 As String, str收费类别 As String
    Dim str医院项目编码 As String, str医院项目名称 As String
    Dim str单价 As String, str数量 As String, str金额 As String
    Dim str处方号 As String, str开单时间 As String, str开单医生 As String
    '----需要初始化为空----
    Dim str包装数量 As String, str包装单位 As String, str含量 As String, str含量单位 As String, str容量 As String, str容量单位 As String, str剂型 As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsVerify As New ADODB.Recordset
    
    '----与内部映射记录集相关----
    Dim strFields As String, strValues As String
    
    On Error GoTo errHand
    '初始化内部记录集
'    1.      string  18      门诊/住院流水号
'    2.      string  20      个人编号
'    3.      string  20      处方交易流水号
'    4.      string  18      结算交易流水号
'    5.      string  20      退单对应处方交易流水号
'    6.      datetime        秒  开方日期
'    7.      string  14      医保项目流水号
'    8.      string  3       项目类别
'    9.      string  20      医院内码
'    10.     string  50      项目名称
'    11.     number  10  4   单价
'    12.     number  8   2   数量
'    13.     number  10  4   金额
'    14.     string  50      剂型
'    15.     number  8   2   包装数量
'    16.     string  40      包装单位
'    17.     string  8       含量
'    18.     string  40      含量单位
'    19.     string  8       容量
'    20.     string  14      容量单位
'    21.     string  3       急诊标志
'    22.     string  18      处方号
'    23.     string  20      开方医生
'    24.     string  20      经办人
'    25.     datetime        秒  经办时间
'    26.     string  3       统筹区号
    Call DebugTool("创建内部记录集")
    strFields = "流水号," & adLongVarChar & "," & 20 & "|个人编号," & adLongVarChar & "," & 20 & _
                "|处方交易流水号," & adLongVarChar & "," & 20 & "|结算交易流水号," & adLongVarChar & "," & 20 & _
                "|退单处方交易流水号," & adLongVarChar & "," & 20 & "|开方时间," & adLongVarChar & "," & 18 & _
                "|项目流水号," & adLongVarChar & "," & 15 & "|项目类别," & adLongVarChar & "," & 3 & _
                "|医院项目编码," & adLongVarChar & "," & 20 & "|医院项目名称," & adLongVarChar & "," & 50 & _
                "|单价," & adLongVarChar & "," & 18 & "|数量," & adLongVarChar & "," & 18 & _
                "|金额," & adLongVarChar & "," & 18 & "|剂型," & adLongVarChar & "," & 50 & _
                "|包装数量," & adLongVarChar & "," & 18 & "|包装单位," & adLongVarChar & "," & 40 & _
                "|含量," & adLongVarChar & "," & 8 & "|含量单位," & adLongVarChar & "," & 40 & _
                "|容量," & adLongVarChar & "," & 8 & "|容量单位," & adLongVarChar & "," & 14 & _
                "|急诊," & adLongVarChar & "," & 3 & "|处方号," & adLongVarChar & "," & 18 & _
                "|开方医生," & adLongVarChar & "," & 20 & "|经办人," & adLongVarChar & "," & 20 & _
                "|经办时间," & adLongVarChar & "," & 18 & "|统筹区号," & adLongVarChar & "," & 3
    Call Record_Init(rsRecipe, strFields)
    strFields = ""
    For intDO = 0 To rsRecipe.Fields.Count - 1
        strFields = strFields & "|" & rsRecipe.Fields(intDO).Name
    Next
    strFields = Mid(strFields, 2)
    
    '将未上传的处方明细产生为交换文件
    With rsDetail
        '提取该病人的流水号和结算流水号
        Call DebugTool("赋初值")
        lng病人ID = !病人ID
        str流水号 = gComInfo_重庆银海版.就诊流水号
        str结算流水号 = gComInfo_重庆银海版.结算流水号
        str统筹区号 = gComInfo_重庆银海版.统筹区号
        str个人编号 = gComInfo_重庆银海版.个人编号
        str业务 = gComInfo_重庆银海版.业务类型
        str处方流水号_UP = ""
        bln急诊 = (str业务 = "14")
        
        '取主页ID
        gstrSQL = "Select Nvl(住院次数,0) AS 主页ID From 病人信息 Where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
        lng主页ID = rsTemp!主页ID
        
        '产生交换文件
        For intDO = 1 To 2
            Call DebugTool("过滤费用明细，先正记录，后负记录")
            If intDO = 1 Then
                If Not bln门诊 Then
                    .Filter = "金额>0"
                Else
                    .Filter = "实收金额>0"
                End If
            Else
                If Not bln门诊 Then
                    .Filter = "金额<0"
                Else
                    .Filter = "实收金额<0"
                End If
            End If
            Do While Not .EOF
                bln上传 = True      '门诊永远是真
                If Not bln门诊 Then bln上传 = (Nvl(!是否上传, 0) = 0)
                
                If bln上传 Then
                    '取处方流水号及退单处方流水号
                    Call DebugTool("取处方流水号及退单处方流水号")
                    If bln门诊 And bln预结算 Then
                        Call Get处方流水号("", "1", "1", .AbsolutePosition, str处方流水号, str退单处方流水号, lng病人ID)
                    Else
                        If bln门诊 Then
                            Call Get处方流水号(!NO, !记录性质, !记录状态, !序号, str处方流水号, str退单处方流水号, lng病人ID)
                        Else
                            Call Get处方流水号(!NO, !记录性质, !记录状态, !序号, str处方流水号, str退单处方流水号)
                        End If
                    End If
                    str处方流水号_UP = str处方流水号_UP & ",'" & str处方流水号 & "'"
                    
                    '将只有药品项目才有的属性清为空
                    str包装数量 = "": str包装单位 = "": str含量 = "": str含量单位 = "": str容量 = "": str容量单位 = "": str剂型 = ""
                    
                    Call DebugTool("获取医保项目编码与名称")
                    gstrSQL = "Select 类别,编码,名称 From 收费细目 Where ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医院项目的编码与名称", CLng(!收费细目ID))
                    str医院项目编码 = rsTemp!编码
                    str医院项目名称 = rsTemp!名称
                    str收费类别 = rsTemp!类别
                    
                    '取该项目的医保信息
                    Call DebugTool("取该项目的医保信息")
                    If !收费类别 = 5 Or !收费类别 = 6 Or !收费类别 = 7 Then
                        bln药品 = True
                        gstrSQL = " Select 流水号,项目类型,剂型,包装数量,包装单位,含量,含量单位,容量,容量单位" & _
                                  " From " & mstrOwner & ".中间库_药品目录 Where 流水号=" & _
                                  "     (Select 项目编码 From 保险支付项目 " & _
                                  "     Where 收费细目ID=" & !收费细目ID & " And 险类=" & TYPE_重庆银海版 & ")"
                    Else
                        bln药品 = False
                        gstrSQL = " Select 流水号,项目类别 项目类型" & _
                                  " From " & mstrOwner & ".中间库_诊疗项目 Where 流水号=" & _
                                  "     (Select 项目编码 From 保险支付项目 " & _
                                  "     Where 收费细目ID=[1] And 险类=[2])"
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费细目医保相关信息", CLng(!收费细目ID), TYPE_重庆银海版)
                    If rsItem.EOF Then
                        MsgBox "[" & str医院项目编码 & "]" & str医院项目名称 & "行的明细记录未找到对应的保险项目，请检查！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    str项目流水号 = Nvl(rsItem!流水号)
                    str项目类别 = Nvl(rsItem!项目类型)
                    If bln药品 Then
                        str剂型 = Nvl(rsItem!剂型)
                        str包装数量 = Nvl(rsItem!包装数量)
                        str含量 = Nvl(rsItem!含量)
                        str含量单位 = Nvl(rsItem!含量单位)
                        str容量 = Nvl(rsItem!容量)
                        str容量单位 = Nvl(rsItem!容量单位)
                    End If
                    
                    If bln门诊 Then
                        str单价 = Format(!实收金额 / Nvl(!数量, 1), "#####0.0000;-#####0.0000; ;")
                        str数量 = Format(!数量, "#####0.00;-#####0.00; ;")
                        str金额 = Format(!实收金额, "#####0.0000;-#####0.0000; ;")
                        
                        str开单时间 = Format(zlDatabase.Currentdate(), "yyyyMMdd HH:mm:ss")
                        str开单医生 = Nvl(!开单人)
                        str处方号 = gComInfo_重庆银海版.就诊流水号
                    Else
                        str单价 = Format(!金额 / !数量, "#####0.0000;-#####0.0000; ;")
                        str数量 = Format(!数量, "#####0.00;-#####0.00; ;")
                        str金额 = Format(!金额, "#####0.0000;-#####0.0000; ;")
                        
                        str开单时间 = Format(!发生时间, "yyyyMMdd HH:mm:ss")
                        str开单医生 = Nvl(!医生)
                        str处方号 = !NO
                    End If
                    
                    '退费时，其开单时间取原始处方明细的开单时间
                    Call DebugTool("退费时，其开单时间取原始处方明细的开单时间")
                    If str退单处方流水号 <> str处方流水号 Then
                        gstrSQL = "Select to_char(开方日期,'yyyyMMdd hh24:mi:ss') 开方日期 From " & mstrOwner & ".中间库_处方明细 Where 处方流水号=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取原始的开单时间", str退单处方流水号)
                        If Not rsTemp.EOF Then      '如果为空，表示住院记帐作废，单据日期不变
                            str开单时间 = rsTemp!开方日期
                        End If
                    End If
                    
                    '如果是住院，且是高收费项目或血液白蛋白，向审核项目表中插入数据，同时此类数据不上传，将上传标志更新为假:bln上传=False
                    '银海处理的有点特殊：处方上传时发现需要审批的项目，不进行计算，自然中间库的处方明细表中没有此类待审批的数据
                    '审核程序就根据待审批表与费用记录发生关联进行审批，然后再进行计算与上传
                    If Not bln门诊 Then
                        Call 调用接口_准备_重庆银海版("100", str项目流水号)
                        If Not 调用接口_重庆银海版() Then Exit Function
                        
                        '20060829 断点
                        bln血液白蛋白 = (gstrReturn_重庆银海版 <> "")
                        bln高收费项目 = (Format(str单价, "#0.00") >= "1000.00" And str收费类别 <> "F")
                        If bln血液白蛋白 Or bln高收费项目 Then
                            '已审核项目的审核标志是不会更新的
                            gstrSQL = "zlYB_审核项目表_UPDATE(" & IIf(bln血液白蛋白, 2, 1) & "," & lng病人ID & "," & lng主页ID & ",'" & str处方流水号 & "',0)"
                            gcn重庆银海版.Execute gstrSQL, , adCmdStoredProc
                        End If
                    End If
                    
                    '产生处方明细文件
                    Call DebugTool("产生处方明细文件")
                    strValues = str流水号 & "|" & str个人编号 & "|" & str处方流水号 & "|" & str结算流水号 & "|" & _
                                str退单处方流水号 & "|" & str开单时间 & "|" & str项目流水号 & "|" & str项目类别 & "|" & _
                                str医院项目编码 & "|" & str医院项目名称 & "|" & str单价 & "|" & str数量 & "|" & str金额 & "|" & _
                                str剂型 & "|" & str包装数量 & "|" & str包装单位 & "|" & str含量 & "|" & str含量单位 & "|" & _
                                str容量 & "|" & str容量单位 & "|" & IIf(bln急诊, "1", "0") & "|" & str处方号 & "|" & _
                                str开单医生 & "|" & gstrUserName & "|" & Format(zlDatabase.Currentdate(), "yyyyMMdd HH:mm:ss") & "|" & _
                                str统筹区号
                    Call Record_Add(rsRecipe, strFields, strValues)
                End If
                .MoveNext
            Loop
        Next
    End With
    
    If Not MakeFile_Recipe2() Then Exit Function
    If str处方流水号_UP <> "" Then str处方流水号_UP = Mid(str处方流水号_UP, 2)
    MakeFile_Recipe = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Get处方流水号(ByVal strNO As String, ByVal str性质 As String, ByVal str状态 As String, _
            ByVal str序号 As String, str处方流水号 As String, str退单处方流水号 As String, _
            Optional ByVal lng病人ID As Long = 0)
    Dim bln作废 As Boolean
    Dim rsHandback As New ADODB.Recordset
    
    Call DebugTool("得到处方流水号")
    '返回的处方流水号共：NO[8]+性质[3]+状态[3]+序号[4] 最大20位，目前只用到18位
    '只要传入了病人ID，则说明是门诊
    If strNO = "" Then      '说明是门诊预结算
        str处方流水号 = ToVarchar(lng病人ID & Mid(Format(zlDatabase.Currentdate(), "YYYYMMDDHHmmss"), 11), 15)
        str处方流水号 = str处方流水号 & str序号
        str退单处方流水号 = str处方流水号
        Exit Sub
    End If
    
    str处方流水号 = strNO & String(3 - Len(str性质), "0") & str性质 & _
                    String(3 - Len(str状态), "0") & IIf(str状态 = "3", "1", str状态) & _
                    String(3 - Len(str序号), "0") & str序号
    If str状态 = 1 Then
        '取该笔明细的单价与金额，如果小于零，则随便取一笔正常记录的流水号作为退单流水号
        gstrSQL = " Select 病人ID,主页ID,收费细目ID,Nvl(标准单价,0) 单价,NVl(实收金额,0) 金额" & _
                  " From 住院费用记录" & _
                  " Where NO=[1] And 记录性质=[2] And 记录状态=[3] And 序号=[4]"
        Set rsHandback = zlDatabase.OpenSQLRecord(gstrSQL, "提取该明细，看是否是负数记帐", strNO, str性质, str状态, str序号)
        If rsHandback!单价 < 0 Or rsHandback!金额 < 0 Then
            str退单处方流水号 = GetSequence(rsHandback!病人ID, rsHandback!主页ID, rsHandback!收费细目ID)
        Else
            str退单处方流水号 = str处方流水号
        End If
    Else
        If lng病人ID <> 0 Then
            '门诊
            '冲帐（从保险结算记录中取出摘要，是其原始的处方流水号）
            gstrSQL = " Select 摘要 From 门诊费用记录 " & _
                      " Where 结帐ID=[1] And 序号=[2]"
            Set rsHandback = zlDatabase.OpenSQLRecord(gstrSQL, "从保险结算记录中取出摘要", gComInfo_重庆银海版.冲销ID, CLng(str序号))
            str退单处方流水号 = rsHandback!摘要
        Else
            str退单处方流水号 = strNO & String(3 - Len(str性质), "0") & str性质 & _
                            String(3 - Len(str状态), "0") & "1" & _
                            String(3 - Len(str序号), "0") & str序号
        End If
    End If
End Sub

Private Function MakeFile_Recipe2() As Boolean
    '根据记录集产生交换文件，其格式和根据流水号提取现有处方明细产生交换文件不一样
    Dim lngCol As Long, lngCols As Long
    Dim strRow As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not CreateExchangeFile(处方明细) Then Exit Function
    
    With rsRecipe
        lngCols = .Fields.Count - 1
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    mobjStream.Close
    
    MakeFile_Recipe2 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile_RecipeCalculated(ByVal str流水号 As String, Optional ByVal str处方流水号 As String) As Boolean
    '将接口处理过的处方明细从中间库中提取出来（本次就诊的所有明细），并产生为交换文件
    Dim lngCol As Long, lngCols As Long
    Dim lng病人ID As Long, lng主页ID As Long
    Dim strRow As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '取病人ID、主页ID
    gstrSQL = " Select A.病人ID,A.住院次数 AS 主页ID From 病人信息 A,保险帐户 B" & _
              " Where A.病人ID=B.病人ID ANd B.流水号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID、主页ID", str流水号)
    lng病人ID = rsTemp!病人ID
    lng主页ID = rsTemp!主页ID
    
    If Not CreateExchangeFile(处方明细) Then Exit Function
    '20060829 断点
    gstrSQL = " SELECT 流水号,个人编号,处方流水号,医保项目编码,医保项目流水号,结算交易流水号,退单交易流水号,医疗类别, " & _
              "     医疗机构代码,项目编码,项目名称,急诊标志,高收费审批编号,审批标志,数量,单价, " & _
              "     金额,先自付金额,处方号,项目类别,最高限价,剂型,包装数量,包装单位,含量, " & _
              "     含量单位,容量,容量单位,开单医生,to_char(开方日期,'yyyyMMdd hh24:mi:ss') 开方日期,经办人,to_char(经办时间,'yyyyMMdd hh24:mi:ss') 经办时间,统筹区号,备注  " & _
              " FROM 中间库_处方明细" & _
              " Where 流水号='" & str流水号 & "' And 处方流水号 Not In " & _
              "     (Select 处方流水号 From 审核项目表 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID & " And 审核标志=0)" & _
              IIf(str处方流水号 = "", "", " And 处方流水号 in (" & str处方流水号 & ")")
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn重庆银海版
        
        lngCols = .Fields.Count - 1
        Do While Not .EOF
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    mobjStream.Close
    
    MakeFile_RecipeCalculated = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile_Deal(ByVal str流水号 As String) As Boolean
    '从中间库中提取指定流水号的待遇信息（只可能有一条），并产生为交换文件
    Dim lngCol As Long, lngCols As Long
    Dim strRow As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not CreateExchangeFile(待遇信息) Then Exit Function
    gstrSQL = "SELECT 审批编号,审批记录序号,流水号,个人编号,医疗类别,人员类别,行政级别,实足年龄, " & _
            "     享受公务员补助,跨年度住院标志,待遇封锁类别,封锁原因,医疗机构等级,转诊前医疗机构代码,病种编码,特病审批编号, " & _
            "     本年特殊是否已住院,本年特殊住院次数,本年特殊住院最高等级,计算起付线累计住院次数,个人帐户余额,起付标准, " & _
            "     本次起付线,需连续计算起付线累计,统筹支付累计,公务员门诊费用累计,特病门诊医保费累计,历史未补助先自付, " & _
            "     转院前已进入统筹,转院前公务员补助起付线,to_char(开始时间,'yyyyMMdd hh24:mi:ss') 开始时间,to_char(结束时间,'yyyyMMdd hh24:mi:ss') 结束时间,统筹区号  " & _
            " FROM 中间库_医疗待遇信息" & _
            " Where 流水号='" & str流水号 & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn重庆银海版
        
        lngCols = .Fields.Count - 1
        Do While Not .EOF
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    mobjStream.Close
    
    MakeFile_Deal = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile_Balance(ByVal str流水号 As String) As Boolean
    '从中间库中提取指定流水号，本次就诊的历次结算记录，并产生为交换文件
    Dim lngCol As Long, lngCols As Long
    Dim strRow As String
    Dim blnEmpty As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not CreateExchangeFile(结算信息) Then Exit Function
    
    blnEmpty = True
    gstrSQL = "SELECT 流水号,结算交易流水号,退单交易流水号,审批记录序号,个人编号,年龄,人员类别, " & _
             "  享受公务员,医疗机构代码,医疗机构等级,医疗类别,病种编码,本次起付线,行政级别, " & _
             "  特殊病症标志,特病审批编号,结算类型,本次住院床日,特病自付比例,本次住院次数增加, " & _
             "  医疗费总额,自费总额,个人帐户支付总额,个人现金支付总额,特治特检自付总额, " & _
             "  乙类药自付总额,公务员补助先自付,公务员补助比例,先自付部分公务员补助, " & _
             "  历史先自付公务员补助,本次实际支付起付线,起付标准自付金额,起付线下公务员比例, " & _
             "  起付线下公务员补助,历史起付线公务员返还,本次普通门诊公务员补助,符合范围医保费, " & _
             "  第一段金额,第一段自付比例,第二段金额,第二段自付比例,第三段金额,第三段自付比例, " & _
             "  分段自付金额,本次基本统筹金额,分段自付公务员补助,进入大病金额,大病支付金额, " & _
             "  转院起付线纳入医保费,转院起付线纳入基本,转院起付线纳入公务员,转院起付线纳入大病, " & _
             "  转院起付线纳入自付,发票号,经办人,to_char(经办时间,'yyyyMMdd hh24:mi:ss') 经办时间,统筹区号  " & _
             " FROM 中间库_结算信息" & _
             " Where trim(流水号)='" & Trim(str流水号) & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn重庆银海版
        
        lngCols = .Fields.Count - 1
        Do While Not .EOF
            blnEmpty = False
            strRow = ""
            For lngCol = 0 To lngCols
                strRow = strRow & IIf(lngCol = 0, "", vbTab) & .Fields(lngCol).Value
            Next
            mobjStream.WriteLine strRow
            .MoveNext
        Loop
    End With
    
    If blnEmpty Then mobjStream.WriteLine ""
    mobjStream.Close
    
    MakeFile_Balance = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_Deal(Optional ByVal blnSave As Boolean = False) As Boolean
    '分析接口返回的待遇文件，并保存到中间库（预结算返回的结果80%不准确，因此建议不保存）
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strDeal As String, strBuffer As String
    Dim lngRow As Long
    Dim arrCol
    
    Const int开始日期  As Integer = 30
    Const int结束日期 As Integer = 31
    On Error GoTo errHand
    
'    医疗待遇信息(审批编号,审批记录序号,流水号,个人编号,医疗类别,人员类别,行政级别,实足年龄,享受公务员补助,跨年度住院标志,
'        待遇封锁类别,封锁原因,医疗机构等级,转诊前医疗机构代码,病种编码,特病审批编号,本年特殊是否已住院,本年特殊住院次数,
'        本年特殊住院最高等级,计算起付线累计住院次数,个人帐户余额,起付标准,本次起付线,需连续计算起付线累计,统筹支付累计,
'        公务员门诊费用累计,特病门诊医保费累计,历史未补助先自付,转院前已进入统筹,转院前公务员补助起付线,开始时间,结束时间,统筹区号)
    Call DebugTool("分析待遇信息文件")
    strData = "ZL_中间库_医疗待遇信息_Insert("
    If Not OpenExchangeFile(待遇信息) Then Exit Function
    
    Do While Not mobjStream.AtEndOfStream
        lngRow = mobjStream.Line
        strBuffer = mobjStream.ReadLine
        strDeal = ""
        arrCol = Split(strBuffer, vbTab)
        lngCols = UBound(arrCol)
        For lngCol = 0 To lngCols
            Select Case lngCol
            Case int开始日期, int结束日期
                '由于日期格式不同，需要转换
                strDate = arrCol(lngCol)
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                strDeal = strDeal & strDate
            Case Else
                strDeal = strDeal & ",'" & arrCol(lngCol) & "'"
            End Select
        Next
        strDeal = strData & Mid(strDeal, 2) & IIf(lngRow = 1, ",1", "") & ")"
        If blnSave Then gcn重庆银海版.Execute strDeal, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_Deal = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_Recipe(Optional ByVal blnSave As Boolean = False) As Boolean
    '分析接口返回的处方明细文件，并保存到中间库（预结算返回的结果80%不准确，因此建议不保存）
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strRecipe As String, strBuffer As String
    Dim arrCol
    
    Const int开方日期 As Integer = 29
    Const int经办日期 As Integer = 31
    On Error GoTo errHand
    
    '处方明细(流水号,个人编号,处方流水号,医保项目编码,医保项目流水号,结算交易流水号,退单交易流水号,
    '       医疗类别,医疗机构代码,项目编码,项目名称,急诊标志,高收费审批编号,审批标志,数量,
    '       单价,金额,先自付金额,处方号,项目类别,最高限价,剂型,包装数量,包装单位,含量,含量单位,
    '       容量,容量单位,开单医生,开方日期,经办人,经办时间,统筹区号,备注)
    Call DebugTool("分析处方明细文件")
    strData = "ZL_中间库_处方明细_Insert("
    If Not OpenExchangeFile(处方明细) Then Exit Function
    
    Do While Not mobjStream.AtEndOfStream
        strBuffer = mobjStream.ReadLine
        strRecipe = ""
        arrCol = Split(strBuffer, vbTab)
        lngCols = UBound(arrCol)
        For lngCol = 0 To lngCols
            Select Case lngCol
            Case int开方日期, int经办日期
                '由于日期格式不同，需要转换
                strDate = arrCol(lngCol)
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                strRecipe = strRecipe & strDate
            Case Else
                strRecipe = strRecipe & ",'" & arrCol(lngCol) & "'"
            End Select
        Next
        strRecipe = strData & Mid(strRecipe, 2) & ")"
        If blnSave Then gcn重庆银海版.Execute strRecipe, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_Recipe = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_Balance(strReturn As String, Optional ByVal blnSave As Boolean = False) As Boolean
    '分析接口返回的结算结果文件，并保存到中间库（预结算返回的结果80%不准确，因此建议不保存）
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strBalance As String, strBuffer As String
    Dim arrCol
    Dim cur医保基金 As Currency, cur公务员补助 As Currency, cur个人帐户 As Currency, cur大病基金 As Currency
    
    Const int费用总额 As Integer = 20
    Const int基本统筹 As Integer = 44
    Const int公务员1 As Integer = 28
    Const int公务员2 As Integer = 29
    Const int公务员3 As Integer = 33
    Const int公务员4 As Integer = 34
    Const int公务员5 As Integer = 35
    Const int公务员6 As Integer = 45
    Const int公务员7 As Integer = 50
    Const int个人帐户 As Integer = 22
    Const int大病统筹 As Integer = 47
    Const int经办时间 As Integer = 55
    On Error GoTo errHand
    
'    结算信息(流水号,结算交易流水号,退单交易流水号,审批记录序号,个人编号,年龄,人员类别,享受公务员,医疗机构代码,
'        医疗机构等级,医疗类别,病种编码,本次起付线,行政级别,特殊病症标志,特病审批编号,结算类型,本次住院床日,
'        特病自付比例,本次住院次数增加,医疗费总额,自费总额,个人帐户支付总额,个人现金支付总额,特治特检自付总额,
'        乙类药自付总额,公务员补助先自付[26],公务员补助比例,先自付部分公务员补助[28],历史先自付公务员补助[29],本次实际支付起付线,
'        起付标准自付金额,起付线下公务员比例,起付线下公务员补助[33],历史起付线公务员返还[34],本次普通门诊公务员补助[35],
'        符合范围医保费[36],第一段金额[37],第一段自付比例,第二段金额[39],第二段自付比例,第三段金额[41],第三段自付比例,
'        分段自付金额,本次基本统筹金额[44],分段自付公务员补助[45],进入大病金额,大病支付金额[47],转院起付线纳入医保费,
'        转院起付线纳入基本,转院起付线纳入公务员,转院起付线纳入大病,转院起付线纳入自付,发票号,经办人,经办时间,统筹区号)
    gComInfo_重庆银海版.总费用_中心 = 0
    Call DebugTool("分析结算信息文件")
    strData = "ZL_中间库_结算信息_Insert("
    If Not OpenExchangeFile(结算信息) Then Exit Function
    
    Do While Not mobjStream.AtEndOfStream
        strBuffer = mobjStream.ReadLine
        strBalance = ""
        If Trim(strBuffer) <> "" Then
            arrCol = Split(strBuffer, vbTab)
            lngCols = UBound(arrCol)
            For lngCol = 0 To lngCols
                Select Case lngCol
                Case int经办时间
                    '由于日期格式不同，需要转换
                    strDate = arrCol(lngCol)
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                    strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                    strBalance = strBalance & strDate
                Case Else
                    strBalance = strBalance & ",'" & arrCol(lngCol) & "'"
                End Select
            Next
            strBalance = strData & Mid(strBalance, 2) & ")"
            If blnSave Then gcn重庆银海版.Execute strBalance, , adCmdStoredProc
        
            '获取每笔记录的医保统筹、公务员补助、及个人帐户支付总额
            gComInfo_重庆银海版.总费用_中心 = gComInfo_重庆银海版.总费用_中心 + Val(arrCol(int费用总额))
            cur医保基金 = cur医保基金 + Val(arrCol(int基本统筹))
            cur公务员补助 = cur公务员补助 + Val(arrCol(int公务员1)) + Val(arrCol(int公务员2)) + Val(arrCol(int公务员3)) + _
                    Val(arrCol(int公务员4)) + Val(arrCol(int公务员5)) + Val(arrCol(int公务员6)) + Val(arrCol(int公务员7))
            cur个人帐户 = cur个人帐户 + Val(arrCol(int个人帐户))
            cur大病基金 = cur大病基金 + Val(arrCol(int大病统筹))
        End If
    Loop
    mobjStream.Close
    
    If cur医保基金 <> 0 Then strReturn = strReturn & "|" & "医保基金;" & cur医保基金 & ";0"
    If cur公务员补助 <> 0 Then strReturn = strReturn & "|" & "公务员补助基金;" & cur公务员补助 & ";0"
    If cur个人帐户 <> 0 Then strReturn = strReturn & "|" & "个人帐户;" & cur个人帐户 & ";0"
    If cur大病基金 <> 0 Then strReturn = strReturn & "|" & "大病基金;" & cur大病基金 & ";0"
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    If strReturn = "" Then strReturn = "个人帐户;0;0"
    AnalyFile_Balance = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenExchangeFile(ByVal int类型 As 操作类型_重庆银海版) As Boolean
    '打开文件
    Dim strFile As String
    On Error GoTo errHand
    
    strFile = GetFileName(int类型)
    If Not mobjFileSystem.FileExists(strFile) Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile(strFile, ForReading, False, TristateMixed)
    
    OpenExchangeFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CreateExchangeFile(ByVal int类型 As 操作类型_重庆银海版) As Boolean
    On Error GoTo errHand
    
    Set mobjStream = mobjFileSystem.CreateTextFile(GetFileName(int类型))
    
    CreateExchangeFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetFileName(ByVal int类型 As 操作类型_重庆银海版) As String
    Select Case int类型
    Case 操作类型_重庆银海版.处方明细
        GetFileName = strRecipe
    Case 操作类型_重庆银海版.待遇信息
        GetFileName = strDeal
    Case 操作类型_重庆银海版.结算信息
        GetFileName = strBalance
    Case 操作类型_重庆银海版.门诊处方明细
        GetFileName = str门诊处方明细
    End Select
    GetFileName = strFolder & "\" & GetFileName
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer
    Dim intRecords As Integer
    '编制人:朱玉宝
    '编制日期:2000-11-02
    '也使用于保存
    Set RecTarget = New ADODB.Recordset
    
    With RecTarget
        If .State = 1 Then .Close
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, adLongVarChar, 100, adFldIsNullable     '0:表示新增
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        Do While Not SourceRec.EOF
            If Nvl(SourceRec!是否上传, 0) = 0 Then
                .AddNew
                For intFields = 0 To SourceRec.Fields.Count - 1
                    .Fields(intFields) = SourceRec.Fields(intFields).Value
                Next
                .Update
            End If
            If Nvl(SourceRec!是否上传, 0) = 0 Then
                intRecords = intRecords + 1
                If intRecords = 15 Then
                    SourceRec.MoveNext
                    Exit Do
                End If
            End If
            SourceRec.MoveNext
        Loop
    End With
    
    Set CopyNewRec = RecTarget
End Function

Public Function 身份标识_重庆银海版(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim str流水号 As String, StrInput As String, strIdentify As String
    Dim blnTrans As Boolean
    Dim strReturn As String
    Dim arrReturn
    On Error GoTo errHand
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    strIdentify = frmIdentify重庆银海版.GetPatient(bytType, lng病人ID)
    If strIdentify = "" Then Exit Function
    If Not (bytType = 1 Or bytType = 0) Then Exit Function
    
    '启动事务
    gcn重庆银海版.BeginTrans
    blnTrans = True
    
    '如果是门诊业务，则调用就诊登记接口
    If bytType = 0 Then
        '1.      string  20      社会保障号
        '2.      string  20      门诊/住院号
        '3.      string  3       医疗类别，见代码表
        '4.      string  30      科室
        '5.      string  20      医生
        '6.      datetime        日  入院日期
        '7.      string  20      入院疾病
        '8.      string  20      经办人
        '9.      string  50      并发症
        '10.     string          处理后的医疗待遇信息文件保存的路径及文件名
        gComInfo_重庆银海版.就诊时间 = Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00"
        str流水号 = ToVarchar(lng病人ID & Format(zlDatabase.Currentdate(), "yyMMddHHmmss"), 18)
        StrInput = gComInfo_重庆银海版.个人编号 & gstrSplit_Col_重庆银海版 & str流水号 & gstrSplit_Col_重庆银海版 & _
                 gComInfo_重庆银海版.业务类型 & gstrSplit_Col_重庆银海版 & "门诊" & gstrSplit_Col_重庆银海版 & _
                 ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.就诊时间 & gstrSplit_Col_重庆银海版 & _
                 gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & _
                 ToVarchar(gComInfo_重庆银海版.并发症, 50) & gstrSplit_Col_重庆银海版 & GetFileName(待遇信息)
        Call 调用接口_准备_重庆银海版("08", StrInput)
        If Not 调用接口_重庆银海版() Then
            gcn重庆银海版.RollbackTrans
            Exit Function
        End If
        strReturn = gstrReturn_重庆银海版
        If Not AnalyFile_Deal(True) Then
            gcn重庆银海版.RollbackTrans
            Exit Function
        End If
        
        '得到就诊流水号和结算流水号（以上步骤正确完成才保存新的流水号）
        arrReturn = Split(strReturn, gstrSplit_Col_重庆银海版)
        gComInfo_重庆银海版.就诊流水号 = arrReturn(1)
        gComInfo_重庆银海版.结算流水号 = arrReturn(0)
        
        '更新结算交易流水号及门诊/住院流水号
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'流水号','''" & gComInfo_重庆银海版.就诊流水号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊流水号")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'结算流水号','''" & gComInfo_重庆银海版.结算流水号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算交易流水号")
    End If
    
    '更新保险帐户相关信息（统筹区号、业务类型）
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'统筹区号','''" & gComInfo_重庆银海版.统筹区号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存统筹区号")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'业务类型','''" & gComInfo_重庆银海版.业务类型 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'并发症','''" & gComInfo_重庆银海版.并发症 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存并发症")
    
    gcn重庆银海版.CommitTrans
    
    '返回病人信息串
    身份标识_重庆银海版 = strIdentify
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn重庆银海版.RollbackTrans
End Function

Public Function 医保初始化_重庆银海版(Optional ByVal blnTest As Boolean = False) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim strServer As String, strUser As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    '检查是否存在交换目录，不存在则创建
    If Not mobjFileSystem.FolderExists(strFolder) Then
        mobjFileSystem.CreateFolder (strFolder)
    End If
    
    If mblnInit = False Then
        If Not blnTest Then '如果是测试，则说明是保险参数设置处调用
            '读出连接医保服务器的配置
            gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_重庆银海版)
            
            Do Until rsTemp.EOF
                Select Case rsTemp("参数名")
                    Case "医保用户名"
                        strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保服务器"
                        strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保用户密码"
                        strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                End Select
                rsTemp.MoveNext
            Loop
            
            mstrOwner = strUser
            If OraDataOpen(gcn重庆银海版, strServer, strUser, strPass, False) = False Then
                MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Set gobjYH = CreateObject("YinHai.ChongQing.MedicareDefray")
        '检查连接是否建立
        If gobjYH Is Nothing Then
            MsgBox "医保初始化失败！", vbInformation, gstrSysName
            '调试重庆医保银海版 204-04-07
            Exit Function
        End If
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_重庆银海版)
        gComInfo_重庆银海版.医院编码 = Nvl(rsTemp!医院编码)
        If Not blnTest Then mblnInit = True
        
        '校正本机时间，因银海接口只取本地时间（银海返回数据精确到秒，格式为：yyyyMMdd HH:mm:ss）
        Call 调用接口_准备_重庆银海版("01")
        On Error Resume Next
        gstrReturn_重庆银海版 = Mid(gstrReturn_重庆银海版, 10)
        If Err = 0 Then Time = gstrReturn_重庆银海版
    End If
    
    医保初始化_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保设置_重庆银海版() As Boolean
    医保设置_重庆银海版 = frmSet重庆银海版.参数设置
End Function

Public Function 医保终止_重庆银海版() As Boolean
    On Error Resume Next
    
    Set gobjYH = Nothing
    gcn重庆银海版.Close
    Set gcn重庆银海版 = Nothing
    
    mblnInit = False
    医保终止_重庆银海版 = True
End Function

Public Sub 调用接口_准备_重庆银海版(ByVal strBusiness As String, Optional ByVal StrInput As String = "", _
    Optional ByVal strOutput As String = "", Optional ByVal strAppMsg As String = "")
    mstrAppMsg = strAppMsg
    mstrBusiness = strBusiness
    mstrInput = StrInput
    gstrReturn_重庆银海版 = strOutput
End Sub

Public Function 调用接口_重庆银海版() As Boolean
    '交易代码    交易名称    是否需要广域网
    '1   获取远程系统时间       '√
    '2   获取药品目录           '
    '3   获取诊疗项目目录       '
    '4   获取病种目录           '
    '5   获取医保定点信息       '
    '6   获取代码对照信息       '
    '7   获取个人基本信息医疗待遇信息    '√
    '8   就诊登记               '√
    '9   就诊信息修改           '√
    '10  处方明细计算           '√
    '11  获取高收费项目审批信息 '√
    '12  处方明细信息上传       '√
    '13  模拟费用结算           '
    '14  费用结算               '√
    '15  就诊登记作废           '√
    '16  核对费用结算信息       '√
    '17  核对处方明细信息       '√
    '18  获取药品目录历史变更信息    '
    '19  获取诊疗项目目录历史变更信息    '
    '20  获取病种目录历史变更信息    '
    '22  获取结算信息（用于上次结算失败，中心成功而HIS失败的情况）
    '23  获取待遇信息
    '33  请假信息写入
    On Error GoTo errHand
    Dim lngResult As Long
    
    Call DebugTool(String(20, "-"))
    Call DebugTool("交易代码：" & mstrBusiness)
    Call DebugTool("入参：" & mstrInput)
    lngResult = gobjYH.passivebusiness(mstrBusiness, mstrInput, gstrReturn_重庆银海版, mstrAppMsg)
    If lngResult < 0 Then               '错误信息
        MsgBox "银海提示：交易类型[" & mstrBusiness & "]错误代码[" & lngResult & "]" & mstrAppMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf lngResult > 0 Then           '仅仅是应用提示信息
        MsgBox "银海提示：" & mstrAppMsg, vbInformation, gstrSysName
    End If
    
    调用接口_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_重庆银海版(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim strFileName As String, StrInput As String, strReturn As String
    On Error GoTo errHand
    '得到本次结算的总费用
    Call DebugTool("得到本次结算的总费用")
    gComInfo_重庆银海版.总费用 = 0
    With rs明细
        Do While Not .EOF
            gComInfo_重庆银海版.总费用 = gComInfo_重庆银海版.总费用 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    '不上传明细，先调用处方明细计算，再调用预结算
    '----处方明细计算----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string          传入的处方明细信息文件保存的路径及文件名
'    OutputString
'    序号    数据类型    长度    精度    说明
'    1.      string          处理后的处方明细信息文件保存的路径及文件名
    Call DebugTool("准备产生处方明细文件，以便接口计算")
    If Not MakeFile_Recipe(rs明细, True, True) Then Exit Function
    strFileName = GetFileName(处方明细)
    Call DebugTool("调用处方明细计算接口")
    Call 调用接口_准备_重庆银海版("10", strFileName & gstrSplit_Col_重庆银海版 & strFileName)
    If Not 调用接口_重庆银海版 Then Exit Function
    If Not AnalyFile_Recipe Then Exit Function
    
    '----模拟结算----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  18      结算交易流水号
'    3.      string          医疗待遇信息文件保存的路径及文件名
'    4.      string          处方明细信息文件保存的路径及文件名
'    5.      string          该次就诊历次费用结算结果文件保存的路径及文件名
'    6.      string          费用结算结果文件保存的路径及文件名
    '先产生入参中涉及到的文件，再调用接口（注意：处方明细文件直接使用处方明细计算后，接口返回的文件即可，不需重新产生）
    Call DebugTool("产生待遇信息文件")
    If Not MakeFile_Deal(gComInfo_重庆银海版.就诊流水号) Then Exit Function
    If Not MakeFile_Balance(gComInfo_重庆银海版.就诊流水号) Then Exit Function
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.结算流水号 & gstrSplit_Col_重庆银海版 & _
             GetFileName(待遇信息) & gstrSplit_Col_重庆银海版 & GetFileName(处方明细) & gstrSplit_Col_重庆银海版 & _
             GetFileName(结算信息) & gstrSplit_Col_重庆银海版 & GetFileName(结算信息)
    Call DebugTool("调用13接口，进行模拟结算")
    Call 调用接口_准备_重庆银海版("13", StrInput)
    If Not 调用接口_重庆银海版 Then Exit Function
    Call DebugTool("分析结算结果文件")
    If Not AnalyFile_Balance(strReturn) Then Exit Function
    str结算方式 = strReturn
    Call AnalyBalance(str结算方式)
    
    门诊虚拟结算_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_重庆银海版(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, Optional ByRef strAdvance As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    Dim LNGDO As Long, lngLoop As Long              '循环次数，用以控制处方明细上传
    Dim intCounter As Integer                       '计数器
    Dim lng病人ID As Long
    Dim blnTrans As Boolean
    Dim StrInput As String, strReturn As String, strBillNO As String, str结算流水号 As String, strFileName As String
    Dim cur医保基金 As Currency, cur大病基金 As Currency, cur公务员补助基金 As Currency, cur现金 As Currency, curMoney As Currency
    Dim intBalance As Integer, intBalances As Integer, str结算方式 As String, arrBalance
    Dim str就诊时间 As String, str出院原因 As String
    Dim objStream As TextStream, objFileSys As New FileSystemObject
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim blnOld As Boolean, blnRevise As Boolean '是否需要填写校正字段，结算结果是否需要校正
    On Error GoTo errHand
    
    '和预结算不一样，需要在调用处方明细计算后，接着调用处方明细上传，最后再调用结算
    gcn重庆银海版.BeginTrans
    blnTrans = True
    '取出本次所有费用明细
    gstrSQL = "Select ID From 门诊费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(记录状态,0)<>0 Order by 序号"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "取出本次所有费用明细")
    '计算循环次数，因为每个文件上传的记录数只有15条
    lngLoop = (rs明细.RecordCount \ 15) + IIf(rs明细.RecordCount Mod 15 = 0, 0, 1)
    
    If Not AnalyFile_Recipe(True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
'    '如果存在未通过审批的高收费项目
'    If Not CheckItem Then
'        If Not frm等待响应_重庆银海版.ShowME() Then
'            gcn重庆银海版.RollbackTrans
'            Exit Function
'        End If
'    End If
'
'    '再次产生明细文件
'    rsRecipe.MoveFirst
'    If Not MakeFile_Recipe2() Then
'        gcn重庆银海版.RollbackTrans
'        Exit Function
'    End If
'
'    '再次调用处方明细计算，以获取新的处方明细文件，以便调用门诊结算
'    strFileName = GetFileName(处方明细)
'    Call 调用接口_准备_重庆银海版("10", strFileName & gstrSplit_Col_重庆银海版 & strFileName)
'    If Not 调用接口_重庆银海版 Then
'        gcn重庆银海版.RollbackTrans
'        Exit Function
'    End If
'    If Not AnalyFile_Recipe(True) Then
'        gcn重庆银海版.RollbackTrans
'        Exit Function
'    End If
    
    '根据处方明细文件更新病人费用记录的摘要（用来保存处方流水号）
    With rsRecipe
        .MoveFirst
        Do While Not .EOF
            gstrSQL = "ZL_病人费用记录_更新医保(" & rs明细!ID & ",NULL,NULL,NULL,NULL,1,'" & rsRecipe!处方交易流水号 & "')"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            .MoveNext
            rs明细.MoveNext
        Loop
    End With
    
    '----处方明细上传----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string          传入的处方明细信息文件保存的路径及文件名
    Set objStream = objFileSys.OpenTextFile(GetFileName(处方明细))
    For LNGDO = 1 To lngLoop
        intCounter = 0
        '由原始文件产生上传处方明细文件（最多20笔明细）
        If Not CreateExchangeFile(门诊处方明细) Then
            gcn重庆银海版.RollbackTrans
            Exit Function
        End If
        
        Do While Not objStream.AtEndOfStream
            If intCounter = 15 Then Exit Do
            mobjStream.WriteLine objStream.ReadLine
            intCounter = intCounter + 1
        Loop
        
        mobjStream.Close
        '上传处方
        Call 调用接口_准备_重庆银海版("12", GetFileName(门诊处方明细))
        If Not 调用接口_重庆银海版 Then
            gcn重庆银海版.RollbackTrans
            objStream.Close
            Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "处方明细上传失败！", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    objStream.Close
    
    '得到开始发票号
    gstrSQL = "Select 病人ID,实际票号 From 门诊费用记录 Where 结帐ID=[1] And 序号=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取发票号", lng结帐ID)
    strBillNO = Nvl(rsTemp!实际票号)
    lng病人ID = rsTemp!病人ID
    
    '----正式结算----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  18      结算交易流水号
'    3.      string  3       结算类型  0-正常结算;1-中途结算;2-转家庭病床结算;3-卡挂失结算
'    4.      string  20      病种编码
'    5.      string  20      发票号
'    6.      string          处方明细信息文件保存的路径及文件名
'    7.      string          该次就诊历次费用结算结果文件保存的路径及文件名
'    8.      string          传出医疗待遇信息文件保存的路径及文件名
'    9.      string          传出费用结算结果文件保存的路径及文件名
'    OutputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      结算交易流水号
    '先产生入参中涉及到的文件，再调用接口（注意：处方明细文件直接使用处方明细计算后，接口返回的文件即可，不需重新产生）
    If Not MakeFile_Deal(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.结算流水号 & gstrSplit_Col_重庆银海版 & _
             "0" & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & _
             strBillNO & gstrSplit_Col_重庆银海版 & GetFileName(处方明细) & gstrSplit_Col_重庆银海版 & _
             GetFileName(结算信息) & gstrSplit_Col_重庆银海版 & GetFileName(待遇信息) & gstrSplit_Col_重庆银海版 & GetFileName(结算信息)
    Call 调用接口_准备_重庆银海版("14", StrInput)
    
    If Not 调用接口_重庆银海版 Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    str结算流水号 = gstrReturn_重庆银海版
    If Not AnalyFile_Deal(True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Balance(strReturn, True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    '分解结算支付信息
    arrBalance = Split(strReturn, "|")
    intBalances = UBound(arrBalance)
    For intBalance = 0 To intBalances
        str结算方式 = Split(arrBalance(intBalance), ";")(0)
        curMoney = Split(arrBalance(intBalance), ";")(1)
        Select Case str结算方式
        Case "个人帐户"
            cur个人帐户 = curMoney
        Case "医保基金"
            cur医保基金 = curMoney
        Case "大病基金"
            cur大病基金 = curMoney
        Case "公务员补助基金"
            cur公务员补助基金 = curMoney
        End Select
    Next
    cur现金 = gComInfo_重庆银海版.总费用 - cur个人帐户 - cur医保基金 - cur大病基金 - cur公务员补助基金
    
    '对结算结果进行核对
    If Not (cur个人帐户 = pre_Balance.cur个人帐户 And cur医保基金 = pre_Balance.cur医保基金 And _
        cur大病基金 = pre_Balance.cur大病基金 And cur公务员补助基金 = pre_Balance.cur公务员补助) Then
        blnRevise = True
        #If gverControl < 2 Then
            blnOld = True
        #End If
    End If
    
    '更新保险帐户（新的结算流水号）
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'结算流水号','''" & str结算流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算交易流水号")
    
    '保存本次结算情况
    '进入统筹金额=公务员补助基金;统筹报销金额=统筹基金;大病自付金额=大病基金
    '备注=业务类型|就诊流水号|结算流水号|就诊时间|疾病编码|并发症
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_重庆银海版 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_重庆银海版.总费用 & "," & cur现金 & "," & 0 & "," & cur公务员补助基金 & "," & cur医保基金 & "," & cur大病基金 & "," & _
        0 & "," & cur个人帐户 & ",null,null,null,'" & gComInfo_重庆银海版.业务类型 & "|" & gComInfo_重庆银海版.就诊流水号 & "|" & gComInfo_重庆银海版.结算流水号 & "|" & gComInfo_重庆银海版.就诊时间 & "|" & gComInfo_重庆银海版.疾病编码 & "|" & gComInfo_重庆银海版.并发症 & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    '就诊登记修改，标识为结束
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  3       医疗类别
'    3.      string  30      科室
'    4.      string  20      医生
'    5.      datetime        日  入院日期
'    6.      string  20      入院疾病
'    7.      string  3       就诊状态
'    8.      datetime        日  出院日期
'    9.      string  20      确诊疾病编码
'    10.     string  3       出院原因
'    11.     string  20      经办人
'    12.     string  50      并发症
    str就诊时间 = Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00"
    str出院原因 = 1
    
    '取业务相关信息
    gstrSQL = " Select A.流水号,A.业务类型,B.编码 疾病编码,A.并发症 From 保险帐户 A,保险病种 B " & _
              " Where A.病人ID=[1] And A.险类=[2] And A.病种ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取其它相关信息", lng病人ID, TYPE_重庆银海版)
    
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.业务类型 & gstrSplit_Col_重庆银海版 & _
            "门诊" & gstrSplit_Col_重庆银海版 & ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & _
            str就诊时间 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & _
            "0" & gstrSplit_Col_重庆银海版 & str就诊时间 & gstrSplit_Col_重庆银海版 & _
            gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & str出院原因 & gstrSplit_Col_重庆银海版 & _
            ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & ToVarchar(gComInfo_重庆银海版.并发症, 50)
    Call 调用接口_准备_重庆银海版("09", StrInput)
    If Not 调用接口_重庆银海版() Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
   '由于预结算的结果80%可能与正式结算的结果不一致（因为银海的预结算接口不会去取病人最新的待遇信息，因此需要修正）
    str结算方式 = ""
    If cur个人帐户 <> 0 Then str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    If cur医保基金 <> 0 Then str结算方式 = str结算方式 & "||医保基金|" & cur医保基金
    If cur大病基金 <> 0 Then str结算方式 = str结算方式 & "||大病基金|" & cur大病基金
    If cur公务员补助基金 <> 0 Then str结算方式 = str结算方式 & "||公务员补助基金|" & cur公务员补助基金
    If str结算方式 <> "" And blnRevise Then
        str结算方式 = Mid(str结算方式, 3)
        If blnOld Then
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    
    gcn重庆银海版.CommitTrans
    
    blnTrans = False
    门诊结算_重庆银海版 = True
'
'    '打印票据
'    Call 调用接口_准备_重庆银海版("21", gComInfo_重庆银海版.结算流水号)
'    Call 调用接口_重庆银海版
    
    Exit Function
errHand:
    If blnTrans Then gcn重庆银海版.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_重庆银海版(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim blnTrans As Boolean, bln上传 As Boolean
    Dim lng冲销ID As Long
    Dim StrInput As String, strReturn As String, strBillNO As String, str结算流水号 As String, strFileName As String
    Dim cur医保基金 As Currency, cur大病基金 As Currency, cur公务员补助基金 As Currency, cur现金 As Currency, curMoney As Currency
    Dim intBalance As Integer, intBalances As Integer, str结算方式 As String, arrBalance
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim LNGDO As Long, lngLoop As Long              '循环次数，用以控制处方明细上传
    Dim intCounter As Integer                       '计数器
    Dim objStream As TextStream, objFileSys As New FileSystemObject
    On Error GoTo errHand
    
    '需要提取上次就诊的相关基本信息（就诊流水号、业务类型、病种编码等）
    '保险结算记录.备注=业务类型|就诊流水号|结算流水号|就诊时间|疾病编码|并发症
    gstrSQL = " Select B.医保号 个人编号,A.支付顺序号,A.备注,B.结算流水号 From 保险结算记录 A,保险帐户 B " & _
              " Where A.性质=1 And A.记录ID=[1] And A.病人ID=B.病人ID And B.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取上次结算时的就诊流水号", lng结帐ID, TYPE_重庆银海版)
    gComInfo_重庆银海版.就诊流水号 = Split(rsTemp!备注, "|")(1)
    gComInfo_重庆银海版.业务类型 = Split(rsTemp!备注, "|")(0)
    gComInfo_重庆银海版.就诊时间 = Split(rsTemp!备注, "|")(3)
    gComInfo_重庆银海版.结算流水号 = rsTemp!结算流水号  '只有此数据取当前的结算流水号
    gComInfo_重庆银海版.疾病编码 = Split(rsTemp!备注, "|")(4)
    gComInfo_重庆银海版.并发症 = Split(rsTemp!备注, "|")(5)
    gComInfo_重庆银海版.个人编号 = rsTemp!个人编号
    gComInfo_重庆银海版.冲销ID = lng结帐ID
    
    '和预结算不一样，需要在调用处方明细计算后，接着调用处方明细上传，最后再调用结算
    gcn重庆银海版.BeginTrans
    blnTrans = True
    gComInfo_重庆银海版.总费用 = 0
    
    '取冲销记录的结帐ID，单据号
    gstrSQL = "select distinct A.结帐ID,A.NO from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    strBillNO = rsTemp!NO
    
    '打开明细记录集
    gstrSQL = " Select A.ID,A.病人ID,A.NO,A.序号,A.记录性质,A.记录状态,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 登记时间,A.收费类别," & _
              " A.开单人,B.名称 开单部门,A.收费细目ID,A.计算单位,C.项目编码 医保项目编码 ,A.实收金额,A.数次*Nvl(A.付数,1) 数量,A.实收金额/(A.数次*Nvl(A.付数,1)) 单价,Nvl(A.是否上传,0) 是否上传" & _
              " From 门诊费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=" & TYPE_重庆银海版 & ") C " & _
              " Where A.记录性质=1 And A.记录状态=2 And A.NO=[1]" & _
              " And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0" & _
              " Order by A.NO,A.病人ID"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "读取费用明细记录", strBillNO)
    With rs明细
        Do While Not .EOF
            gComInfo_重庆银海版.总费用 = gComInfo_重庆银海版.总费用 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        lngLoop = (rs明细.RecordCount \ 20) + IIf(rs明细.RecordCount Mod 20 = 0, 0, 1)
    End With
    
    '产生处方明细计算用的交换文件（因为NO号在预结算时还没有，因此无法使用预结算接口返回的处方明细文件）
    '----处方明细计算----
    strBillNO = "Z9000999"      '退费时无发票号
    If Not MakeFile_Recipe(rs明细, True, False) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    strFileName = GetFileName(处方明细)
    Call 调用接口_准备_重庆银海版("10", strFileName & gstrSplit_Col_重庆银海版 & strFileName)
    If Not 调用接口_重庆银海版 Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Recipe(True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    '----处方明细上传----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string          传入的处方明细信息文件保存的路径及文件名
    Set objStream = objFileSys.OpenTextFile(GetFileName(处方明细))
    For LNGDO = 1 To lngLoop
        intCounter = 0
        '由原始文件产生上传处方明细文件（最多20笔明细）
        If Not CreateExchangeFile(门诊处方明细) Then
            gcn重庆银海版.RollbackTrans
            Exit Function
        End If
        
        Do While Not objStream.AtEndOfStream
            If intCounter = 20 Then Exit Do
            mobjStream.WriteLine objStream.ReadLine
            intCounter = intCounter + 1
        Loop
        
        mobjStream.Close
        '上传处方
        Call 调用接口_准备_重庆银海版("12", GetFileName(门诊处方明细))
        If Not 调用接口_重庆银海版 Then
            gcn重庆银海版.RollbackTrans
            objStream.Close
            Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "处方明细上传失败！", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    objStream.Close
    
    With rs明细
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
        'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
        gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        .MoveNext
    End With
    
    gcn重庆银海版.CommitTrans
    gcn重庆银海版.BeginTrans
    '----正式结算----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  18      结算交易流水号
'    3.      string  3       结算类型  0-正常结算;1-中途结算;2-转家庭病床结算;3-卡挂失结算
'    4.      string  20      病种编码
'    5.      string  20      发票号
'    6.      string          处方明细信息文件保存的路径及文件名
'    7.      string          该次就诊历次费用结算结果文件保存的路径及文件名
'    8.      string          传出医疗待遇信息文件保存的路径及文件名
'    9.      string          传出费用结算结果文件保存的路径及文件名
'    OutputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      结算交易流水号
    '先产生入参中涉及到的文件，再调用接口（注意：处方明细文件直接使用处方明细计算后，接口返回的文件即可，不需重新产生）
    If Not MakeFile_RecipeCalculated(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Deal(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.结算流水号 & gstrSplit_Col_重庆银海版 & _
             "0" & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & _
             strBillNO & gstrSplit_Col_重庆银海版 & GetFileName(处方明细) & gstrSplit_Col_重庆银海版 & _
             GetFileName(结算信息) & gstrSplit_Col_重庆银海版 & GetFileName(待遇信息) & gstrSplit_Col_重庆银海版 & GetFileName(结算信息)
    Call 调用接口_准备_重庆银海版("14", StrInput)
    If Not 调用接口_重庆银海版 Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    str结算流水号 = gstrReturn_重庆银海版
    If Not AnalyFile_Deal(True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Balance(strReturn, True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    '分解结算支付信息
    arrBalance = Split(strReturn, "|")
    intBalances = UBound(arrBalance)
    For intBalance = 0 To intBalances
        str结算方式 = Split(arrBalance(intBalance), ";")(0)
        curMoney = Split(arrBalance(intBalance), ";")(1)
        Select Case str结算方式
        Case "个人帐户"
            cur个人帐户 = curMoney
        Case "医保基金"
            cur医保基金 = curMoney
        Case "大病基金"
            cur大病基金 = curMoney
        Case "公务员补助基金"
            cur公务员补助基金 = curMoney
        End Select
    Next
    cur现金 = gComInfo_重庆银海版.总费用 - cur个人帐户 - cur医保基金 - cur大病基金 - cur公务员补助基金
    
    '更新保险帐户（新的结算流水号）
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'结算流水号','''" & str结算流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算交易流水号")
    
    '保存本次结算情况
    '进入统筹金额=公务员补助基金;统筹报销金额=统筹基金;大病自付金额=大病基金
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_重庆银海版 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_重庆银海版.总费用 & "," & cur现金 & "," & 0 & "," & cur公务员补助基金 & "," & cur医保基金 & "," & cur大病基金 & "," & _
        0 & "," & cur个人帐户 & ",null,null,null,'" & gComInfo_重庆银海版.业务类型 & "|" & gComInfo_重庆银海版.就诊流水号 & "|" & gComInfo_重庆银海版.结算流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    gcn重庆银海版.CommitTrans
    门诊结算冲销_重庆银海版 = True
    
    '打印票据
'    Call 调用接口_准备_重庆银海版("21", gComInfo_重庆银海版.结算流水号)
'    Call 调用接口_重庆银海版
    Exit Function
errHand:
    If blnTrans Then gcn重庆银海版.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 入院登记_重庆银海版(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim str流水号 As String, StrInput As String, strReturn As String
    Dim arrReturn
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '1.      string  20      社会保障号
    '2.      string  20      门诊/住院号
    '3.      string  3       医疗类别，见代码表
    '4.      string  30      科室
    '5.      string  20      医生
    '6.      datetime        日  入院日期
    '7.      string  20      入院疾病
    '8.      string  20      经办人
    '9.      string  50      并发症
    '10.     string          处理后的医疗待遇信息文件保存的路径及文件名
    gcn重庆银海版.BeginTrans
    blnTrans = True
    
    gstrSQL = " Select to_char(A.入院日期,'yyyy-MM-dd') 入院日期,B.名称 科室,A.门诊医师 医生 From 病案主页 A,部门表 B " & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院日期", lng病人ID, lng主页ID)
    
    gComInfo_重庆银海版.就诊时间 = Format(rsTemp!入院日期, "yyyyMMdd") & " 00:00:00"
    str流水号 = ToVarchar(lng病人ID & "_" & lng主页ID, 18)
    StrInput = gComInfo_重庆银海版.个人编号 & gstrSplit_Col_重庆银海版 & str流水号 & gstrSplit_Col_重庆银海版 & _
             gComInfo_重庆银海版.业务类型 & gstrSplit_Col_重庆银海版 & rsTemp!科室 & gstrSplit_Col_重庆银海版 & _
             Nvl(rsTemp!医生, "银海") & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.就诊时间 & gstrSplit_Col_重庆银海版 & _
             gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & _
             ToVarchar(gComInfo_重庆银海版.并发症, 50) & gstrSplit_Col_重庆银海版 & GetFileName(待遇信息)
    Call 调用接口_准备_重庆银海版("08", StrInput)
    If Not 调用接口_重庆银海版() Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    strReturn = gstrReturn_重庆银海版
    If Not AnalyFile_Deal(True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    '得到就诊流水号和结算流水号（以上步骤正确完成才保存新的流水号）
    arrReturn = Split(strReturn, gstrSplit_Col_重庆银海版)
    gComInfo_重庆银海版.就诊流水号 = arrReturn(1)
    gComInfo_重庆银海版.结算流水号 = arrReturn(0)
    
    '更新结算交易流水号及门诊/住院流水号
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'流水号','''" & gComInfo_重庆银海版.就诊流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存就诊流水号")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'结算流水号','''" & gComInfo_重庆银海版.结算流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算交易流水号")
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆银海版 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    gcn重庆银海版.CommitTrans
    
    入院登记_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn重庆银海版.RollbackTrans
End Function

Public Function 入院登记撤销_重庆银海版(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "该医保病人存在未结费用，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    '检查是否已存在费用记录,存在则只允许办理出院登记
    gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存在费用记录", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        MsgBox "该病人已经存在费用记录,只能办理出院手续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '获取原就诊流水号
    gstrSQL = "Select 流水号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊流水号", TYPE_重庆银海版, lng病人ID)
    gComInfo_重庆银海版.就诊流水号 = rsTemp!流水号
    
    '调用就诊登记作废接口
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
    Call 调用接口_准备_重庆银海版("15", gComInfo_重庆银海版.就诊流水号)
    If Not 调用接口_重庆银海版() Then Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆银海版 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_重庆银海版(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim bln医保出院 As Boolean, bln结帐 As Boolean
    Dim str就诊时间 As String, str科室 As String, str医生 As String, str出院原因 As String
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
'    akc195  1   出院原因    康复
'    akc195  2   出院原因    转院
'    akc195  3   出院原因    死亡
'    akc195  4   出院原因    好转
'    akc195  9   出院原因    其他
    
    bln医保出院 = False
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '判断该病人是否结算过，没有结算过的病人费用为零，说明需要调用就诊登记撤销
        bln结帐 = False
        gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(结帐ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否该调用就诊登记撤销", lng病人ID, lng主页ID)
        If Not rsTemp.EOF Then
            bln结帐 = True
        End If
        
        bln医保出院 = True
        If bln结帐 Then
            '取入院相关信息
            gstrSQL = " Select to_char(A.入院日期,'yyyy-MM-dd') 入院日期,B.名称 科室,A.出院方式,A.住院医师 医生 From 病案主页 A,部门表 B " & _
                      " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院日期", lng病人ID, lng主页ID)
            str就诊时间 = Format(rsTemp!入院日期, "yyyyMMdd") & " 00:00:00"
            str科室 = Nvl(rsTemp!科室)
            str医生 = IIf(IsNull(rsTemp!医生), "银海", rsTemp!医生)
            str出院原因 = IIf(IsNull(rsTemp!出院方式), "", rsTemp!出院方式)
            Select Case str出院原因
            Case "正常", "康复"
                str出院原因 = 1
            Case "死亡"
                str出院原因 = 3
            Case "转院"
                str出院原因 = 2
            Case Else
                str出院原因 = 9
            End Select
            
            '取业务相关信息
            gstrSQL = " Select A.流水号,A.业务类型,B.编码 疾病编码,A.并发症 From 保险帐户 A,保险病种 B " & _
                      " Where A.病人ID=[1] And A.险类=[2] And A.病种ID=B.ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取其它相关信息", lng病人ID, TYPE_重庆银海版)
            
            StrInput = rsTemp!流水号 & gstrSplit_Col_重庆银海版 & rsTemp!业务类型 & gstrSplit_Col_重庆银海版 & _
                    str科室 & gstrSplit_Col_重庆银海版 & str医生 & gstrSplit_Col_重庆银海版 & _
                    str就诊时间 & gstrSplit_Col_重庆银海版 & rsTemp!疾病编码 & gstrSplit_Col_重庆银海版 & _
                    "0" & gstrSplit_Col_重庆银海版 & Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00" & gstrSplit_Col_重庆银海版 & _
                    rsTemp!疾病编码 & gstrSplit_Col_重庆银海版 & str出院原因 & gstrSplit_Col_重庆银海版 & _
                    ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & ToVarchar(Nvl(rsTemp!并发症), 50)
            Call 调用接口_准备_重庆银海版("09", StrInput)
            If Not 调用接口_重庆银海版() Then Exit Function
        Else
            '获取原就诊流水号
            gstrSQL = "Select 流水号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊流水号", TYPE_重庆银海版, lng病人ID)
            gComInfo_重庆银海版.就诊流水号 = rsTemp!流水号
            
            '调用就诊登记作废接口
        '    InputString
        '    序号    数据类型    长度    精度    说明
        '    1.      string  18      门诊/住院流水号
            Call 调用接口_准备_重庆银海版("15", gComInfo_重庆银海版.就诊流水号)
            If Not 调用接口_重庆银海版() Then Exit Function
        End If
    End If
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆银海版 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    MsgBox IIf(bln医保出院, "医保出院", "HIS出院") & "办理成功！", vbInformation, gstrSysName
    出院登记_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_重庆银海版(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim StrInput As String
    Dim str就诊时间 As String, str科室 As String, str医生 As String, str出院原因 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '取入院相关信息
        gstrSQL = " Select to_char(A.入院日期,'yyyy-MM-dd') 入院日期,B.名称 科室,A.出院方式,A.住院医师 医生 From 病案主页 A,部门表 B " & _
                  " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院日期", lng病人ID, lng主页ID)
        str就诊时间 = Format(rsTemp!入院日期, "yyyyMMdd") & " 00:00:00"
        str科室 = rsTemp!科室
        str医生 = Nvl(rsTemp!医生, "银海")
        str出院原因 = IIf(IsNull(rsTemp!出院方式), "", rsTemp!出院方式)
        Select Case str出院原因
        Case "正常", "康复"
            str出院原因 = 1
        Case "死亡"
            str出院原因 = 3
        Case "转院"
            str出院原因 = 2
        Case Else
            str出院原因 = 9
        End Select
        
        '取业务相关信息
        gstrSQL = " Select A.流水号,A.业务类型,B.编码 疾病编码,A.并发症 From 保险帐户 A,保险病种 B " & _
                  " Where A.病人ID=[1] And A.险类=[2] And A.病种ID=B.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取其它相关信息", lng病人ID, TYPE_重庆银海版)
        
        StrInput = rsTemp!流水号 & gstrSplit_Col_重庆银海版 & rsTemp!业务类型 & gstrSplit_Col_重庆银海版 & _
                str科室 & gstrSplit_Col_重庆银海版 & str医生 & gstrSplit_Col_重庆银海版 & _
                str就诊时间 & gstrSplit_Col_重庆银海版 & rsTemp!疾病编码 & gstrSplit_Col_重庆银海版 & _
                "1" & gstrSplit_Col_重庆银海版 & Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00" & gstrSplit_Col_重庆银海版 & _
                rsTemp!疾病编码 & gstrSplit_Col_重庆银海版 & str出院原因 & gstrSplit_Col_重庆银海版 & _
                ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & ToVarchar(Nvl(rsTemp!并发症), 50)
        Call 调用接口_准备_重庆银海版("09", StrInput)
        If Not 调用接口_重庆银海版() Then Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重庆银海版 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_重庆银海版 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 个人余额_重庆银海版(strSelfNo As String) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'返回: 返回个人帐户余额的金额
    Dim strReturn As String
    Const int帐户余额 As Integer = 13
    On Error GoTo errHandle
    
    '直接调用身份验证获取个人余额
    Call 调用接口_准备_重庆银海版("07", strSelfNo)
    If Not 调用接口_重庆银海版 Then Exit Function
    strReturn = gstrReturn_重庆银海版
    个人余额_重庆银海版 = Val(Split(strReturn, gstrSplit_Col_重庆银海版)(int帐户余额))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_重庆银海版(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim strFile_Recipe As String, StrInput As String, strReturn As String
    Dim str处方流水号_UP As String                     '记录本次上传明细的流水号
    Dim str处方流水号 As String, str处方退单流水号 As String
    Dim blnTrans As Boolean, bln上传 As Boolean     '是否在事务中,是否存在未上传的记录
    Dim bln更新 As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim gcn上传 As New ADODB.Connection
    Dim rsRecipe As New ADODB.Recordset
    Dim intDO As Integer
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    On Error GoTo errHand
    '对未上传的明细进行计算，并上传所有处方明细，再调用预结算接口
    Call DebugTool("进入住院虚拟结算")
    '取该病人的相关业务信息
    gstrSQL = " Select A.医保号,A.流水号,A.业务类型,A.统筹区号,A.结算流水号,A.并发症,B.编码 疾病编码 From 保险帐户 A,保险病种 B" & _
              " Where A.病人ID=[1] And A.险类=[2] And A.病种ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的相关业务信息", lng病人ID, TYPE_重庆银海版)
    With gComInfo_重庆银海版
        .个人编号 = rsTemp!医保号
        .并发症 = Nvl(rsTemp!并发症)
        .就诊流水号 = rsTemp!流水号
        .结算流水号 = rsTemp!结算流水号
        .疾病编码 = rsTemp!疾病编码
        .业务类型 = rsTemp!业务类型
        .统筹区号 = rsTemp!统筹区号
        .总费用 = 0
    End With
    
    '新打开一个事务用来上传费用明细，避免重复上传
    Set gcn上传 = GetNewConnection
    blnTrans = True
    gcn重庆银海版.BeginTrans
    
    With rsExse
        Call DebugTool("检查是否对码，且汇总费用总额")
        Do While Not .EOF
            If Nvl(!是否上传, 0) = 0 And Not bln上传 Then
                bln上传 = True
                gcn上传.BeginTrans
            End If
            If Nvl(!医保项目编码) = "" Then
                If bln上传 Then gcn上传.RollbackTrans
                gcn重庆银海版.RollbackTrans
                MsgBox "存在未对码的项目，不允许上传！", vbInformation, gstrSysName
                Exit Function
            End If
            gComInfo_重庆银海版.总费用 = gComInfo_重庆银海版.总费用 + Nvl(!金额, 0)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        If bln上传 Then
            Call DebugTool("上传明细")
            strFile_Recipe = GetFileName(处方明细)
        
            For intDO = 1 To 2
                If intDO = 1 Then
                    .Filter = "金额>0"
                Else
                    .Filter = "金额<=0"
                End If
            
                '对未上传的明细进行计算
                Do While Not rsExse.EOF
                    Call DebugTool("拷贝费用明细记录集，用以上传")
                    If Not blnTrans Then gcn重庆银海版.BeginTrans: blnTrans = True
                    If Not bln上传 Then gcn上传.BeginTrans: bln上传 = True
                    Set rsRecipe = CopyNewRec(rsExse)
                    
                    If rsRecipe.RecordCount <> 0 Then
                        Call DebugTool("产生处方明细文件")
                        rsRecipe.Filter = 0
                        rsRecipe.MoveFirst
                        If Not MakeFile_Recipe(rsRecipe, False, False, str处方流水号_UP) Then
                            .Filter = 0
                            rsRecipe.Filter = 0
                            gcn重庆银海版.RollbackTrans
                            gcn上传.RollbackTrans
                            Exit Function
                        End If
                        
                        rsRecipe.Filter = 0
                        Call 调用接口_准备_重庆银海版("10", strFile_Recipe & gstrSplit_Col_重庆银海版 & strFile_Recipe)
                        Call DebugTool("调用处方明细计算")
                        If Not 调用接口_重庆银海版 Then
                            .Filter = 0
                            gcn重庆银海版.RollbackTrans
                            gcn上传.RollbackTrans
                            Exit Function
                        End If
                        If Not AnalyFile_Recipe(True) Then
                            .Filter = 0
                            gcn重庆银海版.RollbackTrans
                            gcn上传.RollbackTrans
                            Exit Function
                        End If
                        
                        '----处方明细上传（仅处理未上传部分的处方明细）----
                        '    InputString
                        '    序号    数据类型    长度    精度    说明
                        '    1.      string          传入的处方明细信息文件保存的路径及文件名
                        Call DebugTool("调用处方明细上传")
                        '产生处方明细文件
                        Call MakeFile_RecipeCalculated(gComInfo_重庆银海版.就诊流水号, str处方流水号_UP)
                        
                        Call 调用接口_准备_重庆银海版("12", strFile_Recipe)
                        If Not 调用接口_重庆银海版 Then
                            .Filter = 0
                            gcn重庆银海版.RollbackTrans
                            gcn上传.RollbackTrans
                            Exit Function
                        End If
                        
                        '将已上传的处方打上传标志
                        Call DebugTool("打上传标记")
                        With rsRecipe
                            If .RecordCount <> 0 Then .MoveFirst
                            Do While Not .EOF
                                If Nvl(!是否上传, 0) = 0 Then
                                    '20060829 断点
                                    bln更新 = True
                                    
                                    Call Get处方流水号(!NO, !记录性质, !记录状态, !序号, str处方流水号, str处方退单流水号)
                                    
                                    '如果不是待审批项目，或审核标志已更新的，更新上传标志
                                    gstrSQL = "Select 审核标志 From 审核项目表 Where 处方流水号='" & str处方流水号 & "'"
                                    Call OpenRecordset_OtherBase(rsTemp, "判断", gstrSQL, gcn重庆银海版)
                                    If rsTemp.RecordCount <> 0 Then
                                        bln更新 = (rsTemp!审核标志 <> 0)
                                    End If
                                    
                                    '仅打上传标志，因为明细必须正确上传后，才能保证结算的正确性
                                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                                    If bln更新 Then
                                        gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                                        gcn上传.Execute gstrSQL, , adCmdStoredProc
                                    End If
                                End If
                                .MoveNext
                            Loop
                        End With
                    End If
                    
                    '保证处方明细计算保存成功
                    gcn上传.CommitTrans
                    gcn重庆银海版.CommitTrans
                    bln上传 = False
                    blnTrans = False
                Loop
            Next
            
            .Filter = 0
        End If
    End With
    
    If blnTrans = False Then gcn重庆银海版.BeginTrans: blnTrans = True
    Call DebugTool("获取高收费项目审批信息")
    Call TestVerifyItem
    
    '将所有处方明细计算后的数据提取出来，并产生为交换文件，准备调用预结算接口
    '----模拟结算----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  18      结算交易流水号
'    3.      string          医疗待遇信息文件保存的路径及文件名
'    4.      string          处方明细信息文件保存的路径及文件名
'    5.      string          该次就诊历次费用结算结果文件保存的路径及文件名
'    6.      string          费用结算结果文件保存的路径及文件名
    '先产生入参中涉及到的文件，再调用接口（注意：处方明细文件直接使用处方明细计算后，接口返回的文件即可，不需重新产生）
    If Not MakeFile_Deal(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_RecipeCalculated(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    Call DebugTool("调用住院虚拟结算接口")
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.结算流水号 & gstrSplit_Col_重庆银海版 & _
             GetFileName(待遇信息) & gstrSplit_Col_重庆银海版 & GetFileName(处方明细) & gstrSplit_Col_重庆银海版 & _
             GetFileName(结算信息) & gstrSplit_Col_重庆银海版 & GetFileName(结算信息)
    Call 调用接口_准备_重庆银海版("13", StrInput)
    If Not 调用接口_重庆银海版 Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not AnalyFile_Balance(strReturn) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    gcn重庆银海版.CommitTrans
    住院虚拟结算_重庆银海版 = strReturn
    Call AnalyBalance(strReturn)
    If 住院虚拟结算_重庆银海版 = "" Then 住院虚拟结算_重庆银海版 = "个人帐户;0;0"
    
    '如果总额不等，仅提示
    If Format(gComInfo_重庆银海版.总费用, "#0.00") <> Format(gComInfo_重庆银海版.总费用_中心, "#0.00") Then
        MsgBox "发现本次未结算总费用与医保中心不一致！" & vbCrLf & _
               "医院：" & Format(gComInfo_重庆银海版.总费用, "#0.00") & Space(10) & "医保中心：" & Format(gComInfo_重庆银海版.总费用_中心, "#0.00"), vbInformation, gstrSysName
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcn重庆银海版.RollbackTrans
    If bln上传 Then gcn上传.RollbackTrans
End Function

Public Function 住院结算_重庆银海版(lng结帐ID As Long, ByVal lng病人ID As Long, Optional ByRef strAdvance As String) As Boolean
    Dim strBillNO As String, StrInput As String, strReturn As String, str结算方式 As String, str结算流水号 As String
    Dim cur个人帐户 As Currency, cur医保基金 As Currency, cur大病基金 As Currency, cur公务员补助基金 As Currency, cur现金 As Currency, curMoney As Currency
    Dim cur个人帐户_OLD As Currency, cur医保基金_OLD As Currency, cur大病基金_OLD As Currency, cur公务员补助基金_OLD As Currency, cur现金_OLD As Currency
    Dim intBalance As Integer, intBalances As Integer, arrBalance
    Dim blnTrans As Boolean, bln单病种 As Boolean, lng病种ID As Long
    Dim intState As Integer
    Dim lng主页ID As Long
    Dim blnOld As Boolean, blnRevise As Boolean '是否需要填写校正字段，结算结果是否需要校正
    Dim rsTemp As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
        '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    On Error GoTo errHand
    
    '提取病种类型，如果存在，说明是单病种结算
    gstrSQL = " Select Nvl(类别,0) AS 类别 From 保险病种 Where ID=" & _
              "     (Select 病种ID From 保险帐户 Where 险类=[1] And 病人ID=[2])" & _
              " And 险类=" & TYPE_重庆银海版
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病种的名称", TYPE_重庆银海版, lng病人ID)
    If rsTemp.RecordCount <> 0 Then
        bln单病种 = (rsTemp!类别 = 4)
    End If
    
    '得到开始发票号
    gstrSQL = "Select 病人ID,实际票号 From 病人结帐记录 Where ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取发票号", lng结帐ID)
    strBillNO = Nvl(rsTemp!实际票号)
    
    '得到主页ID
    gstrSQL = "Select Nvl(住院次数,0) 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的主页ID", lng病人ID)
    lng主页ID = rsTemp!主页ID
    
    '如果存在需要审批的项目却未审批，不允许结算
    '20060829 断点
    gstrSQL = "Select Count(*) From 审核项目表 Where 审核标志=0 And 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
    Call OpenRecordset_OtherBase(rsTemp, "如果存在需要审批的项目却未审批，不允许结算", gstrSQL, gcn重庆银海版)
    If rsTemp.Fields(0).Value > 0 Then
        MsgBox "还有" & rsTemp.Fields(0).Value & "条待审批项目，不允许结算！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '结算类型  0-正常结算;1-中途结算;2-转家庭病床结算;3-卡挂失结算;5-单病种结算
    intState = 1
    gstrSQL = "Select Nvl(当前状态,0) 状态 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人当前状态", lng病人ID, TYPE_重庆银海版)
    '如果是出院结算，设置为正常结算标志
    If rsTemp!状态 = 0 Then intState = 0
    '在银海医保内，单病种病人不存在中途结算的概念，因此不管是中结还是出院结算，都传5
    If bln单病种 Then intState = 5
    
    If intState = 1 Then        '中途结算则提示是否转出院家庭病床结算
        If MsgBox("该病人是否进行转家庭病床结算？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            intState = 2
        End If
    End If
    
    '考虑上次结算可能中心成功，而HIS未成功，因此，需调用交易22获取结算信息，如果费用总额为零，说明上次结算成功，本次按正常流程结算即可
    '----------获取上次结算数据----------
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.结算流水号 & gstrSplit_Col_重庆银海版 & GetFileName(结算信息)
    Call 调用接口_准备_重庆银海版("22", StrInput)
    If Not 调用接口_重庆银海版 Then Exit Function
    If Not AnalyFile_Balance(strReturn) Then Exit Function
    
    '分解结算支付信息
    If gComInfo_重庆银海版.总费用_中心 <> 0 Then
        '如果费用总额不为零，则说明上次已结算，新的结算流水号以返回的为准进行结算
        gComInfo_重庆银海版.结算流水号 = gstrReturn_重庆银海版
        arrBalance = Split(strReturn, "|")
        intBalances = UBound(arrBalance)
        For intBalance = 0 To intBalances
            str结算方式 = Split(arrBalance(intBalance), ";")(0)
            curMoney = Split(arrBalance(intBalance), ";")(1)
            Select Case str结算方式
            Case "个人帐户"
                cur个人帐户_OLD = curMoney
            Case "医保基金"
                cur医保基金_OLD = curMoney
            Case "大病基金"
                cur大病基金_OLD = curMoney
            Case "公务员补助基金"
                cur公务员补助基金_OLD = curMoney
            End Select
        Next
        If Not AnalyFile_Balance(strReturn, True) Then Exit Function
    End If
    
    gcn重庆银海版.BeginTrans
    blnTrans = True
    '----正式结算----
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  18      结算交易流水号
'    3.      string  3       结算类型  0-正常结算;1-中途结算;2-转家庭病床结算;3-卡挂失结算
'    4.      string  20      病种编码
'    5.      string  20      发票号
'    6.      string          处方明细信息文件保存的路径及文件名
'    7.      string          该次就诊历次费用结算结果文件保存的路径及文件名
'    8.      string          传出医疗待遇信息文件保存的路径及文件名
'    9.      string          传出费用结算结果文件保存的路径及文件名
'    OutputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      结算交易流水号
    '先产生入参中涉及到的文件，再调用接口
    '处方明细在预结算时已经产生了，不必再次产生
    If Not MakeFile_Deal(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    If Not MakeFile_Balance(gComInfo_重庆银海版.就诊流水号) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    '----------获取本次结算数据----------
    StrInput = gComInfo_重庆银海版.就诊流水号 & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.结算流水号 & gstrSplit_Col_重庆银海版 & _
             intState & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.疾病编码 & gstrSplit_Col_重庆银海版 & _
             strBillNO & gstrSplit_Col_重庆银海版 & GetFileName(处方明细) & gstrSplit_Col_重庆银海版 & _
             GetFileName(结算信息) & gstrSplit_Col_重庆银海版 & GetFileName(待遇信息) & gstrSplit_Col_重庆银海版 & GetFileName(结算信息)
    Call 调用接口_准备_重庆银海版("14", StrInput)
    
    '如果是单病种，在结算交易使用完其状态后，需要将状态改为中途结算或出院结算
    If intState = 5 Then
        intState = 1
        gstrSQL = "Select Nvl(当前状态,0) 状态 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人当前状态", lng病人ID, TYPE_重庆银海版)
        '如果是出院结算，设置为正常结算标志
        If rsTemp!状态 = 0 Then intState = 0
    End If
    
    If Not 调用接口_重庆银海版 Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    str结算流水号 = gstrReturn_重庆银海版
    If Not AnalyFile_Deal(True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    strReturn = ""
    If Not AnalyFile_Balance(strReturn, True) Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    '分解结算支付信息
    arrBalance = Split(strReturn, "|")
    intBalances = UBound(arrBalance)
    For intBalance = 0 To intBalances
        str结算方式 = Split(arrBalance(intBalance), ";")(0)
        curMoney = Split(arrBalance(intBalance), ";")(1)
        Select Case str结算方式
        Case "个人帐户"
            cur个人帐户 = curMoney
        Case "医保基金"
            cur医保基金 = curMoney
        Case "大病基金"
            cur大病基金 = curMoney
        Case "公务员补助基金"
            cur公务员补助基金 = curMoney
        End Select
    Next
    
    '不应该累加，因银海会自动冲销上次的结算记录，判断条件是使用相同的结算流水号
'    '累加两次的结算数据，为本次的结算结果
'    cur个人帐户 = cur个人帐户 + cur个人帐户_OLD
'    cur医保基金 = cur医保基金 + cur医保基金_OLD
'    cur大病基金 = cur大病基金 + cur大病基金_OLD
'    cur公务员补助基金 = cur公务员补助基金 + cur公务员补助基金_OLD
'    gComInfo_重庆银海版.总费用 = gComInfo_重庆银海版.总费用 + gComInfo_重庆银海版.总费用_中心
'    cur现金 = gComInfo_重庆银海版.总费用 - cur个人帐户 - cur医保基金 - cur大病基金 - cur公务员补助基金
    
    '比较虚拟结算与正式结算结果是否一致
    If Not (cur个人帐户 = pre_Balance.cur个人帐户 And cur医保基金 = pre_Balance.cur医保基金 And _
        cur大病基金 = pre_Balance.cur大病基金 And cur公务员补助基金 = pre_Balance.cur公务员补助) Then
        blnRevise = True
        #If gverControl < 2 Then
            blnOld = True
        #End If
    End If
    
    '更新保险帐户（新的结算流水号）
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'结算流水号','''" & str结算流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算交易流水号")
    
    '保存本次结算情况
    '进入统筹金额=公务员补助基金;统筹报销金额=统筹基金;大病自付金额=大病基金
    '备注=业务类型|就诊流水号|结算流水号|就诊时间|疾病编码|并发症
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_重庆银海版 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_重庆银海版.总费用 & "," & cur现金 & "," & 0 & "," & cur公务员补助基金 & "," & cur医保基金 & "," & cur大病基金 & "," & _
        0 & "," & cur个人帐户 & ",null,null,null,'" & gComInfo_重庆银海版.业务类型 & "|" & gComInfo_重庆银海版.就诊流水号 & "|" & gComInfo_重庆银海版.结算流水号 & "|" & gComInfo_重庆银海版.就诊时间 & "|" & gComInfo_重庆银海版.疾病编码 & "|" & gComInfo_重庆银海版.并发症 & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")

    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将结帐记录打上上传标志")
    
    '由于预结算的结果80%可能与正式结算的结果不一致（因为银海的预结算接口不会去取病人最新的待遇信息，因此需要修正）
    str结算方式 = ""
    If cur个人帐户 <> 0 Then str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    If cur医保基金 <> 0 Then str结算方式 = str结算方式 & "||医保基金|" & cur医保基金
    If cur大病基金 <> 0 Then str结算方式 = str结算方式 & "||大病基金|" & cur大病基金
    If cur公务员补助基金 <> 0 Then str结算方式 = str结算方式 & "||公务员补助基金|" & cur公务员补助基金
    If str结算方式 <> "" And blnRevise Then
        str结算方式 = Mid(str结算方式, 3)
        If blnOld Then
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
        Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    
    '取入院相关信息
    Dim str就诊时间 As String, str科室 As String, str医生 As String, str出院原因 As String
    gstrSQL = " Select to_char(A.入院日期,'yyyy-MM-dd') 入院日期,B.名称 科室,A.出院方式,A.住院医师 医生 From 病案主页 A,部门表 B " & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院日期", lng病人ID, lng主页ID)
    str就诊时间 = Format(rsTemp!入院日期, "yyyyMMdd") & " 00:00:00"
    str科室 = Nvl(rsTemp!科室)
    str医生 = IIf(IsNull(rsTemp!医生), "银海", rsTemp!医生)
    str出院原因 = IIf(IsNull(rsTemp!出院方式), "", rsTemp!出院方式)
    Select Case str出院原因
    Case "正常", "康复"
        str出院原因 = 1
    Case "死亡"
        str出院原因 = 3
    Case "转院"
        str出院原因 = 2
    Case Else
        str出院原因 = 9
    End Select
    
    '取业务相关信息
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
'    2.      string  3       医疗类别
'    3.      string  30      科室
'    4.      string  20      医生
'    5.      datetime        日  入院日期
'    6.      string  20      入院疾病
'    7.      string  3       就诊状态
'    8.      datetime        日  出院日期
'    9.      string  20      确诊疾病编码
'    10.     string  3       出院原因
'    11.     string  20      经办人
'    12.     string  50      并发症
    gstrSQL = " Select A.流水号,A.业务类型,B.编码 疾病编码,A.并发症 From 保险帐户 A,保险病种 B " & _
              " Where A.病人ID=[1] And A.险类=[2] And A.病种ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取其它相关信息", lng病人ID, TYPE_重庆银海版)
    
    StrInput = rsTemp!流水号 & gstrSplit_Col_重庆银海版 & rsTemp!业务类型 & gstrSplit_Col_重庆银海版 & _
            str科室 & gstrSplit_Col_重庆银海版 & Nvl(str医生, "银海") & gstrSplit_Col_重庆银海版 & _
            str就诊时间 & gstrSplit_Col_重庆银海版 & rsTemp!疾病编码 & gstrSplit_Col_重庆银海版 & _
            "1" & gstrSplit_Col_重庆银海版 & Format(zlDatabase.Currentdate(), "yyyyMMdd") & " 00:00:00" & gstrSplit_Col_重庆银海版 & _
            rsTemp!疾病编码 & gstrSplit_Col_重庆银海版 & str出院原因 & gstrSplit_Col_重庆银海版 & _
            ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & ToVarchar(Nvl(rsTemp!并发症), 50)
    Call 调用接口_准备_重庆银海版("09", StrInput)
    If Not 调用接口_重庆银海版() Then
        gcn重庆银海版.RollbackTrans
        Exit Function
    End If
    
    gcn重庆银海版.CommitTrans
    住院结算_重庆银海版 = True
    
    '同时办理出院登记
    If intState = 0 Then Call 出院登记_重庆银海版(lng病人ID, lng主页ID)
    Exit Function
errHand:
    If blnTrans Then gcn重庆银海版.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_重庆银海版(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    MsgBox "医保不支持结帐作废，请直接作废记帐单据后，再结帐！", vbInformation, gstrSysName
    住院结算冲销_重庆银海版 = False
End Function

Private Function Get保险参数_重庆银海版(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名=[1] and A.险类=[2] and A.中心 is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重庆医保", str参数名, TYPE_重庆银海版)
    
    If rsTemp.EOF = False Then
        Get保险参数_重庆银海版 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Function 价格判断_重庆银海版(ByVal dbl医院 As Double, ByVal dbl医保 As Double, ByVal str限价方式 As String, _
                              ByVal bln特价 As Boolean, ByVal dbl特价 As Double) As Boolean
'功能：判断医院的价格是否超过医保规定的单价
    Dim str医院类别 As String
    
    On Error GoTo errHandle
    
    If InStr(str限价方式, "二级") > 0 Then
        str医院类别 = Get保险参数_重庆银海版("医院等级")
        '给出的标准价格为二级医院的最高限价，三级医院的最高限价在此基础上可以上浮10%，一级医院的最高限价在此基础上下调5%
        
        Select Case str医院类别
            Case "三级"
                dbl医保 = dbl医保 * 1.1
            Case "一级"
                dbl医保 = dbl医保 * 0.95
        End Select
    End If
    
    If bln特价 = True And dbl特价 > dbl医保 Then
        '允许使用特价
        dbl医保 = dbl特价
    End If
    
    If dbl医院 > dbl医保 Then
        If MsgBox("医院单价" & Format(dbl医院, "0.000") & " 高于医保中心核准的价格" & Format(dbl医保, "0.000") & "，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    价格判断_重庆银海版 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 更新疾病_重庆银海版(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim int业务类型 As Integer
    Dim lng病种ID As Long
    Dim StrInput As String
    Dim str并发症 As String, str疾病编码 As String
    Dim str流水号 As String, str科室 As String, str医生 As String, str就诊时间 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    
    '获得病人出院病种及并发症
    gstrSQL = " Select B.编码 病种编码,A.并发症,A.业务类型,A.流水号 From 保险帐户 A,保险病种 B " & _
              " Where A.病种ID=B.ID And A.险类=[1] ANd A.险类=B.险类 And A.病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人出院病种及并发症", TYPE_重庆银海版, lng病人ID)
    str疾病编码 = Nvl(rsTemp!病种编码)
    str并发症 = Nvl(rsTemp!并发症)
    int业务类型 = Nvl(rsTemp!业务类型)
    str流水号 = Nvl(rsTemp!流水号)
    
    '取入院相关信息
    gstrSQL = " Select to_char(A.入院日期,'yyyy-MM-dd') 入院日期,B.名称 科室,A.出院方式,A.住院医师 医生 From 病案主页 A,部门表 B " & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院日期", lng病人ID, lng主页ID)
    str就诊时间 = Format(rsTemp!入院日期, "yyyyMMdd") & " 00:00:00"
    str科室 = rsTemp!科室
    str医生 = Nvl(rsTemp!医生, "银海")
    
    '让操作员修改病种信息和并发症
    If frm病种选择_重庆银海版.ShowSelect(frmParent, int业务类型, str疾病编码, str并发症) = False Then
        Exit Function
    End If
    
    '根据病种编码取该病种的ID
    gstrSQL = "Select ID From 保险病种 Where 险类=[1] And 编码=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病种ID", TYPE_重庆银海版, str疾病编码)
    lng病种ID = rsTemp!ID
    
    '更新保险帐户
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'病种ID','" & lng病种ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆银海版 & ",'并发症','''" & str并发症 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新并发症")
    
    '更新病人相关状态
    StrInput = str流水号 & gstrSplit_Col_重庆银海版 & int业务类型 & gstrSplit_Col_重庆银海版 & _
            str科室 & gstrSplit_Col_重庆银海版 & str医生 & gstrSplit_Col_重庆银海版 & _
            str就诊时间 & gstrSplit_Col_重庆银海版 & str疾病编码 & gstrSplit_Col_重庆银海版 & _
            "1" & gstrSplit_Col_重庆银海版 & "" & gstrSplit_Col_重庆银海版 & _
            "" & gstrSplit_Col_重庆银海版 & "" & gstrSplit_Col_重庆银海版 & _
            ToVarchar(gstrUserName, 20) & gstrSplit_Col_重庆银海版 & ToVarchar(str并发症, 50)
    Call 调用接口_准备_重庆银海版("09", StrInput)
    If Not 调用接口_重庆银海版() Then Exit Function
    
    gcnOracle.CommitTrans
    更新疾病_重庆银海版 = True
    
    '可能待遇信息发生了变化，需要重新获取
'    InputString
'    序号    数据类型    长度    精度    说明
'    1   string  18      门诊/住院流水号
'    2   string  20      病种编码（入院诊断）
'    3   datetime        日  入院日期
'    4   string          处理后的医疗待遇信息文件保存的路径及文件名
    StrInput = str流水号 & gstrSplit_Col_重庆银海版 & str疾病编码 & gstrSplit_Col_重庆银海版 & _
            str就诊时间 & gstrSplit_Col_重庆银海版 & GetFileName(待遇信息)
    Call 调用接口_准备_重庆银海版("23", StrInput)
    If Not 调用接口_重庆银海版() Then
        MsgBox "获取待遇信息时发生错误！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not AnalyFile_Deal(True) Then
        MsgBox "分析待遇信息文件时发生错误！", vbInformation, gstrSysName
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Public Sub TestVerifyItem()
    Dim StrInput As String, strReturn As String
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    
    Const int流水号 As Integer = 0
    Const int处方流水号 As Integer = 1
    Const int高收费审批编号 As Integer = 2
    Const int审批标志 As Integer = 3
    Const int先自付金额 As Integer = 4
    '未审批的项目再次调用处方明细计算,并保存返回结果
    gstrSQL = " Select 流水号,处方流水号 From 中间库_处方明细" & _
              " Where 审批标志 = '0' And 流水号='" & gComInfo_重庆银海版.就诊流水号 & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn重庆银海版
    End With
    
    Do While Not rsTemp.EOF
        StrInput = rsTemp!流水号 & gstrSplit_Col_重庆银海版 & rsTemp!处方流水号
        Call 调用接口_准备_重庆银海版("11", StrInput)
        If 调用接口_重庆银海版 Then
            strReturn = gstrReturn_重庆银海版
            '流水号,处方流水号,高收费审批编号,审批标志,先自付金额
            arrData = Split(strReturn, gstrSplit_Col_重庆银海版)
            gcn重庆银海版.Execute "zl_中间库_处方明细_UPDATE(" & _
                            "'" & arrData(int流水号) & "','" & arrData(int处方流水号) & "'," & _
                            "'" & arrData(int高收费审批编号) & "','" & arrData(int审批标志) & "'," & _
                            "'" & arrData(int先自付金额) & "')", , adCmdStoredProc
        End If
        rsTemp.MoveNext
    Loop
End Sub

Public Function CheckItem() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '检查是否存在未通过审批的高收费项目
    gstrSQL = " Select 流水号,处方流水号,审批标志 From 中间库_处方明细" & _
              " Where 审批标志 = '0' And 流水号='" & gComInfo_重庆银海版.就诊流水号 & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn重庆银海版
        CheckItem = (.RecordCount = 0)
    End With
End Function

Public Sub 核对费用结算_重庆银海版()
    '将中间库中保存的数据和中心端的数据进行对比
    Dim StrInput As String, strReturn As String
    Dim strStart As String, strEnd As String
    Dim str就诊流水号_开始 As String, str就诊流水号_结束 As String
    Dim cur基本统筹_医保 As Currency, cur大病统筹_医保 As Currency, cur公务员补助_医保 As Currency
    Dim cur基本统筹_医院 As Currency, cur大病统筹_医院 As Currency, cur公务员补助_医院 As Currency
    Dim arrReturn
    Dim rsTemp As New ADODB.Recordset
    
    If frm日期范围_米易.Show_ME(strStart, strEnd) = False Then Exit Sub
    gstrSQL = " Select min(流水号) 开始流水号,max(流水号) 结束流水号 From " & mstrOwner & ".中间库_结算信息" & _
            " Where 经办时间 Between [1] And [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取指定时间范围内的流水号", CDate(strStart), CDate(strEnd))
    str就诊流水号_开始 = Nvl(rsTemp!开始流水号)
    str就诊流水号_结束 = Nvl(rsTemp!结束流水号)
    If str就诊流水号_结束 = "" And str就诊流水号_开始 = "" Then
        MsgBox "指定的期间内没有发生任何医保费用！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '准备调用核对费用接口
    If Not 医保初始化_重庆银海版 Then Exit Sub
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  3       统筹区号
'    2.      string  14      定点医疗机构编号
'    3.      string  18      需要核对的费用结算信息的起始门诊/住院流水号
'    4.      string  18      需要核对的费用结算信息的截止门诊/住院流水号
'    5.      string  18      结算交易流水号
'    6.      string  18      审批记录序号(结算状态流水号)
'    OutputString
'    序号    数据类型    长度    精度    说明
'    1.      number  15      在核对范围内的所有信息的记录数量
'    2.      number  14  2   在核对范围内的所有记录的个人自付总额累加值
'    3.      number  14  2   在核对范围内的所有记录的基本医疗统筹支付总额累加值
'    4.      number  14  2   在核对范围内的所有记录的公务员补助总额累加值
'    5.      number  14  2   在核对范围内的所有记录的大额理赔总额累加值
    StrInput = "" & gstrSplit_Col_重庆银海版 & gComInfo_重庆银海版.医院编码 & gstrSplit_Col_重庆银海版 & _
               str就诊流水号_开始 & gstrSplit_Col_重庆银海版 & str就诊流水号_结束 & gstrSplit_Col_重庆银海版 & _
               "" & gstrSplit_Col_重庆银海版 & ""
    Call 调用接口_准备_重庆银海版("16", StrInput)
    If Not 调用接口_重庆银海版 Then Exit Sub
    strReturn = gstrReturn_重庆银海版
    
    '分解返回串
    arrReturn = Split(strReturn, gstrSplit_Col_重庆银海版)
    cur基本统筹_医保 = Val(arrReturn(2))
    cur公务员补助_医保 = Val(arrReturn(3))
    cur大病统筹_医保 = Val(arrReturn(4))
    
    '提取中间库中保存的结算信息
    gstrSQL = " Select Sum(Nvl(本次基本统筹金额,0)) 基本统筹,SUM(Nvl(大病支付金额,0)) 大病统筹," & _
              " Sum(Nvl(先自付部分公务员补助,0)+Nvl(历史先自付公务员补助,0)+Nvl(起付线下公务员补助,0)+Nvl(本次普通门诊公务员补助,0)+Nvl(分段自付公务员补助,0)+Nvl(转院起付线纳入公务员,0)) 公务员补助" & _
              " From " & mstrOwner & ".中间库_结算信息" & _
              " Where 流水号>=[1] And 流水号<=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取中间库中的结算信息", str就诊流水号_开始, str就诊流水号_结束)
    cur基本统筹_医院 = Val(Nvl(rsTemp!基本统筹, 0))
    cur大病统筹_医院 = Val(Nvl(rsTemp!大病统筹, 0))
    cur公务员补助_医院 = Val(Nvl(rsTemp!公务员补助, 0))
    
    '检查是否相同
    If Format(cur基本统筹_医保, "#####0.00") = Format(cur基本统筹_医院, "#####0.00") And _
    Format(cur大病统筹_医保, "#####0.00") = Format(cur大病统筹_医院, "#####0.00") And _
    Format(cur公务员补助_医保, "#####0.00") = Format(cur公务员补助_医院, "#####0.00") Then
        MsgBox "基本统筹、大病补助及公务员补助金额与中心一致！", vbInformation, gstrSysName
    Else
        MsgBox "核对金额不一致：" & vbCrLf & _
        "基本统筹：（医保）" & Format(cur基本统筹_医保, "#####0.00") & Space(10) & "（医院）" & Format(cur基本统筹_医院, "#####0.00") & vbCrLf & _
        "大病补助：（医保）" & Format(cur大病统筹_医保, "#####0.00") & Space(10) & "（医院）" & Format(cur大病统筹_医院, "#####0.00") & vbCrLf & _
        "公务员补助：（医保）" & Format(cur公务员补助_医保, "#####0.00") & Space(10) & "（医院）" & Format(cur公务员补助_医院, "#####0.00"), vbInformation, gstrSysName
    End If
End Sub

Private Function GetSequence(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng收费细目ID As Long) As String
    '随机取一条正常记录的流水号（用于负数记帐）
    Dim rsTemp As New ADODB.Recordset
    GetSequence = ""
    
    gstrSQL = " Select NO,记录性质,记录状态,序号 From 住院费用记录" & _
              " Where 收费细目ID=[1] And 病人ID=[2] And 主页ID=[3]" & _
              " And 记录状态=1 And Nvl(实收金额,0)>0 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取流水号", lng收费细目ID, lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then
        GetSequence = rsTemp!NO & String(3 - Len(rsTemp!记录性质), "0") & rsTemp!记录性质 & _
                    String(3 - Len(rsTemp!记录状态), "0") & rsTemp!记录状态 & _
                    String(3 - Len(rsTemp!序号), "0") & rsTemp!序号
    Else
        Call DebugTool("未找到原始处方明细[病人ID:" & lng病人ID & "|主页ID:" & lng主页ID & "|收费细目ID:" & lng收费细目ID)
    End If
End Function

Public Function 转为普通病人_银海(ByVal lng病人ID As Long) As Boolean
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    '获取原就诊流水号
    gstrSQL = "Select 流水号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊流水号", TYPE_重庆银海版, lng病人ID)
    gComInfo_重庆银海版.就诊流水号 = rsTemp!流水号
    
    '调用就诊登记作废接口
'    InputString
'    序号    数据类型    长度    精度    说明
'    1.      string  18      门诊/住院流水号
    Call 调用接口_准备_重庆银海版("15", gComInfo_重庆银海版.就诊流水号)
    If Not 调用接口_重庆银海版() Then Exit Function
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重庆银海版 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    转为普通病人_银海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AnalyBalance(ByVal strBalance As String)
    Dim arrBalance
    Dim STRNAME As String
    Dim dblMoney As Double
    Dim intDO As Integer, intCOUNT As Integer
    '分析结算返回串，将信息填充到结构体中
    
    pre_Balance.cur个人帐户 = 0
    pre_Balance.cur医保基金 = 0
    pre_Balance.cur公务员补助 = 0
    pre_Balance.cur大病基金 = 0
    
    arrBalance = Split(strBalance, "|")
    intCOUNT = UBound(arrBalance)
    For intDO = 0 To intCOUNT
        STRNAME = Split(arrBalance(intDO), ";")(0)
        dblMoney = Val(Split(arrBalance(intDO), ";")(1))
        Select Case STRNAME
        Case "个人帐户"
            pre_Balance.cur个人帐户 = dblMoney
        Case "医保基金"
            pre_Balance.cur医保基金 = dblMoney
        Case "公务员补助基金"
            pre_Balance.cur公务员补助 = dblMoney
        Case "大病基金"
            pre_Balance.cur大病基金 = dblMoney
        End Select
    Next
End Sub
