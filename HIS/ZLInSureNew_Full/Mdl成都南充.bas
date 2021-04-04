Attribute VB_Name = "Mdl成都南充"
Option Explicit
Public Const gstrSplit大类 As String = "♂"
Public Const gstrSplit小类 As String = "♀"
Public Const gstr费用项目 As String = "床位费♂西药费♂中成药费♂中草药费♂手术费♂化验费♂检查费♂治疗费♂材料费♂血费♂氧气费♂器官费♂护理费♂陪伴费♂CT费♂核磁共振♂其他费"
Public Const gstrCol_ENG As String = "BH,ID,ZWMC,JLDW,DJ,YPXM,YPLX,YPZLX,YPXLX,YPSHQF,XZSYFW,YPXMLH"
Public Const gstrCol_CHI As String = "编号,医保项目ID,中文名称,计量单位,单价,药品项目,药品类型,药品中类型,药品小类型,药品使用区分,药品使用范围,药品项目内涵"
Public gcnInterbase As New ADODB.Connection

Private rsTemp As New ADODB.Recordset
Private Const giniPath As String = "c:\his_yb"
Private Const giniFile As String = "his_yb.ini"
Private strSQL As String
Private strProcedure As String
Private intReturn As Integer

Type Bill_Head
    住院号 As String
    处方流水号 As String
    处方时间 As Date
    医生 As String
    质量 As String
    科室 As String
End Type
Type Bill_Body
    处方明细流水号 As Long
    医保收费细目 As Long
    单价 As Currency
    数量 As Single
    费用项目 As String
    剂量单位 As String
    '--以下参数仅对药品有效，可以为空
    商品名 As String
    规格 As String
    类型 As String  '空，甲类，乙类
    进价 As Currency
End Type
Private 处方头 As Bill_Head
Private 处方体 As Bill_Body

Public Function 医保设置_成都南充() As Boolean
'功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
'返回：接口配置成功，返回true；否则，返回false
    Dim strConn As String
    
    On Error GoTo errHand
    If frmSet成都.ShowSet(TYPE_成都南充) = False Then
        Exit Function
    End If
    
    strConn = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
    '重新建立到医保服务器的公共连接
    If gcnInterbase.State = adStateClosed Then
        On Error Resume Next
        gcnInterbase.Open strConn
        If Err = 0 Then
            医保设置_成都南充 = True
        Else
            Err.Clear
        End If
    Else
        医保设置_成都南充 = True
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保初始化_成都南充() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    On Error GoTo errHand
    '建立到医保服务器的公共连接
    Dim strConn As String
    strConn = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
    Err = 0
    On Error Resume Next
    With gcnInterbase
        If .State = adStateOpen Then .Close
        .ConnectionString = strConn
        .Open
        If Err <> 0 Then
            MsgBox "不能建立到医保服务器的连接，无法执行医保交易", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    
    医保初始化_成都南充 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 身份标识_成都南充(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    On Error GoTo errHand
    Dim strTmpIden As String
    
    strTmpIden = frmIdentify成都南充.ShowCard(bytType, lng病人ID)
    身份标识_成都南充 = strTmpIden
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 处方登记_成都南充_指定费用(ByVal lngPatient As Long) As Boolean
    '上传剩下的费用（主要是床位费、护理费等）
    '先写入单据头，再写入单据体
    '记录状态（1-新增;否则为删除），费用处单据只能整张单据删除后，再产生新单据
    On Error GoTo errHand
    处方登记_成都南充_指定费用 = False
    处方头.住院号 = ""
     
    gstrSQL = " Select A.ID 处方明细流水号,A.标识号 as 住院号,A.病人ID,A.记录性质,A.记录状态,A.NO,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 处方时间," & _
              " A.开单人 医生,'' 质量,B.名称 科室,C.项目编码 医保收费细目,A.标准单价 单价,A.数次*Nvl(A.付数,1) 数量,D.类别 费用项目,'['||E.编码||']'||E.名称 收费细目," & _
              " E.类别 收费细目类别,substrb(E.规格,1,60) 规格,E.费用类型 类型,A.标准单价 进价" & _
              " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[1]) C,收费类别 D,收费细目 E " & _
              " Where A.执行部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And A.收费类别=D.编码 And A.收费细目ID=E.ID " & _
              " And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And A.实收金额 Is Not NULL" & _
              " And A.病人ID=[2]" & _
              " Order by A.标识号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方登记", TYPE_成都南充, lngPatient)
 
    
    If Not 上传_处方登记(rsTemp) Then Exit Function
    
    处方登记_成都南充_指定费用 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 处方登记_成都南充(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '先写入单据头，再写入单据体
    '记录状态（1-新增;否则为删除），费用处单据只能整张单据删除后，再产生新单据
    On Error GoTo errHand
    处方登记_成都南充 = False
    处方头.住院号 = ""
    
    gstrSQL = " Select A.ID 处方明细流水号,A.标识号 as 住院号,A.病人ID,A.记录性质,A.记录状态,A.NO,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 处方时间," & _
              " A.开单人 医生,'' 质量,B.名称 科室,C.项目编码 医保收费细目,A.标准单价 单价,A.数次*Nvl(A.付数,1) 数量,D.类别 费用项目,'['||E.编码||']'||E.名称 收费细目," & _
              " E.类别 收费细目类别,substrb(E.规格,1,60) 规格,E.费用类型 类型,A.标准单价 进价" & _
              " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[4]) C,收费类别 D,收费细目 E,保险帐户 F " & _
              " Where A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" & _
              " And A.执行部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And A.收费类别=D.编码 And A.收费细目ID=E.ID And A.病人ID=F.病人ID And F.险类=[4]" & _
              " And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And A.实收金额 Is Not NULL" & _
              " Order by A.标识号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方登记", lng记录性质, lng记录状态, str单据号, TYPE_成都南充)
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到处方记录，向医保服务器传输数据失败！[处方登记]", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If Not 上传_处方登记(rsTemp) Then Exit Function
    
    处方登记_成都南充 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function Get流水号(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, strNO)
    Get流水号 = lng记录性质 & lng记录状态 & (asc(Mid(strNO, 1, 1)) - 55) & Mid(strNO, 2)
End Function

Public Function 处方删除_成都南充(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    Dim blnNew As Boolean
    Dim blnExec As Boolean 'Modified.By.ZYB 2003-01-23 非医保病人不执行上传
    '先写入单据头，再写入单据体
    '记录状态（1-新增;否则为删除），费用处单据只能整张单据删除后，再产生新单据
    On Error GoTo errHand
    处方删除_成都南充 = False
    处方头.住院号 = ""
    
    gcnInterbase.BeginTrans
    gstrSQL = " Select A.ID 处方明细流水号,A.标识号 as 住院号,A.病人ID,A.记录性质,A.记录状态,A.NO,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 处方时间," & _
              " A.开单人 医生,'' 质量,B.名称 科室,C.项目编码 医保收费细目,A.标准单价 单价,A.数次*Nvl(A.付数,1) 数量,D.类别 费用项目,'['||E.编码||']'||E.名称 收费细目" & _
              " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=[4]) C,收费类别 D,收费细目 E,保险帐户 F " & _
              " Where A.记录性质=[1] And A.记录状态=[2] And A.NO=[3]" & _
              " And A.执行部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And A.收费类别=D.编码 And A.收费细目ID=E.ID And A.病人ID=F.病人ID And F.险类=[4]" & _
              " And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And A.实收金额 Is Not NULL" & _
              " Order by A.标识号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处方删除", lng记录性质, lng记录状态, str单据号, TYPE_成都南充)
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到处方记录，向医保服务器传输数据失败！[处方删除]", vbInformation, gstrSysName
            gcnInterbase.RollbackTrans
            Exit Function
        End If
        
        Do While Not .EOF
            '写入处方头
            blnNew = (处方头.住院号 <> Get住院号(rsTemp!病人ID))
            blnExec = IsYBPatient(rsTemp!病人ID)
            If blnNew And blnExec Then
                With 处方头
                    .住院号 = Get住院号(rsTemp!病人ID)
                    .处方流水号 = Get流水号(lng记录性质, 1, str单据号)
                    .处方时间 = Format(rsTemp!处方时间, "yyyy-MM-dd HH:mm:ss")
                    .医生 = IIf(IsNull(rsTemp!医生), "", rsTemp!医生)
                    .质量 = IIf(IsNull(rsTemp!质量), "", rsTemp!质量)
                    .科室 = rsTemp!科室
                End With
                
                strProcedure = "DELETE_CFJLK"
                strSQL = "Execute Procedure DELETE_CFJLK('" & 处方头.住院号 & "'," & 处方头.处方流水号 & ")"
                If Not ExecProc(strSQL) Then gcnInterbase.RollbackTrans: Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    '更新病人费用记录处的上传标志为真
    If Not 更新上传标志(rsTemp) Then
        gcnInterbase.RollbackTrans
        Exit Function
    End If
    
    gcnInterbase.CommitTrans
    处方删除_成都南充 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    gcnInterbase.RollbackTrans
End Function

Public Function 上传_处方登记(ByVal rsTemp As ADODB.Recordset) As Boolean
    Dim blnNew As Boolean
    Dim blnExec As Boolean 'Modified.By.ZYB 2003-01-23 非医保病人不执行上传
    On Error GoTo errHand
    
    上传_处方登记 = False
    gcnInterbase.BeginTrans
    
    With rsTemp
        Do While Not .EOF
            '写入处方头
            blnNew = (处方头.住院号 <> Get住院号(rsTemp!病人ID))
            blnExec = IsYBPatient(rsTemp!病人ID)
            If blnExec Then
                If blnNew Then
                    With 处方头
                        .住院号 = Get住院号(rsTemp!病人ID)
                        .处方流水号 = Get流水号(rsTemp!记录性质, rsTemp!记录状态, rsTemp!NO)
                        .处方时间 = Format(rsTemp!处方时间, "yyyy-MM-dd HH:mm:ss")
                        .医生 = IIf(IsNull(rsTemp!医生), "", rsTemp!医生)
                        .质量 = IIf(IsNull(rsTemp!质量), "", rsTemp!质量)
                        .科室 = rsTemp!科室
                    End With
                    
                    strProcedure = "ADD_CFJLK"
                    strSQL = "Execute Procedure ADD_CFJLK('" & 处方头.住院号 & "'," & 处方头.处方流水号 & _
                             ",'" & 处方头.处方时间 & "','" & 处方头.医生 & "',NULL,'" & 处方头.科室 & "')"
                    If Not ExecProc(strSQL) Then gcnInterbase.RollbackTrans: Exit Function
                End If
                
                '写入处方明细
                With 处方体
                    .处方明细流水号 = rsTemp!处方明细流水号
                    .医保收费细目 = IIf(IsNull(rsTemp!医保收费细目), 0, rsTemp!医保收费细目)
                    .单价 = rsTemp!单价
                    .数量 = rsTemp!数量
                    .剂量单位 = ""
                    .费用项目 = Get费用项目(rsTemp!费用项目)
                    If InStr(1, ",5,6,7,", "," & rsTemp!收费细目类别 & ",") <> 0 Then
                        .商品名 = rsTemp!收费细目
                        .规格 = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                        .类型 = IIf(rsTemp!类型 = "甲类药", "甲类", IIf(rsTemp!类型 = "乙类药", "乙类", ""))
                    Else
                        .商品名 = ""
                        .规格 = ""
                        .类型 = ""
                    End If
                    .进价 = 0
                    
                    If .医保收费细目 = 0 Then
                        MsgBox rsTemp!收费细目 & "未设置对应的医保项目！[上传数据]", vbInformation, gstrSysName
                        gcnInterbase.RollbackTrans
                        Exit Function
                    End If
                    If .费用项目 = "" Then
                        MsgBox "请设置His系统中的收费类别与医保系统中费用项目的对照关系！[保险类别]", vbInformation, gstrSysName
                        gcnInterbase.RollbackTrans
                        Exit Function
                    End If
                End With
                
                strProcedure = "ADD_CFMXK"
                strSQL = "Execute Procedure ADD_CFMXK('" & 处方头.住院号 & "'," & 处方头.处方流水号 & _
                        "," & 处方体.处方明细流水号 & "," & 处方体.医保收费细目 & ",'" & 处方体.剂量单位 & _
                        "'," & 处方体.单价 & "," & 处方体.数量 & ",'" & 处方体.费用项目 & "','" & 处方体.商品名 & _
                        "','" & 处方体.规格 & "','" & 处方体.类型 & "'," & 处方体.进价 & ")"
                If Not ExecProc(strSQL) Then gcnInterbase.RollbackTrans: Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    '更新病人费用记录处的上传标志为真
    If Not 更新上传标志(rsTemp) Then
        gcnInterbase.RollbackTrans
        Exit Function
    End If
    
    gcnInterbase.CommitTrans
    上传_处方登记 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    gcnInterbase.RollbackTrans
End Function

Public Function 入院登记_成都南充(ByVal lngPatient As Long) As Boolean
    Dim strObj As String, str住院号 As String, blnExist As Boolean
    Dim strLine As TextStream, FileSys As New FileSystemObject
    '将入院病人的住院号写入本机的(c:\his_yb\his_yb.ini)
    '格式为：zyh=11111111
    '同时更新保险帐户和病案主页
    
    On Error GoTo errHand
    入院登记_成都南充 = False
    
    '入院登记
    gstrSQL = "zl_保险帐户_入院(" & lngPatient & "," & TYPE_成都南充 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "南充医保")
    
    '如果路径不存在，则生成
    If Not FileSys.FolderExists(giniPath) Then FileSys.CreateFolder (giniPath)
    '如果文件存在，则删除后重新产生
    blnExist = FileSys.FileExists(giniPath & "\" & giniFile)
    If blnExist Then Call FileSys.DeleteFile(giniPath & "\" & giniFile, True)
    str住院号 = Get住院号(lngPatient)
    '查找是否存在该对象
    Set strLine = FileSys.OpenTextFile(giniPath & "\" & giniFile, ForWriting, True)
    
    Call strLine.WriteLine("[String]")  'Modified.By.ZYB 2003-01-23
    Call strLine.WriteLine("ZYH=" & str住院号)
    strLine.Close
    
    入院登记_成都南充 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_成都南充(ByVal lngPatient As Long) As Boolean
    On Error GoTo errHand
    出院登记_成都南充 = False
    
    '上传剩下的费用（主要是床位费、护理费等）
    If Not 处方登记_成都南充_指定费用(lngPatient) Then Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lngPatient & "," & TYPE_成都南充 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "南充医保")
    
    出院登记_成都南充 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记撤销_成都南充(ByVal lngPatient As Long) As Boolean
    On Error GoTo errHand
    出院登记撤销_成都南充 = False

    '恢复在院
    gstrSQL = "zl_保险帐户_入院(" & lngPatient & "," & TYPE_成都南充 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "南充医保")
    
    出院登记撤销_成都南充 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function Get住院号(ByVal lngPatient As Long) As String
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        ''Modified.By.ZYB 2003-01-23 必须每次住院，其住院号都必须唯一，所以加上住院次数
        gstrSQL = "Select 住院号||'_'||住院次数 住院号 From 病人信息 Where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "南充医保", lngPatient)
        
        Get住院号 = !住院号
    End With
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function IsYBPatient(ByVal lngPatient As Long) As Boolean
    Dim rsYbPatient As New ADODB.Recordset
    On Error GoTo errHand
    '判断是否是医保病人
    IsYBPatient = False
    
    gstrSQL = "Select Count(*) Records From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsYbPatient = zlDatabase.OpenSQLRecord(gstrSQL, "南充医保", TYPE_成都南充, lngPatient)
        
    With rsYbPatient
        If .EOF Then Exit Function
        If IsNull(!Records) Then Exit Function
        If !Records = 0 Then Exit Function
    End With
    
    IsYBPatient = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 住院虚拟结算_成都南充(ByVal lngPatient As Long) As String
    On Error GoTo errHand
    Dim rsPay As New ADODB.Recordset, curPay As Currency, str住院号 As String
    住院虚拟结算_成都南充 = ""
    
    str住院号 = Get住院号(lngPatient)
    
    strProcedure = "GET_SBBXJE"
    With rsPay
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Execute Procedure GET_SBBXJE('" & str住院号 & "')", gcnInterbase
        curPay = IIf(IsNull(!BXXJ), 0, !BXXJ)
        intReturn = !SUCC
    End With
    If intReturn <> 0 Then
        Call IsError
        住院虚拟结算_成都南充 = ""
        Exit Function
    End If
    住院虚拟结算_成都南充 = "医保基金;" & curPay & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 住院结算_成都南充(ByVal lng结帐ID As Long, ByVal rsTmp As ADODB.Recordset) As Boolean
    Dim curPay As Currency
    '必须病人在医保数据库中出院，方可调用结算过程GET_SBBXJE
    '不支持结算后退费操作，病人与医保中心接触解决
    
    curPay = Split(住院虚拟结算_成都南充(rsTmp!病人ID), ";")(1)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_成都南充 & "," & rsTmp!病人ID & "," & _
        Int(Format(zlDatabase.Currentdate, "yyyy")) & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & 0 & "," & 0 & "," & 0 & "," & curPay & ",0," & _
        0 & "," & 0 & ",NULL," & 0 & "," & 0 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "南充医保")
    
    住院结算_成都南充 = True
End Function

Public Function Get费用项目(ByVal str收费类别 As String) As String
    On Error GoTo errHand
    Dim str费用项目 As String, arrItem, intItem As Integer
    '获取与收费类别对应的医保费用项目
    Get费用项目 = ""
    str费用项目 = Trim(GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCItem"), ""))
    If str费用项目 = "" Then Exit Function
    
    arrItem = Split(str费用项目, gstrSplit大类)
    For intItem = 0 To UBound(arrItem)
        If Split(arrItem(intItem), gstrSplit小类)(0) = str收费类别 Then
            Get费用项目 = Split(arrItem(intItem), gstrSplit小类)(1)
            Exit For
        End If
    Next
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ExchangeColName(ByVal strCol As String, Optional ByVal blnExchange As Boolean = True) As String
    Dim arrEng, arrChi, arrTemp
    Dim intExchange As Integer, intFind As Integer
    '英文列名与中文列名互相转换
    On Error GoTo errHand
    
    arrEng = Split(gstrCol_ENG, ",")
    arrChi = Split(gstrCol_CHI, ",")
    If blnExchange Then
        arrTemp = arrEng
    Else
        arrTemp = arrChi
    End If
    
    For intExchange = 0 To UBound(arrTemp)
        If arrTemp(intExchange) = strCol Then
            intFind = intExchange
            Exit For
        End If
    Next
    
    If blnExchange Then
        ExchangeColName = arrChi(intFind)
    Else
        ExchangeColName = arrEng(intFind)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 更新上传标志(ByVal rsTemp As ADODB.Recordset) As Boolean
    On Error GoTo errHand
    更新上传标志 = False
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "ZL_病人记帐记录_上传(" & !处方明细流水号 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "南充医保")
            .MoveNext
        Loop
    End With
    更新上传标志 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ExecProc(ByVal strExec As String, Optional ByVal bln提示 As Boolean = True) As Boolean
    Dim rsExecute As New ADODB.Recordset
    On Error GoTo errHand
    With rsExecute
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open strExec, gcnInterbase
        If .RecordCount = 0 Then
            MsgBox "向医保服务器传送数据过程中，发生未知错误！", vbInformation, gstrSysName
            Exit Function
        End If
        intReturn = .Fields(0).Value
    End With
    
    ExecProc = Not IsError(bln提示)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsError(Optional ByVal bln提示 As Boolean = True) As Boolean
    On Error GoTo errHand
    Dim strMsg As String
    IsError = False
    If intReturn = 0 Then Exit Function
    strProcedure = UCase(strProcedure)
    
    Select Case strProcedure
    Case "ADD_CFJLK"
        Select Case intReturn
        Case 1
            strMsg = "必须先申请住院,才能登记处方！"
        Case 2
            strMsg = "该处方已录入！"
        Case 3
            strMsg = "处方时间小于入院时间！"
        Case 4
            strMsg = "费用已审核，不能增加！"
        End Select
    Case "DELETE_CFJLK"
        Select Case intReturn
        Case 1
            strMsg = "必须先录入处方,才能删除！"
        Case 2
            strMsg = "费用已审核，不能删除！"
        Case 3
            strMsg = "处方已审核，不能删除！"
        End Select
    Case "UPDATE_CFJLK"
        Select Case intReturn
        Case 1
            strMsg = "必须先录入处方,才能修改！"
        Case 2
            strMsg = "费用已审核，不能删除！"
        Case 3
            strMsg = "处方已审核，不能删除！"
        Case 4
            strMsg = "处方时间小于入院时间！"
        End Select
    Case "ADD_CFMXK"
        Select Case intReturn
        Case 1
            strMsg = "必须先申请住院,才能录入处方！"
        Case 2
            strMsg = "必须先登记处方,才能登记处方明细！"
        Case 3
            strMsg = "费用已审核，不能增加！"
        Case 4
            strMsg = "处方时间小于入院时间！"
        Case 5
            strMsg = "药品没找到，请更新药品信息库！"
        End Select
    Case "DELETE_CFMXK"
        Select Case intReturn
        Case 1
            strMsg = "必须先录入处方明细,才能删除！"
        Case 2
            strMsg = "费用已审核，不能删除！"
        Case 3
            strMsg = "处方已审核，不能删除！"
        End Select
    Case "UPDATE_CFMXK"
        Select Case intReturn
        Case 1
            strMsg = "必须先录入处方明细,才能修改！"
        Case 2
            strMsg = "费用已审核，不能修改！"
        Case 3
            strMsg = "处方已审核，不能修改！"
        End Select
    Case "GET_SBBXJE"
        Select Case intReturn
        Case 1
            strMsg = "没有申请住院！"
        Case 2
            strMsg = "必须在医保数据库中出院后才能进行结算！"
        End Select
    End Select
    IsError = True
    If bln提示 Then MsgBox strMsg, vbInformation, gstrSysName
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function



