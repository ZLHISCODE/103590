Attribute VB_Name = "mdlSpePrint"
Option Explicit

Private mlng体温不升显示方式 As Long
Private mintBaby As Integer  '是否是婴儿
Private mlngBreatheHeight  '呼吸表格固定高度

Public Function PrintOrPreviewBodyStateNew(objOut As Object, _
                                        ByVal lng病人ID As Long, _
                                        ByVal lng主页ID As Long, _
                                        ByVal lng文件ID As Long, _
                                        ByVal intBaby As Integer, _
                                        ByVal lngSectID As Long, _
                                        ByVal lngBeginY As Long, _
                                        ByVal lngBeginX As Long, _
                                        ByVal objParent As Object, _
                                        Optional ByVal blnKeepOn As Boolean = False, _
                                        Optional ByVal intBeginPage As Integer = -1, _
                                        Optional ByVal intEndPage As Integer = -1, _
                                        Optional ByVal intPageNo As Integer = -1, _
                                        Optional ByVal sngScale As Single = 1, _
                                        Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '功能:打印或预览某七天的专科温度表
    '参数:objOut=输出对象,可以为Printer或一个窗体(窗体中包含控件数组picPage)
    '      lngCaseRecordID=病历记录id
    '      lngBeginY=开始纵坐标
    '      blnKeepOn=是否保持连续
    '      objParent=主调用窗体
    '      intBeginPage=要开始页面序号,当为-1时表示输出所有.
    '      intEndPage=结束页面号如果intEndPage大于实际页数就只打印到实际页数
    '      intPageNO=开始的页码,如果为-1表示不显示页码
    '      sngScale=输出比例
    
    '返回:本次打印操作是否成功
    '******************************************************************************************************************
    Dim strSQL As String, strNewSql As String
    '护理设置参数变量
    Dim intOpDays As Integer  '手术后标注天数
    Dim blnStopFlag As Boolean '再次手术停止前次标注
    Dim intOpFormat As Integer '手术当天缺省格式
    Dim byt未记显示位置 As Byte '未记说明显示位置
    Dim bln婴儿体温单显示出院 As Boolean '婴儿体温单显示出院信息
    Dim bln体温单显示诊断 As Boolean '体温单显示诊断
    Dim intRepairRows As Integer  '表格显示行数
    Dim bln显示皮试 As Boolean '体温单输出显示皮试结果
    Dim bln打印医院名称 As Boolean '体温单是否打印医院名称
    Dim bln入科显示入院 As Boolean
    Dim bln波动 As Boolean
    Dim bln汇总当天 As Boolean '体温单汇总数据时今天汇总今天还是今天汇总昨天的数据
    Dim bln录入小时 As Boolean '体温单全天汇总允许录入和显示汇总小时数
    Dim bln打印脉搏短绌 As Boolean  '体温单打印时是否打印脉搏短绌
    Dim bln不打印心率列 As Boolean '体温单打印时是否打印心率列(仅在心率单独应用是有效，只是不打印刻度列，点正常输出)
    Dim lngCurveRow As Long '体温曲线固定添加行数
    Dim bln出院 As Boolean
    Dim lngSignColor As Long '参数:体温单标记显示颜色
    
    '其他绘图变量
    Dim i As Integer, j As Integer
    Dim lngPicPageIndex As Integer '预览时PIC的索引
    Dim blnPrint As Boolean  '是否打印
    Dim strInfo As String '说明信息
    Dim intAllOpt As Single  '打印的总共步骤
    Dim intCurOpt As Single  '打印进行到第几步
    Dim objDraw As Object '绘图对象
    Dim lngHwnd As Long '句柄
    Dim lngDC As Long  '绘图对象的DC
    Dim lngFont As Long
    Dim lngOldFont As Long
    Dim stdSet As StdFont
    Dim lngLableStep As Long '刻度区域列宽
    Dim lngColStep As Long ' 体温区域列宽
    Dim lngInitRowStep As Long '体温区域列高
    
    Dim lngCountPage As Long '所有页数
    Dim lngPage As Long
    Dim strBeginDate  As String, strBeginDate1 As String '开始时间
    Dim strEndDate As String '终止时间
    Dim strTmpDay As String, strEndDay As String
    Dim dtBegin As Date, dtEnd As Date
    Dim intDrawLineRows As Integer '体温单区域总列数
    Dim intDrawLineCOL As Integer '体温单刻度区域列数
    Dim strTmp As String, strTime As String, strTmp1 As String
    Dim lngValue As Long '住院天数
    Dim T_Rect As RECT
    Dim rsPart As New ADODB.Recordset  '所有体温部位信息
    Dim rsTemp As New ADODB.Recordset  '此记录集请不要顺便使用
    Dim rsTmp As New ADODB.Recordset
    Dim rsItems As New ADODB.Recordset '使用与此病人的所有护理项目信息
    Dim rsDrawItems As New ADODB.Recordset '体温单各个项目信息
    Dim rsPoints As New ADODB.Recordset '所有体温单的集合
    Dim rsNotes As New ADODB.Recordset   '所有说明信息
    Dim rsDownTab As New ADODB.Recordset '表下表格数据信息
    Dim H_16pt As Long, W_16pt As Long
    Dim int心率应用 As Integer
    Dim str心率符号  As String
    Dim arrTmpValue() As Variant, arrTmpNote() As Variant
    Dim arrValues() As String
    Dim strPart As String '部位
    Dim SinX As Single, sinY As Single
    Dim intCOl As Integer
    Dim blnAdd As Boolean, blnAllow As Boolean
    Dim dbl数值 As Double, dblMinValue As Double, dblMaxValue As Double
    Dim lng项目序号 As Long
    Dim str体温说明 As String
    Dim bln呼吸 As Boolean  '呼吸是否为表格
    Dim sngHTab As Single  '表下表格高度
    Dim sngHPrint As Single '可打印区域
    
    Dim strBegin As String, strEnd As String
    Dim str结果 As String
    Dim strItemName As String, strItems As String
    Dim int频次 As Integer
    Dim intCol1 As Integer
    Dim str项目名称 As String
    Dim int项目性质 As Integer, int项目类型 As Integer, int入院首测 As Integer
    Dim int舒张压 As Integer, int收缩压 As Integer, Int列号 As Integer
    Dim blnColor As Boolean

    '病人基本信息
    Dim strPatiInfo As String
    Dim VarPatiInfo As Variant
    Dim lng护理等级 As Long
    
    '--下面三个变量 在记录体温不升时做临时储存对象
    Dim strTmpString0 As String  '记录当前时间
    Dim strTmpString2 As String '记录住院天数
    Dim strTmpString1 As String '记录手术后天数
    Dim strNewTmpString As String
    Dim ArrNewTmpString() As String '记录表格项目的列数和每一列值的信息
    Dim ArrNewString() As String '记录所有表格项目信息
    Dim intDays As String '手术天数
    Dim strOpdays() As String
    Dim strOpValue() As String
    Dim arrOperDay
    Dim strEditors() As Variant    '记录曲线项目信息(项目序号||项目名称||项目单位||项目值域||记录符||记录色||最大值||最小值||临界值）
    Dim ArrComTable() As Variant '记录所有的表下表格项目 (项目序号||部位+项目名称|项目单位||项目值域||记录频次||项目性质||项目表示||入院首测)
    Dim lng次数 As Long  '记录手术次数
    Dim bln术后显示 As Boolean '病人术后不足14天出院标记显示
    Dim str结束时间 As String
    Dim bln入科不转入院 As Boolean
    
    '坐标信息
    Dim lngLeft As Long, lngTop As Long
    Dim lngRight As Long, lngButtom As Long
    Dim X As Long, Y As Long
    Dim lngCurX As Long, lngCurY As Long
    Dim dblSureW As Double, dblSureH As Double
    
    Dim M_DrawClient As DrawClient
    Dim lng刻度宽度 As Long
    
    On Error GoTo ErrPrint
    
    msngTwips = 1
    
    mintBaby = intBaby
    '保存原始值:
    
    M_DrawClient.偏移量X = T_DrawClient.偏移量X
    M_DrawClient.偏移量Y = T_DrawClient.偏移量Y
    M_DrawClient.刻度区域 = T_DrawClient.刻度区域
    M_DrawClient.刻度单位 = T_DrawClient.刻度单位
    M_DrawClient.体温区域 = T_DrawClient.体温区域
    M_DrawClient.行单位 = T_DrawClient.行单位
    M_DrawClient.时间行单位 = T_DrawClient.时间行单位
    M_DrawClient.时间列单位 = T_DrawClient.时间列单位
    M_DrawClient.列单位 = T_DrawClient.列单位
    M_DrawClient.双倍 = T_DrawClient.双倍
    M_DrawClient.总列数 = T_DrawClient.总列数
    M_DrawClient.曲线总区域 = T_DrawClient.曲线总区域
    M_DrawClient.独立曲线总行数 = T_DrawClient.独立曲线总行数
    lng刻度宽度 = T_BodyStyle.lng刻度宽度
    
    mintBmpW = gintBmpW
    mintBmpH = gintBmpH
    '读取体温参数信息
    '------------------------------------------------------------------------------------------------------------------
    intOpDays = Val(zlDatabase.GetPara("手术后标注天数", glngSys, 1255, "10"))
    blnStopFlag = (Val(zlDatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0")) = 1)
    byt未记显示位置 = Abs(Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0")))
    bln婴儿体温单显示出院 = (zlDatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, 1) = 1)
    bln体温单显示诊断 = (zlDatabase.GetPara("体温单显示诊断", glngSys, 1255, 1) = 1)
    intRepairRows = T_BodyStyle.lng表格空行 + GetRows(bln呼吸, T_BodyItem.str表格项目)
    bln显示皮试 = (Val(zlDatabase.GetPara("体温单显示皮试结果", glngSys, 1255, "0")) = 1)
    bln打印医院名称 = (Val(zlDatabase.GetPara("打印医院名称", glngSys, 1255, "1")) = 1)
    bln汇总当天 = (Val(zlDatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, 0)) = 1)
    bln打印脉搏短绌 = (Val(zlDatabase.GetPara("不打印脉搏短绌图形", glngSys, 1255, "0")) = 0)
    bln不打印心率列 = (Val(zlDatabase.GetPara("体温单不打印心率列", glngSys, 1255, "0")) = 1)
    mlng体温不升显示方式 = Val(zlDatabase.GetPara("体温不升显示方式", glngSys, 1255, "0"))
    bln术后显示 = (Val(zlDatabase.GetPara("病人术后不足14天出院标记显示", glngSys, 1255, "0")) = 1)
    bln入科不转入院 = (Val(zlDatabase.GetPara("入科标识不自动转换为入院", glngSys, 1255, "0")) = 0)
    '62989:刘鹏飞,2013-07-24,体温单标记显示颜色
    lngSignColor = Val(zlDatabase.GetPara("体温单标记显示颜色", glngSys, 1255, "255"))
    
    lngCurveRow = T_BodyStyle.lng曲线空行
    
    '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
    bln录入小时 = (Val(zlDatabase.GetPara("全天汇总显示录入时间", glngSys, 1255, 0)) = 1)
    
    '51338,刘鹏飞,2012-07-06
    strTmp = zlDatabase.GetPara("手术当天缺省格式", glngSys, 1255, "2")
    If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
        intOpFormat = Val(strTmp)
    Else
        intOpFormat = 0
    End If
    '病人变动标记显示方法
    '------------------------------------------------------------------------------------------------------------------
    Call InitPara(T_BodyStyle.bln专科)

    blnPrint = TypeName(objOut) = "Printer"
    
    '由于打印机和屏幕的像素不同，此处需要取各自的像素
    If blnPrint = True Then
        T_TwipsPerPixel.X = Printer.TwipsPerPixelX
        T_TwipsPerPixel.Y = Printer.TwipsPerPixelY
        msngTwips = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        Printer.Font.Size = 9
        Printer.FontName = "宋体"
    Else
        T_TwipsPerPixel.X = Screen.TwipsPerPixelX
        T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
        msngTwips = 1
    End If
    
    mlngBreatheHeight = 300 \ T_TwipsPerPixel.Y
    Screen.MousePointer = 11
    intAllOpt = 5
    
    '计算进度处理
    '------------------------------------------------------------------------------------------------------------------
    strInfo = "正在" & IIf(blnPrint, "准备打印体温表", "处理预览") & ",请稍候..."
    Call ShowFlash(strInfo, , objParent)
    
    '打印前的清除
    If blnKeepOn = False Then
        If Not blnPrint Then
            For i = objOut.picPage.UBound To 0 Step -1
                If i = 0 Then
                    objOut.picPage(i).Cls
                Else
                    Unload objOut.picPage(i)
                End If
            Next
            Set objDraw = objOut.picPage(0)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    Else
        If Not blnPrint Then
            i = objOut.picPage.UBound + 1
            Load objOut.picPage(i)
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    End If
    
    bln出院 = False
    '提取婴儿医嘱信息(转科，出院)存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strSQL = getSQLString("提取文件时间范围", blnMoved)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng文件ID, lng病人ID, lng主页ID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngCountPage = DateDiff("d", rsTemp!开始, rsTemp!终止) + 1
        lngCountPage = IIf(lngCountPage / T_BodyStyle.lng天数 = Fix(lngCountPage / T_BodyStyle.lng天数), lngCountPage / T_BodyStyle.lng天数, Fix(lngCountPage / T_BodyStyle.lng天数) + 1)
        strBeginDate = Format(rsTemp!开始, "YYYY-MM-DD HH:MM:SS")
        strBeginDate1 = strBeginDate
        strEndDate = Format(rsTemp!终止, "YYYY-MM-DD HH:MM:SS")
        bln出院 = Not (Val(rsTemp!记录) = 0)
    Else
        CloseRs rsTemp
        GoTo ErrPrint '无数病人变动信息退出
    End If
    
    gbln出院 = bln出院
    If bln出院 = True Then
        '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
        strEndDate = Format(RetrunEndTimeNew(CDate(strBeginDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
    End If
    
    bln入科显示入院 = False
    
    If CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) > CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln入科显示入院 = True
    ElseIf T_BodyFlag.入院 = 0 And CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) = CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln入科显示入院 = True
    End If
            
    intCurOpt = intCurOpt + 1
    
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '第1部份：病人的基本信息
    '读取病人基本信息
    
    '"姓名'年龄'性别'科别'床号'入院日期'住院号:
    strPatiInfo = "''''''"
    VarPatiInfo = Split(strPatiInfo, "'")
    
    strSQL = " Select  NVL(A.姓名,b.姓名) 姓名,A.住院号,A.入院日期 入院时间,NVL(A.性别,b.性别) 性别,NVL(A.年龄,B.年龄) 年龄 From 病人信息 B,病案主页 A Where A.病人ID=B.病人ID And A.病人id=[1] And A.主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人ID, lng主页ID)
    If rsTemp.BOF = False Then
        VarPatiInfo(0) = zlCommFun.Nvl(rsTemp("姓名").Value)
        VarPatiInfo(6) = zlCommFun.Nvl(rsTemp("住院号").Value)
        VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("入院时间").Value), "yyyy-MM-dd")
        VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("性别").Value)
        VarPatiInfo(1) = zlCommFun.Nvl(rsTemp("年龄").Value)
    End If
    
    '入院时间(如果体温单开始时间大于入院时间就以入科时间为准)
    strSQL = "select 开始时间 from 病人变动记录 where 病人id=[1] And 主页id=[2] and 开始原因=2 order by 开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人ID, lng主页ID)
    If rsTemp.BOF = False Then
        If bln入科显示入院 = True Then
            VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("开始时间").Value), "yyyy-MM-dd")
        End If
    End If
    
    If intBaby <> 0 Then
        
        VarPatiInfo(1) = ""
        VarPatiInfo(2) = ""
        
        strSQL = "Select Decode(a.婴儿姓名,Null,NVL(C.姓名,B.姓名) ||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,a.婴儿性别,a.出生时间 " & _
            " From 病人信息 B,病案主页 C,病人新生儿记录 A " & _
            " Where B.病人ID=C.病人ID And C.病人ID=A.病人ID And C.主页ID=A.主页ID And C.病人id=[1] And C.主页id=[2] And a.序号=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人ID, lng主页ID, intBaby)
        If rsTemp.BOF = False Then
            VarPatiInfo(0) = rsTemp("婴儿姓名").Value
            VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("婴儿性别").Value)
            VarPatiInfo(1) = "新生儿"
            If IsNull(rsTemp("出生时间").Value) = False Then VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("出生时间").Value), "yyyy-MM-dd")
        End If
        
    End If
    
    If bln体温单显示诊断 Then ReDim Preserve VarPatiInfo(UBound(VarPatiInfo) + 1)
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '获取病人护理等级
     lng护理等级 = Get护理等级(lng病人ID, lng主页ID)
    
    '提取共用记录集
    Call InitPublicData
    
    '求出心率应用方式
    int心率应用 = 2
    str心率符号 = ""
    strSQL = "Select a.应用方式,b.记录符 From 护理记录项目 a,体温记录项目 b Where a.项目序号=-1 And a.项目序号=b.项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint")
    If rsTemp.BOF = False Then
        int心率应用 = zlCommFun.Nvl(rsTemp("应用方式").Value, 2)
        str心率符号 = zlCommFun.Nvl(rsTemp("记录符").Value, "○")
    Else
        int心率应用 = 0
    End If
    
    Dim int脉搏 As Integer, int心率 As Integer
    
    '-------------------------------------------------------------------------------------------------------------------
    '2提取所有曲线项目(此体温单固定有两行输出所以最高行-2)
    strSQL = getSQLString("提取所有曲线项目", blnMoved)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取所有曲线项目", T_BodyItem.str曲线项目)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        rsTemp.Filter = "记录法=1"
        intDrawLineCOL = rsTemp.RecordCount
        rsTemp.Filter = "项目序号=" & gint心率 & " And 记录法=1"
        If rsTemp.RecordCount > 0 And bln不打印心率列 Then
            rsTemp.Filter = 0
            intDrawLineCOL = intDrawLineCOL - 1
        Else
            rsTemp.Filter = 0
        End If
        If intDrawLineCOL <= 0 Then intDrawLineCOL = 1
    Else
        CloseRs rsTemp
        MsgBox "无任何体温曲线项目！", vbExclamation, gstrSysName
        GoTo ErrExit
    End If
    strEditors = Array()
    int脉搏 = -1: int心率 = -1
    rsTemp.Filter = 0
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            strTmp = Nvl(!项目序号, 0) & "|| " & Nvl(!记录名) & "|| " & Nvl(!单位) & "|| " & Nvl(!项目值域) & "|| " & _
                 Nvl(!记录符) & "|| " & Nvl(!记录色) & "||" & Nvl(!最大值) & "||" & Nvl(!最小值) & "||" & Nvl(!临界值)
                
            ReDim Preserve strEditors(UBound(strEditors) + 1)
            strEditors(UBound(strEditors)) = strTmp
            If zlCommFun.Nvl(!项目序号, 0) = gint脉搏 Then
                int脉搏 = UBound(strEditors)
            End If
        .MoveNext
        Loop
        .MoveFirst
    End With
    If int心率应用 = 2 And int脉搏 <> -1 Then
        ReDim Preserve strEditors(UBound(strEditors) + 1)
        strTmp = "-1||心率||" & Split(strEditors(int脉搏), "||")(2) & "||" & Split(strEditors(int脉搏), "||")(3) & "||" & str心率符号 & "||" & RGB_RED & "||" & _
            Split(strEditors(int脉搏), "||")(6) & "||" & Split(strEditors(int脉搏), "||")(7) & "||" & Split(strEditors(int脉搏), "||")(8)
        strEditors(UBound(strEditors)) = strTmp
    End If
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '3—提取所有特殊项目信息包括活动项目（活动项目可能存在一个项目多个部位也要提取）
    ArrComTable = Array()
    strTmp = ""
    strTime = ""
    
    '提取表格非汇总项目
    strSQL = getSQLString("提取表格非汇总项目", blnMoved)
    If blnMoved Then
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
        strSQL = Replace(strSQL, "病人护理明细", "H病人护理明细")
    End If

    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, "取开始行", lng文件ID, lng病人ID, lng主页ID, intBaby, Int(CDate(Format(strBeginDate, "yyyy-mm-dd hh:mm:ss"))), CDate(Format(strEndDate, "yyyy-mm-dd hh:mm:ss")), lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID, T_BodyItem.str表格项目)
    bln呼吸 = False
    With rsItems
        Do While Not .EOF
            str项目名称 = ""
            If Val(Nvl(!项目性质, 1)) = 2 Then
                str项目名称 = Trim(Nvl(!部位)) & Nvl(!项目名称)
            Else
                str项目名称 = Nvl(!项目名称)
            End If
            
            int频次 = Val(zlCommFun.Nvl(!记录频次))
            
            If zlCommFun.Nvl(!项目表示) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!项目序号))) Then
                If int频次 > 2 Then int频次 = 2
            End If
            
            strTmp = zlCommFun.Nvl(!项目序号) & "||" & Replace(str项目名称, ";", ":") & "||" & zlCommFun.Nvl(!项目单位) & "||" & _
                zlCommFun.Nvl(!项目值域) & "||" & int频次 & "||" & zlCommFun.Nvl(!项目性质, 1) & "||" & _
                zlCommFun.Nvl(!项目表示) & "||" & zlCommFun.Nvl(!项目类型) & "||" & zlCommFun.Nvl(!入院首测, 0)
            If Val(zlCommFun.Nvl(!项目序号)) = gint呼吸 Then
                bln呼吸 = True
            End If
            
            ReDim Preserve ArrComTable(UBound(ArrComTable) + 1)
            ArrComTable(UBound(ArrComTable)) = strTmp
        .MoveNext
        Loop
    End With

    If rsItems.RecordCount > 0 Then rsItems.MoveFirst
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '------------------------------------------------------------------------------------------------------------------
    '4、确定X和Y的坐标位置
    '边界信息(Twip)
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    
    dblSureH = 0
    dblSureW = 0
    If blnPrint = True Then
        '如果是打印预览,应按打印机的可打印的开始处开始预览
        dblSureW = Round(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH), 4)
        dblSureH = Round(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT), 4)
        On Error Resume Next
        dblSureH = (objDraw.Height * dblSureH) / T_TwipsPerPixel.Y
        dblSureW = (objDraw.Width * dblSureW) / T_TwipsPerPixel.X
    End If

    lngRight = gPrinter.lngRight
    lngButtom = gPrinter.lngBottom
     
    lngRight = lngRight * (conRatemmToTwip / T_TwipsPerPixel.X) * sngScale
    If lngRight < dblSureW Then lngRight = dblSureW
    lngButtom = lngButtom * (conRatemmToTwip / T_TwipsPerPixel.Y) * sngScale
    If lngButtom < dblSureH Then lngButtom = dblSureH
    lngLeft = lngBeginX * (conRatemmToTwip / T_TwipsPerPixel.X) * sngScale
    If lngLeft < dblSureW Then lngLeft = dblSureW
    lngTop = (lngBeginY / T_TwipsPerPixel.X) * sngScale
    If lngTop < dblSureH Then lngTop = dblSureH
    
    H_16pt = objDraw.TextHeight("字") / T_TwipsPerPixel.Y
    W_16pt = objDraw.TextWidth("字") / T_TwipsPerPixel.X
    
    X = lngLeft: Y = lngTop
    lngCurX = X: lngCurY = Y
        
    T_DrawClient.刻度区域.Left = lngCurX
    T_DrawClient.刻度区域.Right = lngCurX + T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X * sngScale
    
    lngColStep = (T_BodyStyle.lng曲线列宽 / T_TwipsPerPixel.X) * sngScale
    lngInitRowStep = (T_BodyStyle.lng曲线行高 / T_TwipsPerPixel.Y) * sngScale
    
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Right = T_DrawClient.刻度区域.Right + (T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数 * lngColStep)
    
    Dim sigSign As Single
    sigSign = 1
    If T_DrawClient.体温区域.Right > objDraw.Width / T_TwipsPerPixel.X - lngRight Then
        sigSign = Round((T_DrawClient.体温区域.Right - (objDraw.Width / T_TwipsPerPixel.X - lngRight)) / (T_DrawClient.体温区域.Right - T_DrawClient.刻度区域.Right), 2)
        sigSign = Round((1 - sigSign), 2)
        If sigSign < 0.8 Then sigSign = 0.8
        T_BodyStyle.lng刻度宽度 = Fix(T_BodyStyle.lng刻度宽度 * sigSign)
        lngColStep = Fix(lngColStep * sigSign)
    End If
    If T_BodyStyle.lng曲线列宽 / T_TwipsPerPixel.X > W_16pt Then
        If lngColStep < W_16pt Then lngColStep = W_16pt
    Else
        lngColStep = (T_BodyStyle.lng曲线列宽 / T_TwipsPerPixel.X) * sngScale
    End If
    
    If lngColStep < gintBmpW Then
        mintBmpW = lngColStep
        mintBmpH = lngColStep
    End If
    
    lngLableStep = Fix((T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X / intDrawLineCOL) * sngScale)
    T_DrawClient.刻度单位 = lngLableStep
    T_DrawClient.刻度区域.Right = lngCurX + (T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X * sngScale)
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Right = T_DrawClient.刻度区域.Right + (T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数 * lngColStep) * sngScale
    T_DrawClient.列单位 = lngColStep
    T_DrawClient.行单位 = lngInitRowStep
    T_DrawClient.时间列单位 = T_BodyStyle.lng表格高度 / T_TwipsPerPixel.Y * sngScale
    T_DrawClient.偏移量X = lngLeft

    '------------------------------------------------------------------------------------------------------------------
    '求得首列宽，求左边标尺总共有多少行
    '求得体温表项目的总行数
    intDrawLineRows = Get总行数(dbl数值, lngCurveRow)
    If intDrawLineRows = 0 Then GoTo ErrPrint

    T_DrawClient.总列数 = intDrawLineRows

    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '5、循环：按总页数循环
    intCurOpt = 0
    intAllOpt = 100
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    If blnPrint = False Then
        lngPicPageIndex = objOut.picPage.UBound + 1
    End If
    
    '正式开始第四步，循环每一页
    '------------------------------------------------------------------------------------------------------------------
    For lngPage = 1 To lngCountPage
        strTmpDay = Format(CDate(strBeginDate) + T_BodyStyle.lng天数 * (lngPage - 1), "YYYY-MM-DD")  '求得当前页面的第一天日期与时间
        If CDate(strTmpDay) < CDate(strBeginDate) Then strTmpDay = strBeginDate
        If CDate(strEndDate) < CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")) And Not bln出院 Then strEndDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        strEndDay = Format(CDate(strTmpDay) + T_BodyStyle.lng天数 - 1, "YYYY-MM-DD") & " 23:59:59"
        If CDate(strEndDay) > CDate(strEndDate) Then strEndDay = Format(strEndDate, "YYYY-MM-DD HH:mm:ss")
        intCurOpt = lngPage / lngCountPage
        strInfo = "正在" & IIf(blnPrint, "打印体温表", "预览") & ",请稍候..."
        Call ShowFlash(strInfo, intCurOpt, objParent)
        
        '按页号打印
        If intBeginPage > 0 Then  '只打印指定页码的
            If lngPage >= intBeginPage And lngPage <= intEndPage Then
                If lngPage > intBeginPage Then  '到第二页时开始初始化纸张或页面
                    If Not blnPrint Then
                        Load objOut.picPage(lngPicPageIndex)
                        Set objDraw = objOut.picPage(lngPicPageIndex)
                        objDraw.Cls
                        objDraw.Width = Printer.Width * sngScale
                        objDraw.Height = Printer.Height * sngScale
                        lngPicPageIndex = lngPicPageIndex + 1
                    Else
                        Printer.NewPage
                    End If
                End If
            Else
                GoTo NOPageSub
            End If
        Else  '打印所有时
            If lngPage > 1 Then
                If Not blnPrint Then
                    Load objOut.picPage(lngPicPageIndex)
                    Set objDraw = objOut.picPage(lngPicPageIndex) 'PictureBox
                    objDraw.Cls
                    objDraw.Width = Printer.Width * sngScale
                    objDraw.Height = Printer.Height * sngScale
                    lngPicPageIndex = lngPicPageIndex + 1
                Else
                    Printer.NewPage
                End If
            End If
        End If
        
         '页眉图形输出
        Call frmTendFileRead.PrintRTBData(objDraw, True, lngTop)
        
        '获取对象的DC
        Call ReleaseFontIndirect(objDraw)
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, objDraw.hDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(objDraw.hDC, lngFont)
        lngDC = objDraw.hDC
        '67934:刘鹏飞,2013-12-03,以透明状态进行绘图
        Call SetBkMode(lngDC, TRANSPARENT)
        '创建字体
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '打印质控号
        strTmp = zlDatabase.GetPara("质控号", glngSys, 1255, "")
        Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
        lngCurX = T_DrawClient.体温区域.Right - T_Size.W
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, , , , sngScale)
        Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '是否打印医院名称，有的医院体温单医院名可能存在两个，需要用页眉来实现。此时就不在打印注册文件中的医院信息。
        If bln打印医院名称 = True Then
            '获取医院名称
            Set stdSet = New StdFont
            stdSet.Name = Split(T_BodyStyle.str标题字体, ",")(0)
            stdSet.Size = Split(T_BodyStyle.str标题字体, ",")(1) * sngScale
            If InStr(1, T_BodyStyle.str标题字体, "粗") > 0 Then stdSet.Bold = True
            If InStr(1, T_BodyStyle.str标题字体, "斜") > 0 Then stdSet.Italic = True
            Call SetFontIndirect(stdSet, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            strTmp = IIf(GetUnitName = "-", "", GetUnitName) & IIf(intBaby <> 0, "婴儿", "") & T_BodyStyle.str标题文本
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            lngCurY = T_Size.H \ 2 + lngCurY
            Call GetTextRect(objDraw, 0, lngCurY, strTmp, objDraw.Width / T_TwipsPerPixel.X, True, T_Size.H, sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
            objDraw.Font.Size = 9 * sngScale
            Y = lngCurY + T_Size.H \ 2 + 12 * msngTwips
        Else
            Y = lngCurY + 12 * msngTwips
        End If
        lngCurX = X
        lngCurY = Y
        '读取病人科室、床号等信息
    
        VarPatiInfo(3) = ""
        VarPatiInfo(4) = ""
        strTmp = "": strTime = ""
        
        strSQL = getSQLString("提取科室床号", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病人科室、床号等信息", lng病人ID, lng主页ID, CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")))
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                
                If zlCommFun.Nvl(rsTmp("科室").Value) <> strTmp And zlCommFun.Nvl(rsTmp("科室").Value) <> "" Then
                
                    strTmp = zlCommFun.Nvl(rsTmp("科室").Value)
                    
                    If VarPatiInfo(3) = "" Then
                        VarPatiInfo(3) = strTmp
                    Else
                        VarPatiInfo(3) = VarPatiInfo(3) & "->" & strTmp
                    End If
                    
                End If
    
                If zlCommFun.Nvl(rsTmp("床号").Value) <> strTime And zlCommFun.Nvl(rsTmp("床号").Value) <> "" Then
                
                    strTime = zlCommFun.Nvl(rsTmp("床号").Value)
                    
                    If VarPatiInfo(4) = "" Then
                        VarPatiInfo(4) = strTime
                    Else
                        VarPatiInfo(4) = VarPatiInfo(4) & "->" & strTime
                    End If
                    
                End If
                            
                rsTmp.MoveNext
            Loop
            
            If Left(VarPatiInfo(3), 2) = "->" Then VarPatiInfo(3) = Mid(VarPatiInfo(3), 3)
            If Left(VarPatiInfo(4), 2) = "->" Then VarPatiInfo(4) = Mid(VarPatiInfo(4), 3)
        End If
        
        If bln体温单显示诊断 Then
            '提取诊断的最小时间
            strTmp = GetDiagnoseMinTime(lng病人ID, lng主页ID, CDate(strTmpDay), blnMoved)
            '提取病人诊断信息
            strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As 最后诊断 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "最后诊断", "最后诊断", lng病人ID, lng主页ID, CDate(Format(strTmp, "yyyy-mm-dd hh:mm:ss")))
            If rsTmp.BOF = False Then
                If intBaby = 0 Then
                    VarPatiInfo(UBound(VarPatiInfo)) = zlCommFun.Nvl(rsTmp("最后诊断").Value)
                Else
                    VarPatiInfo(UBound(VarPatiInfo)) = ""
                End If
            Else
                VarPatiInfo(UBound(VarPatiInfo)) = ""
            End If
        End If
        strPatiInfo = Join(VarPatiInfo, "'")
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        stdSet.Italic = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '输出病人信息
        Call DrawPatiInfo(lngDC, objDraw, strPatiInfo, lngCurX, lngCurY, T_DrawClient.体温区域.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '---开始画体温单上表格(住院日期,住院天数,手术,时间)
        Y = lngCurY: lngCurX = X: lngCurY = Y
        '1.提取住院开始天数
        lngValue = 0: strTmp = "": strTime = ""
        strSQL = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As 开始天数 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取住院天数", lng文件ID, lng病人ID, lng主页ID, Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss"))))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("开始天数").Value
        End If
        For i = 0 To T_BodyStyle.lng天数 - 1
            strTmp = Format(CDate(strTmpDay) + i, "YYYY-MM-DD")
            If Right(strTmp, 5) = "01-01" Then
                '一年的第一天
                strTime = strTmp
            ElseIf strTmp = Format(strBeginDate, "yyyy-MM-dd") Then
                '入院第一天，写上年份
                strTime = strTmp
            ElseIf i = 0 Then '每页的第一列
                '70299:刘鹏飞,2014-4-4,每页首列日期显示为年月日(1-年-月-日,0:默认格式:按规则显示)
                If Val(zlDatabase.GetPara("首列日期格式", glngSys, 1255, "0")) = 1 Then
                    strTime = strTmp
                Else
                    strTime = Right(strTmp, 5)
                End If
            ElseIf Right(strTmp, 2) = "01" Then
                strTime = Right(strTmp, 5)
            Else
                strTime = Right(strTmp, 2)
            End If

            strTmpString0 = strTmpString0 & "'" & strTime
            strTmpString2 = strTmpString2 & "'" & lngValue + i
        Next i
        strTmpString0 = Mid(strTmpString0, 2)
        strTmpString2 = Mid(strTmpString2, 2)
        '2.提取手术时间和次数
        strTime = ""
        '显示但前段的手术标记
        strSQL = getSQLString("提取当前手术信息", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取手术标记", lng文件ID, intBaby, lng病人ID, lng主页ID, Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")) - 14), CDate(strEndDay))
        
        ReDim strOpdays(1 To T_BodyStyle.lng天数) As String
        ReDim strOpValue(1 To T_BodyStyle.lng天数) As String
        
        str结束时间 = strEndDay
        Do While Not rsTmp.EOF
            strTime = Format(rsTmp("时间"), "YYYY-MM-DD")
            
            '问题号:56005,李涛,2013-04-27
            If Not rsTmp.EOF Then
                If bln术后显示 And DateDiff("d", CDate(Format(strTime, "YYYY-MM-DD")), str结束时间) < 14 Then
                    strEndDay = Format(DateAdd("D", T_BodyStyle.lng天数 - 1, CDate(strTmpDay)), "YYYY-MM-DD") & " 23:59:59"
                End If
            End If
            
            For i = 1 To T_BodyStyle.lng天数
                If DateDiff("d", strTmpDay, str结束时间) + 1 >= i Then
                    intDays = DateDiff("d", strTime, strTmpDay) + (i - 1)

                    Select Case intDays
                        Case 0 '当前区域内的手术开始时间
                             'Modify 2012-03-05 修改一天可以有多次手术
                            If Trim(strOpdays(i)) <> "" Then
                                strOpdays(i) = strTime & "/" & strOpdays(i)
                            Else
                                strOpdays(i) = strTime
                            End If
                        Case Else
                            If intDays >= 1 And intDays <= intOpDays Then '手术开始天数
                                If blnStopFlag Then '手术标注后天数在次手术时停止前一次标注
                                    strOpValue(i) = intDays
                                Else
                                    If Trim(strOpValue(i)) <> "" Then
                                        If intOpFormat = 3 Then
                                            strOpValue(i) = strOpValue(i) & "/" & intDays
                                        Else
                                            strOpValue(i) = intDays & "/" & strOpValue(i)
                                        End If
                                    Else
                                        strOpValue(i) = intDays
                                    End If
                                End If
                            End If
                    End Select
                End If
            Next i
            rsTmp.MoveNext
        Loop
        
        '提取当前开始日期-14天前的手术记录信息
        strSQL = getSQLString("提取14天之前的手术信息", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取手术标记", lng文件ID, intBaby, lng病人ID, lng主页ID, Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss"))))
        
        lng次数 = 0
        If rsTmp.BOF = False Then lng次数 = Val(rsTmp("次数"))
        
        For i = 1 To T_BodyStyle.lng天数
            If DateDiff("d", Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss"))), Int(CDate(Format(str结束时间, "yyyy-mm-dd hh:mm:ss")))) + 1 >= i Then
                If Trim(strOpdays(i)) <> "" Then
                    arrOperDay = Split(strOpdays(i), "/")
                Else
                    arrOperDay = Split("1", "/")
                End If
                lngValue = lng次数
                If Trim(strOpdays(i)) <> "" And lngValue + UBound(arrOperDay) < 12 Then
                    strTmp = "": strTmp1 = ""
                    For j = UBound(arrOperDay) + 1 To 1 Step -1
                        lng次数 = lngValue + j
                        '问题号:57771,李涛,2013-05-02
                        If intOpFormat = 3 Then
                            strTmp1 = Switch(lng次数 = 1, "术日", lng次数 = 2, "术2", lng次数 = 3, "术3", lng次数 = 4, "术4", lng次数 = 5, "术5", lng次数 = 6, _
                            "术6", lng次数 = 7, "术7", lng次数 = 8, "术8", lng次数 = 9, "术9", lng次数 = 10, "术10", lng次数 = 11, "术11", lng次数 = 12, "术12")
                        Else
                            strTmp1 = Switch(lng次数 = 1, "Ⅰ", lng次数 = 2, "Ⅱ", lng次数 = 3, "Ⅲ", lng次数 = 4, "Ⅳ", lng次数 = 5, "Ⅴ", lng次数 = 6, _
                            "Ⅵ", lng次数 = 7, "Ⅶ", lng次数 = 8, "Ⅷ", lng次数 = 9, "Ⅸ", lng次数 = 10, "Ⅹ", lng次数 = 11, "Ⅺ", lng次数 = 12, "Ⅻ")
                        End If
    
                        If strTmp = "" Then
                            strTmp = strTmp1
                        Else
                            strTmp = strTmp & "/" & strTmp1
                        End If
                        If blnStopFlag Then Exit For
                    Next j
                    lng次数 = lngValue + UBound(arrOperDay) + 1
                    If blnStopFlag Then '手术标注后天数在次手术时停止前一次标注
                        Select Case intOpFormat
                            Case 1 '显示0
                                strOpValue(i) = 0
                            Case 2 '显示手术次数
                                If strTmp = "Ⅰ" Then
                                    strOpValue(i) = 0
                                Else
                                    strOpValue(i) = strTmp & "-0"
                                End If
                            Case 3
                                If strTmp = "术1" Then
                                    strOpValue(i) = "术日"
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case Else '不显示
                                strOpValue(i) = ""
                        End Select
                    Else
                        Select Case intOpFormat
                            Case 1 '显示0
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = 0 & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = 0
                                End If
                            Case 2 '显示手术次数
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strTmp & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case 3
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strOpValue(i) & "/" & strTmp
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case Else '不显示
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strOpValue(i)
                                Else
                                    strOpValue(i) = ""
                                End If
                        End Select
                    End If
                End If
            End If
        Next i
        
        strTmpString1 = Join(strOpValue, "'")
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '3开始输出住院日期，天数，手术信息
        Call DrawUpTableNew(lngDC, objDraw, strTmpString0 & "||" & strTmpString2 & "||" & strTmpString1, lngCurX, lngCurY, T_DrawClient.体温区域.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '----------------------------------------------------------------------------------------------
         '此处计算可打印区域 从而计算体温单打印的行高
        T_DrawClient.时间列单位 = T_BodyStyle.lng下表格高度 / T_TwipsPerPixel.Y * sngScale
        If intRepairRows = 0 Then
            sngHTab = intRepairRows
        Else
            '呼吸固定为300
            sngHTab = intRepairRows * T_DrawClient.时间列单位 + IIf(bln呼吸 = True, mlngBreatheHeight - T_DrawClient.时间列单位, 0)
        End If
        
        sngHTab = sngHTab + msngTwips * 12
        sngHPrint = Format(objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHTab, "#0.00;-#0.00;0.00")
        T_DrawClient.行单位 = (sngHPrint - 2 * T_DrawClient.行单位) / (T_DrawClient.总列数 + T_DrawClient.独立曲线总行数)
        T_DrawClient.行单位 = Round(T_DrawClient.行单位 - 0.05, 1) * sngScale
        If T_DrawClient.行单位 > T_BodyStyle.lng曲线行高 / T_TwipsPerPixel.Y * sngScale Then T_DrawClient.行单位 = T_BodyStyle.lng曲线行高 / T_TwipsPerPixel.Y * sngScale
        If T_DrawClient.行单位 < T_BodyStyle.lng曲线行高 / T_TwipsPerPixel.Y * sngScale Then T_DrawClient.行单位 = T_BodyStyle.lng曲线行高 / T_TwipsPerPixel.Y * sngScale
        
        '计算行高后在计算体温单可打印的表格行数
        If intRepairRows > 0 Then
            sngHPrint = (T_DrawClient.总列数 + T_DrawClient.独立曲线总行数) * T_DrawClient.行单位 + 2 * T_DrawClient.行单位
            sngHTab = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHPrint - (msngTwips * 12)
            sngHTab = sngHTab - IIf(bln呼吸 = True, mlngBreatheHeight - T_DrawClient.时间列单位, 0)
            If Fix(sngHTab / T_DrawClient.时间列单位 + 0.3) < intRepairRows Then intRepairRows = Fix(sngHTab / T_DrawClient.时间列单位 + 0.3)
        End If
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '4开始画刻度区域和体温区域并输出刻度值信息
        T_DrawClient.偏移量Y = lngCurY
        mbln呼吸曲线 = False
        
        rsTemp.Filter = 0
        rsTemp.Sort = "排列序号"
        rsTemp.MoveFirst
        str体温说明 = DrawCanvasNew(lngDC, objDraw, rsTemp, rsDrawItems, bln不打印心率列, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        
        '5.读取病人体温数据和入出转等标记信息
        '初始化 体温点记录集和入出转等标记信息
        
        '所有点的表现集合
        '   重叠是否重叠序号.
        '   重叠项目记录重叠项目
        '   断开的条件:超过一天无数据,存在未记说明
        '   备注:物理降温时记录原值
        '   符号:用来标注体温不升，或者值小于等于项目最小值大于等于项目最大值是的特殊符号.此外默认为空

        gstrFields = "序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|部位," & adLongVarChar & ",200|" & _
             "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|" & _
             "复查," & adDouble & ",1|断开," & adDouble & ",1|重叠项目," & adLongVarChar & ",50|" & _
             "重叠," & adDouble & ",5|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|备注," & adLongVarChar & ",50|" & _
             "符号," & adLongVarChar & ",10|显示," & adDouble & ",1"
        Call Record_Init(rsPoints, gstrFields)
    
        '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,13-出生,99-未记说明)
        '禁用表示信息是否输出
        gstrFields = "时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|类型," & adDouble & ",2|" & _
            "内容," & adLongVarChar & ",200|颜色," & adLongVarChar & ",20|X坐标," & adDouble & ",20|" & _
            "Y坐标," & adDouble & ",20|高度," & adDouble & ",20|打印X坐标," & adDouble & ",20|" & _
            "禁用," & adInteger & ",1|显示," & adDouble & ",1"
        Call Record_Init(rsNotes, gstrFields)
        
        Dim rs脉搏 As New ADODB.Recordset
        Dim strFileds As String, strValues As String
        
        '记录脉搏信息
        strFileds = "项目序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|X坐标," & adDouble & ",5|时间," & adLongVarChar & ",20"
        Call Record_Init(rs脉搏, strFileds)
        
        Dim int标记 As Integer
        
        '----提取所有部位信息
        strSQL = "select 项目序号,部位,缺省项 from 体温部位"
        Call zlDatabase.OpenRecordset(rsPart, strSQL, "体温部位")
        '----读取病人体温数据和未记说明
        strSQL = getSQLString("读取病人体温数据和未记说明", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取曲线项目数据", lng文件ID, lng病人ID, lng主页ID, CDate(strTmpDay), CDate(strEndDay), T_BodyItem.str曲线项目)
         
        strTmpString0 = ""
        strTmpString1 = ""
        strTmpString2 = ""
        With rsTmp
            Do While Not .EOF
                strTmp = ""
                blnAllow = False
                strPart = zlCommFun.Nvl(!体温部位)
                lng项目序号 = Val(zlCommFun.Nvl(!项目序号))
                Select Case lng项目序号
                    Case gint心率
                        int标记 = 1
                    Case Else
                        int标记 = Val(zlCommFun.Nvl(!记录标记))
                End Select
                If strPart = "" Then
                    rsPart.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
                    If rsPart.BOF = False Then
                        strPart = zlCommFun.Nvl(rsPart!部位)
                    Else
                        Select Case lng项目序号
                            Case gint体温
                                strPart = "腋温"
                            Case gint呼吸
                                strPart = "自主呼吸"
                            Case Else
                                strPart = ""
                        End Select
                    End If
                End If
                
                SinX = GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinateNew(SinX, Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinateNew(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                
                '记录所有脉搏信息
                If lng项目序号 = gint脉搏 Then
                    strFileds = "项目序号|数值|X坐标|时间"
                    strValues = lng项目序号 & "|" & zlCommFun.Nvl(!数值) & "|" & SinX & "|" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
                    Call Record_Add(rs脉搏, strFileds, strValues)
                End If
                
                If (Not IsNull(!未记说明)) And zlCommFun.Nvl(!数值) <> "不升" Then
                    rsNotes.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号)) & " AND X坐标=" & SinX
                    blnAdd = (rsNotes.RecordCount = 0)
                    '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,99-未记说明)
                    gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用|显示"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
                    gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & !项目序号 & "|99|" & _
                        !未记说明 & "|" & RGB_BLUE & "|" & SinX & "|0|0|0|0|" & zlCommFun.Nvl(!显示)
                   
                    If blnAdd Then
                        '提取接近中间时间点的值做为本列值
                         Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsNotes!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsNotes!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                             blnAllow = GetCanvasCenterNew(CDate(Format(rsNotes!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
    
                        If blnAllow = True Then
                            If Val(rsNotes!显示) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsNotes, gstrFields, gstrValues, "时间|" & Format(rsNotes!时间, "yyyy-MM-dd HH:mm:ss"))
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                gstrFields = "显示"
                                gstrValues = "2"
                                Call Record_Update(rsNotes, gstrFields, gstrValues, "时间|" & Format(rsNotes!时间, "yyyy-MM-dd HH:mm:ss"))
                            End If
                        End If
                    End If
                Else
                    blnAdd = False
                    
                    rsPoints.Filter = "项目序号=" & lng项目序号 & " AND X坐标=" & SinX & " And 标记=" & int标记
                    
                    blnAdd = (rsPoints.RecordCount = 0)
                    
                    dbl数值 = Val(zlCommFun.Nvl(!数值))
                    
                    dblMinValue = GetMaxMinValue(0, lng项目序号, strEditors)
                    dblMaxValue = GetMaxMinValue(1, lng项目序号, strEditors)

                    '不指定符号，项目数据操作最大值和最小值以项目本身符号显示
                    If dbl数值 <= dblMinValue Then
                        dbl数值 = dblMinValue
                        'strTmp = "·"
                    End If
                    
                    
                    If dbl数值 >= dblMaxValue Then
                        dbl数值 = dblMaxValue
                        'strTmp = "·"
                    End If
                    
                     '体温不升是在显示在35刻度
                    If Trim(Nvl(!数值)) = "不升" And lng项目序号 = gint体温 Then dbl数值 = 35
                    
                    sinY = Val(GetYCoordinate(objDraw, rsDrawItems, !项目序号, dbl数值, lngDC, True))
                    
                    gstrFields = "序号|数值|部位|标记|时间|项目序号|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号|显示"
                    gstrValues = Val(zlCommFun.Nvl(!序号)) & "|" & !数值 & "|" & strPart & "|" & int标记 & "|" & _
                                 Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng项目序号 & "|" & Val(zlCommFun.Nvl(!复试合格, 0)) & "|" & IIf(zlCommFun.Nvl(!数值, 0) = "不升", 1, 0) & "|空|0|" & _
                                 SinX & "|" & sinY & "||" & strTmp & "|" & zlCommFun.Nvl(!显示, 0)
                    If blnAdd Then '添加
                        Call Record_Add(rsPoints, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsPoints!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                            blnAllow = GetCanvasCenterNew(CDate(Format(rsPoints!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                       '提取接近中间时间点的值做为本列值
                        If blnAllow = True Then
                            If Val(rsPoints!显示) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & rsPoints!序号)
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                gstrFields = "显示"
                                gstrValues = "2"
                                Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & rsPoints!序号)
                            End If
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
                
        '上面已经得到了所有项目的数据信息，下来处理物理降温和脉搏和心率数据
        rsPoints.Filter = ""
        arrTmpValue = Array()
        If int心率应用 = 2 Then
            rsPoints.Filter = "项目序号=" & gint心率
            With rsPoints
                Do While Not .EOF
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !X坐标 & ";" & Format(!时间, "yyyy-MM-DD HH:mm:ss")
                .MoveNext
                Loop
            End With
        End If
        
        '心率设为脉搏共用时，检查脉搏是否设置为可用
        If int脉搏 <> -1 Then
            For i = 0 To UBound(arrTmpValue)
                '检查心率是否与脉搏相对应
                rs脉搏.Filter = "项目序号=" & gint脉搏 & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                
                rsPoints.Filter = "项目序号=" & gint脉搏 & " and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                If rsPoints.RecordCount = 0 Then
                    If rs脉搏.RecordCount = 0 Then
                        rsPoints.Filter = ""
                        gstrFields = "项目序号": gstrValues = gint脉搏
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & Val(Split(CStr(arrTmpValue(i)), ";")(0)))
                    Else
                        rsPoints.Filter = "项目序号=" & gint心率 & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                        rsPoints.Delete
                    End If
                End If
            Next i
        End If
        
        If int心率应用 = 2 Then
            Set rs脉搏 = New ADODB.Recordset
            strFileds = "序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|部位," & adLongVarChar & ",200|" & _
                        "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|" & _
                        "复查," & adDouble & ",1|断开," & adDouble & ",1|重叠项目," & adLongVarChar & ",50|" & _
                        "重叠," & adDouble & ",5|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|备注," & adLongVarChar & ",50|" & _
                        "符号," & adLongVarChar & ",10|显示," & adDouble & ",1"
            Call Record_Init(rs脉搏, strFileds)
            
            rsPoints.Filter = "项目序号=" & gint脉搏
            With rsPoints
                Do While Not .EOF
                    rs脉搏.AddNew
                    For i = 0 To .Fields.Count - 1
                        rs脉搏.Fields(.Fields(i).Name).Value = .Fields(i).Value
                    Next i
                    rs脉搏.Update
                .MoveNext
                Loop
            End With
            
            rsPoints.Filter = "项目序号=" & gint脉搏
            Do While Not rsPoints.EOF
                rsPoints.Delete
                rsPoints.MoveNext
            Loop
            
            rs脉搏.Filter = ""
            rs脉搏.Sort = "时间"
            With rs脉搏
                Do While Not .EOF
                    blnAdd = False
                    blnAllow = False
                    
                    SinX = Val(zlCommFun.Nvl(!X坐标))
                    sinY = Val(zlCommFun.Nvl(!Y坐标))
                    rsPoints.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号, 0)) & " AND X坐标=" & SinX
                    blnAdd = IIf(rsPoints.RecordCount = 0, True, False)
                    
                    strFileds = "序号|数值|部位|标记|时间|项目序号|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号|显示"
                    strValues = Val(zlCommFun.Nvl(!序号)) & "|" & !数值 & "|" & zlCommFun.Nvl(!部位) & "|" & Val(zlCommFun.Nvl(!标记, 0)) & "|" & _
                                 Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!项目序号)) & "|0|" & Val(zlCommFun.Nvl(!断开)) & "|空|0|" & _
                                 SinX & "|" & sinY & "||" & zlCommFun.Nvl(!符号) & "|" & Val(zlCommFun.Nvl(!显示, 0))
                    
                    If blnAdd Then '添加
                        Call Record_Add(rsPoints, strFileds, strValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsPoints!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                            blnAllow = GetCanvasCenterNew(CDate(Format(rsPoints!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '提取接近中间时间点的值做为本列值
                        If blnAllow = True Then
                            If Val(rsPoints!显示) = 2 Then
                                arrValues = Split(strValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                strValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, strFileds, strValues, "序号|" & rsPoints!序号)
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                strFileds = "显示"
                                strValues = "2"
                                Call Record_Update(rsPoints, strFileds, strValues, "序号|" & rsPoints!序号)
                            End If
                        End If
                    End If
                .MoveNext
                Loop
            End With
        End If
        
        '处理物理降温数据
        arrTmpValue = Array()
        rsPoints.Filter = "项目序号=1 and 标记=0"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "项目序号=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "项目序号=1 and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount <> 0 Then
                gstrFields = "备注": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & zlCommFun.Nvl(rsPoints!序号))
            End If
        Next i
        
        arrTmpValue = Array()
        rsPoints.Filter = "项目序号=1 and 标记=1"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "项目序号=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "项目序号=1 and 标记=0 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount = 0 Then
                rsPoints.Filter = "项目序号=1 and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                rsPoints.Delete
            End If
        Next i
    
        '删除显示为2的数据
        rsPoints.Filter = ""
        rsPoints.Filter = "显示=2"
        Do While Not rsPoints.EOF
            rsPoints.Delete
        rsPoints.MoveNext
        Loop
        
        rsNotes.Filter = ""
        rsNotes.Filter = "显示=2"
        Do While Not rsNotes.EOF
            rsNotes.Delete
        rsNotes.MoveNext
        Loop
        
        '处理未记说明和曲线数据该显示那一条
        rsNotes.Filter = ""
        rsPoints.Filter = ""
        
        arrTmpValue = Array()
        arrTmpNote = Array()
        rsNotes.Sort = "项目序号,X坐标"
        With rsNotes
            Do While Not .EOF
                SinX = Val(!X坐标)
                blnAllow = False
                rsPoints.Filter = "项目序号=" & Val(!项目序号) & " And X坐标=" & SinX
                If rsPoints.RecordCount > 0 Then
                    If (zlCommFun.Nvl(rsPoints!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(rsPoints!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                        blnAllow = GetCanvasCenterNew(CDate(Format(rsPoints!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                    ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                        blnAllow = True
                    End If
                    If blnAllow = True Then
                        ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                        arrTmpValue(UBound(arrTmpValue)) = !项目序号 & ";" & SinX
                    Else
                        ReDim Preserve arrTmpNote(UBound(arrTmpNote) + 1)
                        arrTmpNote(UBound(arrTmpNote)) = !项目序号 & ";" & SinX
                    End If
                End If
            .MoveNext
            Loop
        End With
        
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "项目序号=" & Val(Split(CStr(arrTmpValue(i)), ";")(0)) & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(1))
            Do While Not rsPoints.EOF
                rsPoints.Delete
            rsPoints.MoveNext
            Loop
        Next i
        
        For i = 0 To UBound(arrTmpNote)
            rsNotes.Filter = "项目序号=" & Val(Split(CStr(arrTmpNote(i)), ";")(0)) & " And X坐标=" & Val(Split(CStr(arrTmpNote(i)), ";")(1))
            Do While Not rsNotes.EOF
                rsNotes.Delete
            rsNotes.MoveNext
            Loop
        Next i
    
'        '处理体温不升 体温为不升需要在35度下纵向输出体温不升二字
        rsPoints.Filter = "项目序号=" & gint体温 & " and 数值='不升' and 标记<>1"
        rsPoints.Sort = "时间"
        With rsPoints
            Do While Not .EOF
                strTmpString0 = strTmpString0 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!项目序号)) & "|99|" & _
                      "不升|" & RGB_BLUE & "|" & !X坐标 & "|0|0|0|0"
                strTmpString2 = strTmpString2 & ";" & !X坐标
            .MoveNext
            Loop
        End With
        
        '--------更新断开标记
        '两点之间有未记说明断开，时间操作一天断开,体温不升断开
        rsPoints.Filter = ""
        
        gstrFields = "断开"
        gstrValues = "1"
        rsNotes.Filter = ""
        
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst
        With rsNotes
            Do While Not .EOF
                If int心率应用 = 2 And !项目序号 = -1 Then
                    rsPoints.Filter = "项目序号=" & gint脉搏 & " And X坐标<=" & !X坐标
                Else
                    If !项目序号 = 1 Then
                        rsPoints.Filter = "项目序号=" & !项目序号 & " And  标记<>1 And X坐标<" & !X坐标
                    Else
                        rsPoints.Filter = "项目序号=" & !项目序号 & " And X坐标<" & !X坐标
                    End If
                End If
                rsPoints.Sort = "时间"
                If rsPoints.RecordCount <> 0 Then
                    rsPoints.MoveLast
                    Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & rsPoints!序号)
                End If
      
            .MoveNext
            Loop
        End With
        '时间超过一天
        strTime = ""
        strTmp = ""
        rsPoints.Filter = ""
        
        rsPoints.Sort = "项目序号,时间,标记"
        With rsPoints
            Do While Not .EOF
                If Not IsNull(!序号) Then
                    If Not (Val(!项目序号) = 1 And Val(!标记) = 1) Then
                        If lng项目序号 <> 0 Then
                            If lng项目序号 <> !项目序号 Then strTime = ""
                        End If
                        lng项目序号 = !项目序号
                        If strTime <> "" Then
                            If DateDiff("D", CDate(strTime), CDate(Format(!时间, "YYYY-MM-DD"))) > 1 Then
                                strTmp = strTmp & "," & lngValue
                            End If
                        End If
                        strTime = Format(rsPoints!时间, "YYYY-MM-DD")
                        lngValue = Val(rsPoints!序号)
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        strTmp = Mid(strTmp, 2)
        For i = 0 To UBound(Split(strTmp, ","))
            Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & Split(strTmp, ",")(i))
        Next i
        
        '处理体温不升的.把前一个点的断开标志设置为1
        rsPoints.Filter = ""
        rsPoints.Filter = "项目序号=" & gint体温 & " and 标记<>1"
        rsPoints.Sort = "时间,标记"
        With rsPoints
            Do While Not .EOF
                If !数值 = "不升" And .AbsolutePosition <> 1 Then
                    .MovePrevious '更新上一行断开标记
                    If Val(!断开) <> 1 Then
                        lngValue = !序号
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "序号|" & lngValue)
                    End If
                    .MoveNext
                End If
            .MoveNext
            Loop
        End With
    
        '重新整理未及说明，同一X坐标有相同的说明值输出一次
        rsNotes.Filter = ""
        rsNotes.Sort = "X坐标"
        With rsNotes
            Do While Not .EOF
                If lngValue = !X坐标 Then
                    If InStr(1, "," & strTmp & ",", "," & zlCommFun.Nvl(!内容) & ",") <> 0 Then
                       rsNotes.Delete
                    Else
                        strTmp = strTmp & "," & zlCommFun.Nvl(!内容)
                    End If
                Else
                    lngValue = !X坐标
                    strTmp = zlCommFun.Nvl(!内容)
                End If
            .MoveNext
            Loop
        End With
        
        '--提取入出院,手术等标记说明
        Dim bytShow As Byte
        Dim str内容 As String
        Dim lng行号 As Long, lngColor As Long
        
        '读取手术、上下标信息
        '-----------------------------------------------------------------------
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
        strSQL = getSQLString("读取手术、上下标信息", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取手术、上下标等信息", lng文件ID, lng病人ID, lng主页ID, Int(CDate(strTmpDay)), CDate(strEndDay), intBaby, lng护理等级)
        With rsTmp
            Do While Not .EOF
                bytShow = 1
                str内容 = Trim(zlCommFun.Nvl(!记录内容))
               
                lng行号 = IIf(!记录类型 = 2, 10, IIf(!记录类型 = 6, 11, 4))
                
                '对于手术显示需要特殊处理
                If !记录类型 = 4 Then
                    str内容 = Trim(zlCommFun.Nvl(!项目名称))
                    
                    If str内容 = "分娩" Then
                        bytShow = T_BodyFlag.分娩
                    ElseIf str内容 = "回室" Then
                        bytShow = T_BodyFlag.回室
                    Else
                        bytShow = T_BodyFlag.手术
                    End If
                    
                    If bytShow = 2 Then
                        str内容 = str内容 & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                    Else
                        str内容 = !项目名称
                    End If
                    lngColor = lngSignColor
                Else
                    lngColor = IIf(Not IsNumeric(Nvl(!未记说明)), RGB_BLUE, Val(Nvl(!未记说明)))
                End If
                
                If bytShow > 0 Then
                    SinX = Val(GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                    
                    rsNotes.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=" & !记录类型 & " And 时间='" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "'"
                    If rsNotes.BOF Then
                        gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|" & !记录类型 & "|" & _
                            str内容 & "|" & lngColor & "|" & SinX & "|0|0|0|0"
                        Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        rsNotes!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                        rsNotes!内容 = str内容
                        rsNotes.Update
                    End If
                End If
                rsNotes.Filter = ""
                .MoveNext
            Loop
        End With
        
        '读取入出转等信息
        '-----------------------------------------------------------------------
        '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,99-未记说明)
        '1-入院；2-入科；3-转科；4-换床
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
        Set rsTmp = GetDataFromHis(lng病人ID, lng主页ID, intBaby, CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), 2)
        With rsTmp
            Do While Not .EOF
                If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                    bytShow = 0
                    lng行号 = Val(!行号)
                    str内容 = zlCommFun.Nvl(!内容)
                    Select Case Val(!行号)
                    Case 5
                        bytShow = T_BodyFlag.入院
                    Case 6, 3 '6转入，3转出
                        bytShow = T_BodyFlag.转出
                    Case 7
                        bytShow = T_BodyFlag.换床
                    Case 8
                        bytShow = T_BodyFlag.出院
                        If intBaby > 0 Then
                            bytShow = IIf(bln婴儿体温单显示出院, bytShow, 0)
                        End If
                    Case 9
                        bytShow = T_BodyFlag.入科
                    End Select
                    
                    If bytShow > 0 Then
                        If lng行号 = 9 And bln入科显示入院 = True And bln入科不转入院 = True Then str内容 = "入院"
                        '目前3，4 针对于转科 3-显示说明和科室 4 显示说明，科室，时间
                        If bytShow = 2 Then
                            str内容 = str内容 & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                        ElseIf bytShow = 3 Then
                            str内容 = str内容 & gstrCaveSplit & zlCommFun.Nvl(!科室)
                        ElseIf bytShow = 4 Then
                            str内容 = str内容 & gstrCaveSplit & zlCommFun.Nvl(!科室) & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                        ElseIf bytShow = 1 Then
                            str内容 = str内容
                        End If
                        
                        SinX = Val(GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                        rsNotes.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=3 And 时间='" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "'"
                        
                        If rsNotes.BOF Then
                            gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|3|" & _
                                str内容 & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
                            Call Record_Add(rsNotes, gstrFields, gstrValues)
                        Else
                            rsNotes!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                            rsNotes!内容 = str内容
                            rsNotes.Update
                        End If
                    End If
                    rsNotes.Filter = ""
                End If
                .MoveNext
            Loop
        End With
        
        '提取婴儿出生信息
        If intBaby > 0 Then
            gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
            Set rsTmp = GetDataFromHis(lng病人ID, lng主页ID, intBaby, CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), 3)
            With rsTmp
                Do While Not .EOF
                    bytShow = 0
                    If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                        lng行号 = 12
                        bytShow = T_BodyFlag.出生
                        If bytShow > 0 Then
                            Select Case bytShow
                                Case 1
                                    str内容 = zlCommFun.Nvl(!内容)
                                Case 2
                                    str内容 = zlCommFun.Nvl(!内容) & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                            End Select
                            
                            SinX = Val(GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                            rsNotes.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=13 And 时间='" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "'"
                            
                            If rsNotes.BOF Then
                                gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|13|" & _
                                    str内容 & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
                                Call Record_Add(rsNotes, gstrFields, gstrValues)
                            Else
                                rsNotes!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                                rsNotes!内容 = str内容
                                rsNotes.Update
                            End If
                        End If
                    End If
                    rsNotes.Filter = ""
                .MoveNext
                Loop
            End With
        End If
        '51512,刘鹏飞,2012-07-11,未记说明显示位置 0-显示在上面,1-显示在下面,2-不显示(新增)
        '大医二院要求未记说明不显示，但标注了未记的两边的体温曲线不连接
        strTmp = ""
        Dim arrString() As String
        '处理体温不升 体温不升始终显示在 35 度下面，只有未记说明显示在下面的情况，才将不升放入未记说明中，其它情况都放在下标中
        If Left(strTmpString0, 1) = ";" Then
            gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"
            If mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 2 Then
                arrString = Split(strTmpString0, "|")
                arrString(3) = "↓ "
                strTmpString0 = Join(arrString, "|")
            End If
            strTmpString0 = Mid(strTmpString0, 2)
            strTmpString2 = Mid(strTmpString2, 2)
            For i = 0 To UBound(Split(strTmpString0, ";"))
                strTmp = Split(strTmpString0, ";")(i)
                rsNotes.Filter = "类型=" & IIf(byt未记显示位置 = 1, 99, 6) & " and X坐标=" & Val(Split(strTmpString2, ";")(i))
                rsNotes.Sort = "项目序号"
                If rsNotes.RecordCount > 0 Then
                    rsNotes!内容 = IIf(mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 2, "↓ ", "不升") & ";" & zlCommFun.Nvl(rsNotes!内容)
                    rsNotes.Update
                Else
                    If mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 2 Then strTmp = Replace(strTmp, "不升", "↓ ")
                    Call Record_Add(rsNotes, gstrFields, strTmp)
                    rsNotes!类型 = IIf(byt未记显示位置 = 1, 99, 6)
                    rsNotes.Update
                End If
            Next i
        End If
        
        '如果未记说明不显示，将取消记录集rsNote中类型为99的记录
        If byt未记显示位置 = 2 Then
            rsNotes.Filter = "类型=99"
            Do While Not rsNotes.EOF
                rsNotes.Delete
                rsNotes.MoveNext
            Loop
            rsNotes.Filter = ""
        End If
        rsPoints.Filter = 0
        '6 计算组织重复的点
        Call GetConverPoint(rsPoints)
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '7 开始点的信息并连线
        strTmp = ShowPointsNew(lngDC, objDraw, rsPoints, strEditors, int心率应用, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '8.心率脉搏短轴连线
        rsPoints.Filter = ""
        If strTmp <> "" And bln打印脉搏短绌 = True Then Call CreatePolyNew(rsPoints, objDraw, lngDC, strTmpDay, strTmp, int心率应用 = 2)
        '9输出说明信息
        '先处理未及说明和下标说明
        Dim strText As String
        Dim SinY35 As Single, SinY42 As Single
        Dim intAscCharNum As Integer
        
        strTime = ""
        strTmp = ""
        blnAllow = False
        SinX = 0: sinY = 0
        SinY35 = GetYCoordinate(objDraw, rsDrawItems, gint体温, 35, lngDC)
        SinY42 = GetYCoordinate(objDraw, rsDrawItems, gint体温, 42, lngDC)
        
        rsNotes.Filter = ""
        rsNotes.Sort = "X坐标,项目序号"
        With rsNotes
            Do While Not .EOF
                strTmp = ""
                For i = 0 To UBound(Split(!内容, ";"))
                    If Not (Split(!内容, ";")(i) = "不升" And byt未记显示位置 = 0 And Nvl(!类型) = 99) And Split(!内容, ";")(i) <> "" Then
                        If InStr(1, strTmp, Split(!内容, ";")(i)) = 0 Then
                            strTmp = strTmp & ";" & Split(!内容, ";")(i)
                        End If
                    End If
                Next i
                If Left(strTmp, "1") = ";" Then strTmp = Mid(strTmp, 2)
                If strTmp <> "" Then
                    strTime = Replace(strTmp, ";", " ")
                    If zlCommFun.Nvl(!类型) = 99 Then
                        If byt未记显示位置 = 1 Then '显示在体温的下面
                            If blnAllow = True Then
                                If Val(zlCommFun.Nvl(!X坐标)) <> SinX Then
                                    sinY = SinY35
                                Else
                                    strTime = " " & strTime
                                End If
                            Else
                                sinY = SinY35
                            End If
                            SinX = Val(zlCommFun.Nvl(!X坐标))
                            For i = 1 To Len(strTime)
                                If sinY < T_DrawClient.刻度区域.Bottom Then
                                    strText = Mid(strTime, i, 1)
                                    T_Size.H = objDraw.TextHeight(strText) / T_TwipsPerPixel.Y
                                    T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                                    If T_DrawClient.刻度区域.Bottom - sinY > T_Size.H Then
                                        Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(!颜色))
                                    End If
                                    If Asc(strText) < 0 Then
                                        sinY = sinY + T_Size.H
                                    Else
                                        sinY = sinY + T_Size.H / 2
                                    End If
                                End If
                            Next i
                            rsNotes!禁用 = 1
                            blnAllow = True
                        Else
                            rsNotes!内容 = strTime
                            rsNotes!Y坐标 = SinY42
                            blnAllow = False
                        End If
                    ElseIf zlCommFun.Nvl(!类型) = 6 Then
                        If blnAllow = True Then
                            If Val(zlCommFun.Nvl(!X坐标)) <> SinX Then
                                sinY = SinY35
                            Else
                                strTime = " " & strTime
                            End If
                        Else
                            sinY = SinY35
                        End If
                        SinX = Val(zlCommFun.Nvl(!X坐标))
                        For i = 1 To Len(strTime)
                            If i < 3 Then intAscCharNum = 0
                            If sinY < T_DrawClient.刻度区域.Bottom Then
                                strText = Mid(strTime, i, 1)
                                T_Size.H = objDraw.TextHeight(strText) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                                If Asc(strText) < 0 Then
                                    If intAscCharNum Mod 2 = 1 Then sinY = sinY + T_Size.H / 2
                                End If
                                '输出字体信息
                                If T_DrawClient.刻度区域.Bottom - sinY > T_Size.H Then
                                    Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(zlCommFun.Nvl(!颜色)))
                                End If
                                If Asc(strText) < 0 Then
                                    sinY = sinY + T_Size.H
                                    intAscCharNum = 0
                                Else
                                    sinY = sinY + T_Size.H / 2
                                    intAscCharNum = intAscCharNum + 1
                                End If
                            End If
                        Next i
                        rsNotes!禁用 = 1
                        blnAllow = False
                        sinY = 0
                    Else
                        '入出转等标记信息 开始Y轴坐标均更新为42
                        rsNotes!Y坐标 = SinY42
                    End If
                End If
            .MoveNext
            Loop
        End With
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst: rsNotes.Update
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call OutPutTextNew(objDraw, rsDrawItems, lngDC, rsNotes, strTmpDay, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '开始处理表下表格（特殊项目栏）
        ReDim ArrNewString(0)
        Dim arrTmpString0() As String, arrTmpString1() As String, arrTmpString2() As String
        
        '组织提取表下表格信息
        For i = 0 To UBound(ArrComTable)
            lng项目序号 = Val(Split(ArrComTable(i), "||")(0))
            str项目名称 = Split(ArrComTable(i), "||")(1)
            If lng项目序号 <> 4 Then
                strItemName = str项目名称
                If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                    strItems = strItems & ",'" & strItemName & "'"
                End If
            End If
        Next i
        
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        If Not mbln呼吸曲线 Then strItems = strItems & ",'呼吸'"
        strItems = strItems & ",'收缩压','舒张压'"
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        
        dtBegin = Int(CDate(strTmpDay) - 1)
        dtEnd = CDate(CDate(Format(strEndDay, "YYYY-MM-DD HH:mm:ss")) + 1)
        If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtBegin = CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss"))
        If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtEnd = CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss"))

        
        '提取所有表格项目数据信息
        strSQL = getSQLString("提取所有表格项目数据信息", blnMoved, strItems)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Print", _
                                            lng文件ID, _
                                            lng病人ID, _
                                            lng主页ID, _
                                            CDate(dtBegin), _
                                            CDate(dtEnd), _
                                            strItems, intBaby, lng护理等级, IIf(intBaby = 0, 1, 2), lngSectID, T_BodyItem.str表格内容)
                                                    
        ReDim Preserve ArrNewString(UBound(ArrComTable))
        For i = 0 To UBound(ArrComTable)
            If Split(ArrComTable(i), "||")(0) = 3 Then '呼吸项目
                lng项目序号 = Val(Split(ArrComTable(i), "||")(0))
                strNewTmpString = String(T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数, ";")
                arrTmpString0 = Split(String(T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数, ";"), ";")
                arrTmpString1 = Split(String(T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数, ";"), ";")
                arrTmpString2 = Split(String(T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数, ";"), ";")
                
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                rsTmp.Filter = "项目序号=" & gint呼吸
                With rsTmp
                    Do While Not .EOF
                        blnAdd = False
                        If CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")) Then
                            intCOl = GetCurveColumnNew(CDate(!时间), CDate(strTmpDay), gintHourBegin)
                            If intCOl > LBound(ArrNewTmpString) And intCOl <= UBound(ArrNewTmpString) Then
                            
                                If arrTmpString1(intCOl) <> "" Then
                                    If (Val(arrTmpString2(intCOl)) = 0 And Val(zlCommFun.Nvl(!显示, 0)) = 0) Or _
                                        (Val(arrTmpString2(intCOl)) = 1 And Val(zlCommFun.Nvl(!显示, 0)) = 1) Then
                                        
                                        '检查那个离重点时间更近
                                        SinX = GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                                        blnAdd = GetCanvasCenterNew(CDate(Format(arrTmpString1(intCOl), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                                    ElseIf Val(arrTmpString2(intCOl)) = 1 Then
                                        blnAdd = False
                                    Else
                                        blnAdd = True
                                    End If
                                    If blnAdd = True Then
                                        If Val(arrTmpString2(intCOl)) = 2 Then
                                            arrTmpString0(intCOl) = zlCommFun.Nvl(!结果) & "," & zlCommFun.Nvl(!体温部位)
                                            arrTmpString1(intCOl) = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    Else
                                        If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    End If
                                Else
                                    blnAdd = True
                                End If
                                
                                If blnAdd = True Then
                                    arrTmpString0(intCOl) = zlCommFun.Nvl(!结果) & "," & zlCommFun.Nvl(!体温部位)
                                    arrTmpString1(intCOl) = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                                    arrTmpString2(intCOl) = Val(zlCommFun.Nvl(!显示, 0))
                                End If
                                
                            End If
                        End If
ErrNext:
                    .MoveNext
                    Loop

                    For intCOl = LBound(ArrNewTmpString) To UBound(ArrNewTmpString)
                        ArrNewTmpString(intCOl) = IIf(Val(arrTmpString2(intCOl)) = 2, "", arrTmpString0(intCOl))
                    Next intCOl
                    
                    strNewTmpString = Join(ArrNewTmpString, "||")
                End With
                ArrNewString(i) = strNewTmpString
            Else
                blnColor = False
                int频次 = Val(Split(ArrComTable(i), "||")(4))
                strTmp = Val(Split(ArrComTable(i), "||")(6)) '项目表示 4表示汇总项目
                lng项目序号 = Val(Split(ArrComTable(i), "||")(0))
                str项目名称 = Split(ArrComTable(i), "||")(1)
                int项目性质 = Val(Split(ArrComTable(i), "||")(5))
                int项目类型 = Val(Split(ArrComTable(i), "||")(7))
                int入院首测 = Val(Split(ArrComTable(i), "||")(8))
                
                If Val(strTmp) = 4 Or IsWaveItem(lng项目序号) Then
                    If int频次 > 2 Then int频次 = 2 '汇总/波动项目频次只能是 1 、 2
                End If
                
                blnColor = (int项目性质 = 2 And int项目类型 = 1 And Val(strTmp) = 0)
                strNewTmpString = String(Val(int频次) * T_BodyStyle.lng天数, ";")
              
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                For j = 0 To T_BodyStyle.lng天数 - 1
                    strBegin = DateAdd("D", j, CDate(strTmpDay))
                    If CDate(strBegin) > CDate(strEndDay) Then strBegin = strEndDay
                    int舒张压 = 0
                    int收缩压 = 0
                    Int列号 = 0
                    strTime = ""
                    intCOl = 0
                    
                    Set rsDownTab = ReturnItemRecord(rsTmp, Int(CDate(strBegin)), CDate(strBeginDate), lng项目序号 & ";" & str项目名称 & ";" & _
                                    int频次 & ";" & Val(strTmp) & ";" & int项目性质 & ";" & int入院首测, bln汇总当天, bln录入小时)
                    If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                    rsDownTab.Sort = "时间,项目序号,序号"
                    With rsDownTab
                        Do While Not .EOF
                            lngColor = 0
                            str结果 = zlCommFun.Nvl(!记录内容)
                            intCOl = Val(!序号)
                            intCOl = intCOl + j * int频次
                            If blnColor Then lngColor = Val(zlCommFun.Nvl(!未记说明))
                            
                            Select Case zlCommFun.Nvl(!项目名称)
                                Case "舒张压"
                                    If int舒张压 <> intCOl Then
                                        If Trim(ArrNewTmpString(intCOl)) <> "" Or str结果 <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = Trim(Split(ArrNewTmpString(intCOl), "/")(0)) & "/" & str结果
                                            Else
                                                ArrNewTmpString(intCOl) = "/" & str结果
                                            End If
                                            
                                            mrsCurInfo.Filter = "名称='" & str结果 & "'"
                                            If Not mrsCurInfo.EOF Then ArrNewTmpString(intCOl) = str结果
                                        End If
                                         int舒张压 = intCOl
                                         If ArrNewTmpString(intCOl) = "/" Then ArrNewTmpString(intCOl) = ""
                                    End If
                                Case "收缩压"
                                    If int收缩压 <> intCOl Then
                                        If ArrNewTmpString(intCOl) <> "" Or str结果 <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = str结果 & "/" & Trim(Split(ArrNewTmpString(intCOl), "/")(1))
                                            Else
                                                ArrNewTmpString(intCOl) = str结果 & "/"
                                            End If
                                        End If
                                        int收缩压 = intCOl
                                    End If
                                Case Else
                                    If Int列号 <> intCOl Then
                                        ArrNewTmpString(intCOl) = Replace(str结果, "-#", "") & "-#" & lngColor
                                        Int列号 = intCOl
                                    End If
                            End Select
                        .MoveNext
                        Loop
                    End With
                    
                    If Format(strBegin, "YYYY-MM-DD") = Format(strEndDay, "YYYY-MM-DD") Then Exit For
                Next j
                strNewTmpString = Join(ArrNewTmpString, "||")
                ArrNewString(i) = strNewTmpString
            End If
        Next i
        
        '项目序号||部位+项目名称||项目单位||项目值域||记录频次||项目性质||项目表示
        For i = 0 To UBound(ArrComTable)
            strTmpString0 = ""

            If Trim(CStr(Split(ArrComTable(i), "||")(2))) <> "" Then
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1))) & "(" & Trim(CStr(Split(ArrComTable(i), "||")(2))) & ")"
            Else
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1)))
            End If
           
            ArrNewString(i) = Trim(CStr(Split(ArrComTable(i), "||")(0))) & ";" & strTmpString0 & ";" & ArrNewString(i)
        Next i
        
        '显示皮试结果
        If bln显示皮试 = True Then
            strSQL = getSQLString("显示皮试结果", blnMoved)
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取病人过敏记录信息", lng病人ID, lng主页ID, intBaby, CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")))

            strNewTmpString = String(T_BodyStyle.lng天数, ";")
            ArrNewTmpString = Split(strNewTmpString, ";")
            intCOl = 0

            Do While Not rsTmp.EOF
                intCOl = DateDiff("D", CDate(Format(strTmpDay, "YYYY-MM-DD")), CDate(Format(rsTmp!时间, "YYYY-MM-DD"))) + 1
                ArrNewTmpString(intCOl) = Nvl(rsTmp!药物名)
                rsTmp.MoveNext
            Loop
            strNewTmpString = Join(ArrNewTmpString, "||")
            ReDim Preserve ArrNewString(UBound(ArrNewString) + 1)
            ArrNewString(UBound(ArrNewString)) = "-999;皮试结果" & ";" & strNewTmpString
        End If
        
        lngCurX = X

        '开始绘画表格项目并展示数据
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call DrawBodyRecordItemNew(lngDC, objDraw, ArrNewString, rsItems, lngCurX, T_DrawClient.体温区域.Bottom, T_DrawClient.体温区域.Right, intRepairRows, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        
        lngCurX = X
        lngCurY = lngCurY
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '开始打印 页数 住院周数 和 体温说明信息
        Call DrawBodyPageFooterNew(lngDC, objDraw, lngCurX, lngCurY, T_DrawClient.体温区域.Right, intPageNo, intEndPage, str体温说明, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '页脚图形输出
        Call frmTendFileRead.PrintRTBData(objDraw, False, lngButtom)
        
        If Not blnPrint Then objDraw.Refresh
NOPageSub:  Next

    If blnPrint = False Then Call DrawDeviceCapsNew(lngDC, objDraw)
     
    Call ShowFlash
    PrintOrPreviewBodyStateNew = True
    Screen.MousePointer = 0
    Set stdSet = Nothing
    GoTo ErrClare
    Exit Function
ErrPrint:
    Call ShowFlash
    Screen.MousePointer = 0
    
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo ErrClare
    Call SaveErrLog
ErrExit:
    Call ShowFlash
    Screen.MousePointer = 0
    msngTwips = 1
    Err.Clear
    PrintOrPreviewBodyStateNew = False
    Set stdSet = Nothing
    GoTo ErrClare
ErrClare:
    Call ClearData(M_DrawClient.偏移量X, M_DrawClient.偏移量Y, M_DrawClient.刻度单位, M_DrawClient.行单位, M_DrawClient.时间行单位, M_DrawClient.时间列单位, _
                    M_DrawClient.列单位, M_DrawClient.双倍, M_DrawClient.总列数, M_DrawClient.独立曲线总行数, lng刻度宽度)
    T_DrawClient.刻度区域 = M_DrawClient.刻度区域
    T_DrawClient.体温区域 = M_DrawClient.体温区域
    T_DrawClient.曲线总区域 = M_DrawClient.曲线总区域
    Call ErrEmpty
    Set stdSet = Nothing
End Function

Public Sub DrawUpTableNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmpString As String, _
    ByVal lngX As Long, ByVal lngY As Long, ByVal lngLeft As Long, lngOutY As Long, Optional sngScale As Single)
'-----------------------------------------------------------------------------------------------------------------------
'输出一般项目栏信息（包括 住院日期,天数,手术后天数和时间栏）
'参数:lngDC 绘图对象的DC，strTmpString 有住院日期，天数 和术后天数组成的字符串
'     lngX 左边距,lngY上边距,lngLeft 右边距(可以绘图的最大右边距)
'出参:lngOutY 返回绘图后的上边距
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim ArrCode() As String
    Dim strTmp As String
    Dim arrTmpTime() As String '住院时间
    Dim arrTmpDay() As String  '住院天数
    Dim arrOptDay() As String '术后天数
    Dim lngCurX As Long, lngCurY As Long, lngStartY As Long, lngStartX As Long, lngTmpX As Long
    Dim lngColor As Long
    Dim intBold As Integer, intFine As Integer
    Dim str日期 As String
    Dim str住院天数 As String
    Dim str手术后天数 As String
    Dim str时间 As String
    
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    str日期 = Split(T_BodyStyle.str列头名称, "@")(0)
    str住院天数 = Split(T_BodyStyle.str列头名称, "@")(1)
    str手术后天数 = Split(T_BodyStyle.str列头名称, "@")(2)
    str时间 = Split(T_BodyStyle.str列头名称, "@")(3)
    
    ArrCode = Split(strTmpString, "||")
    strTmp = strTmpString & String(2 - UBound(ArrCode), "||")
    ArrCode = Split(strTmp, "||")
    arrOptDay = Split(ArrCode(2), "'")
    arrTmpTime = Split(ArrCode(0), "'")
    arrTmpDay = Split(ArrCode(1), "'")

    lngCurX = lngX: lngStartX = lngX
    lngCurY = lngY: lngStartY = lngY
    
    '开始画表格栏
    
    'X
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.时间列单位 + 6
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    
    'Y
    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = T_DrawClient.刻度区域.Right

    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)

    For i = 0 To T_BodyStyle.lng天数 - 1
        lngCurX = lngCurX + T_DrawClient.列单位 * T_BodyStyle.lng监测次数
        Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    Next i
    
    lngCurX = T_DrawClient.刻度区域.Right
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 3
    '时间
    For i = 0 To T_BodyStyle.lng天数 - 1
        lngCurX = T_DrawClient.刻度区域.Right + i * T_DrawClient.列单位 * T_BodyStyle.lng监测次数
        For j = 1 To T_BodyStyle.lng监测次数 - 1
            lngCurX = lngCurX + T_DrawClient.列单位
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + T_DrawClient.时间列单位 + 6, PS_SOLID, intFine, RGB_BLACK)
        Next j
    Next i
    
    '开始输出信息
    '日期信息
    lngCurY = lngStartY
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, str日期, Len(str日期), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, str日期, T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, str日期, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    For i = 0 To UBound(arrTmpTime)
        lngCurX = T_DrawClient.刻度区域.Right + i * T_BodyStyle.lng监测次数 * T_DrawClient.列单位
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpTime(i)), Len(CStr(arrTmpTime(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.时间列单位 / 2, CStr(arrTmpTime(i)), T_DrawClient.列单位 * T_BodyStyle.lng监测次数, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpTime(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 1
    '住院天数
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, str住院天数, Len(str住院天数), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, str住院天数, T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, str住院天数, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    
    For i = 0 To UBound(arrTmpDay)
        lngCurX = T_DrawClient.刻度区域.Right + i * T_BodyStyle.lng监测次数 * T_DrawClient.列单位
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpDay(i)), Len(CStr(arrTmpDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.时间列单位 / 2, CStr(arrTmpDay(i)), T_DrawClient.列单位 * T_BodyStyle.lng监测次数, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    '术/娩后天数
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 2
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, str手术后天数, Len(str手术后天数), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, str手术后天数, T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, str手术后天数, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    
    '51283,刘鹏飞,2012-07-11,手术天数颜色
    lngColor = Val(zlDatabase.GetPara("手术天数显示颜色", glngSys, 1255, "255"))
    For i = 0 To UBound(arrOptDay)
        lngCurX = T_DrawClient.刻度区域.Right + i * T_BodyStyle.lng监测次数 * T_DrawClient.列单位
        Call SetTextColor(lngDC, lngColor)
        Call GetTextExtentPoint32(lngDC, CStr(arrOptDay(i)), Len(CStr(arrOptDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.时间列单位 / 2, CStr(arrOptDay(i)), T_DrawClient.列单位 * T_BodyStyle.lng监测次数, True, , sngScale)
        Call DrawText(lngDC, CStr(arrOptDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    lngColor = 0
    '时间
    lngCurY = lngStartY + T_DrawClient.时间列单位 * 3
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, str时间, Len(str时间), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.时间列单位 / 2, str时间, T_DrawClient.刻度区域.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, str时间, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.刻度区域.Right
    
    For i = 0 To T_BodyStyle.lng天数 - 1
        lngCurX = T_DrawClient.刻度区域.Right + i * T_BodyStyle.lng监测次数 * T_DrawClient.列单位
        '输出上午下午时间
        For j = 0 To T_BodyStyle.lng监测次数 - 1
            strTmp = ""
            
            strTmp = gintHourBegin + T_BodyStyle.lng时间间隔 * j

            lngColor = GetTimeColor(Val(strTmp))
            lngTmpX = lngCurX + T_DrawClient.列单位 * j
            Call SetTextColor(lngDC, lngColor)
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            Call GetTextRect(objDraw, lngTmpX - 1, lngCurY + (T_DrawClient.时间列单位 + 6) / 2, strTmp, T_DrawClient.列单位, True, , sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Next j
    Next i
    lngOutY = lngStartY + T_DrawClient.时间列单位 * 4 + 6
End Sub


Public Function DrawCanvasNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, rsDrawItems As ADODB.Recordset, Optional ByVal bln不打印心率列 As Boolean = False, Optional sngScale As Single = 1) As String
'------------------------------------------------------------------------------------------------------
'功能:画刻度区域和体温区域并输出刻度值信息
'参数:lngDC 绘图对象的DC，objDraw 绘画对象.rsTemp:体温曲线项目记录集(A.项目序号,A.排列序号,A.记录名,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,C.项目单位 单位,A.最高行-2 AS 最高行,B.部位)
'出参:返回各个曲线的具体信息包括( "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色")
'返回说明信息(项目的符号)
'-------------------------------------------------------------------------------------------------------
    Dim str说明 As String
    Static SlngMaxY As Long                 '记录上一次的最大高度，以决定本次是否需要重画
    Dim lngCurX     As Long, lngCurY As Long   '当前位置
    Dim lngMaxX     As Long, lngMaxY As Long   '边界
    Dim lngCurAlerY As Long '警戒线坐标
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln双行 As Boolean                  '此参数由用户指定,bln双行=TRUE表示只显示五行;否则显示十行
    Dim bln粗线 As Boolean                  '此参数由用户指定,大行分界是粗线还是细线
    Dim blnAche As Boolean                  '是否是疼痛独立曲线
    '以下都是标准尺度
    Dim intLineMode   As Integer
    Dim blnDoubleRow  As Boolean             '贰行做为一行打印输出
    Dim sinAlertness  As Single              '警戒线,起辅助作用
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sinRowStep As Single, lngInitRowStep As Long
    Dim arrTemp()     As String
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngFont As Long, lngOldFont As Long
    Dim sinY单位 As Single '曲线单位输出的Bottom
    Dim lngY As Long, lngCurveRows As Long, lng刻度宽度 As Long, lngX As Long
    
    '以下与绘图区域相关(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    Dim sin刻度 As Single, bln显示刻度 As Boolean, blnFirst As Boolean
    Dim sin刻度间隔 As Single, sinBegin刻度 As Single, dbl单位值 As Double
    Dim sinCurAlerY As Single
    
    Dim str最大值坐标 As String, str最小值坐标 As String

    On Error GoTo Errhand
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '所有曲线项目的作图区域(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    gstrFields = "项目序号," & adDouble & ",18|最大值," & adDouble & ",18|最小值," & adDouble & ",18|" & "单位值," & adDouble & _
        ",18|最大值坐标," & adLongVarChar & ",20|最小值坐标," & adLongVarChar & ",20|" & "单位刻度," & adLongVarChar & ",20|显示模式," & adDouble & ",5|颜色," & adDouble & ",18"
    Call Record_Init(rsDrawItems, gstrFields)
    '------------------------------------------------------------------------------------------------------------------
    '赋初值
    intLineMode = PS_SOLID
    lngColStep = T_DrawClient.列单位
    lngInitRowStep = T_BodyStyle.lng曲线行高 / T_TwipsPerPixel.Y * sngScale
    sinRowStep = T_DrawClient.行单位
    lngLableStep = T_DrawClient.刻度单位
    lng刻度宽度 = T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X * sngScale
    
    '体温单以单格显示(不勾此选项以双格显示，没两个刻度显示一次) 1：单格显示 0：双格显示
    If zlDatabase.GetPara("体温单显示格式", glngSys, 1255, 0) = 1 Then
        bln双行 = False
    Else
        bln双行 = True
    End If
    'True表示贰行只输出一行,效果是一个刻度只显示了五行;否则一个刻度显示十行,由用户调整参数决定,与blnDoubleRow无关
    bln粗线 = True
    
    If Not bln粗线 Then intLineMode = PS_DASHDOTDOT
    '画表格
    rsTemp.Filter = "记录法=1"
    intLables = rsTemp.RecordCount
    rsTemp.Filter = "项目序号=" & gint心率 & " And 记录法=1"
        If rsTemp.RecordCount > 0 And bln不打印心率列 = True Then
        rsTemp.Filter = 0
        intLables = intLables - 1
    Else
        rsTemp.Filter = 0
    End If
    If intLables <= 0 Then intLables = 1
    
    lngCurX = T_DrawClient.偏移量X
    lngCurY = T_DrawClient.偏移量Y
    lngMaxX = (intLables * lngLableStep) + (T_BodyStyle.lng天数 * T_BodyStyle.lng监测次数 * lngColStep) + T_DrawClient.偏移量X    '刻度+7*宽度+偏移量X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.总列数 * sinRowStep + T_DrawClient.偏移量Y '（为表格大小，还需加上起始Y坐标）
       
    str说明 = ""
        
    SlngMaxY = lngMaxY
    T_DrawClient.刻度单位 = lngLableStep
    T_DrawClient.行单位 = sinRowStep
    T_DrawClient.列单位 = lngColStep
    T_DrawClient.双倍 = blnDoubleRow
    
    For lngRow = 1 To intLables
        lngX = lngCurX
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow = 1, intBold, intFine), RGB_BLACK)
        If lngRow = intLables Then
            lngCurX = lngCurX + lng刻度宽度 - ((intLables - 1) * lngLableStep)
        Else
            lngCurX = lngCurX + lngLableStep
        End If
        
        Call DrawLine(lngDC, lngX, lngCurY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDC, lngX, lngMaxY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        If lngRow = intLables Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        End If
    Next
    
    T_DrawClient.刻度区域.Left = T_DrawClient.偏移量X
    T_DrawClient.刻度区域.Top = lngCurY
    T_DrawClient.刻度区域.Right = lngCurX
    T_DrawClient.刻度区域.Bottom = lngMaxY
    
    '默认添加一行用于显示项目名称
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(lngDC, T_DrawClient.偏移量X, lngCurY, lngMaxX, lngCurY, PS_SOLID, intFine, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    '画体温单所有行
    For lngRow = 0 To T_DrawClient.总列数 - 1
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '画体温单的所有行
        If ((blnDoubleRow Or bln双行) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln双行) Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln粗线, intBold, intFine), RGB_BLACK)
        End If
    Next
    
    lngCurY = T_DrawClient.刻度区域.Top
    
    '画体温单所有列
    For lngRow = 1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
        lngCurX = lngCurX + lngColStep
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, intBold, intFine), IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, RGB_RED, RGB_BLACK))
    Next
        
    lngCurX = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Top = T_DrawClient.刻度区域.Top
    T_DrawClient.体温区域.Right = lngMaxX
    T_DrawClient.体温区域.Bottom = lngMaxY
    
    '画体温区域底线
    Call DrawLine(lngDC, T_DrawClient.体温区域.Left, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)

    '画刻度框的标尺（从固定不变的10行开始标识）
    intLables = 1
    rsTemp.Filter = "记录法=1"
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            If Not (bln不打印心率列 = True And !项目序号 = gint心率) Then
                '显示刻度框项目的名称及符号,如体温×
                lngCurX = T_DrawClient.刻度区域.Left + ((intLables - 1) * T_DrawClient.刻度单位)
                If .AbsolutePosition = .RecordCount Then
                    lngLableStep = (T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) - ((intLables - 1) * T_DrawClient.刻度单位)
                Else
                    lngLableStep = T_DrawClient.刻度单位
                End If
                lngCurY = T_DrawClient.刻度区域.Top
                Set gstdSet = New StdFont
                gstdSet.Name = "宋体"
                gstdSet.Size = 9 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                '输出体温项目的名称
                Call SetTextColor(lngDC, zlCommFun.Nvl(!记录色, RGB_BLACK))
                Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.TextHeight(zlCommFun.Nvl(!记录名)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!记录名)), lngLableStep, , , sngScale)
'                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(zlCommFun.Nvl(!记录名)), zlCommFun.Nvl(!记录色, RGB_BLACK))
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!记录名)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                '设置字体大小
                Set gstdSet = New StdFont
                gstdSet.Name = "宋体"
                gstdSet.Size = 8 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
    
                '输出项目单位
                Call GetTextRect(objDraw, lngCurX, lngCurY + lngInitRowStep * 2 + objDraw.TextHeight(zlCommFun.Nvl(!单位)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!单位)), lngLableStep, , , sngScale)
'                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(zlCommFun.Nvl(!单位, 0)), zlCommFun.Nvl(!记录色, RGB_BLACK))
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!单位, 0)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                sinY单位 = T_LableRect.Bottom
                intLables = intLables + 1
            End If
            objDraw.Font.Size = 9 * sngScale
            '强制设定体温曲线项目的显示模式
            Select Case !项目序号

                Case gint体温  '体温整数时输出刻度
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 1)
                    dbl单位值 = 0.1
                    sinAlertness = zlCommFun.Nvl(!警示线, 37)
                    arrTemp = Split(zlCommFun.Nvl(!记录符, "·,×,○,△"), ",")
                    str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(口温" & arrTemp(0) & ",腋温" & arrTemp(1) & ",肛温" & arrTemp(2) & ",耳温" & arrTemp(3) & ")"

                Case gint脉搏, gint心率  '脉搏/心跳按10的倍数输出刻度
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 10)
                    dbl单位值 = 2
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)

                    If !项目序号 = gint脉搏 Then
                        str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(缺省记录符" & zlCommFun.Nvl(!记录符, "+") & ",起搏器H)"
                    Else
                        str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(" & zlCommFun.Nvl(!记录符, "Ο") & ")"
                    End If

                Case gint呼吸  '呼吸按5的倍数输出刻度
                    mbln呼吸曲线 = True
                    dbl单位值 = 1
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 5)
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(自主呼吸" & zlCommFun.Nvl(!记录符, "*") & ",呼吸机R)"
                Case Else
                    dbl单位值 = Val(zlCommFun.Nvl(!单位值, 0))
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, Val(zlCommFun.Nvl(!单位值, 0)) * 10)
                    If sin刻度间隔 > Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值)) Then
                        sin刻度间隔 = Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值))
                    End If
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    str说明 = str说明 & "、" & zlCommFun.Nvl(!记录名) & "(" & zlCommFun.Nvl(!记录符, "*") & ")"
            End Select

            '赋初值
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow) '固定前?行的高度不输出刻度

            '根据最高行定位到有效位置
            lngCurY = lngCurY + (T_DrawClient.行单位 * zlCommFun.Nvl(!最高行, 0))
            blnFirst = False
            Do While True
                bln显示刻度 = False
                If blnFirst = False Then     '刚进入循环，此时取的最大值
                    sin刻度 = zlCommFun.Nvl(!最大值, 0)
                    sinBegin刻度 = sin刻度
                    str最大值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                    blnFirst = True
                Else                    '计算得到每个刻度的值
                    sin刻度 = sin刻度 - dbl单位值     '如果目前显示模式为双倍，则按双倍累计
                End If
                
                '根据设置的刻度间隔显示刻度值
                If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - IIf(T_DrawClient.双倍, sin刻度间隔 * 2, sin刻度间隔)
                If sinBegin刻度 < Val(Format(zlCommFun.Nvl(!最小值), "#0.00")) Then sinBegin刻度 = Val(Format(zlCommFun.Nvl(!最小值), "#0.00"))
                
                If bln显示刻度 And Not (bln不打印心率列 = True And !项目序号 = gint心率) Then
                    '控制最大值不与曲线单位重复
                    If sin刻度 = Val(Nvl(!最大值, 0)) And lngCurY < sinY单位 Then
                        Call GetTextRect(objDraw, lngCurX, sinY单位, Format(sin刻度, "#0"), lngLableStep, , , sngScale)
                    ElseIf lngCurY = T_DrawClient.刻度区域.Bottom Then
                        Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y), Format(sin刻度, "#0"), lngLableStep, , , sngScale)
                    Else
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Format(sin刻度, "#0"), lngLableStep, , , sngScale)
                    End If
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Format(sin刻度, "#0"), zlCommFun.Nvl(!记录色, RGB_BLACK))
                    Call DrawText(lngDC, Format(sin刻度, "#0"), -1, T_LableRect, DT_CENTER)
                End If
                '如果不在有效范围内，或者超出画布则退出
                If Val(Format(sin刻度, "#0.00")) <= Val(Format(zlCommFun.Nvl(!最小值), "#0.00")) Or Format(lngCurY, "#0") > T_DrawClient.刻度区域.Bottom Then
                    str最小值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                    '添加该项目(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
                    gstrFields = "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色"
                    gstrValues = zlCommFun.Nvl(!项目序号) & "|" & zlCommFun.Nvl(!最大值, 0) & "|" & zlCommFun.Nvl(!最小值, 0) & _
                    "|" & dbl单位值 & "|" & str最大值坐标 & "|" & str最小值坐标 & "|" & T_DrawClient.行单位 & "," & T_DrawClient.列单位 & "|" & sin刻度间隔 & "|" & !记录色
                    Call Record_Add(rsDrawItems, gstrFields, gstrValues)
                    
                    '辅助线或警示线
                    If blnDoubleRow = False And (sinAlertness < Val(Nvl(!最大值)) And sinAlertness > Val(Nvl(!最小值))) Then
                        lngCurAlerY = Val(GetYCoordinate(objDraw, rsDrawItems, Val(Nvl(!项目序号)), sinAlertness))
                        Call DrawLine(lngDC, T_DrawClient.体温区域.Left, lngCurAlerY, lngMaxX, lngCurAlerY, intLineMode, intBold, RGB_RED)
                    End If
                    
                    Exit Do
                End If
                lngCurY = lngCurY + T_DrawClient.行单位
            Loop
            sinBegin刻度 = 0
            sin刻度 = 0                 '控制从第一行开始输出
            .MoveNext
        Loop
    End With
    
    '完成独立曲线部分的输出
    rsTemp.Filter = "记录法=3"
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            lngY = lngMaxY
            lngCurY = lngY
            lngCurX = T_DrawClient.偏移量X
            lngCurveRows = ((Val(Nvl(!最大值, 0)) - Val(Nvl(!最小值, 0))) / Val(Nvl(!单位值)))
            If Val(Nvl(!最高行, 0)) > 0 Then lngCurveRows = lngCurveRows + Val(Nvl(!最高行, 0))
            If lngCurveRows Mod 2 = 1 Then lngCurveRows = lngCurveRows + 1
            If lngCurveRows > 0 Then
                lngMaxY = lngCurveRows * sinRowStep + lngCurY
                '完成刻度区域的绘制
                Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, lngCurX + Fix(T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X * sngScale / 2), lngCurY, lngCurX + Fix(T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X * sngScale / 2), lngMaxY, PS_SOLID, intFine, RGB_BLACK)
                Call DrawLine(lngDC, lngCurX + lng刻度宽度, lngCurY, lngCurX + lng刻度宽度, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, lngCurX, lngMaxY, lngCurX + lng刻度宽度, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                blnAche = Nvl(!记录名) Like "疼痛强度*"
                '完成所有行的绘制
                lngCurX = lngCurX + lng刻度宽度
                For lngRow = 1 To lngCurveRows
                    '画体温单的所有行
                    If lngRow <> 0 Then
                        lngCurY = lngCurY + sinRowStep
                    End If
                    If ((blnDoubleRow Or bln双行) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln双行) Then
                        If blnAche = True Then
                            Call DrawLine(lngDC, lngCurX - lng刻度宽度 + Fix(T_BodyStyle.lng刻度宽度 / T_TwipsPerPixel.X * sngScale / 2) + 1, lngCurY, lngCurX, lngCurY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_BLACK)
                        End If
                        Call DrawLine(lngDC, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0 And blnAche = False, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln粗线 And blnAche = False, intBold, intFine), RGB_BLACK)
                    End If
                Next
                '画底线
                Call DrawLine(lngDC, lngCurX, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                lngCurY = lngY
                '画体温单所有列
                For lngRow = 1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
                    lngCurX = lngCurX + lngColStep
                    Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, intBold, intFine), IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, RGB_RED, RGB_BLACK))
                Next
                
                '完成项目名称和刻度的输出
                lngX = T_DrawClient.刻度区域.Left
                lngCurX = lngX
                lngCurY = lngY
                '输出体温项目的名称
                '创建字体
                Set gstdSet = New StdFont
                gstdSet.Name = "宋体"
                gstdSet.Size = 9 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                Call SetTextColor(lngDC, Nvl(!记录色, RGB_BLACK))
                T_Size.H = objDraw.ScaleY(objDraw.TextHeight("刘"), vbTwips, vbPixels)
                If T_Size.H * Len(Nvl(!记录名)) >= lngCurveRows * sinRowStep Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * sinRowStep) - (T_Size.H * Len(Nvl(!记录名)))) \ 2
                End If
                For lngRow = 1 To Len(Nvl(!记录名))
                    Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Nvl(!记录名), lngRow, 1), lng刻度宽度 \ 2, False)
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Mid(Nvl(!记录名), lngRow, 1), Nvl(!记录色, RGB_BLACK))
                    Call DrawText(lngDC, Mid(Nvl(!记录名), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                '输出项目单位
                lngCurY = lngY: If Nvl(!记录名) <> "" Then lngCurX = T_LableRect.Right
                If Trim(Nvl(!单位)) <> "" And Nvl(!记录名) <> "" Then
                    '设置字体大小
                    Set gstdSet = New StdFont
                    gstdSet.Name = "宋体"
                    gstdSet.Size = 8 * sngScale
                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDC, lngFont)
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight("刘"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(Nvl(!单位))) >= lngCurveRows * sinRowStep Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * sinRowStep) - (T_Size.H * Len(Nvl(!单位)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(Nvl(!单位)))
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Trim(Nvl(!单位)), lngRow, 1), 0, False)
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Mid(Trim(Nvl(!单位)), lngRow, 1), Nvl(!记录色, RGB_BLACK))
                        Call DrawText(lngDC, Mid(Trim(Nvl(!单位)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call SelectObject(lngDC, lngOldFont)
                    Call DeleteObject(lngFont)
                    Call ReleaseFontIndirect(objDraw)
                End If
                objDraw.Font.Size = 9 * sngScale
                '设置字体大小
                dbl单位值 = Val(Nvl(!单位值, 0))
                sin刻度间隔 = Nvl(!刻度间隔, Val(Nvl(!单位值, 0)) * 10)
                If sin刻度间隔 > Val(Nvl(!最大值)) - Val(Nvl(!最小值)) Then
                    sin刻度间隔 = Val(Nvl(!最大值)) - Val(Nvl(!最小值))
                End If
                sinAlertness = Nvl(!警示线, 0)
                str说明 = str说明 & "、" & Nvl(!记录名) & "(" & Nvl(!记录符, "*") & ")"
                lngCurY = lngY + (sinRowStep * Val(Nvl(!最高行, 0)))
                blnFirst = False
                Do While True
                    bln显示刻度 = False
                    If blnFirst = False Then     '刚进入循环，此时取的最大值
                        sin刻度 = Nvl(!最大值, 0)
                        sinBegin刻度 = sin刻度
                        str最大值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                        blnFirst = True
                    Else                    '计算得到每个刻度的值
                        sin刻度 = sin刻度 - dbl单位值     '如果目前显示模式为双倍，则按双倍累计
                    End If
    
                    '根据设置的刻度间隔显示刻度值
                    If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                    If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - sin刻度间隔
                    If sinBegin刻度 < Val(Format(Nvl(!最小值), "#0.00")) Then sinBegin刻度 = Val(Format(Nvl(!最小值), "#0.00"))
    
                    If bln显示刻度 Then
                        '控制最大值不与曲线单位重复
'                        lngCurX = lngX + lng刻度宽度 - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin刻度, "#0.0"))), vbTwips, vbPixels)
'                        lngCurX = lngCurX - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        lngCurX = lngX + lng刻度宽度 - Fix(lng刻度宽度 / 4) - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin刻度, "#0.0"))) / 2, vbTwips, vbPixels)
                        If sin刻度 = Val(Nvl(!最大值, 0)) And lngCurY = lngY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY + (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")))
                        ElseIf lngCurY = lngMaxY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")))
                        Else
                            Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin刻度, "#0.0")))
                        End If
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Val(Format(sin刻度, "#0.0")), Nvl(!记录色, RGB_BLACK))
                        Call DrawText(lngDC, Val(Format(sin刻度, "#0.0")), -1, T_LableRect, DT_CENTER)
                    End If
                    If Val(Format(sin刻度, "#0.00")) <= Val(Format(Nvl(!最小值), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                        str最小值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                        '添加该项目(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
                        gstrFields = "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色"
                        gstrValues = zlCommFun.Nvl(!项目序号) & "|" & zlCommFun.Nvl(!最大值, 0) & "|" & zlCommFun.Nvl(!最小值, 0) & _
                        "|" & dbl单位值 & "|" & str最大值坐标 & "|" & str最小值坐标 & "|" & T_DrawClient.行单位 & "," & T_DrawClient.列单位 & "|" & sin刻度间隔 & "|" & !记录色
                        Call Record_Add(rsDrawItems, gstrFields, gstrValues)
                    
                        '输出警戒线
                        If blnDoubleRow = False And sinAlertness > Val(Nvl(!最小值)) And sinAlertness < Val(Nvl(!最大值)) Then
                            '根据最大值与当前值之间的差额,以及最小值,计算得到相差多少个刻度,再根据单位刻度得到实际坐标
                            lngCurAlerY = Val(GetYCoordinate(objDraw, rsDrawItems, Val(Nvl(!项目序号)), sinAlertness))
                            Call DrawLine(lngDC, lngX + lng刻度宽度, lngCurAlerY, lngMaxX, lngCurAlerY, PS_SOLID, 1, RGB_RED)
                        End If
                        Exit Do
                    End If
                    lngCurY = lngCurY + T_DrawClient.行单位
                Loop
                sinBegin刻度 = 0
                sin刻度 = 0
            End If
        .MoveNext
        Loop
       T_DrawClient.体温区域.Bottom = 2 * mintNullRow * lngInitRowStep + (T_DrawClient.总列数 + T_DrawClient.独立曲线总行数) * sinRowStep + T_DrawClient.偏移量Y
    End With
    str说明 = "说明:" & Mid(str说明, 2)
    
    DrawCanvasNew = str说明
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowPointsNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsPoint As ADODB.Recordset, _
    strEditors() As Variant, Optional int心率引用 As Integer = 1, Optional ByVal sngScale As Single = 1) As String
'-------------------------------------------------------------------------------------
'功能:输出体温项目的连线和图形输出
'参数::lngDC 绘图对象的DC，objDraw 绘画对象.rsPoint 所有项目点的集合(序号|数值|部位|标记|时间|项目序号|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号)
'strEditors 体温，心率，呼吸，脉搏的信息(项目序号||项目名称||项目单位||项目值域||记录符||记录色)
'返回:心率点的集合 !X坐标 & ";" & !Y坐标 & "," & !X坐标 & ";" & !Y坐标
'-------------------------------------------------------------------------------------
    Dim sin原X As Single, sin原Y As Single
    Dim lng项目序号 As Long
    Dim SinX As Single, sinY As Single  '物理降温使用
    Dim dblvalue As Double
    Dim dblMaxValue As Double, dblMinValue As Double
    Dim lngRGB As Long
    Dim strChar As String, str部位 As String, strTmp As String, strPic As String
    Dim str心率 As String
    Dim lngCount As Long '重叠项目数量
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnLine As Boolean
    Dim i As Integer
    Dim X1 As Single
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim bln不升符号 As Boolean
    Dim lngWith As Long
    Dim bln符号 As Boolean
    
    On Error GoTo Errhand
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    
    If blnPrinter = True Then
        intBold = 4
        intFine = 4
    Else
        intBold = 2
        intFine = 1
    End If
    rsPoint.Filter = ""
    rsPoint.Sort = "项目序号,时间"
    '首先进行连线
    With rsPoint
        Do While Not .EOF
            For i = 0 To UBound(strEditors)
                If Val(Split(strEditors(i), "||")(0)) = Val(zlCommFun.Nvl(!项目序号)) Then
                     Exit For
                End If
            Next i
            If Not (zlCommFun.Nvl(!项目序号) = gint体温 And Val(zlCommFun.Nvl(!标记)) = 1) Then
                If zlCommFun.Nvl(!项目序号) <> lng项目序号 Then
                    sin原X = 0
                    sin原Y = 0
                    lngRGB = Split(CStr(strEditors(i)), "||")(5)
                    lng项目序号 = zlCommFun.Nvl(!项目序号)
                End If
                If int心率引用 = 2 Then
                    If !项目序号 = -1 Then
                        blnLine = False
                    Else
                        blnLine = True
                    End If
                Else
                    blnLine = True
                End If
                
                '问题号:56886,李涛,2013-05-06,圆圈符号不穿中心
                bln符号 = Get符号(!重叠, !重叠项目, !项目序号, !符号, !部位, strEditors, !标记)
                lngWith = 0
                If bln符号 Then
                    lngWith = objDraw.TextWidth("○") / 4 / T_TwipsPerPixel.X
                End If
                
                If sin原X <> 0 And blnLine Then
                    Call DrawLine(lngDC, sin原X + T_DrawClient.列单位 / 2, sin原Y, !X坐标 + T_DrawClient.列单位 / 2 - lngWith, !Y坐标, PS_SOLID, intFine, lngRGB)
                End If
                If !断开 = 0 Then
                    sin原X = zlCommFun.Nvl(!X坐标, 0) + lngWith
                    sin原Y = zlCommFun.Nvl(!Y坐标, 0)
                Else
                    sin原X = 0
                End If
                
                If !项目序号 = gint体温 Then
                    If zlCommFun.Nvl(!复查) = 1 Then '复试合格
                        Call SetTextColor(lngDC, lngRGB)
                        Call GetTextRect(objDraw, !X坐标, !Y坐标 - T_DrawClient.行单位, "v", T_DrawClient.列单位, True, , sngScale)
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), "v", lngRGB)
                        Call DrawText(lngDC, "v", -1, T_LableRect, DT_CENTER)
                    End If
                End If
                
                dblMinValue = GetMaxMinValue(0, Val(zlCommFun.Nvl(!项目序号)), strEditors)
                dblMaxValue = GetMaxMinValue(1, Val(zlCommFun.Nvl(!项目序号)), strEditors)
                    
                If Not (Val(Nvl(!项目序号)) = gint体温 And Trim(Nvl(!数值)) = "不升") Then
                    dblvalue = Val(zlCommFun.Nvl(!数值))
                    If dblvalue > dblMaxValue Then
                        Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - T_DrawClient.行单位 * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, intFine, lngRGB, True)
                    ElseIf dblvalue < dblMinValue Then
                        Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + T_DrawClient.行单位 * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, intFine, lngRGB, True)
                    End If
                End If
            Else
                '体温的物理降温
                dblvalue = Split(!备注, ",")(0)
                SinX = Val(Split(Split(!备注, ",")(1), ";")(0))
                sinY = Val(Split(Split(!备注, ",")(1), ";")(1))
                T_Size.H = objDraw.TextHeight("○") / T_TwipsPerPixel.Y

                If Val(!数值) > Val(dblvalue) Then
                    '物理降温失败，画带箭头的红色实线，字符固定用○
                    'Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, SinX + T_DrawClient.列单位 / 2, sinY, PS_SOLID, intFine, RGB_RED, True)
                    '现在失败也为虚线(医院要求)
                    Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + (T_Size.H / 4), SinX + T_DrawClient.列单位 / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                ElseIf Val(!数值) < Val(dblvalue) Then
                    '物理降温成功，画红色虚线，字符固定用○
                    Call DrawLine(lngDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - (T_Size.H / 2), SinX + T_DrawClient.列单位 / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                End If
            End If
            .MoveNext
        Loop
    End With
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    '输出所有点的图形
    With rsPoint
        Do While Not .EOF
            str部位 = ""
            strTmp = ""
            For i = 0 To UBound(strEditors)
                If Split(CStr(strEditors(i)), "||")(0) = Val(zlCommFun.Nvl(!项目序号)) Then
                     Exit For
                End If
            Next i
            If zlCommFun.Nvl(!重叠) = 0 And zlCommFun.Nvl(!重叠项目) = "空" Then '未重叠的项目
                lngRGB = Split(CStr(strEditors(i)), "||")(5)
                If zlCommFun.Nvl(!项目序号) = -1 And int心率引用 = 2 Then lngRGB = RGB_RED
                str部位 = zlCommFun.Nvl(!部位)
                If str部位 = "" Then
                    Select Case lng项目序号
                        Case gint体温
                            str部位 = "腋温"
                        Case gint呼吸
                            str部位 = "自主呼吸"
                        Case Else
                            str部位 = ""
                    End Select
                End If
                strTmp = Split(CStr(strEditors(i)), "||")(4)
                strPic = ""
                strChar = ""
                Select Case zlCommFun.Nvl(!项目序号)
                    Case gint体温
                        strTmp = strTmp & String(3 - UBound(Split(strTmp, ",")), ",")
                        If str部位 = "口温" Then
                            strChar = Split(strTmp, ",")(0)
                        ElseIf str部位 = "腋温" Then
                            strChar = Split(strTmp, ",")(1)
                        ElseIf str部位 = "肛温" Then
                            strChar = Split(strTmp, ",")(2)
                        Else
                            strChar = Split(strTmp, ",")(3)
                        End If
                        If zlCommFun.Nvl(!标记) = 1 Then '物理降温符号
                            lngRGB = RGB_RED
                            strChar = "○"
                        Else
                            If strChar = "" Then strChar = "×"
                        End If
                    Case gint心率
                        strChar = IIf(strTmp = "", "Ο", strTmp)
                    Case gint脉搏
                        If str部位 = "起搏器" Then
                            strPic = "PACEMAKER"
                        Else
                            strChar = IIf(strTmp = "", "+", strTmp)
                        End If
                    Case gint呼吸
                        If str部位 = "自主呼吸" Then
                            strChar = IIf(strTmp = "", "*", strTmp)
                        Else
                            strPic = "BREATH"
                        End If
                    Case Else
                        strChar = strTmp
                End Select
                If Trim(zlCommFun.Nvl(!符号)) <> "" Then
                    strChar = Trim(zlCommFun.Nvl(!符号))
                    strPic = ""
                End If
                
                If !项目序号 = gint体温 And Trim(Nvl(!数值)) = "不升" And (mlng体温不升显示方式 = 0 Or mlng体温不升显示方式 = 1) Then
                    bln不升符号 = False
                Else
                    bln不升符号 = True
                End If
                                
                If strPic = "" And bln不升符号 Then
                    Call SetTextColor(lngDC, lngRGB)
                    Call GetTextRect(objDraw, !X坐标, !Y坐标, Trim(strChar), T_DrawClient.列单位, True, , sngScale)
                    T_LableRect.Left = T_LableRect.Left - 1
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(strChar), lngRGB)
                    Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                    'Debug.Print T_LableRect.Left & ";" & T_LableRect.Right
                Else
                    Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                        objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), True)
                End If
            
            Else  '展示重叠部位图标
                strPic = ""
                strChar = ""
                If zlCommFun.Nvl(!重叠项目) <> "空" Then '重叠=1的不做任何处理
                    lngCount = UBound(Split(zlCommFun.Nvl(!重叠项目), ","))
                    strTmp = zlCommFun.Nvl(!重叠项目)
                    If Trim(strTmp) <> "" Then
                        str部位 = zlCommFun.Nvl(!部位)
                        lngCount = lngCount + 2
                        strTmp = zlCommFun.Nvl(!项目序号) & "," & strTmp
                        If InStr(1, "," & strTmp & ",", ",1,") <> 0 Then

                            strSQL = "SELECT A.序号,A.标记符号,A.标记颜色" & vbNewLine & _
                                    " FROM 体温重叠标记 A," & vbNewLine & _
                                    "     (SELECT 上级序号, COUNT(*) 数量" & vbNewLine & _
                                    "     FROM 体温重叠标记" & vbNewLine & _
                                    "     WHERE 项目序号 IN (" & strTmp & ")" & vbNewLine & _
                                    "     GROUP BY 上级序号) B" & vbNewLine & _
                                    " WHERE A.重叠数目 = B.数量" & vbNewLine & _
                                    " AND A.序号 = B.上级序号 AND A.序号=[1]"
                        Else
                            strSQL = "Select A.序号, A.标记符号, A.标记颜色" & vbNewLine & _
                                "  From 体温重叠标记 A," & vbNewLine & _
                                "       (Select 上级序号, Count(1) 数量" & vbNewLine & _
                                "          from 体温重叠标记" & vbNewLine & _
                                "         where 项目序号 in (" & strTmp & ")" & vbNewLine & _
                                "         group by 上级序号) B" & vbNewLine & _
                                " Where A.重叠数目 = B.数量" & vbNewLine & _
                                "   And A.序号 = B.上级序号"
                        End If
                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "重叠", Val(str部位))
                        
                        If rsTmp.RecordCount > 0 Then
                            If IsNull(rsTmp!标记符号) Then
                                strPic = zlBlobRead(9, zlCommFun.Nvl(rsTmp!序号))
                            Else
                                strChar = Trim(zlCommFun.Nvl(rsTmp!标记符号))
                                lngRGB = Val(zlCommFun.Nvl(rsTmp!标记颜色, 0))
                            End If
                            If strPic = "" Then
                                Call SetTextColor(lngDC, lngRGB)
                                Call GetTextRect(objDraw, !X坐标 - 1, !Y坐标, Trim(strChar), T_DrawClient.列单位, True, , sngScale)
'                                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(strChar), lngRGB)
                                Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                            Else
                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                                    objDraw.ScaleX(!X坐标 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y坐标 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), False)
                                
                                Call FileSystem.Kill(strPic)
                            End If
                        End If
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '提取所有心率的信息
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    rsPoint.Filter = "项目序号=" & gint心率
    With rsPoint
        Do While Not .EOF
            str心率 = str心率 & "," & !X坐标 & ";" & !Y坐标
        .MoveNext
        Loop
    End With
    If str心率 <> "" Then str心率 = Mid(str心率, 2)
    
    ShowPointsNew = str心率
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub DrawBodyRecordItemNew(ByVal lngDC As Long, ByVal objDraw As Object, strValue() As String, ByVal rsItems As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long, _
    ByVal lngLeft As Long, ByVal intRepairRows As Integer, lngOutY As Long, Optional sngScale As Single = 1)
'-----------------------------------------------------------------------------------------------------------------------
'输出病人基本信息
'参数:lngDC 绘图对象的DC，strValue() 所有表格项目的信息 (格式（呼吸）:项目序号;名称;内容,部位||内容,部位/(其他) 项目序号;名称;内容||内容) 内容和部位组成的数组表示此项目有多少列
'    rsItems 所有体温表格护理项目, lngX 左边距,lngY上边距,lngLeft 右边距(可以绘图的最大右边距),intRepairRows 要打印表格项目的总行数
'出参:lngOutY 返回绘图后的上边距
'-----------------------------------------------------------------------------------------------------------------------
    Dim lngX1 As Long, lngY1 As Long, lngCurY As Long, lngCurX As Long
    Dim lngRowHeiht As Long, lngTestisHeight As Long, arrTestis
    Dim arrTmpString0() As String, arrTmpString1() As String
    Dim arrTmp() As String, arrText() As String, arrData
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer, j As Integer
    Dim int呼吸表格输出格式 As Integer
    Dim bln灌肠大便以分子分母显示 As Boolean
    Dim strTmp As String, strPart As String
    Dim strPic As String
    Dim blnValue As Boolean
    Dim intValue As Integer, int呼吸位置 As Integer
    Dim intRowCount As Integer
    Dim int频次 As Integer '记录频次
    Dim blnDataTrue As Boolean
    Dim lngColor As Long
    Dim intNum As Integer
    Dim blnOutText As Boolean '是否输出文本
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim sgnSize As Single
    Dim sngLen As Single, lngLen As Long
    Dim LPoint As T_LPoint
    Dim lngFont As Long, lngOldFont As Long
    Dim bln显示皮试 As Boolean
    
    If UBound(strValue) < 0 Then Exit Sub
    If IsEmpty(strValue) = True Then Exit Sub
    
    On Error GoTo Errhand
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
        intBold = 6
        intFine = 2
    Else
        msngTwips = 1
        intBold = 2
        intFine = 1
    End If
    
    lngCurY = lngY
    lngCurX = lngX
    blnValue = False
    intValue = 0
    int呼吸位置 = 0
    int呼吸表格输出格式 = zlDatabase.GetPara("呼吸表格输出", glngSys, 1255, 0)
    bln灌肠大便以分子分母显示 = (Val(zlDatabase.GetPara("灌肠后大便显示格式", glngSys, 1255, 0)) = 1)
    
    strPic = ""
    If InStr(1, strValue(0), ";") > 0 Then
        bln显示皮试 = IIf(Split(strValue(UBound(strValue)), ";")(0) = "-999", True, False)
        
        For intRow = LBound(strValue) To UBound(strValue)
            arrTmpString0 = Split(strValue(intRow), ";")
            arrTmpString1 = Split(arrTmpString0(2), "||")
            
            If intRepairRows > 0 And intRepairRows > intRowCount Then
            
                If arrTmpString0(0) = "3" Then '呼吸项目
                    '提取表格颜色
                    rsItems.Filter = 0
                    rsItems.Filter = "项目序号=" & gint呼吸
                    If rsItems.RecordCount > 0 Then
                        lngColor = Val(Nvl(rsItems!记录色, RGB_RED))
                    Else
                        lngColor = RGB_RED
                    End If
                    intRowCount = intRowCount + 1
                    arrTmpString1 = Split(arrTmpString0(2), "||")
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '表头
                            Call SetTextColor(lngDC, RGB_BLACK)
                            T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                            
                            LPoint.X = lngX
                            LPoint.Y = lngY
                            LPoint.W = T_DrawClient.刻度区域.Right - lngX
                            LPoint.H = mlngBreatheHeight
                            
                            Call DrawTabTextNew(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                            Call DrawLine(lngDC, lngX, lngY, lngX, lngY + mlngBreatheHeight, PS_SOLID, intBold, RGB_BLACK)
                            Call DrawLine(lngDC, lngX, lngY + mlngBreatheHeight, T_DrawClient.刻度区域.Right, _
                                lngY + mlngBreatheHeight, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                            Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY, T_DrawClient.刻度区域.Right, _
                                lngY + mlngBreatheHeight, PS_SOLID, intBold, RGB_BLACK)
                            lngX1 = T_DrawClient.刻度区域.Right
                            lngY1 = lngCurY
                        Else
                            arrTmpString1(intCOl) = arrTmpString1(intCOl) & String(1 - UBound(Split(arrTmpString1(intCOl), ",")), ",")
                            strTmp = Split(arrTmpString1(intCOl), ",")(0)
                            strPart = Split(arrTmpString1(intCOl), ",")(1)
                            If strPart = "" Then strPart = "自主呼吸"
                            strPic = ""
                            '打印呼吸值（间隔错开打印） 第一行始终在上面
                            If IsNumeric(strTmp) Then
                                If strPart = "自主呼吸" Then
                                    Call SetTextColor(lngDC, lngColor)
                                    T_Size.H = objDraw.TextHeight(strTmp) / T_TwipsPerPixel.Y
                                    T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                Else
                                    strPic = "BREATH"
                                End If
                                
                                If blnValue = False Then
                                    intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                                    blnValue = True
                                    int呼吸位置 = 2
                                End If
                                
                                If int呼吸表格输出格式 = 0 Then '顺序上下显示
                                    If intCOl Mod 2 = intValue Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.列单位
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.列单位
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), _
                                                vbPixels, vbTwips), objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                Else        '有数据时数据之间上下显示
                                    If int呼吸位置 = 2 Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.列单位
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.列单位
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.列单位 - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                   
                                    int呼吸位置 = int呼吸位置 + 1
                                    If int呼吸位置 > 2 Then int呼吸位置 = 1
                                End If
                                
                            End If
                            lngX1 = lngX1 + T_DrawClient.列单位
                        End If
                    Next intCOl
                    lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位
                    lngY1 = lngY + mlngBreatheHeight
                    
                    '画呼吸栏所有的列
                    For intCOl = 1 To T_BodyStyle.lng天数 * T_BodyStyle.lng监测次数
                        If intCOl Mod T_BodyStyle.lng监测次数 = 0 Then
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Else
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intFine, RGB_BLACK)
                        End If
                        lngX1 = lngX1 + T_DrawClient.列单位
                    Next intCOl
                    Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                    
                    '当前Y轴坐标
                    lngCurY = lngY1
                ElseIf arrTmpString0(0) <> "-999" Then '不是皮试结果
                    
                    rsItems.Filter = ""
                    rsItems.Filter = "序号=" & intRow
                    If rsItems.RecordCount > 0 Then
                        int频次 = CInt(zlCommFun.Nvl(rsItems!记录频次, 2))
                        If Val(Nvl(rsItems!项目表示)) = 4 Or IsWaveItem(Val(Nvl(rsItems!项目序号))) Then
                            If int频次 > 2 Then int频次 = 2 '汇总/波动项目频次只能是 1 、 2
                        End If
                        '活动项目检查是否存在数据，不存在就不打印此行
                        If zlCommFun.Nvl(rsItems!项目性质) = 2 Then
                            
                            If Trim(Replace(arrTmpString0(2), "||", "")) = "" Then
                                blnDataTrue = False
                            Else
                                blnDataTrue = True
                            End If
                        Else
                            blnDataTrue = True
                        End If
                    Else
                        blnDataTrue = False
                    End If
                    
                    If blnDataTrue = True Then
                        lngY1 = lngCurY
                        lngX1 = lngCurX
                        
                        '根据频次计算要打印的表格行数是否超出用户设置的表格行数
                        
                        intNum = 0
                        Select Case int频次
                            Case 1, 2, 6
                                intRowCount = intRowCount + 1
                            Case 3
                                intRowCount = intRowCount + 3
                            Case 4
                                intRowCount = intRowCount + 2
                            Case Else
                                intRowCount = intRowCount + 1
                        End Select
                        
                        If intRowCount > intRepairRows Then
                            intNum = intRowCount - intRepairRows
                            intRowCount = intRepairRows
                        End If
                        blnOutText = False
                        
                        For intCOl = 0 To UBound(arrTmpString1)
                            If intCOl = 0 Then '开始画表头信息包括标题的输出
                                Select Case int频次
                                    Case 1, 2, 6
                                        lngY1 = lngY1 + T_DrawClient.时间列单位
                                        lngRowHeiht = T_DrawClient.时间列单位 / 2
                                    Case 3
                                        lngY1 = lngY1 + T_DrawClient.时间列单位 * (3 - intNum)
                                        lngRowHeiht = (T_DrawClient.时间列单位 * (3 - intNum)) / 2
                                    Case 4
                                        lngY1 = lngY1 + T_DrawClient.时间列单位 * (2 - intNum)
                                        lngRowHeiht = (T_DrawClient.时间列单位 * (2 - intNum)) / 2
                                End Select

                                Call SetTextColor(lngDC, RGB_BLACK)
                                T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                            
                                LPoint.X = lngX1
                                LPoint.Y = lngY1 - lngRowHeiht * 2
                                LPoint.W = T_DrawClient.刻度区域.Right - lngX1
                                LPoint.H = lngRowHeiht * 2
                                Call DrawTabTextNew(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                
                                lngY1 = lngCurY
                                lngX1 = T_DrawClient.刻度区域.Right
                            Else  '开始进行画表格线
                                strTmp = CStr(arrTmpString1(intCOl))
                               
                                If InStr(1, strTmp, "-#") <> 0 Then
                                    If Not IsNumeric(Split(strTmp, "-#")(1)) Then
                                        lngColor = 0
                                    Else
                                        lngColor = Val(Split(strTmp, "-#")(1))
                                        strTmp = Split(strTmp, "-#")(0)
                                    End If
                                Else
                                    lngColor = 0
                                End If
                                
                                If strTmp = "*" And Val(arrTmpString0(0)) = gint大便 Then strTmp = "※"
                                
                                Call SetTextColor(lngDC, lngColor)
                                
                                T_Size.H = objDraw.TextHeight(strTmp) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                blnOutText = True
                                
                                If InStr(1, ",3,4,", "," & int频次 & ",") = 0 Then
                                    LPoint.X = lngX1
                                    LPoint.Y = lngCurY
                                    LPoint.W = T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                                    lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                                ElseIf int频次 = 3 Then
                                    LPoint.W = T_DrawClient.列单位 * T_BodyStyle.lng监测次数
                                    If intCOl Mod int频次 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位 * 2
                                        If intNum <> 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.列单位 * T_BodyStyle.lng监测次数
                                    ElseIf intCOl Mod int频次 = 2 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位
                                        If intNum > 1 Then blnOutText = False
                                    Else
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY
                                    End If
                                    
                                ElseIf int频次 = 4 Then
                                    LPoint.W = T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                    If intCOl Mod 4 = 3 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                    ElseIf intCOl Mod 4 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.时间列单位
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                    ElseIf intCOl Mod 2 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY
                                        lngX1 = lngX1 - T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                    ElseIf intCOl Mod 4 = 1 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY
                                        lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                    End If
                                End If
                                LPoint.H = T_DrawClient.时间列单位
                                
                                If blnOutText = True Then
                                    If AnsyGrade(Val(arrTmpString0(0)), strTmp, arrText) = True Then
                                        Call DrawAnsyGrade(lngDC, objDraw, arrText, LPoint, lngColor, bln灌肠大便以分子分母显示, sngScale)
                                    Else
                                        Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale)
                                    End If
                                End If
                   
                            End If
                        Next intCOl
                        
                        '画单元格竖线
                        If InStr(1, ",2,3,4,", "," & int频次 & ",") = 0 Then
                            lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                            lngY1 = lngCurY + T_DrawClient.时间列单位
                            For intCOl = 1 To int频次 * T_BodyStyle.lng天数
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod int频次 = 0, intBold, intFine), RGB_BLACK)
                                lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                            Next intCOl
                            Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                        ElseIf int频次 = 3 Then
                            intRowCount = intRowCount - (int频次 - intNum)
                            intValue = intRowCount
                            For i = 1 To 3 - intNum
                                lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位 * T_BodyStyle.lng监测次数
                                lngY1 = lngCurY + T_DrawClient.时间列单位
                                For intCOl = 1 To T_BodyStyle.lng天数
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.列单位 * T_BodyStyle.lng监测次数
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                
                                lngCurY = lngY1
                            Next i
                        ElseIf InStr(1, ",2,4,", "," & int频次 & ",") <> 0 Then
                            intRowCount = intRowCount - (int频次 / 2 - intNum)
                            intValue = intRowCount
                            For i = 1 To (int频次 / 2 - intNum)
                                lngY1 = lngCurY + T_DrawClient.时间列单位
                                lngX1 = T_DrawClient.刻度区域.Right + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                For intCOl = 1 To T_BodyStyle.lng天数 * 2
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, intBold, intFine), RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / 2)
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                lngCurY = lngY1
                            Next i
                        End If
                        
                        lngCurY = lngY1
                    End If
                End If
                
                intNum = 0
                arrTestis = Array()
                '皮试结果,只输出标题和内容，表格在不空行中处理。
                If arrTmpString0(0) = "-999" Then
                    lngY1 = lngCurY
                    lngX1 = lngCurX
                    int频次 = 1
                    
                    arrTestis = Array(0) '皮试结果分行显示，用于存放每一行皮试结果的最大高度
                    arrTestis(0) = Val(Format(T_DrawClient.时间列单位 * T_TwipsPerPixel.Y, "#0"))
                    
                    lngTestisHeight = Val(Format(T_DrawClient.时间列单位 * T_TwipsPerPixel.Y, "#0")) '皮试结果占用的最大高度
                    '得到皮试结果占用的总行数
                    LPoint.W = T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                    For intCOl = 1 To UBound(arrTmpString1)
                        intNum = 1
                        strTmp = CStr(arrTmpString1(intCOl))
                        If strTmp = "" Then strTmp = "-#"
                        arrTmp = Split(strTmp, ",")
                        T_Size.H = 0
                        If UBound(arrTmp) > UBound(arrTestis) Then
                            ReDim Preserve arrTestis(UBound(arrTmp))
                        End If
                        For i = LBound(arrTmp) To UBound(arrTmp)
                            strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") '皮试结果
                            If Trim(strTmp) <> "" Then
                                sgnSize = GetFontSize(objDraw, CStr(strTmp) & "L", LPoint.W, sngScale)
                                With frmTendFileRead.txtLength
                                    .Width = Val(Format(LPoint.W * T_TwipsPerPixel.X, "#0")) + IIf(blnPrinter, 12, 0)
                                    .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    .FontName = "宋体"
                                    .FontSize = sgnSize * sngScale
                                    .FontBold = False
                                    .FontItalic = False
                                End With
                                
                                arrData = GetData(frmTendFileRead.txtLength.Text, frmTendFileRead.txtLength)
                                '计算某一天皮试结果的高度
                                If Val(objDraw.TextHeight("刘") * (UBound(arrData) + 1)) < Val(Format(T_DrawClient.时间列单位 * T_TwipsPerPixel.Y, "#0")) Then
                                    lngRowHeiht = Val(Format(T_DrawClient.时间列单位 * T_TwipsPerPixel.Y, "#0"))
                                Else
                                    lngRowHeiht = objDraw.TextHeight("刘") * (UBound(arrData) + 1)
                                End If
                                T_Size.H = T_Size.H + lngRowHeiht
                                If Val(arrTestis(i)) < lngRowHeiht Then arrTestis(i) = lngRowHeiht
                                intNum = intNum + 1
                                If intRowCount + intNum > intRepairRows Then Exit For
                            End If
                        Next i
                        If lngTestisHeight < T_Size.H Then lngTestisHeight = T_Size.H
                    Next intCOl
                    Call ReleaseFontIndirect(objDraw)
                    lngTestisHeight = Val(Format(lngTestisHeight / T_TwipsPerPixel.Y, "#0"))
                    
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '开始画表头信息包括标题的输出
                            lngY1 = lngY1 + lngTestisHeight
                            lngRowHeiht = lngTestisHeight / 2
                               
                            Call SetTextColor(lngDC, RGB_BLACK)
                            T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                
                            LPoint.X = lngX1
                            LPoint.Y = lngY1 - lngTestisHeight
                            LPoint.W = T_DrawClient.刻度区域.Right - lngX1
                            LPoint.H = lngTestisHeight
                            Call DrawTabTextNew(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                            
                            lngY1 = lngCurY
                            lngX1 = T_DrawClient.刻度区域.Right
                        Else  '开始进行画表格线
                            intNum = 1
                            strTmp = CStr(arrTmpString1(intCOl))
                            If strTmp = "" Then strTmp = "-#"
                            LPoint.X = lngX1
                            LPoint.Y = lngCurY
                            LPoint.W = T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                            '开始计算是否需要换行
                            strPart = ""
                            
                            arrTmp = Split(strTmp, ",")
                            
                            For i = LBound(arrTmp) To UBound(arrTmp)
                                lngColor = Val(Split(arrTmp(i), "-#")(0))
                                '设置字体颜色
                                Call SetTextColor(lngDC, lngColor)
                                strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") '皮试结果
                                If Trim(strTmp) <> "" Then
                                    sgnSize = GetFontSize(objDraw, CStr(strTmp) & "L", LPoint.W, sngScale)
                                    '计算皮试结果输出的实际行数
                                    With frmTendFileRead.txtLength
                                        .Width = Val(Format(LPoint.W * T_TwipsPerPixel.X, "#0")) + IIf(blnPrinter, 12, 0)
                                        .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        .FontName = "宋体"
                                        .FontSize = sgnSize
                                        .FontBold = False
                                        .FontItalic = False
                                    End With
                                    arrData = GetData(frmTendFileRead.txtLength.Text, frmTendFileRead.txtLength)
                                    
                                    Set gstdSet = New StdFont
                                    gstdSet.Name = "宋体"
                                    gstdSet.Size = sgnSize
                                    gstdSet.Bold = False
                                    gstdSet.Italic = False
                                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                                    lngFont = CreateFontIndirect(T_Font)
                                    lngOldFont = SelectObject(lngDC, lngFont)
                                    lngY1 = LPoint.Y
                                    If Val((UBound(arrData) + 1) * objDraw.TextHeight("刘")) < Val(arrTestis(i)) Then
                                        LPoint.Y = LPoint.Y + (Val(arrTestis(i)) - ((UBound(arrData) + 1) * objDraw.TextHeight("刘"))) / T_TwipsPerPixel.Y / 2
                                    End If
                                    
                                    '开始输出内容
                                    For j = 0 To UBound(arrData)
                                        Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(arrData(j)), , False, , sngScale)
                                        Call DrawText(lngDC, CStr(arrData(j)), -1, T_LableRect, DT_CENTER)
                                        LPoint.X = lngX1
                                        LPoint.Y = LPoint.Y + Val(Format(objDraw.TextHeight("刘") / T_TwipsPerPixel.Y, "#0"))
                                    Next j
                                    LPoint.Y = lngY1 + Val(Format(Val(arrTestis(i)) / T_TwipsPerPixel.Y, "#0"))
                                    Call SelectObject(lngDC, lngOldFont)
                                    Call DeleteObject(lngFont)
                                    Call ReleaseFontIndirect(objDraw)
                                    
                                    intNum = intNum + 1
                                    If intRowCount + intNum > intRepairRows Then GoTo ErrNext
                                End If
                            Next i
ErrNext:
                            lngX1 = lngX1 + T_DrawClient.列单位 * (T_BodyStyle.lng监测次数 / int频次)
                        End If
                    Next intCOl
                End If
            End If
        Next intRow
        '画皮试结果表格线条
        arrData = Array()
        lngTestisHeight = 0
        For i = 0 To UBound(arrTestis)
            '说明有皮试结果
            If Val(arrTestis(i)) >= Val(Format(T_DrawClient.时间列单位 * T_TwipsPerPixel.Y, "#0")) Then
                ReDim Preserve arrData(UBound(arrData) + 1)
                arrData(UBound(arrData)) = Val(Format(Val(arrTestis(i)) / T_TwipsPerPixel.Y, "#0"))
                lngTestisHeight = lngTestisHeight + Val(arrData(UBound(arrData)))
            End If
        Next i
        lngX1 = lngCurX
        lngY1 = lngCurY
        For i = 0 To UBound(arrData)
            If i = 0 Then
                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, T_DrawClient.体温区域.Right, lngCurY, T_DrawClient.体温区域.Right, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, lngX1, lngCurY + lngTestisHeight, T_DrawClient.刻度区域.Right, lngCurY + lngTestisHeight, PS_SOLID, IIf(intRowCount + UBound(arrData) + 1 = intRepairRows, intBold, intFine), RGB_BLACK)
            End If
            For intCOl = 0 To T_BodyStyle.lng天数 - 1
                If intCOl = 0 Then
                    Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1 + Val(arrData(i)), T_DrawClient.体温区域.Right, lngY1 + Val(arrData(i)), PS_SOLID, IIf(intRowCount + i + 1 = intRepairRows, intBold, intFine), RGB_BLACK)
                Else
                    lngX1 = T_DrawClient.刻度区域.Right + (T_DrawClient.列单位 * T_BodyStyle.lng监测次数) * intCOl
                    Call DrawLine(lngDC, lngX1, lngY1, lngX1, lngY1 + Val(arrData(i)), PS_SOLID, intBold, RGB_BLACK)
                End If
            Next intCOl
            lngY1 = lngY1 + Val(arrData(i))
        Next i
        
        lngCurY = lngCurY + lngTestisHeight
        intRowCount = intRowCount + UBound(arrData) + 1
        
        '补空行
        If intRepairRows > 0 And intRepairRows > intRowCount Then
            intRowCount = intRowCount + 1
            For intRow = intRowCount To intRepairRows
                lngX1 = lngCurX
                lngY1 = lngCurY + T_DrawClient.时间列单位
                
                '空格每行1列
                For intCOl = 0 To T_BodyStyle.lng天数
                    If intCOl = 0 Then
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                    Else
                        
                        lngX1 = T_DrawClient.刻度区域.Right + (T_DrawClient.列单位 * T_BodyStyle.lng监测次数) * intCOl
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        If intCOl = T_BodyStyle.lng天数 Then
                            Call DrawLine(lngDC, T_DrawClient.刻度区域.Right, lngY1, T_DrawClient.体温区域.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        End If
                    End If
                Next intCOl
                lngCurY = lngY1
            Next intRow
        End If
        
        lngOutY = lngCurY + 2 * msngTwips
    Else
        lngOutY = lngCurY + 2 * msngTwips
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DrawBodyPageFooterNew(ByVal lngDC As Long, objDraw As Object, X As Long, Y As Long, ByVal LeftX As Long, ByVal intPageNo As Integer, _
    ByVal intBeginPage As Integer, Optional ByVal strInfo As String, Optional ByVal sngScale As Single = 1)
    '--------------------------------------------------------------------------------------------------------------------------------
    '功能：画出最底部说明
    '参数:intPageNO=页码
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim blnWeek As Boolean
    Dim blnPageNo As Boolean
    Dim blnOper As Boolean
    Dim blnPrintCurveInfo As Boolean
    Dim strNOPage As String
    Dim lngX As Long
    Dim blnPrinter As Boolean
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    blnPrintCurveInfo = (Val(zlDatabase.GetPara("体温单不打印曲线说明", glngSys, 1255, "0")) = 1)
    If blnPrintCurveInfo = False Then
        '打印体温说明信息
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextExtentPoint32(lngDC, strInfo, Len(strInfo), T_Size)
        Call GetTextRect(objDraw, X, Y, strInfo, 0, False, , sngScale)
        Call DrawText(lngDC, strInfo, -1, T_LableRect, DT_CENTER)
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 14
    Else
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 6
    End If
    
    blnWeek = (Val(zlDatabase.GetPara("打印周数", glngSys, 1255, "0")) = 1)
    blnPageNo = (Val(zlDatabase.GetPara("打印页号", glngSys, 1255, "1")) = 1)
    '67405:刘鹏飞,2013-11-25,添加"打印打印人"
    blnOper = (Val(zlDatabase.GetPara("打印打印人", glngSys, 1255, "0")) = 1)
    
    '打印页码
    '------------------------------------------------------------------------------------------------------------------
    If intPageNo > -1 And blnPageNo Then
        intPageNo = intPageNo + intBeginPage - 1
        strNOPage = "第   " & CStr(intPageNo) & "   页"
    End If
    
    If blnWeek Then
        If strNOPage = "" Then
            strNOPage = "第   " & CStr(intBeginPage) & "   周"
        Else
            strNOPage = strNOPage & "(第 " & CStr(intBeginPage) & " 周)"
        End If
    End If
    
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strNOPage, Len(strNOPage), T_Size)
    Call GetTextRect(objDraw, 0, Y, strNOPage, objDraw.Width / T_TwipsPerPixel.X, False, , sngScale)
    Call DrawText(lngDC, strNOPage, -1, T_LableRect, DT_CENTER)
    
    '输出打印人,即当前操作员姓名
    '------------------------------------------------------------------------------------------------------------------
    If blnOper = True Then
        strNOPage = "打印人:" & gstrUserName
    
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextExtentPoint32(lngDC, strNOPage, Len(strNOPage), T_Size)
        Call GetTextRect(objDraw, LeftX - objDraw.TextWidth(strNOPage) / T_TwipsPerPixel.X, Y, strNOPage, 0, False, , sngScale)
        Call DrawText(lngDC, strNOPage, -1, T_LableRect, DT_CENTER)
    End If

    Y = Y + T_Size.H / 2
    '--------------------------------------------------------------------------------------------------------------------------------
End Sub


Private Sub DrawDeviceCapsNew(ByVal lngDC As Long, ByVal objDraw As Object)
    Dim dblSureW As Double, dblSureH As Double
    '如果是打印预览,应按打印机的可打印的开始处开始预览
    dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
    dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    On Error Resume Next
    Call DrawRect(lngDC, (objDraw.Width * dblSureW) / T_TwipsPerPixel.X, (objDraw.Height * (1 - dblSureH)) / T_TwipsPerPixel.Y, _
    (objDraw.Width * (1 - dblSureW)) / T_TwipsPerPixel.X, objDraw.Height * dblSureH / T_TwipsPerPixel.Y, PS_DOT, 1, RGB_FleetGRAY)
End Sub

Private Sub CloseRs(RS As ADODB.Recordset)
    '功能：关闭Recordset对象
    On Error Resume Next
    If RS.State = ADODB.adStateOpen Then RS.Close
    Set RS = Nothing
End Sub

Private Sub ErrEmpty()
    msngTwips = 1
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
End Sub

Public Function GetFontSize(ByVal objDraw As Object, ByVal strTmp As String, sinWidth As Single, Optional sngScale As Single = 1) As Single
'---------------------------------------------------
'功能 处理表下表格字体输出
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, sgnSize As Single
    Dim stdSet As StdFont
    Dim sngD As Single
    Dim blnChage As Boolean
    Dim arrText, blnGrade As Boolean
    
    On Error GoTo Errhand
    blnChage = False
    
    sgnSize = 9
    objDraw.Font.Size = sgnSize * sngScale
    objDraw.Font.Name = "宋体"
    objDraw.Font.Bold = False
    objDraw.Font.Italic = False
    
    If objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X > sinWidth Then
ErrGoTo:
        sngD = Round((objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X - sinWidth) / sinWidth, 4)
        If sngD > 0 Then
            sgnSize = CInt(Round((1 - sngD), 2) * sgnSize - 0.5)
            If sgnSize < 7 Then sgnSize = 7
            objDraw.Font.Size = sgnSize * sngScale
            If (objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X) > sinWidth And sgnSize > 7 Then GoTo ErrGoTo
        End If
    Else
        sgnSize = 9
    End If
    
    objDraw.Font.Size = sgnSize * sngScale
    
    GetFontSize = sgnSize
    Exit Function
Errhand:
    objDraw.Font.Size = 9 * sngScale
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawTabTextNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmp As String, ByVal nCount As Long, ByVal wFormat As Long, LPoint As T_LPoint, Optional sngScale As Single = 1, Optional ByVal bytCenterType As Byte = 2, Optional sgnFontSize As Single = 9)
'---------------------------------------------------
'功能 处理表下表格字体输出
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, sgnSize As Single
    Dim stdSet As StdFont
    Dim sngD As Single
    Dim blnChage As Boolean
    Dim arrText, blnGrade As Boolean
    Dim arrData, i As Integer
    Dim lngFontHeight As Long
    Dim lngCurX As Long, lngCurY As Long
    
    On Error GoTo Errhand
    blnChage = False
    
    sgnSize = sgnFontSize
    objDraw.Font.Size = sgnSize * sngScale
    objDraw.Font.Name = "宋体"
    objDraw.Font.Bold = False
    objDraw.Font.Italic = False
    
    If objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X > LPoint.W Then
ErrGoTo:
        sngD = Round((objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X - LPoint.W) / LPoint.W, 4)
        If sngD > 0 Then
            sgnSize = Int(Round((1 - sngD), 2) * sgnSize - 0.5)
            If sgnSize < 7 Then sgnSize = 7
            objDraw.Font.Size = sgnSize * sngScale
            If (objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X) > LPoint.W And sgnSize > 7 Then GoTo ErrGoTo
            blnChage = True
        End If
    Else
        sgnSize = sgnFontSize
    End If
    
    arrData = Array()
    If blnChage = True Then
        With frmTendFileRead.txtLength
            .Width = Val(Format(LPoint.W * T_TwipsPerPixel.X, "#0")) + IIf(TypeName(objDraw) = "Printer", 12, 0)
            .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
            .FontName = "宋体"
            .FontSize = sgnSize * sngScale
            .FontBold = False
            .FontItalic = False
        End With
        arrData = GetData(frmTendFileRead.txtLength.Text, frmTendFileRead.txtLength)
        lngFontHeight = Val(Format((objDraw.TextHeight("刘") / T_TwipsPerPixel.Y) * (UBound(arrData) + 1), "#0"))
    Else
        lngFontHeight = Val(Format(objDraw.TextHeight("刘") / T_TwipsPerPixel.Y, "#0"))
    End If
    
    Set stdSet = New StdFont
    stdSet.Name = "宋体"
    stdSet.Size = sgnSize * sngScale
    stdSet.Bold = False
    stdSet.Italic = False
    Call SetFontIndirect(stdSet, lngDC, objDraw)
    lngFont = CreateFontIndirect(T_Font)
    lngOldFont = SelectObject(lngDC, lngFont)
    
    Select Case bytCenterType
        Case 1 '居上
            lngCurY = LPoint.Y
        Case 2 '居中
            If lngFontHeight < LPoint.H Then
                lngCurY = LPoint.Y + (LPoint.H - lngFontHeight) / 2
            Else
                lngCurY = LPoint.Y
            End If
        Case 3 '居下
            If lngFontHeight < LPoint.H Then
                lngCurY = LPoint.Y + (LPoint.H - lngFontHeight)
            Else
                lngCurY = LPoint.Y
            End If
    End Select
    lngCurX = LPoint.X
    
    '输出字体
    If UBound(arrData) > 0 Then
        For i = 0 To UBound(arrData)
            Call GetTextRect(objDraw, lngCurX, lngCurY, CStr(arrData(i)), , False, , sngScale)
            Call DrawText(lngDC, CStr(arrData(i)), nCount, T_LableRect, wFormat)
            lngCurY = lngCurY + Val(Format(objDraw.TextHeight("刘") / T_TwipsPerPixel.Y, "#0"))
        Next i
    Else
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, LPoint.W, False, , sngScale)
        Call DrawText(lngDC, strTmp, nCount, T_LableRect, wFormat)
    End If
    
    Call SelectObject(lngDC, lngOldFont)
    Call DeleteObject(lngFont)
    Call ReleaseFontIndirect(objDraw)
    Set stdSet = Nothing
    Exit Sub
Errhand:
    objDraw.Font.Size = 9 * sngScale
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub DrawAnsyGrade(ByVal lngDC As Long, ByVal objDraw As Object, arrText() As String, LPoint As T_LPoint, ByVal lngColor As Long, Optional ByVal blnFormat As Boolean = False, Optional sngScale As Single = 1)
'---------------------------------------------------
'功能 大便次数输出
'说明 AnsyGrade=True才能调用此函数
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer
    Dim stdSet As StdFont, stdOldset As StdFont
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    Dim lngMaxWidth As Long
    
    On Error GoTo Errhand
    
    If UBound(arrText) < 2 Then Exit Sub
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        '60529:刘鹏飞,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
            strTmp = str1 & str2
        Else
            strTmp = str1 & str3
        End If
    Else
        strTmp = str1 & str2 & "/" & str3
    End If
    intSize = 9
    objDraw.Font.Size = 9 * sngScale
    Set stdSet = New StdFont
    stdSet.Name = "宋体"
    stdSet.Size = intSize * sngScale
    stdSet.Bold = False
    Set stdOldset = stdSet
    
    LPoint.Y = LPoint.Y + Val(Format(LPoint.H / 2, "#0"))
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , sngScale)
    '输出左边
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        Call DrawText(lngDC, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + msngTwips
        Call ReleaseFontIndirect(objDraw)
    Else
        lngX = T_LableRect.Left
    End If
    
    If blnFormat = True Then '分子分母显示
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        '60529:刘鹏飞,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
            lngMaxWidth = objDraw.TextWidth(str2) / T_TwipsPerPixel.X
        Else
            lngMaxWidth = objDraw.TextWidth(str3) / T_TwipsPerPixel.X
        End If
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = intSize * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str2) / T_TwipsPerPixel.X) \ 2
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        If T_LableRect.Top < LPoint.Y - Val(Format(LPoint.H / 2, "#0")) Then T_LableRect.Top = LPoint.Y - Val(Format(LPoint.H / 2, "#0"))
        T_LableRect.Bottom = LPoint.Y + Val(Format(LPoint.H / 2, "#0"))
        Call DrawText(lngDC, str2, -1, T_LableRect, 0)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '画横线
        objDraw.Font.Size = 9 * sngScale
        Call DrawLine(lngDC, lngX, lngY, lngX + lngMaxWidth, lngY)
        '输出分母
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        lngY = lngY
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str3) / T_TwipsPerPixel.X) \ 2
        T_LableRect.Top = lngY
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = intSize * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        Call DrawText(lngDC, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
    Else
        If str1 <> "" Then
            '输出上标
            intSize = 7
            objDraw.Font.Size = intSize * sngScale
            Set stdSet = New StdFont
            stdSet.Name = "宋体"
            stdSet.Size = intSize * sngScale
            Call SetFontIndirect(stdSet, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < (LPoint.Y - Val(Format(LPoint.H / 2, "#0"))) Then T_LableRect.Top = (LPoint.Y - Val(Format(LPoint.H / 2, "#0")))
            Call DrawText(lngDC, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            Call ReleaseFontIndirect(objDraw)
            '输出后半部分
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDC, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        Else
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            Call DrawText(lngDC, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        End If
    End If
    
    objDraw.Font.Size = 9 * sngScale
    Set stdSet = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Public Function GetXCoordinateNew(ByVal strInput As String, ByVal strBeginDate As String, Optional ByVal bln坐标 As Boolean = True) As String

    '根据时间得到X坐标或根据X坐标转换为时间范围
    Dim SinX   As Single

    Dim intDO  As Integer, intMax As Integer

    Dim intDay As Integer, intTime As Integer

    Dim strDay As String, strTime As String
    
    Dim int监测次数 As Integer

    On Error GoTo Errhand
    
    int监测次数 = T_BodyStyle.lng监测次数
    
    If bln坐标 Then
        '第一天是0,第七天是6
        strDay = Split(strInput, " ")(0)

        If InStr(1, strInput, " ") <> 0 Then
            strTime = Split(strInput, " ")(1)
        Else
            strTime = "00:00:00"
        End If

        intDay = DateDiff("d", CDate(strBeginDate), CDate(strInput))
        
        '得到当天的刻度
        intMax = int监测次数 - 1

        For intDO = 0 To intMax

            If strTime >= Split(gvarTime(intDO), ",")(0) And strTime <= Split(gvarTime(intDO), ",")(1) Then
                intTime = intDO
                Exit For
            End If
        Next
        
        '计算得到X坐标(每天6列,以列数*列单位得到坐标)
        SinX = Format(T_DrawClient.体温区域.Left + (T_DrawClient.列单位 * (intDay * int监测次数 + intTime)), "#0.0")
        GetXCoordinateNew = SinX
    Else
        '计算得到相差多少个刻度
        SinX = Val(strInput)
        intTime = (SinX - T_DrawClient.体温区域.Left) \ T_DrawClient.列单位
        intDay = intTime \ int监测次数
        intTime = intTime Mod int监测次数
        
        strDay = Format(DateAdd("d", intDay, strBeginDate), "yyyy-MM-dd")
        strTime = gvarTime(intTime)
        GetXCoordinateNew = strDay & " " & Split(gvarTime(intTime), ",")(0) & "," & strDay & " " & Split(gvarTime(intTime), ",")(1)
    End If
    
    Exit Function

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetCurveDateNew(ByVal intCOl As Integer, _
                             ByVal dtBeginDateTime As Date, _
                             Optional ByVal intHourBegin As Integer = 4) As String

    '-------------------------------------------------------------------------------------
    '功能:根据列计算出时间范围
    '参数 intCol 当前列    dtBeginDateTime 起始时间
    '返回格式为:开始时间;终止时间
    '-------------------------------------------------------------------------------------
    Dim varTime  As Variant

    Dim intDays  As Integer

    Dim strBegin As String

    Dim strEnd   As String

    Dim lngLoop  As Long

    Dim lng列号  As Long
    
    Dim int监测次数 As Integer

    On Error GoTo Errhand
    
    GetCurveDateNew = -1
    
    int监测次数 = T_BodyStyle.lng监测次数
    
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin, int监测次数, T_BodyStyle.lng时间间隔)
    
    '结算当前列和开始时间 相差的天数,并重新计算列的开始时间
    intDays = (intCOl - 1) \ int监测次数
    strBegin = DateAdd("d", intDays, Int(dtBeginDateTime))
    strEnd = strBegin
    
    '结算列所在的时间范围
    lng列号 = (intCOl - 1) Mod int监测次数
    
    strBegin = Format(strBegin & " " & Split(varTime(lng列号), ",")(0), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(strEnd & " " & Split(varTime(lng列号), ",")(1), "YYYY-MM-DD HH:mm:ss")

    GetCurveDateNew = strBegin & ";" & strEnd

    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function



Public Function GetCurveColumnNew(ByVal dtDateTime As Date, _
                               ByVal dtBeginDateTime As Date, _
                               Optional ByVal intHourBegin As Integer = 4) As Integer

    '******************************************************************************************************************
    '功能： 从时间计算出列
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim varTime As Variant

    Dim strTmp  As String

    Dim intDays As Integer

    Dim intLoop As Integer
    
    Dim int监测次数 As Integer
    
    Dim int天数 As Integer
    On Error GoTo Errhand
    
    GetCurveColumnNew = -1
    
    int监测次数 = T_BodyStyle.lng监测次数
    int天数 = T_BodyStyle.lng天数
    
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin, T_BodyStyle.lng监测次数, T_BodyStyle.lng时间间隔)

    '计算当前天的时间是在一天的第几格位置上
    strTmp = Format(dtDateTime, "HH:mm:ss")
    
    For intLoop = 0 To int监测次数
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next
    
    If intLoop < int监测次数 Then
        '计算当天在当前体温单页上是第几天（0表示第一天；1表示第二天.....）
        intDays = DateDiff("d", Int(dtBeginDateTime), Int(dtDateTime))
        GetCurveColumnNew = intDays * int监测次数 + intLoop + 1
    End If
    
    Exit Function

Errhand:

    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CalcMinMaxColNew(ByVal strDate As String, _
                              MinCol As Integer, _
                              MaxCol As Integer) As Boolean

    '------------------------------------------------------------------------------------------------------------------
    '功能： 获得最小最大时间范围
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String

    Dim dtTmp      As Date

    Dim strTmp     As String
    
    'If mvarEdit = False Then Exit Function
    
    aryValue = Split(strDate, ";")
    gintHourBegin = T_BodyStyle.lng开始时点
    MinCol = GetCurveColumnNew(CDate(aryValue(0)), CDate(aryValue(0)), gintHourBegin)
    MaxCol = GetCurveColumnNew(CDate(aryValue(1)), CDate(aryValue(0)), gintHourBegin)
    
End Function


Public Sub CreatePolyNew(rsPoint As ADODB.Recordset, ByVal objDraw As Object, ByVal lngDC As Long, ByVal strBeginDate As String, ByVal str心率坐标 As String, ByVal bln脉搏共用 As Boolean)

'rsPoint 记录集 必须包括  项目序号,X坐标,Y坐标
    Dim arrData, arrPt
    Dim bln区域 As Boolean      '不是区域就是点对点,心率必须对应脉搏才能形成区域或连线
    Dim bln左 As Boolean, bln右 As Boolean, bln当前 As Boolean, bln断开 As Boolean, bln有效 As Boolean
    Dim intDO   As Integer, intMax As Integer             'intLast记录最后一个有效的心率
    Dim recttmp As RECT, SinX As Single, sinY As Single, sin左X As Single, sin右X As Single
    Dim str当前 As String, str左 As String, str右 As String
    Dim str脉搏 As String, str心率 As String
    Dim PtIn脉搏() As POINTAPI
    Dim PtIn心率() As POINTAPI
    Dim lng填充方式 As Long

    Dim PtInPoly() As POINTAPI, intCOl As Integer, intCols As Integer, intCount As Integer
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngWith As Long
    Dim i As Integer, j As Integer
    On Error GoTo Errhand

    '1个心率对应1至3个脉搏,脉搏必须在每一天都有值,否则不形成区域
    '形成的区域集合必须是连续的,所以,先装入脉搏,再倒起装入心率,形成完整的一个区域
    '由点组成的封闭区域,在DrawPoly中完成封闭区域的连线
    
    lng填充方式 = Val(zlDatabase.GetPara("脉搏短绌填充方式", glngSys, 1255, "0"))
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 4
        intFine = 4
        blnPrinter = True
    Else
        intBold = 2
        intFine = 1
        blnPrinter = False
    End If
    
    rsPoint.Sort = "项目序号,时间"
    arrData = Split(str心率坐标, ",")
    intMax = UBound(arrData)
    

'
    For intDO = 0 To intMax

        SinX = Val(Split(arrData(intDO), ";")(0))
        sinY = Val(Split(arrData(intDO), ";")(1))
        '将当前心率加入区域集合
        intCount = intCount + 1
        ReDim Preserve PtInPoly(intCount)
        str心率 = str心率 & "," & SinX + T_DrawClient.列单位 / 2 & ";" & sinY
        
        '如果左边有,则与左列的脉搏连线
        If Not bln区域 Then
            bln左 = False
            rsPoint.Filter = "项目序号=" & gint脉搏 & " And X坐标<" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
               rsPoint.Sort = "X坐标 DESC"
                bln断开 = (rsPoint!断开 = 1)
                If Not bln断开 Then
                    rsPoint.Sort = "X坐标 DESC"
                    sin左X = rsPoint!X坐标
                
                    '根据当前坐标获取时间
                    str左 = GetXCoordinateNew(sin左X, strBeginDate, False)
                    str当前 = GetXCoordinateNew(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                    '当前点和前一时间点间隔一天没有数据就断开
                    If DateDiff("d", CDate(Split(str左, ",")(0)), CDate(Split(str当前, ",")(0))) < 2 Then
                        recttmp.Left = rsPoint!X坐标
                        recttmp.Top = rsPoint!Y坐标
                        '将左脉搏加入区域集合
                        intCount = intCount + 1
                        ReDim Preserve PtInPoly(intCount)
                        str脉搏 = str脉搏 & "," & rsPoint!X坐标 + T_DrawClient.列单位 / 2 & ";" & rsPoint!Y坐标
                        bln左 = True
                    End If
                End If
            End If
        End If
        
        bln当前 = False
        '缺省是和当前列的脉搏连线
        rsPoint.Filter = "项目序号=" & gint脉搏 & " And X坐标=" & Val(Split(arrData(intDO), ";")(0))
        bln当前 = (rsPoint.RecordCount <> 0)

        If bln当前 Then
            If Not bln左 Then
                recttmp.Left = rsPoint!X坐标
                recttmp.Top = rsPoint!Y坐标
            End If

            bln断开 = (rsPoint!断开 = 1)

            '将当前脉搏加入区域集合
            If Not bln区域 Then
                intCount = intCount + 1
                ReDim Preserve PtInPoly(intCount)
                str脉搏 = str脉搏 & "," & rsPoint!X坐标 + T_DrawClient.列单位 / 2 & ";" & rsPoint!Y坐标
            End If
        End If

        bln右 = False

        If Not bln断开 Then
            rsPoint.Filter = "项目序号=" & gint脉搏 & " And X坐标>" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
                rsPoint.Sort = "X坐标 ASC"
                sin右X = rsPoint!X坐标
            
                '根据当前坐标获取时间
                str右 = GetXCoordinateNew(sin右X, strBeginDate, False)
                str当前 = GetXCoordinateNew(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                '当前点和下一时间点间隔一天没有数据就断开
                If DateDiff("d", CDate(Split(str当前, ",")(0)), CDate(Split(str右, ",")(0))) < 2 Then
                    bln右 = True
                    recttmp.Right = rsPoint!X坐标
                    recttmp.Bottom = rsPoint!Y坐标
                    '将右脉搏加入区域集合
                    intCount = intCount + 1
                    ReDim Preserve PtInPoly(intCount)
                    str脉搏 = str脉搏 & "," & rsPoint!X坐标 + T_DrawClient.列单位 / 2 & ";" & rsPoint!Y坐标
                End If
            End If
        End If
        
        '先把左边封闭
        If bln区域 = False Then
            If bln当前 = True Then
                '与左列或当前列的脉搏连线
                Call DrawLine(lngDC, recttmp.Left + T_DrawClient.列单位 / 2, recttmp.Top, SinX + T_DrawClient.列单位 / 2, sinY, PS_SOLID, intFine, RGB_RED)
            End If

            bln区域 = (bln左 Or bln右) And bln当前
        End If
        
        '找到右边的封闭区进行连线
        If bln区域 Then
            bln区域 = False

            If bln右 = True Then
                '判断当前心率对应的下一个脉搏和下一个心率X坐标是否相等,不相等就封闭区域
                If intDO < intMax Then
                    If recttmp.Right = Val(Split(arrData(intDO + 1), ";")(0)) Then
                        bln区域 = True
                    End If
                End If
            End If
            
            
            If Not bln区域 Then
                '组织区域,从脉搏开始,然后转到心率(心率从最后开始,再回到之前的心率,再回到第一个脉搏,形成封闭区域)
                intCount = 1
                str脉搏 = Mid(str脉搏, 2)
                arrPt = Split(str脉搏, ",")
                intCols = UBound(arrPt)
                i = 0
                ReDim Preserve PtIn脉搏(intCols)
                For intCOl = 0 To intCols
                    PtIn脉搏(i).X = Split(arrPt(intCOl), ";")(0)
                    PtIn脉搏(i).Y = Split(arrPt(intCOl), ";")(1)
                    i = i + 1
                 Next
                
           
                For intCOl = 0 To intCols
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

                str心率 = Mid(str心率, 2)
                arrPt = Split(str心率, ",")
                intCols = UBound(arrPt)
                
                i = 0
                ReDim Preserve PtIn心率(intCols)
                For intCOl = 0 To intCols
                    PtIn心率(i).X = Split(arrPt(intCOl), ";")(0)
                    PtIn心率(i).Y = Split(arrPt(intCOl), ";")(1)
                    i = i + 1
                Next

                For intCOl = intCols To 0 Step -1
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

'                '加上起点形成封闭区域
                ReDim Preserve PtInPoly(intCount)
                PtInPoly(intCount).X = PtInPoly(1).X
                PtInPoly(intCount).Y = PtInPoly(1).Y
                
                '填充该区域
                Call DrawPoly(lngDC, PtInPoly, lng填充方式, UBound(Split(str脉搏, ",")) + 1)
                '76697,LPF,处理66628问题产生的错误
                '问题号：66628,修改人：李涛,脉搏短轴共用时不填充图形，直接脉搏和心率点进行连线。
                If lng填充方式 = 2 And bln脉搏共用 Then
                    i = 0: j = 0
                    For i = 0 To UBound(PtIn心率)
                        For j = 0 To UBound(PtIn脉搏)
                            If PtIn脉搏(j).X = PtIn心率(i).X Then
                                Call DrawLine(lngDC, PtIn脉搏(j).X, PtIn脉搏(j).Y, PtIn心率(i).X, PtIn心率(i).Y, PS_SOLID, intFine, RGB_RED)
                            End If
                        Next
                    Next
                End If
            End If
          
        End If

        If Not bln区域 Then
            intCount = 0
            str脉搏 = ""
            str心率 = ""
            ReDim Preserve PtInPoly(intCount)
            ReDim Preserve PtIn脉搏(intCount)
            ReDim Preserve PtIn心率(intCount)
        End If

    Next
    
    rsPoint.Filter = ""

    Exit Sub

Errhand:

    If ErrCenter() = 1 Then

        Resume

    End If

End Sub


Public Function GetCanvasCenterNew(ByVal dtBegin As Date, ByVal dtEnd As Date, ByVal dtBeginDate As Date, ByVal SinX As Single) As Boolean
'---------------------------------------------------------
'功能:判断该时间点是否是中间值
'参数:dtbegin:被比较的时间段.  dtend:要比较的时间段 . dtBeginDate 本页体温单的开始时间 .sinx当前点的X坐标
'---------------------------------------------------------
    Dim blnTrue As Boolean
    Dim strTime As String, strTmp As String
    Dim intDay As Integer, intTime As Integer, strDay As String
    Dim int监测次数 As Integer
    Dim int时间间隔 As Integer
    
    int监测次数 = T_BodyStyle.lng监测次数
    int时间间隔 = T_BodyStyle.lng时间间隔
    
    intTime = (SinX - T_DrawClient.体温区域.Left) \ T_DrawClient.列单位
    intDay = intTime \ int监测次数
    intTime = intTime Mod int监测次数
        
    strDay = Format(DateAdd("d", intDay, dtBeginDate), "yyyy-MM-dd")
    strTmp = strDay & " " & Split(gvarTime(intTime), ",")(0) & "," & strDay & " " & Split(gvarTime(intTime), ",")(1)
    
    If intTime <= UBound(gvarTime) Then
        If gintHourBegin + intTime * int时间间隔 = 24 Then
            strTime = Format(Format(strDay, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(strDay, "YYYY-MM-DD") & " " & gintHourBegin + intTime * 4 & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If CDate(strTime) > CDate(Split(strTmp, ",")(1)) Then strTime = Format(Split(strTmp, ",")(1), "YYYY-MM-DD HH:mm:ss")
    
    If Abs(DateDiff("s", Format(dtBegin, "YYYY-MM-DD HH:mm:ss"), Format(strTime, "YYYY-MM-DD HH:mm:ss"))) > _
        Abs(DateDiff("s", Format(dtEnd, "YYYY-MM-DD HH:mm:ss"), Format(strTime, "YYYY-MM-DD HH:mm:ss"))) Then
        blnTrue = True
    End If

    GetCanvasCenterNew = blnTrue
End Function

Public Function RetrunEndTimeNew(ByVal dtBegin As Date, ByVal dtEnd As Date, Optional ByVal intHourBegin As Integer = 4) As Date
'**********************************************************************************
'功能：检查体温单终止时间和开始时间是否在同一单元格，如果在同一单元格需要将终止时间移到下一单元格
'参数：strBegin 体温单开始时间,strEnd 体温单终止时间(病人出院时间)
'返回值：体温单终止时间
'**********************************************************************************
'需求：对于病人出院和入院时间在同一个格子，既要录入入院体温，也要录入出院体温，将出院体温录入到下一个格子。

    Dim varTime As Variant
    Dim intLoop As Integer, strTmp As String
    Dim intBegin As Integer, intEnd As Integer
    Dim strEnd As String
    Dim int监测次数 As Integer
    Dim int时间间隔 As Integer
    
    int监测次数 = T_BodyStyle.lng监测次数
    int时间间隔 = T_BodyStyle.lng时间间隔
    RetrunEndTimeNew = dtEnd
    If Format(dtBegin, "YYYY-MM-DD") <> Format(dtEnd, "YYYY-MM-DD") Then Exit Function
    '初始化时间范围划分
    Call InitDateTimeRange(varTime, intHourBegin, int监测次数, int时间间隔)
    '1/计算开始时间和终止时间在第几个格子
    strTmp = Format(dtBegin, "HH:mm:ss")
    For intLoop = 0 To int监测次数
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            intBegin = intLoop
            Exit For
        End If
    Next
    strTmp = Format(dtEnd, "HH:mm:ss")
    For intLoop = 0 To int监测次数
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            intEnd = intLoop
            Exit For
        End If
    Next
    '2 不在同一列就退出
    If intBegin <> intEnd Then Exit Function
    If intEnd > int监测次数 - 1 Then Exit Function
    '3 完成终止时间的重新赋值
    If intEnd > int监测次数 - 2 Then
        strEnd = Format(DateAdd("D", 1, dtEnd), "YYYY-MM-DD") & " " & Format(Split(varTime(0), ",")(1), "HH:mm:ss")
    Else
        strEnd = Format(dtEnd, "YYYY-MM-DD") & " " & Format(Split(varTime(intEnd + 1), ",")(1), "HH:mm:ss")
    End If
    
    RetrunEndTimeNew = CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss"))
End Function



Public Sub OutPutTextNew(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal lngDC As Long, ByVal rsNote As ADODB.Recordset, ByVal mstrBeginDate As String, Optional ByVal sngScale As Single = 1)

    'rsDrawItems  记录项目的最大坐标 单位值等基本信息
    'rsNote 所有说的信息
    'mstrBeginDate 体温单每页开始时间
    '输出以下信息:入院,入科,转科,出院,手术分娩,未记说明,上标说明及出生
    '未记说明及上标说明,在没有入出转手术分娩及出生的信息时,打印在42-40之间;否则从40开始向下打印
    '除未记说明及上标说明外,入出转等信息当一个刻度发生多个时,依次写入各个刻度中,如其它刻度也有信息,顺移
    Dim lngMaxX As Long     '体温单最大X坐标
    Dim lngX    As Long '第一列的X坐标
    Dim lngY    As Long 'Y坐标
    Dim lngY1   As Long '40 度固定坐标
    Dim i       As Integer, j As Integer
    Dim X, Y As Long '输出内容时的坐标
    Dim strComment    As String, strText As String
    Dim intAscCharNum As Integer
    Dim rsTemp  As New ADODB.Recordset
    Dim strDate As String
    Dim bln上标 As Boolean
    Dim bln事件显示规则 As Boolean '参数:体温标志按顺序当天排列
    Dim blnLessenSize As Boolean  '参数:体温标志超出42刻度缩小字体显示
    Dim arrX, arrCurX
    Dim blnBigSize As Boolean '是否以九号字体显示
    Dim lngFont As Long, lngOldFont As Long
    Dim dblCurveHeight As Double  '体温单42到40的高度
    Dim dblHeight As Double
    Dim blnCenter As Boolean
    
    On Error GoTo Errhand
    
    arrX = Array(): arrCurX = Array()
    bln事件显示规则 = (Val(zlDatabase.GetPara("体温标志按顺序当天排列", glngSys, 1255, 0)) = 1)
    blnLessenSize = (Val(zlDatabase.GetPara("体温标志超出40刻度缩小字体显示", glngSys, 1255, 0)) = 1)
    
    lngMaxX = T_DrawClient.体温区域.Right - T_DrawClient.列单位
    dblCurveHeight = Format(GetYCoordinate(objDraw, rsDrawItems, gint体温, 40) - GetYCoordinate(objDraw, rsDrawItems, gint体温, 42), "#0.00")
    
    rsNote.Filter = "禁用<>1"

    '首先检查更新入出转，手术分娩信息
    If rsNote.RecordCount = 0 Then Exit Sub
    
    '70228:刘鹏飞,2014-02-18,体温自动标识显示修改。
    '规则：
    '   1、体温标志按顺序当天排列=True，每页按循序排列依次显示当天标记，一个时点最多显示两个标记(缩小字体处理).如果在当天显示不完，剩余标记不进行显示。
    '   2、体温标志按顺序当天排列=False,每页按顺序排列依次显示，一个时点只显示一个。如果在本页最后一列还显示不完，则在最后一列纵向显示剩余标记。
    rsNote.Sort = "X坐标,时间,项目序号"
    lngX = rsNote!X坐标
    j = 1
    With rsNote
        Do While Not .EOF
            If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                If Not (!类型 = 2 Or !类型 = 99) Then
                    '体温标志按顺序当天排列
                    If bln事件显示规则 = True Then
                        If Val(!X坐标) > lngX Then j = 1
                        If lngX <= lngMaxX Then
                            strDate = Format(Split(GetXCoordinateNew(lngX, mstrBeginDate, False), ",")(0), "YYYY-MM-DD")
                            If CDate(strDate) > CDate(Format(!时间, "YYYY-MM-DD")) Then
                                !禁用 = 1
                            End If
                        Else
                            lngX = lngMaxX
                            !禁用 = 1
                        End If
                    Else
                        '控制x坐标，如果超过体温最大x坐标，则进行校正
                        If lngX > lngMaxX Then lngX = lngMaxX
                    End If
                    
                    !打印X坐标 = IIf(lngX <= Val(!X坐标), !X坐标, lngX)
                    !高度 = GetFontHeight(objDraw, zlCommFun.Nvl(!内容))
                    .Update
                    
                    If lngX <= !X坐标 Then lngX = !X坐标
                    
                    '70228:某列存在多个标记，最多显示两个(处理X坐标)
                    If Not (bln事件显示规则 = True And j Mod 2 = 1) Then
                        ReDim Preserve arrX(UBound(arrX) + 1)
                        arrX(UBound(arrX)) = lngX
                        lngX = lngX + T_DrawClient.列单位
                        j = 0
                    End If
                    If bln事件显示规则 = True Then j = j + 1
                Else
                    !高度 = GetFontHeight(objDraw, zlCommFun.Nvl(!内容))
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        '重新整理自动标志的高度
        '规则：如果一列要输出两个标志，则缩小字体。否则检查是否勾选了参数"体温标志超出40刻度缩小字体显示"，勾选则缩小字体
        .Filter = "禁用<>1"
        .Sort = "X坐标,时间,项目序号"
        Do While Not .EOF
            If Not (!类型 = 2 Or !类型 = 99) Then
                blnBigSize = True
                If bln事件显示规则 = True Then
                    For i = 0 To UBound(arrX)
                        If Val(arrX(i)) = Val(Nvl(!打印X坐标)) Then
                            blnBigSize = False
                            Exit For
                        End If
                    Next i
                End If
                
                If blnBigSize = True And blnLessenSize = True Then
                    If GetFontHeight(objDraw, zlCommFun.Nvl(!内容)) > dblCurveHeight Then
                        blnBigSize = False
                    End If
                End If
                
                If blnBigSize = False Then
                    gstdSet.Name = "宋体"
                    gstdSet.Size = 7.5
                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDC, lngFont)
                    dblHeight = GetFontHeight(objDraw, zlCommFun.Nvl(!内容))
                    Call SelectObject(lngDC, lngOldFont)
                    Call DeleteObject(lngFont)
                    '还原字体
                    gstdSet.Name = "宋体"
                    gstdSet.Size = 9
                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    Call SelectObject(lngDC, lngFont)
                    !高度 = dblHeight
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        lngY = GetYCoordinate(objDraw, rsDrawItems, gint体温, 42)
        '调整入出转 手术，分娩到达最大X坐标有多列式的Y坐标
        .Filter = "打印X坐标=" & lngMaxX & " And 禁用<>1"
        .Sort = "时间,项目序号"

        Do While Not .EOF
            !Y坐标 = lngY
            .Update
            lngY = lngY + Val(!高度) + T_DrawClient.行单位
            .MoveNext
        Loop
        
        '更新未记说明，上标的显示位置(Y坐标).
        '说明:在没有入出转，手术信息的情况下 打印在 42-40度之间，否则打印在40度以下打印
        .Filter = "禁用<>1"
        .MoveFirst
        .Sort = "X坐标,时间,项目序号"
        Set rsTemp = .Clone

        Do While Not .EOF
            If (!类型 = 2 Or !类型 = 99) Then
                
                rsTemp.Filter = "(打印X坐标=" & !X坐标 & " And 禁用<>1 and 类型=99) or (打印X坐标=" & !X坐标 & " And 禁用<>1 and 类型=2)"
                
                If rsTemp.BOF Then
                    rsTemp.Filter = "打印X坐标=" & !X坐标 & " And 禁用<>1"
                End If
                
                If rsTemp.RecordCount > 0 Then
                    bln上标 = False
                    lngY = 0
                    Do While Not rsTemp.EOF
                        If bln上标 = False Then
                            bln上标 = IIf(rsTemp!类型 = 2 Or rsTemp!类型 = 99, True, False)
                            lngY1 = Val(rsTemp!Y坐标)
                        End If
                        
                        If lngY < lngY1 + rsTemp!高度 + T_DrawClient.行单位 Then lngY = lngY1 + rsTemp!高度 + T_DrawClient.行单位
                        lngY1 = lngY
                        
                        rsTemp.MoveNext
                    Loop
                    
                    lngY1 = GetYCoordinate(objDraw, rsDrawItems, gint体温, 40)

                    If lngY > lngY1 Or bln上标 Then lngY1 = lngY
                    
                Else '不存在任何信息 从42开始打印
                    lngY1 = Val(!Y坐标)
                End If
                
                !Y坐标 = lngY1
                !打印X坐标 = !X坐标
                .Update
            End If

            .MoveNext
        Loop
        
        '70228:整理一列显示两个标记的打印X坐标，这段代码必须放在处理上标显示位置的后面
        '设置字体为7.5
        If bln事件显示规则 = True Then
            gstdSet.Name = "宋体"
            gstdSet.Size = 7.5
            objDraw.Font.Name = gstdSet.Name
            objDraw.Font.Size = gstdSet.Size
            
            For i = 0 To UBound(arrX)
                .Filter = "打印X坐标=" & Val(arrX(i)) & " And 禁用<>1 And 类型<>2 And 类型<>99"
                .Sort = "X坐标,时间,项目序号"
                If .RecordCount > 1 Then
                    lngX = !打印X坐标 - Abs(T_DrawClient.列单位 - (objDraw.TextWidth("刘") / T_TwipsPerPixel.X) * 2) / 2
                    !打印X坐标 = lngX
                    .Update
                    ReDim Preserve arrCurX(UBound(arrCurX) + 1)
                    arrCurX(UBound(arrCurX)) = !类型 & "," & !项目序号 & "," & !打印X坐标 & "," & Format(!时间, "yyyy-MM-dd HH:mm:ss")
                    .MoveNext
                    !打印X坐标 = lngX + objDraw.TextWidth("刘") / T_TwipsPerPixel.X
                    .Update
                    ReDim Preserve arrCurX(UBound(arrCurX) + 1)
                    arrCurX(UBound(arrCurX)) = !类型 & "," & !项目序号 & "," & !打印X坐标 & "," & Format(!时间, "yyyy-MM-dd HH:mm:ss")
                End If
            Next i
            '还原字体为9号字体
            gstdSet.Name = "宋体"
            gstdSet.Size = 9
            objDraw.Font.Name = gstdSet.Name
            objDraw.Font.Size = gstdSet.Size
        End If
        '开始输出内容
        .Filter = "禁用<>1"
        .MoveFirst
        .Sort = "X坐标,时间,项目序号"
        Dim sigNum As Single
        Do While Not .EOF
            '输出内容
            strComment = Trim(zlCommFun.Nvl(!内容))

            If strComment <> "" Then
                X = Val(IIf(Trim(!打印X坐标) <> "", !打印X坐标, !X坐标))
                Y = Val(!Y坐标)
                intAscCharNum = 0
                
                '70228:一列显示两个标记进行字体缩小处理
                blnBigSize = True
                blnCenter = True
                If bln事件显示规则 = True Then
                    For i = 0 To UBound(arrCurX)
                        If !类型 & "," & !项目序号 & "," & !打印X坐标 & "," & Format(!时间, "yyyy-MM-dd HH:mm:ss") = CStr(arrCurX(i)) Then
                            blnBigSize = False
                            Exit For
                        End If
                    Next i
                End If
                blnCenter = blnBigSize
                '如果一列只有一个标记，并且标记内容超出40刻度，则缩小字体。
                If blnBigSize = True And blnLessenSize = True And Not (!类型 = 2 Or !类型 = 99) Then
                    If GetFontHeight(objDraw, strComment) > dblCurveHeight Then
                        blnBigSize = False
                    End If
                End If
                
                gstdSet.Name = "宋体"
                gstdSet.Size = IIf(blnBigSize = True, 9, 7.5)
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                T_Size.H = objDraw.TextHeight("1") / T_TwipsPerPixel.Y
                    
                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.刻度区域.Bottom Then
                        strText = Mid(strComment, i, 1)
                        
                        If Asc(strText) < 0 Then
                            If intAscCharNum Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If

                        '输出字体信息
                        Call DrawRotateText(objDraw, lngDC, X, Y, strText, !颜色, sngScale, IIf(blnCenter = True, -999, objDraw.TextWidth("刘") / T_TwipsPerPixel.X))

                        If Asc(strText) < 0 Then
                            Y = Y + T_Size.H
                            intAscCharNum = 0
                        Else
                            Y = Y + T_Size.H / 2
                            intAscCharNum = intAscCharNum + 1
                        End If
                    End If
                Next i
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                
                gstdSet.Name = "宋体"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                Call SelectObject(lngDC, lngFont)
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub





Public Function GetAppendGridItemNew(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int护理等级 As Integer, ByVal int婴儿 As Long, dt开始时间 As Date, dt结束时间 As Date, ByVal byt适用病人 As Byte, ByVal lng科室ID As Long, ByVal str表格项目 As String, Optional blnMove As Boolean = False) As ADODB.Recordset
    '**************************************************************************
    '功能:提取活动有数据的体温表格项目以及固定表格项目
    '**************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Errhand
    
    Set rsTemp = GetGridItem(int护理等级, byt适用病人, lng科室ID, 2)
    If rsTemp.RecordCount = 0 Then
        '不存在活动项目直接提取固定表格项目
        Set rsTemp = GetGridItemNew(str表格项目)
        Set GetAppendGridItemNew = rsTemp
        Exit Function
    End If
    With rsTemp
        Do While Not .EOF
            strSQL = IIf(strSQL = "", "select " & !项目序号 & " 项目序号 from dual", strSQL & " UNION ALL select " & !项目序号 & "  项目序号 from dual ")
            .MoveNext
        Loop
    End With
    
    strSQL = "(" & strSQL & ") F"
    '提取活动项目
    gstrSQL = "Select distinct D.排列序号,D.项目序号,C.体温部位,C.体温部位 || D.记录名  记录名,D.记录法,D.记录符,D.记录色,D.最大值,D.最小值,D.单位值,nvl(D.记录频次,2) 记录频次,D.入院首测," & _
        "   E.项目性质,E.分组名,E.项目值域,E.项目表示,E.项目类型,E.项目长度,E.项目小数,E.项目单位 单位" & _
        "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E," & strSQL & _
        "   Where  B.ID=A.文件ID And A.ID = c.记录ID  AND B.ID=[1]  AND Nvl(B.婴儿,0)=[5]  AND B.病人id=[2]    AND B.主页id=[3] AND d.项目序号=C.项目序号 " & _
        "   AND c.记录类型=1 And E.项目性质=2  AND E.项目序号=D.项目序号  AND E.护理等级>=[4]   AND a.发生时间 BETWEEN [6] And [7] And c.终止版本 Is Null " & _
        "   AND d.记录法=2 and D.项目序号=F.项目序号"
    
    '提取固定表格项目
    strSQL = "Select A.排列序号,A.项目序号,'' 体温部位,A.记录名,A.记录法,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,nvl(D.C2,2) 记录频次,A.入院首测,B.项目性质," & _
        "   B.分组名,B.项目值域,B.项目表示,B.项目类型,B.项目长度,B.项目小数,B.项目单位 单位" & _
        "   From 体温记录项目 A,护理记录项目 B,诊治所见项目 C,TABLE(CAST(F_NUM2LIST2([8]) AS ZLTOOLS.T_NUMLIST2)) D" & _
        "   Where A.项目序号=B.项目序号 And B.项目ID=C.Id(+) And B.项目序号=D.C1 And A.记录法=2 And B.项目性质=1"



    
    gstrSQL = "Select /*+ Rule*/ 排列序号,项目序号,体温部位,记录名,记录法,记录符,记录色,最大值,最小值,单位值,记录频次,入院首测,项目性质," & _
        "   分组名,项目值域,项目表示,项目类型,项目长度,项目小数,单位" & _
        "   From (" & gstrSQL & vbCrLf & " UNION ALL " & vbCrLf & strSQL & ") order by Decode(项目序号,3 ,0,1 ),排列序号,记录名"
    If blnMove Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lng文件ID, lng病人ID, lng主页ID, int护理等级, int婴儿, CDate(Format(dt开始时间, "yyyy-mm-dd hh:mm:ss")), CDate(Format(dt结束时间, "yyyy-mm-dd hh:mm:ss")), str表格项目)
    
    Set GetAppendGridItemNew = rsTemp

    Exit Function

Errhand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetGridItemNew(ByVal str表格项目 As String) As ADODB.Recordset

    '**********************************************************************************
    '功能:提取专科体温表格项目
    '**********************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo Errhand
    
   '提取表格活动项目
   gstrSQL = "Select A.排列序号,A.项目序号,'' 体温部位,A.记录名,A.记录法,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,nvl(D.C2,2) 记录频次,A.入院首测,B.项目性质," & _
        "   B.分组名,B.项目值域,B.项目表示,B.项目类型,B.项目长度,B.项目小数,B.项目单位 单位" & _
        "   From 体温记录项目 A,护理记录项目 B,诊治所见项目 C,TABLE(CAST(F_NUM2LIST2([1]) AS ZLTOOLS.T_NUMLIST2)) D" & _
        "   Where A.项目序号=B.项目序号 And B.项目ID=C.Id(+) And B.项目序号=D.C1 And A.记录法=2 And B.项目性质=1" & _
        "   order by Decode(项目序号,3 ,0,1 ),排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取固定体温表格项目", str表格项目)
    Set GetGridItemNew = rsTemp

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetRows(bln呼吸 As Boolean, ByVal strValue As String) As Long
    Dim strOld() As String
    Dim intRow As Integer, i As Integer
    Dim intRows As Integer
    strOld = Split(strValue, ",")
    For i = 0 To UBound(strOld)
        If InStr(1, strOld(i), ":") > 0 Then
            If Split(strOld(i), ":")(0) = 3 Then
                bln呼吸 = True
                intRows = intRows + 1
            Else
                If Split(strOld(i), ":")(0) <> 5 Then
                    Select Case Split(strOld(i), ":")(1)
                        Case 1
                            intRow = 1
                        Case 2
                            intRow = 1
                        Case 3
                            intRow = 3
                        Case 4
                            intRow = 2
                        Case 6
                            intRow = 1
                    End Select
                    intRows = intRows + intRow
                End If
            End If
        End If
    Next
    GetRows = intRows
End Function


Private Function getSQLString(ByVal strText As String, ByVal blnMoved As Boolean, Optional ByVal strItems As String) As String
    Dim strNewSql As String
    Dim strSQL As String
    Dim strSQLText As String
    Dim lngColor As Long
    Select Case strText
        Case "提取文件时间范围"

             strNewSql = "   (SELECT 病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0  AND C.类别 = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.操作类型 = COLUMN_VALUE) And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
            '提取病人出院前的时间信息
            '------------------------------------------------------------------------------------------------------------------
            strSQL = " SELECT /*+ RULE */ DECODE(D.开始时间,NULL,DECODE(B.出生时间, NULL, A.开始, B.出生时间)," & vbNewLine & _
                "               DECODE(SIGN(D.开始时间 - DECODE(B.出生时间, NULL, A.开始, B.出生时间))," & vbNewLine & _
                "                      1," & vbNewLine & _
                "                      D.开始时间," & vbNewLine & _
                "                      DECODE(B.出生时间, NULL, A.开始, B.出生时间))) AS 开始," & vbNewLine & _
                "       DECODE(D.结束时间," & vbNewLine & _
                "               NULL," & vbNewLine & _
                "               DECODE(E.记录," & vbNewLine & _
                "                      0," & vbNewLine & _
                "                      DECODE(SIGN(NVL(E.婴儿时间, A.终止) - D.发生时间), 1, NVL(E.婴儿时间, A.终止), D.发生时间)," & vbNewLine & _
                "                      NVL(E.婴儿时间, A.终止))," & vbNewLine & _
                "               DECODE(SIGN(NVL(E.婴儿时间, A.终止) - D.结束时间), 1, D.结束时间, NVL(E.婴儿时间, A.终止))) 终止," & vbNewLine & _
                "       DECODE(D.结束时间, NULL, E.记录, 1) 记录" & vbNewLine & _
                " FROM (SELECT 病人ID, 主页ID, MIN(开始时间) AS 开始, MAX(NVL(终止时间, SYSDATE)) AS 终止" & vbNewLine & _
                "       FROM 病人变动记录" & vbNewLine & _
                "       WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID = [3]" & vbNewLine & _
                "       GROUP BY 病人ID, 主页ID) A," & vbNewLine & _
                "     (SELECT 病人ID, 主页ID, 出生时间 FROM 病人新生儿记录 WHERE 病人ID = [2] AND 主页ID = [3] AND 序号 = [4]) B," & vbNewLine & _
                "     (SELECT NVL(发生时间, SYSDATE) 发生时间, 开始时间, 结束时间" & vbNewLine & _
                "       FROM (SELECT MAX(B.发生时间) 发生时间, MAX(A.开始时间) 开始时间, MAX(A.结束时间) 结束时间" & vbNewLine & _
                "              FROM 病人护理文件 A, 病人护理数据 B" & vbNewLine & _
                "              WHERE A.ID = B.文件ID(+) AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND A.婴儿 = [4])) D," & vbNewLine & _
                "  " & strNewSql & vbNewLine & _
                " WHERE A.病人ID = E.病人ID AND A.主页ID = E.主页ID AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)"
            
            strSQLText = strSQL
        
        Case "提取所有曲线项目"
            
            strSQL = " Select A.项目序号,A.排列序号,A.记录名,C.项目值域,A.记录法,A.记录符,A.记录色,nvl(A.最大值,0) 最大值 ,nvl(A.最小值,0) 最小值,A.临界值," & _
                "nvl(A.单位值,0) 单位值,A.刻度间隔,A.警示线,C.项目单位 单位,Decode(记录法,3,A.最高行,nvl(A.最高行,2)-2) AS 最高行,B.部位 " & _
                " From 体温记录项目 A,体温部位 B,护理记录项目 C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) D" & _
                " Where A.项目序号=B.项目序号(+) And B.缺省项(+)=1" & _
                " And A.项目序号=C.项目序号 And A.记录法<>2 And NOT (NVL(C.应用方式,0)=2 And C.项目序号=-1) And C.项目序号=D.COLUMN_VALUE" & _
                " Order by 排列序号"
                
            strSQLText = strSQL
        Case "提取表格非汇总项目"
            
            gstrSQL = _
            " Select A.排列序号,A.项目序号,A.记录名,A.记录法,A.记录符,A.记录色,B.项目值域,nvl(D.C2,2) 记录频次,A.入院首测,B.项目性质,'' 部位," & _
            "   B.项目类型,B.项目长度,B.项目表示,B.项目小数,B.项目单位 项目单位" & _
            "   From 体温记录项目 A,护理记录项目 B,诊治所见项目 C,TABLE(CAST(F_NUM2LIST2([10]) AS ZLTOOLS.T_NUMLIST2)) D" & _
            "   Where A.项目序号=B.项目序号 And B.项目ID=C.Id(+) And B.项目序号=D.C1 And A.记录法=2 And B.项目性质=1" & _
            "   UNION ALL " & _
            " Select Distinct  B.排列序号,B.项目序号,B.记录名,B.记录法,B.记录符,B.记录色,C.项目值域,nvl(B.记录频次,2) 记录频次,B.入院首测,C.项目性质, A.部位," & _
                "   C.项目类型,C.项目长度,C.项目表示,C.项目小数,C.项目单位 项目单位" & _
                "            From (Select 项目序号, DECODE(项目序号,3,'',体温部位) 部位" & vbNewLine & _
                "                           From 病人护理文件 a, 病人护理数据 b, 病人护理明细 c" & vbNewLine & _
                "                           Where a.Id = b.文件id And b.Id = c.记录id And a.Id = [1] And Nvl(a.婴儿, 0) = [4] And a.病人id = [2] And" & vbNewLine & _
                "                                       a.主页id = [3] And c.记录类型 = 1 And b.发生时间 Between [5] And [6] And 终止版本 Is Null) a, 体温记录项目 b," & vbNewLine & _
                "                       护理记录项目 c" & vbNewLine & _
                "            Where b.项目序号 = a.项目序号 And b.项目序号 = c.项目序号 And b.记录法 = 2 And C.项目性质=2" & _
                "   And nvl(C.应用方式,0)=1 And nvl(C.护理等级,0)>=[7] And nvl(C.适用病人,0) In (0,[8])" & _
                "   And (C.适用科室=1 Or (C.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=C.项目序号 And D.科室id=[9])))"
        
            strSQL = "Select Rownum-1 序号 ,项目序号, Decode(项目序号, 4, '血压',记录名) 项目名称,记录色,项目单位,项目值域, 部位,记录频次,入院首测,项目性质,项目表示,项目类型 " & _
                " From ( select 排列序号, 项目序号, 记录名, 记录法, 记录符, 记录色, 项目值域, Nvl(记录频次, 2) 记录频次, 入院首测, 项目性质, 部位, 项目类型, 项目长度,项目表示, 项目小数, 项目单位 项目单位 " & _
                      "From(" & gstrSQL & ") A where 项目序号<>5 order by Decode(A.项目序号,3 ,0,1 ),A.排列序号,a.记录名,a.部位) "
                      
            strSQLText = strSQL
        Case "读取病人体温数据和未记说明"
                
            strSQL = _
                    " SELECT /*+ Rule*/ C.ID 序号, a.发生时间 As 时间,C.显示,C.记录内容 As 数值,C.体温部位,c.复试合格,D.记录名,E.保留项目,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明 " & _
                    " FROM 病人护理文件 B,病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E,Table(Cast(f_num2list([6]) As zlTools.t_Numlist)) F " & _
                    " Where B.ID=A.文件ID  " & _
                    "   AND A.ID = C.记录ID " & _
                    "   AND B.ID=[1] " & _
                    "   AND B.病人id=[2] " & _
                    "   AND B.主页id=[3] " & _
                    "   AND D.项目序号=C.项目序号 " & _
                    "   AND C.记录类型=1 " & _
                    "   AND E.项目序号=D.项目序号 " & _
                    "   AND E.项目序号=F.COLUMN_VALUE " & _
                    "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null And D.记录法<>2 " & _
                    " Order By a.发生时间,DECODE(D.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记)"
            
            strSQLText = strSQL
        Case "提取所有表格项目数据信息"
            
            strSQL = "SELECT /*+ Rule*/  C.Id,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & vbNewLine & _
                "  DECODE(E.项目性质,2,C.体温部位 || D.记录名,D.记录名) 项目名称,D.项目序号,C.来源ID,C.共用,E.项目性质 " & _
                "  FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E" & _
                "  Where B.ID = A.文件ID" & vbNewLine & _
                "  AND A.ID = C.记录ID" & vbNewLine & _
                "  AND B.ID = [1]" & vbNewLine & _
                "  AND B.病人id = [2]" & vbNewLine & _
                "  AND B.主页id = [3]" & vbNewLine & _
                "  AND Nvl(B.婴儿, 0) = [7]" & vbNewLine & _
                "  AND INSTR([6], DECODE(E.项目性质, 2,C.体温部位 || D.记录名, D.记录名)) > 0" & vbNewLine & _
                "  AND D.项目序号 = C.项目序号" & vbNewLine & _
                "  AND Mod(c.记录类型,10) = 1" & vbNewLine & _
                "  AND E.项目序号 = D.项目序号" & vbNewLine & _
                "  AND E.护理等级 >= [8]" & vbNewLine & _
                "  AND A.发生时间 BETWEEN [4] And [5]" & vbNewLine & _
                "  AND D.记录法 = 2" & vbNewLine & _
                "  UNION ALL "
             '提取非体温表格的汇总项目（体温表格汇总项目子项可能存在非体温项目）
            strSQL = strSQL & vbNewLine & _
                "  SELECT C.ID,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & _
                "   D.项目名称,D.项目序号,C.来源ID,C.共用,D.项目性质" & _
                "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,(SELECT A.项目序号,A.项目名称, 1 项目性质,B.父序号 FROM 护理记录项目 A,护理汇总项目 B" & vbNewLine & _
                "       WHERE A.项目序号=B.序号 AND NOT EXISTS (SELECT C.COLUMN_VALUE FROM Table(Cast(f_num2list([11]) As zlTools.t_Numlist)) C,护理汇总项目 E WHERE C.COLUMN_VALUE=E.序号 AND C.COLUMN_VALUE=A.项目序号)" & vbNewLine & _
                "       AND NVL(A.应用方式,0)=1 AND NVL(A.护理等级,0)>=[8] AND NVL(A.适用病人,0) IN (0,[9])" & vbNewLine & _
                "       AND (A.适用科室=1 OR (A.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D WHERE D.项目序号=A.项目序号 AND D.科室ID=[10])))) D" & _
                "   Where B.ID=A.文件ID And A.ID = C.记录ID   AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
                "   AND B.病人id=[2]  AND B.主页id=[3]  AND D.项目序号=C.项目序号  AND C.记录类型=1" & _
                "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null"
                
            strSQL = _
                "   Select ID,时间,记录类型,显示,结果,体温部位,未记说明,数据来源,项目名称,项目序号,来源ID,共用,项目性质 From (" & strSQL & ")" & _
                "   Order By  Decode(项目名称,'收缩压',0,1)," & strItems & ",时间"
                
            strSQLText = strSQL
        Case "读取手术、上下标信息"
            strSQL = "" & _
                 " Select B.发生时间 AS 时间,C.记录类型,C.项目序号,C.记录内容,C.项目名称,C.未记说明" & _
                 " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & _
                 " Where A.ID=B.文件ID and  B.ID = C.记录ID AND A.ID=[1]   AND Nvl(A.婴儿, 0)=[6] AND A.病人id=[2] AND A.主页id=[3] And c.终止版本 Is Null" & _
                 " AND mod(c.记录类型,10) <> 1  AND B.发生时间 BETWEEN [4]  And [5]"
            strSQLText = strSQL
        Case "显示皮试结果"
            lngColor = RGB(0, 0, 255)
            strSQL = _
               "SELECT 时间,F_LIST2STR(CAST(COLLECT(药物名) AS T_STRLIST)) 药物名 FROM (" & vbNewLine & _
                "   SELECT TO_CHAR(A.开始执行时间,'YYYY-MM-DD') 时间,DECODE(皮试结果,'(+)',255,'(阳性)',255," & lngColor & ") || '-#' || REPLACE(REPLACE(REPLACE(DECODE(B.试管编码,NULL,A.医嘱内容,B.试管编码),',',''),'-#',''),'皮试','') || A.皮试结果  药物名" & vbNewLine & _
                "   FROM 病人医嘱记录 A,诊疗项目目录 B " & vbNewLine & _
                "   WHERE  A.病人ID=[1] AND A.主页ID=[2] AND A.婴儿=[3] AND A.皮试结果 IS NOT NULL" & vbNewLine & _
                "   AND A.开始执行时间  BETWEEN [4] AND [5] AND A.诊疗项目id=b.id(+)" & vbNewLine & _
                "   ORDER BY A.开始执行时间,A.皮试结果)" & vbNewLine & _
                "GROUP BY 时间"
            strSQLText = strSQL
        Case "提取科室床号"
            strSQL = " Select  c.名称 As 科室,b.名称 As 病区,a.床号,a.开始原因 " & _
                " From 病人变动记录 a,部门表 b,部门表 c " & _
                " Where a.病人id=[1] And a.主页id=[2] And a.科室id Is Not Null And a.病区id=b.id and a.科室id=c.id And NVL(A.附加床位,0)=0 " & _
                " And a.开始时间-" & T_BodyStyle.lng时间间隔 & "/24<=[3] And Nvl(a.终止时间,Sysdate)>=[4] Order By a.开始时间"
            strSQLText = strSQL
        Case "提取当前手术信息"
            strSQL = "Select B.发生时间 时间" & vbNewLine & _
                " From 病人护理文件 A,病人护理数据 B,病人护理明细 C" & vbNewLine & _
                " Where A.Id=B.文件ID And B.Id=C.记录ID And A.Id=[1] And  nvl(A.婴儿,0)=[2]" & vbNewLine & _
                " And A.病人ID=[3] and A.主页ID=[4] and C.记录类型=4 And NVL(C.复试合格,0)<>1 and C.终止版本 is null" & vbNewLine & _
                " And B.发生时间 between [5] and [6] order by B.发生时间"
            strSQLText = strSQL
        Case "提取14天之前的手术信息"
            strSQL = "select Nvl(Count(B.发生时间),0) 次数" & _
                "   from 病人护理文件 A, 病人护理数据 B,病人护理明细 C" & _
                "   where A.ID=B.文件ID and B.ID=C.记录ID and A.ID=[1] and nvl(A.婴儿,0)=[2]" & _
                "   and A.病人ID=[3] and A.主页ID=[4] and C.记录类型=4 And NVL(C.复试合格,0)<>1 and C.终止版本 is null" & _
                "   and B.发生时间 <[5] "
            strSQLText = strSQL
    End Select
    If blnMoved Then
        strSQLText = Replace(strSQLText, "病人护理文件", "H病人护理文件")
        strSQLText = Replace(strSQLText, "病人护理数据", "H病人护理数据")
        strSQLText = Replace(strSQLText, "病人护理明细", "H病人护理明细")
        strSQLText = Replace(strSQLText, "病人过敏记录", "H病人过敏记录")
    End If
    getSQLString = strSQLText
End Function

Public Function GetDiagnoseMinTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strTime As Date, Optional ByVal blnMoved As Boolean = False) As String
'功能:获取最小诊断时间
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, strSQL As String
    On Error GoTo Errhand
    strTmp = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    strSQL = "SELECT /*+Rule */" & vbNewLine & _
        " MIN(记录日期) 诊断日期" & vbNewLine & _
        " FROM 病人诊断记录 a, TABLE(CAST(f_Num2list('1,2') AS Zltools.t_Numlist)) b" & vbNewLine & _
        " WHERE MOD(a.诊断类型, 10) = b.Column_Value AND a.病人id = [1] AND a.主页id = [2] And a.记录来源>1"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取最小诊断时间", lng病人ID, lng主页ID)
    If rsTmp.BOF = False Then
        If IsDate(Nvl(rsTmp!诊断日期)) Then
            If CDate(rsTmp!诊断日期) >= CDate(strTmp) Then
                strTmp = Format(rsTmp!诊断日期, "yyyy-MM-dd HH:mm:ss")
                strTmp = DateAdd("s", 1, CDate(strTmp))
            End If
        End If
    End If
    GetDiagnoseMinTime = strTmp
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub ClearData(ByVal lng偏移量X As Long, ByVal lng偏移量Y As Long, ByVal lng刻度单位 As Long, _
                        ByVal sin行单位 As Single, ByVal sin时间行单位 As Single, ByVal sin时间列单位 As Single, ByVal lng列单位 As Long, ByVal bln双倍 As Boolean, ByVal lng总列数 As Long, _
                        ByVal lng独立曲线总行数 As Long, ByVal lng刻度宽度 As Long)
    
    T_DrawClient.偏移量X = lng偏移量X
    T_DrawClient.偏移量Y = lng偏移量Y
    T_DrawClient.刻度单位 = lng刻度单位
    T_DrawClient.行单位 = sin行单位
    T_DrawClient.时间行单位 = sin时间行单位
    T_DrawClient.时间列单位 = sin时间列单位
    T_DrawClient.列单位 = lng列单位
    T_DrawClient.双倍 = bln双倍
    T_DrawClient.总列数 = lng总列数
    T_DrawClient.独立曲线总行数 = lng独立曲线总行数
    T_BodyStyle.lng刻度宽度 = lng刻度宽度
End Sub


Private Function Get护理等级(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lng护理等级 As Long
    
    
    strSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "护理等级", lng病人ID, lng主页ID)
    If rsTemp.BOF = False Then lng护理等级 = zlCommFun.Nvl(rsTemp("护理等级"), 0)
    
    Get护理等级 = lng护理等级
End Function


Private Function Get总行数(ByVal dbl数值 As Double, ByVal lngCurveRow As Long) As Integer
    Dim intDrawLineRows As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRows As Integer
    
    strSQL = "Select Count(A.项目序号) 记录数 " & _
    "   From 体温记录项目 A,护理记录项目 B,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) C " & _
    "   Where A.项目序号=B.项目序号 And B.项目序号=C.COLUMN_VALUE AND A.记录法<>2 AND NOT (NVL(B.应用方式,0)=2 And B.项目序号=-1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", T_BodyItem.str曲线项目)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        intDrawLineRows = zlCommFun.Nvl(rsTmp!记录数, 0)
    Else
        CloseRs rsTmp
        Get总行数 = 0
        Exit Function
    End If

    If intDrawLineRows < 1 Then
        Get总行数 = 0
        Exit Function
    End If
    
     strSQL = "Select nvl(A.最大值,0) 最大值,nvl(A.最小值,0) 最小值 ,nvl(A.单位值,0.1) 单位值 ,Decode(记录法,3,A.最高行,nvl(A.最高行,2)-2) AS 最高行,A.记录法,A.项目序号" & _
        "   From 体温记录项目 A,护理记录项目 B,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) C " & _
        "   Where A.项目序号=B.项目序号 And b.项目序号=c.Column_value AND A.记录法<>2 AND NOT (NVL(B.应用方式,0)=2 And B.项目序号=-1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", T_BodyItem.str曲线项目)

    rsTmp.Filter = "记录法=1 And 项目序号=" & gint体温
    If rsTmp.RecordCount > 0 Then
        '修改问题：51442
        dbl数值 = Val(zlCommFun.Nvl(rsTmp!最小值, 0))
        intDrawLineRows = (Val(rsTmp!最大值) - IIf(dbl数值 > 34, 35, dbl数值)) / 0.1 + IIf(Val(rsTmp!最高行) < 0, 0, Val(rsTmp!最高行))
        intDrawLineRows = intDrawLineRows + lngCurveRow
    Else
        intDrawLineRows = glngMaxRows + lngCurveRow
    End If
    
       
    T_DrawClient.独立曲线总行数 = 0
    rsTmp.Filter = "记录法=3"
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
             '修改问题：51442
            intRows = (Val(rsTmp!最大值) - Val(zlCommFun.Nvl(rsTmp!最小值, 0))) / Val(zlCommFun.Nvl(rsTmp!单位值, 0)) + IIf(Val(rsTmp!最高行) < 0, 0, Val(rsTmp!最高行))
            If intRows Mod 2 = 1 Then intRows = intRows + 1
            T_DrawClient.独立曲线总行数 = T_DrawClient.独立曲线总行数 + intRows
            rsTmp.MoveNext
        Loop
    End If
    
    Get总行数 = intDrawLineRows
End Function





Private Function Get符号(ByVal lng重叠 As Long, ByVal str重叠项目 As String, ByVal lng项目序号 As Long, ByVal str符号 As String, ByVal strPosition As String, strEditors() As Variant, ByVal lng标记 As Long) As Boolean
    '获取体温符号
    Dim str部位 As String
    Dim strTmp As String
    Dim strChar As String
    Dim strPic As String
    Dim i As Integer
    
    On Error GoTo Errhand
    
    
    If lng重叠 = 0 And str重叠项目 = "空" Then '未重叠的项目
         For i = 0 To UBound(strEditors)
            If Split(CStr(strEditors(i)), "||")(0) = lng项目序号 Then
                 Exit For
            End If
        Next i
        str部位 = strPosition
        If str部位 = "" Then
            Select Case lng项目序号
                Case gint体温
                    str部位 = "腋温"
                Case gint呼吸
                    str部位 = "自主呼吸"
                Case Else
                    str部位 = ""
            End Select
        End If
        strTmp = Split(CStr(strEditors(i)), "||")(4)
        strPic = ""
        strChar = ""
        Select Case lng项目序号
            Case gint体温
                strTmp = strTmp & String(3 - UBound(Split(strTmp, ",")), ",")
                If str部位 = "口温" Then
                    strChar = Split(strTmp, ",")(0)
                ElseIf str部位 = "腋温" Then
                    strChar = Split(strTmp, ",")(1)
                ElseIf str部位 = "肛温" Then
                    strChar = Split(strTmp, ",")(2)
                Else
                    strChar = Split(strTmp, ",")(3)
                End If
                If lng标记 = 1 Then '物理降温符号
                    strChar = "○"
                Else
                    If strChar = "" Then strChar = "×"
                End If
            Case gint心率
                strChar = IIf(strTmp = "", "Ο", strTmp)
            Case gint脉搏
                If str部位 = "起搏器" Then
                    strPic = "PACEMAKER"
                Else
                    strChar = IIf(strTmp = "", "+", strTmp)
                End If
            Case gint呼吸
                If str部位 = "自主呼吸" Then
                    strChar = IIf(strTmp = "", "*", strTmp)
                Else
                    strPic = "BREATH"
                End If
            Case Else
                strChar = strTmp
        End Select
        If Trim(str符号) <> "" Then
            strChar = Trim(str符号)
            strPic = ""
        End If
    End If
    
    If strChar <> "○" Then
        Get符号 = False
    Else
        Get符号 = True
    End If
        
        Exit Function
Errhand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function


Public Sub DrawTextPrint(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    
    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = X
        .CurrentY = Y
        objDraw.FontTransparent = True
        objDraw.Print Text
        .ForeColor = lngSaveForeColor
    End With
End Sub

Private Function GetMaxMinValue(ByVal bytType As Byte, ByVal lngNO As Long, arrEditors() As Variant) As Double
'功能:获取曲线项目的临界值(描点的最大值和最小值)
'参数:bytType=0 最小值,1-最大值
'     arrEditors:'记录曲线项目信息(项目序号||项目名称||项目单位||项目值域||记录符||记录色||最大值||最小值||临界值）
    Dim dblvalue As Double
    Dim dblMax As Double, dblMin As Double
    Dim strValue As String
    Dim i As Integer
    
    For i = 0 To UBound(arrEditors)
        If Val(Split(arrEditors(i), "||")(0)) = lngNO Then
             Exit For
        End If
    Next i
    
    If i <= UBound(arrEditors) Then
        dblMax = Val(Split(arrEditors(i), "||")(6))
        dblMin = Val(Split(arrEditors(i), "||")(7))
    End If
    
    strValue = Split(arrEditors(i), "||")(8)
    If bytType = 0 Then
        dblvalue = dblMin
        If InStr(1, strValue, ";") <> 0 Then
            strValue = Split(strValue, ";")(0)
        Else
            strValue = ""
        End If
        If IsNumeric(strValue) = True And Val(strValue) <= dblMax And Val(strValue) >= dblMin Then
            dblvalue = Val(strValue)
        Else
            '体温如果最小值无效，则输出的最小值为35
            If lngNO = gint体温 And dblvalue < 35 Then dblvalue = 35
        End If
    Else
        dblvalue = dblMax
        If InStr(1, strValue, ";") <> 0 Then strValue = Split(strValue, ";")(1)
        If IsNumeric(strValue) = True And Val(strValue) <= dblMax And Val(strValue) >= dblMin Then dblvalue = Val(strValue)
    End If
    
    GetMaxMinValue = dblvalue
End Function
