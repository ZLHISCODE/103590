Attribute VB_Name = "mdlRegEventSelf"
Option Explicit

Public Function GetRoom(str号别 As String) As String
'周燕川2003年1月6日调整到公用函数
'功能：根据号别的分诊方式获取号别的诊室
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    gstrSQL = "Select ID,Nvl(分诊方式,0) as 分诊 From 挂号安排 Where 号码=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取挂号安排", str号别)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!分诊 = 0 Then Exit Function '不分诊
    
    '处理分诊
    If rsTmp!分诊 = 1 Then
        '指定诊室
        '指定诊室
        Dim lng号表ID As Long
        lng号表ID = Val(Nvl(rsTmp!ID))
        Set rsTmp = GetRs挂号诊室
        If rsTmp Is Nothing Then
            gstrSQL = "Select 号表ID,门诊诊室 From 挂号安排诊室 Where 号表ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "指定诊室", lng号表ID)
        End If
        rsTmp.Filter = "号表ID=" & lng号表ID
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
        rsTmp.Filter = 0
        
    ElseIf rsTmp!分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        gstrSQL = _
        " Select 门诊诊室,Sum(NUM) as NUM  " & _
        " From ( Select 门诊诊室,0 as NUM From 挂号安排诊室 Where 号表ID=[1]" & _
        "        Union ALL" & _
        "       Select 发药窗口,Count(发药窗口) as NUM From 门诊费用记录" & _
        "       Where 记录性质=4 And 记录状态=1 And 序号=1 And Nvl(执行状态,0)=0" & _
        "               And 登记时间 Between Trunc(Sysdate) And Sysdate And 计算单位=[2]" & _
        "               And 发药窗口 IN(Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1])" & _
        "       Group by 发药窗口) " & _
        " Group by 门诊诊室 " & _
        " Order by Num"
        
       Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊人数最少诊室", Val(Nvl(rsTmp!ID)), str号别)
       If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        gstrSQL = "Select 号表ID,门诊诊室,当前分配 From 挂号安排诊室 Where 号表ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "平均分配诊室", Val(Nvl(rsTmp!ID)))
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!当前分配), 0, rsTmp!当前分配) = 1 Then
                    GetRoom = rsTmp!门诊诊室
                    rsTmp!当前分配 = 0
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If GetRoom = "" Then
                rsTmp.MoveFirst
                GetRoom = rsTmp!门诊诊室
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRegistPrice(ByVal lng项目ID) As Variant
    '******************************************************************************************************************
    '功能：返回指定挂号类型，在指定时间的价格二维（六列）数组。
    '   第一列为价格，第二列表示收入项目ID，第三列填写收入项目,第四列为计算单位,第五列为数次,第六列为收费细目ID,第七列(价格序号),第八列(从属父号)
    '参数：lng项目ID=挂号项目ID(收费细目ID)
    '返回：数组
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim aryTmp(), i As Integer
    Dim int性质 As Integer, int父号 As Integer, lng收入项目ID As Long
    On Error GoTo errH

    gstrSQL = "Select 1 as 性质,A.类别,A.ID as 项目ID,A.计算单位,B.收入项目ID,1 as 数次,C.收据费目,B.现价" & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=[1] " & _
        " And ((To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS') Between To_Char(B.执行日期,'YYYY-MM-DD HH24:MI:SS') And To_Char(B.终止日期,'YYYY-MM-DD HH24:MI:SS')) or (To_Char(Sysdate,'YYYY-MM-DD HH24:MI:SS')>=To_Char(B.执行日期,'YYYY-MM-DD HH24:MI:SS') And (B.终止日期 is NULL Or B.终止日期=To_Date('3000-01-01','YYYY-MM-DD'))))"
    gstrSQL = gstrSQL & " Union ALL " & _
        "Select 2 as 性质,A.类别,A.ID as 项目ID,A.计算单位,C.ID as 收入项目ID,D.从项数次 as 数次,C.收据费目,B.现价" & _
        " From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D" & _
        " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.从项ID And D.主项ID=[1]" & _
        "        And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & _
        ""
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lng项目ID)
    If rs.EOF Then
        GetRegistPrice = Null
    Else
        ReDim aryTmp(rs.RecordCount - 1, 8)
        int性质 = 0: lng收入项目ID = 0
        For i = 1 To rs.RecordCount
            If lng项目ID = Val(Nvl(rs!项目ID)) Then
                If lng收入项目ID <> Val(Nvl(rs!收入项目ID)) Then
                    int性质 = 1: int父号 = i:
                     lng收入项目ID = Val(Nvl(rs!收入项目ID))
                End If
            Else
                int性质 = 2
            End If
            
            aryTmp(i - 1, 0) = zlCommFun.Nvl(rs("现价").Value, 0)
            aryTmp(i - 1, 1) = zlCommFun.Nvl(rs("收入项目ID").Value, 0)
            aryTmp(i - 1, 2) = zlCommFun.Nvl(rs("收据费目").Value)
            aryTmp(i - 1, 3) = zlCommFun.Nvl(rs("计算单位").Value)
            aryTmp(i - 1, 4) = zlCommFun.Nvl(rs("数次").Value)
            aryTmp(i - 1, 5) = zlCommFun.Nvl(rs("项目ID").Value)
            aryTmp(i - 1, 6) = IIf(int性质 = 1 And i <> int父号, int父号, 0)
            aryTmp(i - 1, 7) = IIf(int性质 = 2 And i <> int父号, int父号, 0)
            rs.MoveNext
        Next
        GetRegistPrice = aryTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GetRegistPrice = Null
End Function

Public Function ActualMoney(ByVal str费别 As String, ByVal lng收入ID As Long, ByVal cur应收 As Currency) As Currency
'功能：根据指定的费别和收入项目,计算指定金额的实际收款金额
'参数：
'   str费别   ：费别
'   lng收入ID  ：收入项目ID
'   cur应收：应收金额值
'返回：实际应收的金额
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    gstrSQL = "Select 实收比率 From 费别明细 " & _
        " Where 费别=[1] And 收入项目ID= [2] " & _
        " And ABS([3]) Between 应收段首值 And 应收段尾值" & _
        " Order by 应收段尾值"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", str费别, CStr(lng收入ID), CStr(cur应收))
    
    If rsTmp.EOF Then
        ActualMoney = cur应收
    Else
        ActualMoney = cur应收 * rsTmp!实收比率 / 100
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim rsPar As New ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errH
    
    '卡号显示方式
    gblnShowCard = Not (-Abs(Val(zlDatabase.GetPara(12, glngSys))))
    
    '挂号票据号码长度
    strTmp = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbyt挂号 = Val(Split(strTmp, "|")(3))
    
    gstrSQL = "Select 卡号长度 From 医疗卡类别 where 名称='就诊卡' and nvl(是否固定,0)=1"
    gbytCardNOLen = 7
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取就诊卡卡号长度")
    If Not rsTemp.EOF Then
        gbytCardNOLen = Val(Nvl(rsTemp!卡号长度))
    End If
    '票号严格控制
    strTmp = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnBill挂号 = (Mid(strTmp, 4, 1) = "1")
    
    '日报统计时间允许
    gblnDailyTime = zlDatabase.GetPara(22, glngSys, , 0)
     
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InitSysPar = False
End Function

Public Sub InitLocPar()
'功能：初始化费用本机参数
    '可选输入值
    '本地共用预交票据批次ID
    
    glng挂号ID = Val(zlDatabase.GetPara("共用挂号票据批次", glngSys, 1111, 0))

    If glng挂号ID > 0 Then
        If Not ExistBill(glng挂号ID, 4) Then
            
            Call zlDatabase.SetPara("共用挂号票据批次", 0, glngSys, 1111)

            glng挂号ID = 0
        End If
    End If
    
End Sub
Public Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'功能：判断是否存在指定的票据领用
    Dim rsTmp As New ADODB.Recordset

    On Error GoTo errH

    gstrSQL = "Select ID From 票据领用记录 Where ID= [1] And 票种= [2] "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lngID, bytKind)
    
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNext号别() As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    gstrSQL = "Select Max(号码) as 号码 From 挂号安排 Where Length(号码)=(Select Max(Length(号码)) From 挂号安排)"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf")
    
    If Not rsTmp.EOF Then GetNext号别 = IncStr(IIf(IsNull(rsTmp!号码), "", rsTmp!号码))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, Optional ByVal strBill As String) As Long
'功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
'参数：bytKind=票种
'      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
'      strBill=要检查范围的票据号
'说明：
'    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
'    2.在检查范围时,长度也在检查范围之内。
'    3.当有多批自用时,缺省按少的先用,先领先用原则取
'返回：
'      正常：票据领用ID>0
'      0=失败
'      -1:没有自用(用完或未领用)、也没有共用(未设置)
'      -2:设置的共用已用完
'      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
     '病人有剩余的自用票据集
    gstrSQL = _
        " Select " & zlGetFeeFields("票据领用记录") & " From 票据领用记录 Where 票种=[1]" & _
        " And 使用方式=1 And 剩余数量>0 And 领用人=[2]" & _
        " Order by 剩余数量,登记时间"
        
    Set rsSelf = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", bytKind, UserInfo.姓名)
    
    If lng领用ID = 0 Then
        '程序中第一次检查,且没有设置本地共用
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '也没有自用票据
        '有自用票据,按优先原则返回
        CheckUsedBill = rsSelf!ID
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        gstrSQL = "Select " & zlGetFeeFields("票据领用记录") & " From 票据领用记录 Where 票种=[1] And ID= [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", bytKind, lng领用ID)
        If rsTmp!使用方式 = 2 Then '共用,要先看有没有自用
            If Not rsSelf.EOF Then
                '有自用的，优先
                CheckUsedBill = rsSelf!ID
            Else
                '没有自用取共用
                If rsTmp!剩余数量 = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                CheckUsedBill = rsTmp!ID
                blnTmp = True
            End If
        Else
            '自用票据
            If rsTmp!剩余数量 > 0 Then
                '有剩余
                CheckUsedBill = rsTmp!ID
            Else
                '其它有剩余的自用
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '其它自用也没有剩余
                CheckUsedBill = rsSelf!ID
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)))) <> UCase(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)) Then
                CheckUsedBill = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!开始号码) And UCase(strBill) <= UCase(rsTmp!终止号码) And Len(strBill) = Len(rsTmp!开始号码)) Then
                CheckUsedBill = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            rsSelf.Filter = "ID=" & CheckUsedBill
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                CheckUsedBill = -3
                rsSelf.Filter = "ID<>" & CheckUsedBill
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then CheckUsedBill = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckUsedBill = 0
End Function


Public Function GetNextBill(lng领用ID As Long) As String
'功能：根据领用批次ID,获取下一个实际票据号
'说明：当取不到范围内的有效票据时,返回空由用户输入
    Dim rsTmp As New ADODB.Recordset
    Dim strBill As String
    
    On Error GoTo errH
    
    gstrSQL = "Select 前缀文本,开始号码,终止号码,当前号码 From 票据领用记录 Where 剩余数量>0 And ID=[1] "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlRegEventSelf", lng领用ID)
    
    If rsTmp.EOF Then Exit Function
    
    If IsNull(rsTmp!当前号码) Then
        strBill = UCase(rsTmp!开始号码)
    Else
        strBill = UCase(IncStr(rsTmp!当前号码))
    End If
    '检查范围
    If Left(strBill, Len(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本))) <> UCase(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)) Then
        Exit Function
    ElseIf Not (strBill >= UCase(rsTmp!开始号码) And strBill <= UCase(rsTmp!终止号码)) Then
        Exit Function
    End If
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte

    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function


'下面为添加的函数
Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, _
    Optional ByVal TW As Long, Optional ByVal TH As Long, Optional BorderColor As Long, _
    Optional ForeColor As Long, Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont, _
    Optional Border As String = "1111", Optional HAlign As Byte, Optional VAlign As Byte = 1, Optional Warp As Boolean = True, _
    Optional Ratio As Single = 1) As Boolean
'功能：在指定设备上按指定格式集输出文字或图象
'参数：
'   Dev=输出设备,为Printer或PictureBox对象
'   Data=输出内容,为线条(x)、字符串("xxx")或图象(stdPicture)。字符串不包含vbCrLf,当Data类型为数字型时,表示输出线条
'   TW,TH=输出的限定范围,超过这个范围则自动取消或缩小,为0时无效
'   Border=边框定义,上下左右,"1111"表示全画
'   Align=文字对齐,0=左,1=中,2=右,分水平对齐及垂直对齐
'   Warp=当输出内容为字符串时,表示是否自动换行。不自动换行时,超宽部份不输出。
'   Ratio=输出比例,对字体,坐标都有影响,缺省为1(100%)
'说明：1.在使用该函数之前,应该没有改变设备的作图初始值
'      2.输出后定位光标位置在本次输出范围的右上角
    Dim i As Long, Text As String, arrText() As String
    Dim LINE_W As Integer, blnW As Boolean, blnH As Boolean
    
    On Error GoTo errH
    
    DrawCell = True
    
    '范围限定
    If TW > 0 Then
        If X > TW Then Exit Function
        If X + W > TW Then W = TW - X
    End If
    If TH > 0 Then
        If Y > TH Then Exit Function
        If Y + H > TH Then H = TH - Y
    End If
    
    If TypeName(Data) = "Integer" Then
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B '矩形
        Else
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF '实心矩形(线条)
        End If
    ElseIf TypeName(Data) = "String" Then
        '字体
        If Font Is Nothing Then
            Set Font = New StdFont
            Font.Name = "宋体"
            Font.Size = 9
        End If
        '千万不要用Set Dev.Font=Font,不知为何,用的是ByVal
        Dev.Font.Name = Font.Name
        Dev.Font.Size = Font.Size
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        
        '因缩放后可能字体比例不对,判断时以原始大小为准
        If H >= Dev.TextHeight(Replace(Data, vbCrLf, "")) Then blnH = True '高度是否够用(加回车的算一行高度)
        If W >= Dev.TextWidth(Data) Then blnW = True And InStr(Data, vbCrLf) = 0 '宽度是否够用(加回车的为不够用,以便拆行)
        
        '缩变
        LINE_W = 30 * Ratio '边线间隔宽度(输出时用,判断时不用)
        X = -Int(-X * Ratio): Y = -Int(-Y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        
        '背景填充
        Dev.Line (X, Y)-(X + W, Y + H), BackColor, BF
        
        Dev.ForeColor = ForeColor
        '输出文字(边框之内再隔一线)
        '超出高度范围则不输出
        If blnH Then
            If blnW Then
                Select Case HAlign
                    Case 0
                        Dev.CurrentX = X + LINE_W
                    Case 1
                        Dev.CurrentX = X + (W - Dev.TextWidth(Data)) / 2
                    Case 2
                        Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Data)
                End Select
                Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Data)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Data)
                End Select
                Dev.Print Data
            Else
                If Not Warp Then
                    '不自动拆行时超宽部分不输出
                    For i = 1 To Len(Data)
                        If Dev.TextWidth(Text & Mid(Data, i, 1)) > W Then Exit For
                        Text = Text & Mid(Data, i, 1)
                    Next
                    Select Case HAlign
                        Case 0
                            Dev.CurrentX = X + LINE_W
                        Case 1
                            Dev.CurrentX = X + (W - Dev.TextWidth(Text)) / 2
                        Case 2
                            Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Text)
                    End Select
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + LINE_W
                        Case 1
                            Dev.CurrentY = Y + (H - Dev.TextHeight(Text)) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Text)
                    End Select
                    '输出截取部份
                    Dev.Print Text
                Else
                    '拆分文字成多行(在宽高范围内)
                    ReDim arrText(0) '在此,第一行不可能超高
                    Data = Replace(Data, vbCrLf, vbCr)
                    Data = Replace(Data, vbLf, vbCr)
                    For i = 1 To Len(Data)
                        If Mid(Data, i, 1) = vbCr Then
                            '多行超高则退出,超高部份不输出
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) > W Then
                            '多行超高则退出,超高部份不输出
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '有可能一行一个字符宽度都不够
                        If Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) <= W And Mid(Data, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(Data, i, 1)
                        End If
                    Next
                    
                    '输出起始坐标
                    Select Case VAlign
                        Case 0
                            Dev.CurrentY = Y + LINE_W
                        Case 1
                            Dev.CurrentY = Y + (H - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2
                        Case 2
                            Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)
                    End Select
                    
                    '输出各行
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                            Case 0
                                Dev.CurrentX = X + LINE_W
                            Case 1
                                Dev.CurrentX = X + (W - Dev.TextWidth(arrText(i))) / 2
                            Case 2
                                Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(arrText(i))
                        End Select
                        Dev.Print arrText(i)
                    Next
                End If
            End If
        End If
    ElseIf Not Data Is Nothing Then
        LINE_W = 30 * Ratio '边线间隔宽度(输出时用,判断时不用)
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        
        '图形(边框之内)
        Dev.PaintPicture Data, X + 15, Y + 15, W - LINE_W, H - LINE_W
    End If
    
    If TypeName(Data) <> "Integer" Then
        '最后处理边框
        If Mid(Border, 1, 1) Then Dev.Line (X, Y)-(X + W, Y), BorderColor
        If Mid(Border, 2, 1) Then Dev.Line (X, Y + H)-(X + W, Y + H), BorderColor
        If Mid(Border, 3, 1) Then Dev.Line (X, Y)-(X, Y + H), BorderColor
        If Mid(Border, 4, 1) Then Dev.Line (X + W, Y)-(X + W, Y + H), BorderColor
    End If
    Exit Function
errH:
    DrawCell = False
End Function

