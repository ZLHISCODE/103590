Attribute VB_Name = "mdlTransfusion"
Option Explicit
Public gblnShowInTaskBar As Boolean         '是否显示窗体在任务条上

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrSysName As String                '系统名称
Public gstrProductName As String            'OEM产品名称
Public glngSys As Long                      '系统编号
Public glngModul As Long                    '模块号
Public gstrPrivs As String                  '当前用户具有的当前模块的功能
Public gcolPrivs As Collection              '记录内部模块的权限

Public gstrDBUser As String                 '当前数据库用户
Public gstrUnitName As String               '用户单位名称
'系统参数
Public gbytCardLen As Byte '就诊卡号长度
'Public gblnCardHide As Boolean '就诊卡号密文显示

Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gint挂号天数 As Integer '挂号单有效天数
Public gbln病区科室独立 As Boolean
Public gint诊断来源 As Integer '1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Public gint诊断输入 As Integer '1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Public gbln执行后审核 As Boolean    '执行后自动审核划价单
Public gbln消费验证 As Boolean '门诊一卡通消费减少剩余款额时是否需要验证
Public gobjPlugIn As Object
Public gstr药品价格等级 As String '院区的药品价格等级
Public gstr卫材价格等级 As String '院区的卫材价格等级
Public gstr普通项目价格等级 As String '院区的普通项目价格等级

Public gstr医嘱核对 As String    '输血皮试医嘱需要核对 按位存取11，第一位为 输血医嘱，第二位为 皮试医嘱

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum enuCardProperty
    短名 = 0
    全名 = 1
    可读卡 = 2
    卡类别ID = 3
    卡号长度 = 4
    缺省类别 = 5
    存在帐户 = 6
    卡号密文显示 = 7
End Enum

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.部门ID = zlCommFun.NVL(rsTmp!部门ID, 0)
            UserInfo.简码 = zlCommFun.NVL(rsTmp!简码)
            UserInfo.姓名 = zlCommFun.NVL(rsTmp!姓名)
            GetUserInfo = True
        End If
    End If
End Function

Public Function GetSquareCardInfo(ByVal strSquareCards As String, ByVal strCardName As String, ByVal intElement As Integer) As String
'功能：取一卡通卡号的长度
'参数：
'  strSquareCards：一卡通卡号信息
'  strCardName：指定提取的卡号名称
'  intElement：指定取一卡通信息串的元素
'返回：卡号长度
    
    If strSquareCards = "" Then Exit Function
    
    Dim i As Integer
    Dim arrInfo As Variant
    Dim strTmp As String
    
    GetSquareCardInfo = ""
    
    On Error GoTo errHandle
    arrInfo = Split(strSquareCards, ";")
    For i = LBound(arrInfo) To UBound(arrInfo)
        strTmp = Split(arrInfo(i), "|")(enuCardProperty.全名)
        If strCardName = strTmp Then
            GetSquareCardInfo = Split(arrInfo(i), "|")(intElement)
            Exit For
        End If
    Next
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPara As String
    
    On Error GoTo errH
        
    '就诊卡号码的长度
    gbytCardLen = 7
    strSQL = "select 卡号长度 from 医疗卡类别 where 名称=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "取就诊卡长度", "就诊卡")
    If Not rsTmp.EOF Then
        gbytCardLen = IIf(IsNull(rsTmp!卡号长度), 7, rsTmp!卡号长度)
    End If
    
    'HIS系统参数
    
    '挂号有效天数
    strPara = zlDatabase.GetPara(21, glngSys)
    gint挂号天数 = zlCommFun.NVL(strPara, 0)
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    gbytBillOpt = zlCommFun.NVL(zlDatabase.GetPara(23, glngSys), 0)
    
    '诊断输入来源
    gint诊断来源 = zlCommFun.NVL(zlDatabase.GetPara(55, glngSys), 1)
    
    '诊断输入方式
    gint诊断输入 = zlCommFun.NVL(zlDatabase.GetPara(65, glngSys), 1)
    '病区和科室是否独立管理
    gbln病区科室独立 = Val(zlDatabase.GetPara(99, glngSys)) <> 0
    '一卡通消费验证
    gbln消费验证 = Val(zlDatabase.GetPara(28, glngSys)) <> 0 '门诊病人消费时需要刷卡验证
    
    '项目执行前必须先收费或先记帐审核
    
    '输血和皮试医嘱执行后需要核对
    gstr医嘱核对 = zlDatabase.GetPara(186, glngSys)
    
    '执行后自动审核
    gbln执行后审核 = Val(zlDatabase.GetPara(81, glngSys)) <> 0 '执行后自动审核划价单
    Call InitPriceLevel
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    InitSysPar = False
End Function

Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

'Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
''功能：获取指定控件在屏幕中的位置(Twip)
'    Dim vRect As RECT
'    Call GetWindowRect(lngHwnd, vRect)
'    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
'    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
'    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
'    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
'    GetControlRect = vRect
'End Function

Public Function CacleTransTime(ByVal 液体总量 As Long, ByVal 滴系数 As Long, ByVal 每分钟滴速 As Integer) As Integer
    '计算输液时间
    '输液时间(分钟)=(液体总量(ml)×滴系数)/(每分钟滴速)
    If 每分钟滴速 > 0 Then
        CacleTransTime = (液体总量 * 滴系数) / 每分钟滴速
    End If
End Function

Public Function GetAdvicePause(ByVal lng医嘱ID As Long) As String
'功能：获取指定医嘱的暂停时间段记录
'返回："暂停时间,开始时间;...."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    strSQL = "Select 操作类型,操作时间 From 病人医嘱状态" & _
        " Where 操作类型 IN('6','7') And 医嘱ID=[1]" & _
        " Order by 操作时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!操作类型 = "6" Then
            strTmp = strTmp & ";" & Format(rsTmp!操作时间, "yyyy-MM-dd HH:mm:ss") & ","
        ElseIf rsTmp!操作类型 = "7" Then
            '启用的那一秒不在暂停的范围之内
            strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!操作时间), "yyyy-MM-dd HH:mm:ss")
        End If
        rsTmp.MoveNext
    Next
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个日期是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
'说明：不按时点判断,对暂停日期按算始不算止规则判断
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '可能尚未启用或暂停的时候被停止
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function Between(X, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function Calc本周期开始时间(ByVal dat开始执行时间 As Date, ByVal dat某次执行时间 As Date, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String) As Date
'功能：根据长嘱的某次执行时间，得到它在该周期内的开始基准时间
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat开始执行时间
    datBegin = datCurr
    If str间隔单位 = "周" Then datCurr = Format(datCurr - (Weekday(datCurr, vbMonday) - 1), "yyyy-MM-dd 00:00:00")
    
    Do While datCurr <= dat某次执行时间
        datBegin = datCurr
        If str间隔单位 = "周" Then
            datCurr = datCurr + 7
        ElseIf str间隔单位 = "天" Then
            datCurr = datCurr + int频率间隔
        ElseIf str间隔单位 = "小时" Then
            datCurr = DateAdd("h", int频率间隔, datCurr)
        End If
    Loop
    Calc本周期开始时间 = datBegin
End Function

Public Function Calc段内分解时间(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal dat首日日期 As Date) As String
'功能：按时间段计算各次的分解执行时间及次数
'参数：datBegin-datEnd=要计算的时间段,其中datBegin应为每个周期的开始基准时间
'      strPause=暂停的时间段
'      dat首日日期=用于首日时间计算参照
'返回："时间1,时间2,...."(yyyy-MM-dd HH:mm:ss),时间个数即为次数
'说明：1.时间段内要排除暂停的时间段,次数可能因此而减少
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(str执行时间, ",") > 0 Then
        arrNormal = Split(Split(str执行时间, ",")(1), "-")
        arrFirst = Split(Split(str执行时间, ",")(0), "-")
    Else
        arrNormal = Split(str执行时间, "-")
        arrFirst = Array()
    End If
        
    vCurTime = datBegin
    
    If str间隔单位 = "周" Then
        vCurTime = zlCommFun.GetWeekBase(datBegin)
        If dat首日日期 <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = zlCommFun.GetWeekBase(dat首日日期))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False
                        
            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                If i - 1 <= UBound(arrTime) Then '首周可能次数不足
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "天" Then
        If dat首日日期 <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat首日日期))
        Else
            blnFirst = False
        End If
        
        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False
            
            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + int频率间隔, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= Format(datBegin, "yyyy-MM-dd HH:mm:ss") And vTmpTime <= Format(datEnd, "yyyy-MM-dd HH:mm:ss") Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + int频率间隔 / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str间隔单位 = "分钟" Then
        '无执行时间
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime
            
            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + int频率间隔 / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If
    
    Calc段内分解时间 = Mid(strDetailTime, 2)
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '可能尚未启用或暂停的时候被停止
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

'Public Function GetOwner(ByVal lngSys As Long) As String
''功能：获取指点系统的所有者
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL  As String
'
'    On Error GoTo errH
'    strSQL = "Select 所有者 From zlSystems Where 编号=[1]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetOwner", lngSys)
'    If Not rsTmp.EOF Then
'        GetOwner = rsTmp!所有者
'    End If
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '功能：由用户输入的部份单号，返回全部的单号。
    '参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intNum)
    If Not rsTmp.EOF Then
        intType = zlCommFun.NVL(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01"), "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'Public Function Decode(ParamArray arrPar() As Variant) As Variant
''功能：模拟Oracle的Decode函数
'    Dim varValue As Variant, i As Integer
'
'    i = 1
'    varValue = arrPar(0)
'    Do While i <= UBound(arrPar)
'        If i = UBound(arrPar) Then
'            Decode = arrPar(i): Exit Function
'        ElseIf varValue = arrPar(i) Then
'            Decode = arrPar(i + 1): Exit Function
'        Else
'            i = i + 2
'        End If
'    Loop
'End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function DelInvalidChar(ByVal strChar As String, Optional ByVal strInvalidChar As String) As String
    '删除非法字符
    'strChar: 要处理的字符
    'strInvalidChar：非法字符串，如果为空，则为~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,否则按传入的字符处理
    Dim strBit As String, i As Integer, strWord As String
    strWord = "~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>"
    If strInvalidChar <> "" Then strWord = strInvalidChar
    If Len(strChar) > 0 Then
        For i = 1 To Len(strChar)
            strBit = Mid$(strChar, i, 1)
            If InStr(strWord, strBit) <= 0 Then
                DelInvalidChar = DelInvalidChar & strBit
            End If
        Next
    End If
End Function

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
End Function

'Public Function GetAdviceMoney(ByVal str组ID As String, ByVal str医嘱ID As String, ByVal str发送号 As String, _
'    str类别 As String, str类别名 As String, ByVal bln单独执行 As Boolean, ByVal byt来源 As Byte) As Currency
''功能：根据指定的医嘱ID串，获取医嘱对应未审核的记帐费用合计
''参数：str组ID,str医嘱ID,str发送号="ID1,ID2,..."
''      bln单独执行=检验项目单独执行，这时只有一个医嘱ID
''      byt来源，1:门诊，2-住院
''返回：str类别,str类别名=用于报警提示
''说明：当系统参数为执行后审核费用时才返回。
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL As String, curMoney As Currency
'    Dim strTab As String
'
'    str类别 = "": str类别名 = ""
'
'    On Error GoTo errH
'
'    If zldatabase.GetPara(81, glngSys) <> "1" Then Exit Function
'    strTab = IIf(byt来源 = 1, "门诊费用记录", "住院费用记录")
'
'    If bln单独执行 Then
'        strSQL = _
'            " Select B.编码,B.名称,Sum(A.实收金额) as 金额" & _
'            " From " & strTab & " A,收费项目类别 B" & _
'            " Where A.医嘱序号 + 0 = [2] And (A.记录性质, A.NO) In" & _
'            "      (Select 记录性质, NO From 病人医嘱附费 Where 医嘱id = [2] And 发送号 + 0 = [3]" & _
'            "       Union All" & _
'            "       Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = [2] And 发送号 + 0 = [3])" & _
'            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
'            " Group by B.编码,B.名称"
'    Else
'        strSQL = _
'            " Select B.编码,B.名称,Sum(A.实收金额) as 金额" & _
'            " From " & strTab & " A,收费项目类别 B" & _
'            " Where A.医嘱序号 + 0 In" & _
'            "      (Select ID From 病人医嘱记录" & _
'            "       Where ID In (Select Column_Value From Table(f_Num2list([1])))" & _
'            "       Union All" & _
'            "       Select ID From 病人医嘱记录" & _
'            "       Where 相关id In (Select Column_Value From Table(f_Num2list([1]))))" & _
'            "  And (A.记录性质, A.NO) In" & _
'            "      (Select 记录性质, NO From 病人医嘱附费" & _
'            "       Where 医嘱id In" & _
'                "      (Select ID From 病人医嘱记录" & _
'                "       Where ID In (Select Column_Value From Table(f_Num2list([1])))" & _
'                "       Union All" & _
'                "       Select ID From 病人医嘱记录" & _
'                "       Where 相关id In (Select Column_Value From Table(f_Num2list([1]))))" & _
'            "         And 发送号 + 0 In (Select Column_Value From Table(f_Num2list([3])))" & _
'            "       Union All" & _
'            "       Select 记录性质, NO From 病人医嘱发送" & _
'            "       Where 医嘱id In (Select Column_Value From Table(f_Num2list([2])))" & _
'            "         And 发送号 + 0 In (Select Column_Value From Table(f_Num2list([3]))))" & _
'            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
'            " Group by B.编码,B.名称"
'    End If
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str组ID, str医嘱ID, str发送号, glngSys)
'
'    curMoney = 0
'    Do While Not rsTmp.EOF
'        curMoney = curMoney + Val("" & rsTmp!金额)
'        str类别 = str类别 & rsTmp!编码
'        str类别名 = str类别名 & "," & rsTmp!名称
'        rsTmp.MoveNext
'    Loop
'
'    str类别名 = Mid(str类别名, 2)
'    GetAdviceMoney = curMoney
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function OneCardCheck(ByVal lng医嘱ID_IN As Long, ByVal lng发送号_IN As Long, _
                             Optional frmMain As Object, Optional objCardSquare As Object) As Integer
    '一卡通处理函数
    '返回: 0- 按老流程走，1-按新流程走，成功，2-按新流程走，失败
    
    Dim lng病人ID As Long, curMoney As Currency, strSQL As String
    Dim str类别 As String, str类别名 As String, strNO As String, lng记录性质 As Long
    Dim rsTmp As ADODB.Recordset
    On Error GoTo hErr
    OneCardCheck = 2    '默认新流程失败
    
    strSQL = "Select 病人ID,诊疗类别 From 病人医嘱记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "医嘱执行完成", lng医嘱ID_IN)
    If Not rsTmp.EOF Then
        lng病人ID = Val("" & rsTmp!病人ID)
        str类别 = Trim("" & rsTmp!诊疗类别)
    End If
    
    strSQL = "Select No,记录性质 From 病人医嘱发送 Where 医嘱id=[1] And 发送号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "医嘱执行完成", lng医嘱ID_IN, lng发送号_IN)
    If Not rsTmp.EOF Then
        strNO = Trim("" & rsTmp!NO)
        lng记录性质 = Val("" & rsTmp!记录性质)
    End If
    
    If zlDatabase.GetPara("项目执行前必须先收费或先记帐审核", glngSys) = 1 Then
        If objCardSquare Is Nothing Then
            '一卡通消费部件未创建成功！
            MsgBox "一卡通消费部件未创建成功！", vbQuestion, frmMain.Caption
            Exit Function
        End If
        '新的一卡通流程
        
        '1-记帐模式
        If lng记录性质 = 2 Then
            If gbln执行后审核 Then
                '原来的功能处理
                OneCardCheck = 0
            Else
                '1.刷卡提取病人
                '2.是否存在未审核的划价单
                'If ItemHaveCash(1, False, lng医嘱ID_IN, lng医嘱ID_IN, lng发送号_IN, str类别, strNO, lng记录性质, 0, 0) Then
                    '记帐审核 zlSquareAffirm
'                    frmMain Object  In  传入调用对象
'                    lngModule   Long    IN  调用的模块号
'                    strPrivs    String  In  权限串
'                    lngPatiID   Long    In  病人ID,可以不传,在本接口窗体中刷卡!
                     If Not objCardSquare.zlSquareAffirm(frmMain, glngModul, gstrPrivs, lng病人ID, , False, , , lng医嘱ID_IN) Then
                        MsgBox "交易失败，不能执行后面的操作！", vbInformation, frmMain.Caption
                        Exit Function
                     End If
                'End If
            End If
        ElseIf lng记录性质 = 1 Then
            '刷卡提取病人，是否存在未收费的划价单
            'Dim strIDs As String
            'If ItemHaveCash(1, False, lng医嘱ID_IN, lng医嘱ID_IN, lng发送号_IN, str类别, strNO, lng记录性质, 0, 0, , , strIDs) Then

                If Not objCardSquare.zlSquareAffirm(frmMain, glngModul, gstrPrivs, lng病人ID, , False, , , lng医嘱ID_IN) Then
                                                                                              
                    MsgBox "交易失败，不能执行后面的操作！", vbInformation, frmMain.Caption
                    Exit Function
                End If
            'End If
        End If
        OneCardCheck = 1    '新流程成功
    Else
        OneCardCheck = 0    '项目执行前必须先收费或先记帐审核 参数未勾，走老流程
    End If
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Private Function ItemHaveCash(ByVal int病人来源 As Integer, ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, _
'    ByVal lng发送号 As Long, ByVal str类别 As String, ByVal str单据号 As String, ByVal int记录性质 As Integer, ByVal int门诊记帐 As Integer, ByVal int方式 As Integer, _
'    Optional ByVal blnMove As Boolean, Optional ByVal dat发送时间 As Date, Optional ByRef str医嘱IDs As String, Optional ByRef strNOs As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
''功能：判断当前的执行医嘱是否已收费或记帐划价单是否已审核
''参数：int病人来源=1-门诊,2-住院
''      str类别=诊疗类别，用于从一组医嘱中区分分开执行的内容
''      int方式=0-检查是否存在未收费记录
''              1-检查是否存在已收费记录
''      int门诊记帐=1=住院发送到门诊记帐
''      返回：str医嘱IDs=该医嘱及相关的医嘱ID,NOs=医嘱发送的单据号和补的附费中的单据号
'    Dim rsTmp As New ADODB.Recordset
'    Dim strSQL As String, strTab As String
'
'    If int病人来源 = 2 And int记录性质 = 2 And int门诊记帐 = 0 Then
'        strTab = "住院费用记录"
'    Else
'        strTab = "门诊费用记录"
'    End If
'    ItemHaveCash = True
'    str医嘱IDs = ""
'    strNOs = ""
'
'    '对应的费用中是否存在未收费[或已作废]的内容
'    '和清单只显示已收费不同：
'    '1.检查了医嘱附费(不加记录性质的条件，因为可能补收费单或记帐单)
'    '2.记帐划价也显示为未收(清单需要先显出来执行后审核)
'    '3.按NO对应到相关医嘱的费用检查(清单是按显示的医嘱ID)
'    strSQL = _
'        " Select A.记录状态,Nvl(B.相关ID,B.ID) as 医嘱ID,B.诊疗类别,A.执行状态,A.NO" & _
'        " From " & strTab & " A,病人医嘱记录 B" & _
'        " Where A.NO=[4] And A.记录状态 IN(0,1,3) And A.医嘱序号+0=B.ID And A.记录性质=[5]" & IIf(bln单独执行, " And B.ID=[2]", "") & _
'        " Union ALL " & _
'        " Select B.记录状态,Nvl(C.相关ID,C.ID) as 医嘱ID,C.诊疗类别,B.执行状态,A.NO" & _
'        " From 病人医嘱记录 C," & strTab & " B,病人医嘱附费 A" & _
'        " Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱ID=B.医嘱序号+0" & IIf(bln单独执行, " And A.医嘱ID=[2]", _
'            " And A.医嘱ID IN (Select ID From 病人医嘱记录 Where (ID=[1] Or 相关ID=[1]) And 诊疗类别=[6])") & _
'        " And A.发送号=[3] And B.记录状态 IN(0,1,3) And A.医嘱ID=C.ID And A.记录性质=[5]"
'    If blnMove Then
'        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
'        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
'        strSQL = Replace(strSQL, strTab, "H" & strTab)
'    ElseIf zldatabase.DateMoved(dat发送时间) Then
'        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
'    End If
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID), lng医嘱ID, lng发送号, str单据号, int记录性质, str类别)
'    If Not rsTmp.EOF Then
'        If int方式 = 0 Then
'            rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 执行状态=9"
'            If Not rsTmp.EOF Then
'                blnIsAbnormal = True
'                ItemHaveCash = False
'            Else
'                rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态=0"
'                If Not rsTmp.EOF Then ItemHaveCash = False
'            End If
'
'            While Not rsTmp.EOF
'                If InStr("," & str医嘱IDs & ",", "," & rsTmp!医嘱ID & ",") = 0 Then
'                    str医嘱IDs = str医嘱IDs & "," & rsTmp!医嘱ID
'                End If
'                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
'                    strNOs = strNOs & "," & rsTmp!NO
'                End If
'                rsTmp.MoveNext
'            Wend
'            strNOs = Mid(strNOs, 2)
'            str医嘱IDs = Mid(str医嘱IDs, 2)
'        ElseIf int方式 = 1 Then
'            rsTmp.Filter = "医嘱ID=" & IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID) & " And 诊疗类别='" & str类别 & "' And 记录状态<>0 And 执行状态<>9"
'            If rsTmp.EOF Then ItemHaveCash = False
'        End If
'    ElseIf int方式 = 1 Then
'        ItemHaveCash = False
'    End If
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function GetMaxNoAddOne(ByVal strField As String, ByVal strTableAndWhere As String) As String
    '2012-06-04
    '获取指定表，指定字段的最大值加一，一般用于初始化时自动产生的编号
    Dim strSQL As String, rsTmp As ADODB.Recordset, strMaxNO As String
    On Error GoTo hErr
    strMaxNO = ""
    strSQL = "Select Max(" & strField & ") as MaxNo From " & strTableAndWhere
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxNoAddOne")
    Do Until rsTmp.EOF
        strMaxNO = Trim$("" & rsTmp!MaxNo)
        rsTmp.MoveNext
    Loop
    If strMaxNO <> "" Then
        If IsNumeric(strMaxNO) Then
            strMaxNO = Format(Val(strMaxNO) + 1, String(Len(strMaxNO), "0"))
        Else
            strMaxNO = zlCommFun.IncStr(strMaxNO)
        End If
    Else
        strMaxNO = "001"
    End If
    If strMaxNO <> "" Then GetMaxNoAddOne = strMaxNO
    Exit Function
hErr:
    GetMaxNoAddOne = ""
End Function

'Public Function GetDeptInListPara(ByVal strParaName As String, ByVal lngDeptID As Long) As Boolean
'    '获取无线输液中的几个开关参数
''        无线输液_标准穿刺列表
''        无线输液_呼叫科室列表
''        无线输液_简单穿刺列表
''        无线输液_配液科室列表
''        无线输液_巡视科室列表
'    Dim strTmp As String
'    strTmp = zlDatabase.GetPara(strParaName, glngSys)
'    If strTmp <> "" Then
'        If Left(strTmp, 1) <> "," Then strTmp = "," & strTmp
'        If Right(strTmp, 1) <> "," Then strTmp = strTmp & ","
'        GetDeptInListPara = InStr(strTmp, "," & lngDeptID & ",") > 0
'    Else
'        GetDeptInListPara = False
'    End If
'End Function

Public Function AllocationDesks(ByVal lngDeptID As Long, ByVal objPati As cPatient, _
        ByRef strSeqNo As String, ByRef strErr As String) As Boolean
    '配液时 分配穿刺台
    '传入：
    '   lngDeptID  :科室ID
    '   objPati    :病人信息对象
    '传出：
    '   strSeqNo   :分配的穿刺台序号
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSeqNo = "": strErr = ""
    strSQL = "Zl_门诊穿刺台_Liquid(" & lngDeptID & "," & objPati.病人ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "分配穿刺台")
    
    strSQL = "select 穿刺台, 状态 From 排队记录 Where 科室ID=[1] And 病人id=[2] order by 日期 desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "分配穿刺台", lngDeptID, objPati.病人ID)

    If rsTmp.EOF = False Then
        If zlCommFun.NVL(rsTmp!穿刺台) = "" Then
            strErr = "没有穿刺台，请先设置穿刺台后再使用此功能！"
        Else
            strSeqNo = CStr(rsTmp!穿刺台)
            SaveOperLog lngDeptID, objPati, QUEUE, "分配到" & strSeqNo & "号穿刺台,当前状态为" & Trim$("" & rsTmp!状态)
        End If
    End If
    
    If strSeqNo = "" Then
        strErr = "分配失败，请稍后再试！"
    Else
        AllocationDesks = True
    End If
    Exit Function
    
hErr:
    AllocationDesks = False
    strErr = Err.Description
End Function

Public Function CurDayHaveItem(ByVal objPati As cPatient, ByVal lngDeptID As Long) As Boolean
    '判断当天是否接过输液的单
    '反回：　True 接过输液单　Ｆalse　未接过输液单
    Dim objExe As New ExecRecord
    Dim dateS As Date, dateE As Date
    Dim i As Integer, Y As Integer
    Dim blnNoCall As Boolean
   
    Dim blnHaveItem As Boolean
    dateS = Format(objPati.挂号时间, "yyyy-MM-dd 00:00:00")
    dateE = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    
    
    blnHaveItem = False
    Call objExe.GetExecGroups(objPati, lngDeptID, 1, dateS, dateE)
    
    dateS = Format(dateE, "yyyy-MM-dd 00:00:00")
    For i = 1 To objExe.Count - 1
        If objExe.Item(i).执行分类 = "1-输液" And objExe.Item(i).执行时间 >= dateS And objExe.Item(i).执行时间 <= dateE And objPati.接受时间 >= dateS And objPati.接受时间 <= dateE Then
            '当天接过输液的单，就不调 状态了
            blnHaveItem = True
            Exit For
        End If
    Next
    
    CurDayHaveItem = blnHaveItem
    
End Function

Public Function CurDayNoCall(ByVal lngDeptID As Long, ByVal objPatients As cPatients, ByVal objCurPati As cPatient) As Boolean
    '2012-09-25 检查当天是否已经有挂两次号的情况
    '返回：True 当天已经接过输液单　（不需要呼叫，不改状态）　False：当天未接过输液单（需要呼叫，需要重新改状态）
    Dim dateS As Date, dateE As Date
    Dim blnNoCall As Boolean, objPati As cPatient
    
    dateS = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dateE = Format(dateS, "yyyy-MM-dd 23:59:59")
    blnNoCall = False
    For Each objPati In objPatients
        If objPati.病人ID = objCurPati.病人ID And _
            objPati.挂号单 <> objCurPati.挂号单 And _
            objPati.接受时间 >= dateS And objPati.接受时间 <= dateE And _
            (Val(objPati.排队状态) = 5 Or Val(objPati.排队状态) = 7) Then
            If Val(objPati.排队状态) = 5 Then
                blnNoCall = True
            Else
                blnNoCall = CurDayHaveItem(objPati, lngDeptID)
            End If
            SaveOperLog lngDeptID, objCurPati, QUEUE, "此人当天已经有" & objPati.排队状态 & "记录，挂号单为" & objPati.挂号单
        End If
    Next
End Function

Public Function Liquid(ByVal lngDept As Long, ByVal strNO As String, ByVal objPatiList As cPatients, ByRef strErr As String) As String
'配液过程
    
    Dim strSQL As String, strSeqNo As String
    Dim blnNoCall As Boolean
    Dim objPati As cPatient
    
    strErr = ""
    If strNO <> "" Then
        Err.Clear
        On Error Resume Next
        Set objPati = objPatiList.Item(strNO)
        If objPati Is Nothing Or Err.Number <> 0 Then
            '进入本逻辑，表明“排队状态=4：结束；=3：退号；=2：弃号”这类数据。因为FetchPatients只取1，5，6，7病人状态的排队数据
            Liquid = "5-待穿刺"
            Err.Clear
            SaveOperLog lngDept, strNO, QUEUE, "5-待穿刺"
            Exit Function
        End If
        
        strSeqNo = objPati.挂号单
        If Err.Number <> 0 Then
            Liquid = "5-待穿刺"
            Err.Clear
            SaveOperLog lngDept, strSeqNo, QUEUE, "挂号单"
            Exit Function
        End If
        
        blnNoCall = CurDayNoCall(lngDept, objPatiList, objPati)
        
        If Err.Number <> 0 Then
            Liquid = "5-待穿刺"
            SaveOperLog lngDept, strNO, QUEUE, "Call异常"
            Exit Function
        End If
        
        If Not blnNoCall Then
            '分配穿刺台功能提前到接单中处理 2012-10-10
'            If Not AllocationDesks(lngDept, strNo, strSeqNo, strErr) Then
'                Exit Function
'            End If
            Liquid = "5-待穿刺"
        Else
            Liquid = "7-执行中"
        End If
    Else
        strErr = "请选择一条记录后再执行此操作!"
    End If

End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, strLabel As String, intResult As Integer)
'功能：获取皮试标注和结果
'参数：strScript=皮试结果描述串，如"阳性(+),大阳性(++);阴性(-)"
'      strSelect=所选择的皮试结果中文名，如"阳性"
'返回：strLabel = 皮试结果标注，如"(+)"
'      intResult=皮试结果：0-阴性，1-阳性
    Dim arr阳性 As Variant, arr阴性 As Variant
    Dim i As Integer
    
    strLabel = "": intResult = 0
    
    arr阳性 = Split(Split(strScript, ";")(0), ",")
    arr阴性 = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr阳性)
        If arr阳性(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr阳性(i), Len(strSelect) + 1)
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr阴性)
        If arr阴性(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr阴性(i), Len(strSelect) + 1)
            intResult = 0: Exit Sub
        End If
    Next
End Sub

Public Function GetPriceGradeSQL(ByVal str药品价格等级 As String, ByVal str卫材价格等级 As String, ByVal str普通项目价格等级 As String, ByVal strTableTmpA As String, ByVal strTableTmpB As String, _
           ByVal strParNum药品 As String, ByVal strParNum卫材 As String, ByVal strParNum普通项目 As String) As String
'功能：病人价格等级获得批量获取价格的SQL
'参数：str药品价格等级  '病人的药品价格等级
'      str卫材价格等级  '病人的卫材价格等级
'      str普通项目价格等级  '病人的普通项目价格等级
'     strTableTmpA   收费项目目录 表的as 标志,strTableTmpB  收费价目表 的As标志；
'     strParNum药品  药品价格等级SQL参数序号,strParNum卫材  卫材价格等级SQL参数序号,strParNum普通项目  普通项目价格等级SQL参数序号
    Dim strSQL As String
    
    If str药品价格等级 = "" And str卫材价格等级 = "" And str普通项目价格等级 = "" Then
        strSQL = " And " & strTableTmpB & ".价格等级 is Null "
    Else
        strSQL = " And" & vbNewLine & _
                "      ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And " & strTableTmpB & ".价格等级 = [" & strParNum药品 & "]) Or" & vbNewLine & _
                "      (Instr(';4;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And " & strTableTmpB & ".价格等级 = [" & strParNum卫材 & "]) Or" & vbNewLine & _
                "      (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') = 0 And " & strTableTmpB & ".价格等级 = [" & strParNum普通项目 & "]) Or" & vbNewLine & _
                "      (" & strTableTmpB & ".价格等级 Is Null And Not Exists" & vbNewLine & _
                "       (Select 1" & vbNewLine & _
                "         From 收费价目" & vbNewLine & _
                "         Where " & strTableTmpA & ".Id = 收费细目id  And" & vbNewLine & _
                "               ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And 价格等级 = [" & strParNum药品 & "]) Or" & vbNewLine & _
                "               (Instr(';4;', ';' || " & strTableTmpA & ".类别 || ';') > 0 And 价格等级 = [" & strParNum卫材 & "]) Or" & vbNewLine & _
                "               (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".类别 || ';') = 0 And 价格等级 = [" & strParNum普通项目 & "]))))) "

    End If
    
    GetPriceGradeSQL = strSQL
End Function

Private Sub InitPriceLevel()
'功能：初化始价格等级
    Dim objTmpExpense As Object
    
    If objTmpExpense Is Nothing Then
        On Error Resume Next
        Set objTmpExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Not objTmpExpense Is Nothing Then
            Call objTmpExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not objTmpExpense Is Nothing Then
        Call objTmpExpense.zlGetPriceGrade(zl9ComLib.gstrNodeNo, 0, 0, "", gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级)
    End If
End Sub

Public Sub PlugInFunc()
    '外挂程序对象初始化
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, 1264, -1)
        Call zlPlugInErrH(Err, "Initialize")
    End If
End Sub

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub
