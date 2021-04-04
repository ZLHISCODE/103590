Attribute VB_Name = "mdlMedical"
Option Explicit

Public Enum COLOR
    
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E0E0
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
End Enum


Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long

'医保变量
'Public gclsInsure As New clsInsure
Public gblnInsure As Boolean '是否连接医保
Public gintInsure As Integer

Public gblnStrictCtrl As Boolean
Public glngShareUseID As Long

Public gblnBill结帐 As Boolean
Public glng结帐ID As Long
Public gbytBalanceRows As Byte '结帐收据总行次

Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrUnitName As String               '用户单位名称
Public gfrmMain As Object

Public gstrSQL As String
Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号
Public gblnOK As Boolean

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO
Public glngTXTProc As Long '保存默认的消息函数的地址

'HIS系统参数
Public glngOld As Long, glngFormW As Long, glngFormH As Long

Public Type SYS_PARAM_INFO
    费用金额小数位数 As Integer
    收费诊疗项目匹配 As String
    结帐票据号长度 As Integer
    就诊卡号码长度 As Integer
    就诊卡字母前缀 As String
    就诊卡密文显示 As Boolean
    项目输入匹配方式 As Integer '0-双向;1-从左
End Type

Public ParamInfo As SYS_PARAM_INFO

Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"


Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29      '出院接口中是否要与接口商进行交易
End Enum

Public Enum CHECKFORMAT
    电子邮件
    日期
    身份证号
    数值
    自定义
End Enum

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select A.ID,C.部门ID,A.编号,A.简码,A.姓名,B.用户名" & _
        " From 人员表 A,上机人员表 B,部门人员 C" & _
        " Where A.ID = B.人员ID And A.ID = C.人员ID And C.缺省 = 1 And Upper(B.用户名) = USER"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function


Public Sub NewColumn(msf As Object, ByVal vText As String, Optional ByVal vWidth As Single = 1200, Optional ByVal vAlignment As Byte = 9)
    Dim i As Long
    
    msf.Cols = msf.Cols + 1
    i = msf.Cols - 1
    
    msf.TextMatrix(0, i) = vText
    msf.ColWidth(i) = vWidth
    msf.ColAlignment(i) = vAlignment
    
    On Error Resume Next
    msf.ColAlignmentFixed(i) = vAlignment
    
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    X = objPoint.X * 15 + objBill.CellLeft
    Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    Dim blnForeColor As Boolean
    Dim blnBkColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    blnBkColor = (rsData("背景色").Name = "背景色")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
        lngRow = 0
    Else
        
        If Val(objMsf.RowData(objMsf.Rows - 1)) <= 0 Then
            lngRow = objMsf.Rows - 2
        Else
            lngRow = objMsf.Rows - 1
        End If
                
    End If
    
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngLoop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If strMask <> "" Then
                    objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.NVL(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
                Else
                    objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.NVL(rsData(objMsf.TextMatrix(0, lngLoop)))
                End If
            End If
            
        Next
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        If blnBkColor Then objMsf.Cell(flexcpBackColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("背景色").Value)
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    Dim lngLoop As Long
    Dim strChar As String
    
    strChar = "ZXCVBNMASDFGHJKLQWERTYUIOPzxcvbnmasdfghjklqwertyuiop"
    
    Select Case bytMode
    Case 1          '全数字
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 2          '全字母
    
        For lngLoop = 1 To Len(Text)
            If InStr(strChar, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
        
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    gstrSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    gstrSQL = "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical")
    GetMaxLength = rs.Fields(0).DefinedSize

End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'功能: 装载数据入指定的组合下拉框或网格中的下拉框中
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1.Fields(2).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
End Sub

Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Sub ResetVsf(objVsf As Object)
    '
    objVsf.Rows = 2
    objVsf.RowData(1) = ""
    objVsf.Cell(flexcpText, 1, 0, 1, objVsf.Cols - 1) = ""
    
    On Error Resume Next
    
    Set objVsf.Cell(flexcpPicture, 1, 0, 1, objVsf.Cols - 1) = Nothing
End Sub

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '-----------------------------------------------------------------------------------------
    '功能:获取特殊时间
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Public Sub ClearGrid(vsf As Object, Optional ByVal Row As Long = 1)
    '--------------------------------------------------------------------------------------------------------
    '功能:清除表格数据
    '--------------------------------------------------------------------------------------------------------
    vsf.Rows = Row + 1
    vsf.RowData(Row) = 0
    vsf.Cell(flexcpText, Row, 0, Row, vsf.Cols - 1) = ""
    
End Sub

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub DeleteRecord(rs As ADODB.Recordset)
    '-----------------------------------------------------------------------------------
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '-----------------------------------------------------------------------------------
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
    
errHand:
End Sub

Public Sub CopyRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset)
    '-----------------------------------------------------------------------------------
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '-----------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Set rsTo = New ADODB.Recordset
    For lngLoop = 0 To rsFrom.Fields.Count - 1
        rsTo.Fields.Append rsFrom.Fields(lngLoop).Name, rsFrom.Fields(lngLoop).Type, rsFrom.Fields(lngLoop).DefinedSize
    Next
    rsTo.Open
    
    If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields(lngLoop).Value = rsFrom.Fields(lngLoop).Value
        Next
        rsFrom.MoveNext
    Wend
    
errHand:
    
End Sub

Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long, Optional ByVal lngBackColor As Long = -1)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    On Error Resume Next
    
    If lngBackColor = -1 Then
        lngColor = objVsf.BackColorSel
    Else
        lngColor = lngBackColor
    End If
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = lngColor
    End If
    
End Sub

Public Sub DrawLine(pic As PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
End Sub

Public Function FilterRecord(rsTmp As ADODB.Recordset, ByVal strFilter As String) As Boolean
    rsTmp.Filter = ""
    rsTmp.Filter = strFilter
    
    FilterRecord = True
End Function

'Public Sub zlDatabase.ExecuteProcedure(ByVal strSQL As String, ByVal strCaption As String)
''功能：执行SQL语句
'    Call SQLTest(App.ProductName, strCaption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
'End Sub

Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo errHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngLoop = 0 To UBound(varArray)
        varItem = Split(varArray(lngLoop), ",")
                
        objVsf.Cols = objVsf.Cols + 1
        i = objVsf.Cols - 1
    
        objVsf.TextMatrix(0, i) = varItem(0)
        objVsf.ColWidth(i) = Val(varItem(1))
        objVsf.ColAlignment(i) = Val(varItem(2))
        objVsf.ColHidden(i) = (Val(varItem(4)) = 0)
        objVsf.Cell(flexcpData, 0, i) = IIf(varItem(5) = "", varItem(0), varItem(5))
        
    Next
    
    CreateVsf = True
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function AppendSapceRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    lngTop = objVsf.Cell(flexcpTop, objVsf.Rows - 1, 0) + objVsf.RowHeight(objVsf.Rows - 1)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendSapceRows = True
    
    Exit Function
    
errHand:
    
End Function

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Public Function GetNextCode(ByVal strTable As String, Optional ByVal strField As String = "编码", Optional ByVal strFilter As String = "") As String
    Dim rs As New ADODB.Recordset
    Dim strFormat As String
    
    GetNextCode = "1"
    strFormat = "00000000000000000000"
    gstrSQL = "select nvl(max(" & strField & "),0) as 编码 from " & strTable & IIf(strFilter = "", "", " where " & strFilter)

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical")
    If rs.BOF = False Then
        strFormat = IIf(rs!编码 = 0, "0000", Mid(strFormat, 1, Len(rs!编码)))
        GetNextCode = Format(rs!编码 + 1, strFormat)
    End If
End Function

Public Function FillLvw(ByRef objLvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '-------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    LockWindowUpdate objLvw.hWnd
    
    Do While Not rs.EOF
        
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("名称").Value, rs("图标").Value, rs("图标").Value)
        For lngLoop = 2 To objLvw.ColumnHeaders.Count
            objItem.SubItems(lngLoop - 1) = zlCommFun.NVL(rs(objLvw.ColumnHeaders(lngLoop).Text).Value)
        Next
                        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillLvw = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function RestoreRow(ByRef objVsf As Object, ByVal strKey As String) As Boolean
    
    Dim lngLoop As Long
        
    For lngLoop = 1 To objVsf.Rows - 1
        If objVsf.RowData(lngLoop) = strKey Then
            objVsf.Row = lngLoop
            Exit Function
        End If
    Next
End Function

Public Function LoadGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    
    On Error GoTo 0
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = zlCommFun.NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '下一列
        Next
        
pointNext:
        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        
        rsData.MoveNext
    Loop
    
    LoadGrid = True
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetCol(ByVal objVsf As Object, ByVal strData As String) As Long
    
    Dim lngLoop As Long
    
    GetCol = -1
    For lngLoop = 0 To objVsf.Cols - 1
        If objVsf.Cell(flexcpData, 0, lngLoop) = strData Then
            GetCol = lngLoop
            Exit Function
        End If
    Next
    
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function GetNextNo(intBillID As Integer) As Variant

    GetNextNo = zlDatabase.GetNextNo(intBillID)
    
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, Optional ByVal strBill As String) As Long
'功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
'参数：bytKind=票种
'      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
'      strBill=要检查范围的票据号
'说明：
'    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
'    2.在检查范围时,长度也在检查范围之内。
'    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
'返回：
'      正常：票据领用ID>0
'      0=失败
'      -1:没有自用(用完或未领用)、也没有共用(未设置)
'      -2:设置的共用已用完
'      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '操作员有剩余的自用票据集
    strSQL = _
        "Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
        "From 票据领用记录" & vbNewLine & _
        "Where 票种 = [1] And 使用方式 = 1 And 剩余数量 > 0 And 领用人 = [2]" & vbNewLine & _
        "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, 开始号码"
    Set rsSelf = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, UserInfo.姓名)
    
    If lng领用ID = 0 Then
        '程序中第一次检查,且没有设置本地共用
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '也没有自用票据
        '有自用票据,按优先原则返回
        lngReturn = rsSelf!ID
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        strSQL = "Select ID,使用方式,剩余数量,前缀文本,开始号码,终止号码 From 票据领用记录 Where 票种=[1] And ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, lng领用ID)
        If rsTmp.BOF = False Then
            If rsTmp!使用方式 = 2 Then '共用,要先看有没有自用
                If Not rsSelf.EOF Then
                    '有自用的，优先
                    lngReturn = rsSelf!ID
                Else
                    '没有自用取共用
                    If rsTmp!剩余数量 = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                    lngReturn = rsTmp!ID
                    blnTmp = True
                End If
            Else
                '自用票据
                If rsTmp!剩余数量 > 0 Then
                    '有剩余
                    lngReturn = rsTmp!ID
                Else
                    '其它有剩余的自用
                    If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '其它自用也没有剩余
                    lngReturn = rsSelf!ID
                End If
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)))) <> UCase(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!开始号码) And UCase(strBill) <= UCase(rsTmp!终止号码) And Len(strBill) = Len(rsTmp!开始号码)) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng领用ID As Long) As String
    '功能：根据领用批次ID,获取下一个实际票据号
    '说明：1.当取不到范围内的有效票据时,返回空由用户输入
    '      2.排开已报损的号码
    Dim rsMain As New ADODB.Recordset
    Dim rsDelete As New ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo errH
    
    strSQL = "Select 前缀文本,开始号码,终止号码,当前号码" & _
        " From 票据领用记录 Where Nvl(剩余数量,0)>0 And ID=[1]"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng领用ID)
    
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!当前号码) Then
        strBill = UCase(rsMain!开始号码)
    Else
        strBill = UCase(IncStr(rsMain!当前号码))
    End If
    
    strSQL = "Select Upper(号码) as 号码 From 票据使用明细" & _
        " Where 性质=1 And 原因=5 And 号码>=[2] And 领用ID=[1] Order by 号码"

    Set rsDelete = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng领用ID, strBill)
    
    Do While True
        '检查范围
        If Left(strBill, Len(zlCommFun.NVL(rsMain!前缀文本))) <> UCase(zlCommFun.NVL(rsMain!前缀文本)) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!开始号码) And strBill <= UCase(rsMain!终止号码)) Then
            Exit Function
        End If
                
        '排开报损号
        rsDelete.Filter = "号码='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
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

Public Function IntEx(vNumber As Variant) As Variant
    '------------------------------------------------------------------------------------------------------------------
    '功能：取大于指定数值的最小整数
    '------------------------------------------------------------------------------------------------------------------
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function RePrintBalance(strNo As String, frmParent As Object, lng结帐ID As Long, ByVal bytKind As Byte) As Boolean
    '功能：当前收款记录重新打印一张票据
    Dim strSQL As String
    Dim strInvoice As String
    Dim lng领用ID As Long
    Dim blnValid As Boolean
    Dim blnDo As Boolean
    
    '如果严格控制票据使用
    If gblnStrictCtrl Then
        lng领用ID = GetInvoiceGroupID(bytKind, 1, 0, glngShareUseID)
        Select Case lng领用ID
            Case -1
                MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        If lng领用ID <= 0 Then Exit Function
    End If
    
    blnDo = ReportPrintSet(gcnOracle, glngSys, "ZL1_BILL_1862", frmParent)

    If blnDo Then
        '取下一个票据号码
        If Not gblnStrictCtrl Then
            '非严格控制时直接从本地读取
            strInvoice = UCase(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前结帐票据号", ""))
            If strInvoice = "" Then
                '有可能是第一次使用
                Do
                    strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                    vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                        
                    '用户取消输入,允许打印
                    If strInvoice = "" Then
                        If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        blnValid = True
                    Else
                        '检查输入有效性
                        If zlCommFun.ActualLen(strInvoice) <> ParamInfo.结帐票据号长度 Then
                            MsgBox "输入的票据号码长度应该为 " & ParamInfo.结帐票据号长度 & " 位！", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    End If
                Loop While Not blnValid
            Else
                strInvoice = IncStr(strInvoice)
            End If
        Else
            '根据票据领用读取
            strInvoice = GetNextBill(lng领用ID)
            If strInvoice = "" Then
                '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                Do
                    strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                        
                    '用户取消输入,不打印
                    If strInvoice = "" Then Exit Function
                    
                    '检查输入有效性
                    If GetInvoiceGroupID(bytKind, 1, lng领用ID, glngShareUseID, strInvoice) = -3 Then
                        MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Loop While Not blnValid
            End If
        End If
        
        Call frmPrint.ReportPrint(2, strNo, lng结帐ID, lng领用ID, strInvoice, , , , bytKind)
       
        RePrintBalance = True
    End If
End Function

Public Function GetMaxFact(ByVal strNo As String, ByVal bytKind As Byte) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能：获取指定结帐单据发出的最大票据号
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    '应取最后一次打印的最大号码
    strSQL = "Select Max(ID) From 票据打印内容 Where 数据性质=[1] And NO=[2]"
    strSQL = "Select Max(号码) as 号码 From 票据使用明细" & _
        " Where 票种=[1] And 性质=1 And 打印ID=(" & strSQL & ")"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", bytKind, strNo)
    
    If Not rsTmp.EOF Then GetMaxFact = zlCommFun.NVL(rsTmp!号码)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSysParameter(ByVal lngNo As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能;获取系统参数
    '参数:lngNo     参数号
    '返回:参数值
    '------------------------------------------------------------------------------------------------------------------

    GetSysParameter = zlDatabase.GetPara(lngNo, glngSys, , "")

End Function

Public Function ShowTxtFilter(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于文本框控件)
    '------------------------------------------------------------------------------------------------------------------
    Dim objPoint As POINTAPI
    Dim strInput As String
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand

    If rsData.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
            
    If rsData.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
    
    '参数初始化
    strInput = "'%" & UCase(objTxt.Text) & "%'"
    Call ClientToScreen(objTxt.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        

    
    If frmSelectDialog.ShowSelect(frmParent, 2, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False) Then GoTo Over
   
    
    Exit Function
    
Over:
    
    Set rsResult = rsData
    
    ShowTxtFilter = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowTxtSelect(ByVal frmParent As Object, _
                                    ByVal objTxt As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional ByVal bytStyle As Byte = 3) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim objPoint As POINTAPI
    
    On Error GoTo errHand

    If rsData.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objTxt.hWnd, objPoint)
                
    lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
    lngY = objTxt.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
    
    If frmSelectDialog.ShowSelect(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objTxt.Height, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rsData
        ShowTxtSelect = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Function ShowGrdFilter(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True, _
                                    Optional ByVal blnMuli As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于表格控件)
    '------------------------------------------------------------------------------------------------------------------

    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo errHand


    If rsData.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
    If rsData.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
        
    Call ClientToScreen(objVsf.hWnd, objPoint)
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight

    If frmSelectDialog.ShowSelect(frmParent, 2, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False, blnMuli) Then GoTo Over
    
    Exit Function
    
Over:
    
    Set rsResult = rsData
    
    ShowGrdFilter = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowGrdSelect(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional ByVal blnMuliSel As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构,应用于表格控件
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objVsf.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
    
    If frmSelectDialog.ShowSelect(frmParent, 3, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, , strSavePath, , False, blnMuliSel) Then
                            
        Set rsResult = rsData
        ShowGrdSelect = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Function FillTreeData(ByRef objTvw As Object, ByVal rs As ADODB.Recordset, Optional ByVal blnExpand As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim objNode As Node
    
    On Error GoTo errHand
    
    LockWindowUpdate objTvw.hWnd
    
    Do While Not rs.EOF
        
        If IsNull(rs("上级id").Value) Then
            Set objNode = objTvw.Nodes.Add(, , "K" & zlCommFun.NVL(rs("ID").Value, 0), zlCommFun.NVL(rs("名称").Value), rs("图标").Value)
        Else
            Set objNode = objTvw.Nodes.Add("K" & rs("上级id").Value, tvwChild, "K" & zlCommFun.NVL(rs("ID").Value, 0), zlCommFun.NVL(rs("名称").Value), rs("图标").Value)
        End If
        
        objNode.ExpandedImage = rs("打开图标").Value
        objNode.Expanded = blnExpand
        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillTreeData = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function MedicalItemsRecord(ByRef rs As ADODB.Recordset, Optional ByVal bytMode As Byte = 1) As Boolean
    '创建记录集,用于保存选择的体检项目
    Set rs = New ADODB.Recordset
    
    With rs
        If bytMode = 1 Then
            .Fields.Append "组别", adVarChar, 50
            .Fields.Append "ID", adVarChar, 18
            .Fields.Append "清单id", adVarChar, 18
            .Fields.Append "类别", adVarChar, 30
            .Fields.Append "名称", adVarChar, 50
            .Fields.Append "基本价格", adVarChar, 50
            .Fields.Append "体检价格", adVarChar, 50
            .Fields.Append "折扣", adVarChar, 50
            .Fields.Append "体检类型", adVarChar, 50
            .Fields.Append "结算方式", adVarChar, 50
            .Fields.Append "执行科室", adVarChar, 50
            .Fields.Append "采集科室", adVarChar, 50
            .Fields.Append "执行科室id", adVarChar, 18
            .Fields.Append "采集方式", adVarChar, 50
            .Fields.Append "采集方式id", adVarChar, 18
            .Fields.Append "采集科室id", adVarChar, 18
            .Fields.Append "检验标本", adVarChar, 50
            .Fields.Append "检查部位", adVarChar, 2000
            .Fields.Append "检查部位id", adVarChar, 18
            .Fields.Append "新加", adVarChar, 1
            .Fields.Append "前景色", adVarChar, 20
            .Fields.Append "删除", adVarChar, 1
            .Fields.Append "公共", adVarChar, 1
            .Fields.Append "计费明细", adVarChar, 4000
            .Fields.Append "选择", adVarChar, 1
        Else
            .Fields.Append "组别", adVarChar, 50
            .Fields.Append "ID", adVarChar, 18
            .Fields.Append "IC卡号", adVarChar, 18
            .Fields.Append "病人id", adBigInt, 18, adFldKeyColumn
            .Fields.Append "姓名", adVarChar, 50
            .Fields.Append "门诊号", adBigInt, 18
            .Fields.Append "健康号", adVarChar, 20
            .Fields.Append "性别", adVarChar, 50
            .Fields.Append "身份证号", adVarChar, 50
            .Fields.Append "婚姻状况", adVarChar, 50
            .Fields.Append "出生日期", adVarChar, 18
            .Fields.Append "身份证", adVarChar, 30
            .Fields.Append "年龄", adVarChar, 50
            .Fields.Append "民族", adVarChar, 50
            .Fields.Append "国籍", adVarChar, 50
            .Fields.Append "学历", adVarChar, 50
            .Fields.Append "职业", adVarChar, 50
            .Fields.Append "身份", adVarChar, 50
            .Fields.Append "联系人姓名", adVarChar, 50
            .Fields.Append "联系人电话", adVarChar, 50
            .Fields.Append "登记时间", adVarChar, 30
            .Fields.Append "电子邮件", adVarChar, 50
            .Fields.Append "联系人地址", adVarChar, 100
            .Fields.Append "工作单位", adVarChar, 100
            .Fields.Append "就诊卡号", adVarChar, 10
            .Fields.Append "前景色", adVarChar, 30
            .Fields.Append "背景色", adVarChar, 30
            .Fields.Append "删除", adVarChar, 1
            .Fields.Append "新加", adVarChar, 1
        End If
        .Open
    End With
End Function

Public Sub CopyGrid(ByVal objFrom As Object, ByRef objTo As Object, Optional ByVal lngStartCol As Long = 0)
    
    Dim lngRow As Long
    Dim lngCol As Long
        
    objTo.Rows = objFrom.Rows
    objTo.Cols = objFrom.Cols - lngStartCol
    
    For lngCol = lngStartCol To objFrom.Cols - 1
        objTo.ColWidth(lngCol - lngStartCol) = objFrom.ColWidth(lngCol)
        objTo.MergeCol(lngCol - lngStartCol) = objFrom.MergeCol(lngCol)
    Next
    
    For lngRow = 0 To objFrom.Rows - 1
        objTo.MergeRow(lngRow) = objFrom.MergeRow(lngRow)
        For lngCol = lngStartCol To objFrom.Cols - 1
            objTo.TextMatrix(lngRow, lngCol - lngStartCol) = objFrom.TextMatrix(lngRow, lngCol)
        Next
    Next
    
End Sub

Public Function CheckAllowMedical(ByVal lngKey As Long) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '检查是否有团体
    gstrSQL = "SELECT 0,是否团体,合约单位id FROM 体检登记记录 WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        
        If rs("是否团体").Value = 1 Then
            If zlCommFun.NVL(rs("合约单位id").Value, 0) = 0 Then
                CheckAllowMedical = 1
'                strPrompt = "当前体检还没有确定团体信息！"
                Exit Function
            End If
        End If
        
    End If
    
    '检查是否有人员
    gstrSQL = "SELECT NVL(COUNT(1),0) AS 人数 FROM 体检人员档案 WHERE 病人id IS NOT NULL AND 登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        If rs("人数").Value = 0 Then
            'strPrompt = "当前体检还没有确定体检人员！"
            CheckAllowMedical = 2
            Exit Function
        End If
    End If
    
    '检查是否有体检项目
    gstrSQL = "SELECT B.组别名称,Sum(Decode(a.登记id,Null,0,1)) AS 项数 FROM 体检项目清单 A,体检组别 B WHERE A.登记id(+)=B.登记id AND A.组别名称(+)=B.组别名称 AND B.登记id=[1] GROUP BY B.组别名称"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            If rs("项数").Value = 0 Then
                CheckAllowMedical = 3
    '            strPrompt = "当前体检的“" & rs("组别名称").Value & "”组别还没有确定体检项目！"
                Exit Function
            End If
            rs.MoveNext
        Loop
    Else
        CheckAllowMedical = 3
'        strPrompt = "当前体检还没有确定体检项目！"
        Exit Function
    End If
    
    gstrSQL = "SELECT 1 FROM 体检人员档案 WHERE 组别名称 NOT IN (SELECT 组别名称 FROM 体检组别 WHERE 登记id=[1]) AND 登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        CheckAllowMedical = 4
        Exit Function
    End If
    
    CheckAllowMedical = 0
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '功能：初始化参数
    '
    '
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim str费用小数位 As String
    Dim strTmp As String
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '费用金额保留位数
    '表示费用金额核算到小数点后第多少位?
    ParamInfo.费用金额小数位数 = Val(zlDatabase.GetPara(9, glngSys, , "2"))
    If ParamInfo.费用金额小数位数 > 0 Then
        gstrDec = "0." & String(ParamInfo.费用金额小数位数, "0")
    Else
        gstrDec = "0"
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '收费诊疗项目输入匹配
    '第1位1-全数字只查编码,第2位1-全字母只查简码
    ParamInfo.收费诊疗项目匹配 = zlDatabase.GetPara(44, glngSys, , "11")

    
    '------------------------------------------------------------------------------------------------------------------
    '就诊卡字母前缀
    ParamInfo.就诊卡字母前缀 = zlDatabase.GetPara(27, glngSys, , "")
    
    '------------------------------------------------------------------------------------------------------------------
    '就诊卡号码长度,结帐票据号长度
    strTmp = zlDatabase.GetPara(20, glngSys, , "")
    If strTmp <> "" Then
        If UBound(Split(strTmp, "|")) >= 4 Then ParamInfo.就诊卡号码长度 = Val(Split(strTmp, "|")(4))
        If UBound(Split(strTmp, "|")) >= 3 Then ParamInfo.结帐票据号长度 = Val(Split(strTmp, "|")(3))
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '本地参数
    ParamInfo.项目输入匹配方式 = Val(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0"))
    
errHand:
    
End Function

Public Function WriteItems(ByVal rs As ADODB.Recordset, ByRef rsItem As ADODB.Recordset, Optional ByVal bytDo As Byte = 0, Optional ByVal bytMode As Byte = 1) As Boolean
    
    '读取体检项目
    On Error GoTo errHand
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If bytMode = 1 Then
                rsItem.AddNew
                rsItem("组别").Value = zlCommFun.NVL(rs("组别名称").Value)
                rsItem("ID").Value = zlCommFun.NVL(rs("ID").Value)
                rsItem("清单id").Value = Val(zlCommFun.NVL(rs("清单id").Value))
                rsItem("类别").Value = zlCommFun.NVL(rs("类别").Value)
                rsItem("名称").Value = zlCommFun.NVL(rs("名称").Value)
                rsItem("执行科室").Value = zlCommFun.NVL(rs("执行科室").Value)
                rsItem("结算方式").Value = zlCommFun.NVL(rs("结算方式").Value, "1")
                rsItem("体检类型").Value = zlCommFun.NVL(rs("体检类型").Value)
                rsItem("基本价格").Value = Format(zlCommFun.NVL(rs("基本价格").Value), "0.00##")
                rsItem("体检价格").Value = Format(zlCommFun.NVL(rs("体检价格").Value), "0.00##")
                rsItem("折扣").Value = zlCommFun.NVL(rs("折扣").Value)
                rsItem("执行科室id").Value = zlCommFun.NVL(rs("执行科室id").Value)
                rsItem("采集方式").Value = zlCommFun.NVL(rs("采集方式").Value)
                rsItem("采集方式id").Value = zlCommFun.NVL(rs("采集方式id").Value)
                rsItem("采集科室id").Value = zlCommFun.NVL(rs("采集科室id").Value)
                rsItem("采集科室").Value = zlCommFun.NVL(rs("采集科室").Value)
                rsItem("检验标本").Value = zlCommFun.NVL(rs("检验标本").Value)
                rsItem("检查部位").Value = zlCommFun.NVL(rs("检查部位").Value)
                rsItem("检查部位id").Value = zlCommFun.NVL(rs("检查部位id").Value)
                rsItem("计费明细").Value = GetPriceList(zlCommFun.NVL(rs("清单id").Value))
                
                If bytDo = 1 Then
                    rsItem("新加").Value = "1"
                    rsItem("前景色").Value = "16711680"
                    rsItem("删除").Value = ""
                End If
                
                If bytDo = 2 Then
                    rsItem("新加").Value = "1"
                    rsItem("前景色").Value = IIf(Val(zlCommFun.NVL(rs("复查清单id").Value)) = 0, "0", "255")
                    rsItem("删除").Value = ""
                    rsItem("公共").Value = zlCommFun.NVL(rs("公共").Value)
                End If
            Else
                rsItem.AddNew
                
                rsItem("组别").Value = zlCommFun.NVL(rs("组别").Value)
                rsItem("姓名").Value = zlCommFun.NVL(rs("姓名").Value)
                rsItem("IC卡号").Value = zlCommFun.NVL(rs("IC卡号").Value)
                rsItem("健康号").Value = zlCommFun.NVL(rs("健康号").Value)
                rsItem("门诊号").Value = zlCommFun.NVL(rs("门诊号").Value)
                rsItem("就诊卡号").Value = zlCommFun.NVL(rs("就诊卡号").Value)
                rsItem("身份证").Value = zlCommFun.NVL(rs("身份证").Value)
                rsItem("性别").Value = zlCommFun.NVL(rs("性别").Value)
                rsItem("年龄").Value = zlCommFun.NVL(rs("年龄").Value)
                rsItem("出生日期").Value = zlCommFun.NVL(rs("出生日期").Value)
                rsItem("婚姻状况").Value = zlCommFun.NVL(rs("婚姻状况").Value)
                rsItem("病人id").Value = zlCommFun.NVL(rs("病人id").Value)
                rsItem("民族").Value = zlCommFun.NVL(rs("民族").Value)
                rsItem("国籍").Value = zlCommFun.NVL(rs("国籍").Value)
                rsItem("学历").Value = zlCommFun.NVL(rs("学历").Value)
                rsItem("职业").Value = zlCommFun.NVL(rs("职业").Value)
                rsItem("身份").Value = zlCommFun.NVL(rs("身份").Value)
                rsItem("联系人姓名").Value = zlCommFun.NVL(rs("联系人姓名").Value)
                rsItem("联系人电话").Value = zlCommFun.NVL(rs("联系人电话").Value)
                rsItem("电子邮件").Value = zlCommFun.NVL(rs("电子邮件").Value)
                rsItem("联系人地址").Value = zlCommFun.NVL(rs("联系人地址").Value)
                rsItem("工作单位").Value = zlCommFun.NVL(rs("工作单位").Value)
                rsItem("登记时间").Value = zlCommFun.NVL(rs("登记时间").Value)
                
                If bytDo = 1 Then
                    rsItem("新加").Value = "1"
                    rsItem("前景色").Value = "16711680"
                    rsItem("删除").Value = ""
                End If
                
                If bytDo = 2 Then
                    rsItem("新加").Value = "1"
                    rsItem("前景色").Value = "8388736"
                    rsItem("删除").Value = ""
                End If
            End If
            
            rs.MoveNext
        Loop
    End If
    
    WriteItems = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetPriceList(ByVal lngKey As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
                
    strSQL = "Select x.*,y.名称,y.计算单位,z.数次,z.计价性质,z.执行科室id,t.名称 As 执行科室,y.类别,Decode(x.标准单价,0,0,Null,0,10*x.单价/x.标准单价) As 折扣 " & _
            "From  " & _
                "(Select a.清单id,a.收费细目id,Sum(a.标准单价) As 标准单价,Sum(a.单价) As 单价 " & _
                "From 体检项目计价 a " & _
                "Where a.清单id = [1] " & _
                "Group By a.清单id,a.收费细目id " & _
                ") x, " & _
                "收费项目目录 y, " & _
                "体检项目计价 z,部门表 t " & _
            "Where x.清单id = z.清单id and t.id(+)=z.执行科室id " & _
                  "and x.收费细目id=y.id and x.收费细目id=z.收费细目id "
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If strTmp <> "" Then strTmp = strTmp & ";"
            strTmp = strTmp & zlCommFun.NVL(rs("名称")) & ":" & _
                    zlCommFun.NVL(rs("计算单位")) & ":" & _
                    zlCommFun.NVL(rs("数次")) & ":" & _
                    zlCommFun.NVL(rs("标准单价")) & ":" & _
                    zlCommFun.NVL(rs("单价")) & ":" & _
                    zlCommFun.NVL(rs("收费细目id")) & ":" & _
                    zlCommFun.NVL(rs("计价性质")) & ":" & _
                    zlCommFun.NVL(rs("执行科室")) & ":" & _
                    zlCommFun.NVL(rs("执行科室id")) & ":" & _
                    zlCommFun.NVL(rs("类别")) & ":" & _
                    zlCommFun.NVL(rs("折扣"))
            
            rs.MoveNext
        Loop
    End If
                    
    GetPriceList = strTmp
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetTypePriceList(ByVal lngNo As Long, ByVal lngKey As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select z.*,y.名称,y.计算单位,x.现价,x.体检单价,z.计价性质,z.折扣 " & _
                "From " & _
                "( Select a.序号,a.诊疗项目id,a.收费细目id,Sum(c.现价) As 现价,Sum(c.现价*Nvl(a.折扣,1)) As 体检单价 " & _
                  "From 收费价目 c, " & _
                       "体检类型计价 a " & _
                  "Where a.收费细目id = c.收费细目id " & _
                        "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                        "and A.序号=[1] " & _
                        "and A.诊疗项目id=[2] " & _
                  "Group by a.序号,a.诊疗项目id,a.收费细目id " & _
                ") x, " & _
                "收费项目目录 y, " & _
                "体检类型计价 z " & _
                "Where x.收费细目id = y.ID " & _
                      "and z.序号=x.序号 " & _
                      "and z.诊疗项目id=x.诊疗项目id " & _
                      "and z.收费细目id=x.收费细目id"

    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngNo, lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If strTmp <> "" Then strTmp = strTmp & ";"
            strTmp = strTmp & zlCommFun.NVL(rs("名称")) & ":" & _
                    zlCommFun.NVL(rs("计算单位")) & ":" & _
                    zlCommFun.NVL(rs("数次")) & ":" & _
                    zlCommFun.NVL(rs("现价")) & ":" & _
                    zlCommFun.NVL(rs("收费细目id")) & ":" & _
                    zlCommFun.NVL(rs("计价性质")) & ":" & _
                    zlCommFun.NVL(rs("体检单价")) & ":" & _
                    10 * zlCommFun.NVL(rs("折扣"), 0)
            
            rs.MoveNext
        Loop
    End If
                    
    GetTypePriceList = strTmp
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetCombList(ByVal strSQL As String) As String
    
    Dim rs As New ADODB.Recordset
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical")
    If rs.BOF = False Then
        Do While Not rs.EOF
            GetCombList = GetCombList & "|" & zlCommFun.NVL(rs.Fields(0).Value)
            rs.MoveNext
        Loop
    End If
    If GetCombList = "" Then
        GetCombList = " |"
    Else
        GetCombList = Mid(GetCombList, 2)
    End If
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function GetBirth(ByVal intYear As Integer, ByRef strStart As String, ByRef strEnd As String) As Boolean
        
    strStart = Format(DateAdd("yyyy", 0 - intYear - 1, Now), "yyyy-MM-dd")
    strEnd = Format(DateAdd("yyyy", 0 - intYear, Now), "yyyy-MM-dd")
    
End Function

Public Function CheckStrValid(ByVal Text As String, ByVal bytMode As CHECKFORMAT, Optional ByVal KeyCustom As String, Optional ByVal intLen As Integer = 0, Optional ByVal intDec As Integer = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    Select Case bytMode
    Case CHECKFORMAT.电子邮件
        
        If Trim(Text) <> "" Then
            If InStr(Text, "@") = 0 Then Exit Function
            If InStr(Text, "@") = 1 Then Exit Function
            If InStr(Text, "@") = Len(Text) Then Exit Function
        End If
        
    Case CHECKFORMAT.日期
    
        If Trim(Text) <> "" Then
            If IsDate(Trim(Text)) = False Then Exit Function
        End If
        
    Case CHECKFORMAT.身份证号
        
        '只能包含 0,1,2,3,4,5,6,7,8,9,X 字符
        
        If Trim(Text) <> "" Then
            If Len(Text) <> 15 And Len(Text) <> 18 Then Exit Function
            
            For lngLoop = 1 To Len(Text)
                If InStr("0123456789X", UCase(Mid(Text, lngLoop, 1))) = 0 Then Exit Function
            Next
            
        End If
    Case CHECKFORMAT.数值
        
    Case CHECKFORMAT.自定义
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then Exit Function
        Next
    End Select
    
    CheckStrValid = True
End Function

Public Function Lpad(ByVal strText As String, ByVal lngLen As Long, ByVal strReplace As String) As String
    Dim lngL As Long
    
    lngL = Len(strText)
    If lngL > lngLen Then
        Lpad = Left(strText, lngLen)
    ElseIf lngL < lngLen Then
        Lpad = String(lngLen - lngL, strReplace) & strText
    Else
        Lpad = strText
    End If
End Function

Public Function EnterFocus(obj As Object) As Boolean
    
    On Error Resume Next
    
    obj.SetFocus
    
End Function

Public Function HaveExcel() As Boolean
    '------------------------------------------------
    '功能：判断本机上装有EXCEL没有
    '参数：
    '返回：有则返回True
    '------------------------------------------------

    On Error GoTo errHandle
    
    Dim objTemp  As Object
    
    Set objTemp = CreateObject("Excel.Application") '打开一个EXCEL程序
    
    Set objTemp = Nothing
    
    HaveExcel = True
    
    Exit Function

errHandle:
    Set objTemp = Nothing
    HaveExcel = False
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "EXECUTE", adTinyInt
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String) As Boolean
    
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Sub DrawPicture(pic As Object, objPic As StdPicture, ByVal W As Long, ByVal H As Long)
'功能：在PictureBox中央按适当比例画一幅图
'参数：W,H=要作图的尺寸
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= pic.ScaleWidth And H <= pic.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / pic.ScaleWidth
        sngH = H / pic.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    pic.Cls
    On Error Resume Next
    pic.PaintPicture objPic, (pic.ScaleWidth - lngW) / 2, (pic.ScaleHeight - lngH) / 2, lngW, lngH
    
End Sub

Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
'-------------------------------------------------------------
'功能：将指定的记录集图形字段复制为图形临时文件
'参数：
'       rsTable   图形存储记录集
'       strField  图形字段
'       strFile   用户定义的文件名（可选项）
'返回：
'-------------------------------------------------------------
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    Dim strTempFile As String
    
    On Error GoTo errH
    lngFileSize = rsTable.Fields(strField).ActualSize
    If lngFileSize = 0 Then
        '未读取有效数据
        Exit Function
    End If
    
    FileNum = FreeFile
    If strFile = "" Then
        '当用户并没定义文件名时
'        j = 0
        
        strFile = CreateTmpFile
        
'        Do While True
'            strTempFile = CurDir & "\zlNewPicture" & CStr(j) & ".pic"
'            If Len(Dir(strTempFile)) = 0 Then Exit Do
'            j = j + 1
'        Loop
'        strFile = strTempFile
    End If
    Open strFile For Binary As FileNum
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    rsTable.Move 0
    For j = 0 To intBolcks
        If j = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim aryChunk(lngCurSize - 1) As Byte
        aryChunk() = rsTable.Fields(strField).GetChunk(lngCurSize)
        Put FileNum, , aryChunk()
    Next
    Close FileNum
    ReadPicture = strFile
    Exit Function

errH:
    Close FileNum
'    Kill strFile
    ReadPicture = ""

End Function

Public Function GetTmpPath() As String
    
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

Public Function CreateTmpFile(Optional ByVal strFileType As String = "tmp") As String
    '------------------------------------------------------------------------------------------------------------------
    '
    '功能:
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strFileTemp As String
       
    
    strFileTemp = GetTmpPath
    
    strFileTemp = strFileTemp & "zlNewPic" & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    
    CreateTmpFile = strFileTemp
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'功能：判断是否存在指定处方单据类型的入出类别
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = zlCommFun.NVL(rsTmp!类别ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatientID(ByVal strIC As String) As Long
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病人id From 病人信息 Where IC卡号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", strIC)
    If Not rsTmp.EOF Then GetPatientID = zlCommFun.NVL(rsTmp!病人id, 0)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcCharge(ByVal rsSource As ADODB.Recordset, ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim dbTmp实收金额 As Double
    Dim dbTmp已结金额 As Double
    
    Dim db应收金额_收 As Double
    Dim db应收金额_记 As Double
    
    Dim db实收金额 As Double
    Dim db记帐金额 As Double
    Dim db收费金额 As Double
    Dim db未结金额 As Double
    Dim db未收金额 As Double
    Dim db未结算合计 As Double
    Dim db已结金额 As Double
    
    On Error GoTo errH
    
    If rsSource.BOF Then Exit Function
    
    Set rs = New ADODB.Recordset
    With rs
    
        .Fields.Append "应收金额_记", adVarChar, 30
        .Fields.Append "应收金额_收", adVarChar, 30
        
        .Fields.Append "实收金额", adVarChar, 30
        .Fields.Append "记帐金额", adVarChar, 30
        .Fields.Append "收费金额", adVarChar, 30
        
        .Fields.Append "未结算合计", adVarChar, 30
        .Fields.Append "未结金额", adVarChar, 30
        .Fields.Append "未收金额", adVarChar, 30
        .Open
    End With
    
    Do While Not rsSource.EOF
        
        dbTmp实收金额 = zlCommFun.NVL(rsSource("实收金额").Value, 0)
        dbTmp已结金额 = zlCommFun.NVL(rsSource("结帐金额").Value, 0)
        
        db实收金额 = db实收金额 + dbTmp实收金额
        db已结金额 = db已结金额 + dbTmp已结金额
        
        If zlCommFun.NVL(rsSource("记帐费用").Value, 0) = 1 Then
            db记帐金额 = db记帐金额 + dbTmp实收金额
            db未结金额 = db未结金额 + (dbTmp实收金额 - dbTmp已结金额)
            
            db应收金额_记 = db应收金额_记 + zlCommFun.NVL(rsSource("应收金额").Value, 0)
        Else
            db应收金额_收 = db应收金额_收 + zlCommFun.NVL(rsSource("应收金额").Value, 0)
        End If
        
        rsSource.MoveNext
    Loop
    
    db收费金额 = db实收金额 - db记帐金额
    db未结算合计 = db实收金额 - db已结金额
    db未收金额 = db未结算合计 - db未结金额
    
    
    rs.AddNew
    
    rs("应收金额_记").Value = db应收金额_记
    rs("应收金额_收").Value = db应收金额_收
    
    rs("实收金额").Value = db实收金额
    rs("记帐金额").Value = db记帐金额
    rs("收费金额").Value = db收费金额
    
    rs("未结算合计").Value = db未结算合计
    rs("未结金额").Value = db未结金额
    rs("未收金额").Value = db未收金额
    rs.Update
    
    CalcCharge = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InputIsCard(ByVal strText As String, KeyAscii As Integer) As Boolean
    '******************************************************************************************************************
    '功能：判断指定文本框中当前输入是否在刷卡,根据处理密文显示
    '参数：
    '返回：
    '******************************************************************************************************************

'    Dim strText As String
    Dim blnCard As Boolean
    Dim arrMask As Variant
    Dim intLoop As Integer

    '当前键入后显示的内容(还未显示出来)
'    strText = txtInput.Text
'    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
'    Debug.Print strText
        
    '判断是否在刷卡
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf ParamInfo.就诊卡字母前缀 <> "" Then
        arrMask = Split(ParamInfo.就诊卡字母前缀, "|")
        For intLoop = 0 To UBound(arrMask)
            If strText Like arrMask(intLoop) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(intLoop)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(intLoop)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    '刷卡时卡号是否密文显示
'    If blnCard Then
'        txtInput.PasswordChar = IIf(gblnShowCard, "", "*")
'    Else
'        txtInput.PasswordChar = ""
'    End If
    
    InputIsCard = blnCard
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, Optional ByVal strBill As String) As Long
'功能：获取张数够用并且指定票据在其可用范围内的领用ID
'参数：bytKind      =   票种
'      intNum       =   要打印的票据张数
'      lngLastUseID =   上次使用的领用ID
'      lngShareUseID=   本地参数指定的共用ID
'      strBill      =   当前票据号，用于检查领用批次的票据范围
'返回：
'      >0   =   成功，可用的领用ID
'      =0   =   失败
'      -1   =   没有自用(用完或不够，或未领用),未设置共用
'      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
'      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
'      -4   =   指定批次的票据不够用
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
        strSQL = "Select 前缀文本,开始号码,终止号码" & vbNewLine & _
                 "From 票据领用记录 Where 票种=[1] And 剩余数量>=[2] And ID=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, intNum, lngLastUseID)
        With rsTmp
            If .RecordCount > 0 Then    '目前的票据号可能和上次不同，所以需要检查范围
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '可能没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
        
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    strSQL = "Select ID, 前缀文本, 开始号码, 终止号码" & vbNewLine & _
        "From 票据领用记录" & vbNewLine & _
        "Where 票种 = [1] And 剩余数量 >= [2] And 领用人 = [3] And 使用方式 = 1" & vbNewLine & _
        "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, 剩余数量, 登记时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, intNum, UserInfo.姓名)
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = "" & !前缀文本
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        strSQL = "Select 前缀文本,开始号码,终止号码" & vbNewLine & _
                 "From 票据领用记录 Where 票种=[1] And 剩余数量>=[2] And ID=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, intNum, lngShareUseID)
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '第一次使用时没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    
    GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStorage(ByVal lngKey As Long, ByVal lngDeptKey As Long) As Single
    '----------------------------------------------------------------------
    '功能:获取药品库存
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    GetStorage = 0
        
    If lngKey = 0 Then Exit Function
    
    On Error GoTo errHand
    
    strSQL = "SELECT I.是否变价,S.药房分批 AS 药房分批核算,S.剂量系数 FROM 收费项目目录 I,药品规格 S WHERE I.ID=S.药品id AND S.药品id=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lngKey)
    If rs.BOF = False Then
                        
        GetStorage = CalcStorage(lngKey, lngDeptKey, IIf(zlCommFun.NVL(rs("是否变价").Value, 0) = 0, False, True), IIf(zlCommFun.NVL(rs("药房分批核算").Value, 0) = 0, False, True))
    
    End If
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetStock(ByVal lng药品ID As Long, ByVal lng药房ID As Long) As Double
'功能：获取指定药房指定药品库存(以零售单位)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
    strSQL = _
        " Select Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
        " Where (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        " And A.性质=1 And A.药品ID=[1] And A.库房ID=[2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng药品ID, lng药房ID)
    If Not rsTmp.EOF Then GetStock = rsTmp!库存
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcStorage(ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal vChangePrice As Boolean, ByVal vBatch As Boolean) As Single

    '功能：获取指定药房指定药品库存(以零售单位)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
    strSQL = _
        " Select Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
        " Where (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        " And A.性质=1 And A.药品ID=[1] And A.库房ID=[2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng药品ID, lng库房ID)
    If Not rsTmp.EOF Then CalcStorage = rsTmp!库存
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
'    Dim rs As New ADODB.Recordset
'    Dim strSQL As String
'
'
'    If lng药品ID = 0 Then Exit Function
'
'    If vChangePrice And vBatch = False Then
'        '只是实价药品
'
'        strSQL = "SELECT NVL(A.可用数量,0) AS 可用数量 FROM 药品库存 A WHERE A.药品id=[1] AND A.库房ID=[2]"
'
'    ElseIf vChangePrice = False And vBatch Then
'        '只是药房分批核算药品
'
'        strSQL = "Select Sum(Nvl(可用数量,0)) as 可用数量 From 药品库存" & _
'                    " Where 性质=1 " & _
'                    " And (效期 Is NULL Or 效期>Trunc(Sysdate)) " & _
'                    " And 库房ID=[2]" & _
'                    " And 药品ID=[1]"
'
'    ElseIf vChangePrice And vBatch Then
'        '既是实价药品又是药房分批核算药品
'
'        strSQL = "Select Sum(Nvl(可用数量,0)) as 可用数量 From 药品库存" & _
'                    " Where 性质=1 " & _
'                    " And (效期 Is NULL Or 效期>Trunc(Sysdate)) " & _
'                    " And 库房ID=[2]" & _
'                    " And 药品ID=[1]"
'
'    Else
'        '既不是实价药品又不是药房分批核算药品,和只是实价药品一样的
'
'        strSQL = "SELECT NVL(A.可用数量,0) AS 可用数量 FROM 药品库存 A WHERE A.药品id=[1] AND A.库房ID=[2]"
'
'    End If
'
'    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lng药品ID, lng库房ID)
'
'    If rs.BOF = False Then CalcStorage = zlCommFun.NVL(rs("可用数量").Value, 0)

End Function

Public Function PromptStorageWarn(ByVal dbInput As Double, _
                                    ByVal dbStorage As Double, _
                                    ByVal strDrugName As String, _
                                    ByVal strExecuteDept As String, _
                                    ByVal strUnit As String, _
                                    Optional ByVal bytWarn As Byte = 1, _
                                    Optional ByVal bytApply As Byte = 1) As Integer
    '******************************************************************************************************************
    '功能：
    '参数：bytWarn：0-不检查;1-检查,不足提醒;2-检查，不足禁
    '返回：
    '******************************************************************************************************************

    If dbInput > 0 And dbInput > dbStorage Then
        
        If bytApply = 1 Then
            Call ShowSimpleMsg("药品（" & strDrugName & "）在库房（" & strExecuteDept & "）只有" & dbStorage & strUnit & "！")
            bytWarn = 0
        Else
            Select Case bytWarn
            Case 0
                
            Case 1
                If MsgBox("药品（" & strDrugName & "）在库房（" & strExecuteDept & "）只有" & dbStorage & strUnit & "，是否继续？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                    bytWarn = 0
                Else
                    bytWarn = 1
                End If
            Case 2
                MsgBox "药品（" & strDrugName & "）在库房（" & strExecuteDept & "）只有" & dbStorage & strUnit & "，不足禁止！", vbOKOnly + vbCritical, gstrSysName
                bytWarn = 1
            End Select
        End If
        
    End If
    
    PromptStorageWarn = bytWarn
    
End Function

Public Function MakeMedicalCharge(ByRef rsSQL As ADODB.Recordset, ByVal lng登记id As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim lngCount As Long
    Dim dbSum As Double
    Dim db应收金额 As Double
    Dim db实收金额 As Double
    Dim db库存数量 As Double
    Dim int库存检查 As Double
    Dim int序号 As Integer
    Dim lng类别id As Long
    Dim strNow As String
    Dim int单据 As Integer
    Dim int跟踪在用 As Integer
    Dim obj As Object
    
    On Error GoTo errHand
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    strSQL = "Select x.*,y.名称 As 执行科室 From (SELECT d.类别 As 收费类别,d.计算单位,a.收费细目id,Decode(a.执行科室id,b.执行科室id,c.执行科室id,a.执行科室id) As 执行科室id," & _
                    "a.标准单价*A.数次 As 应收金额,a.单价*A.数次 As 实收金额,a.单价,Nvl(a.数次,1) As 收费数量,a.标准单价,a.计价性质, " & _
                    "e.ID As 医嘱id,e.诊疗类别,Decode(e.诊疗类别,'E',c.采集No,c.No) As No,b.结算途径,f.病人id,f.姓名,f.性别,f.年龄,f.费别,f.门诊号,e.开嘱科室id,e.医嘱内容,d.名称 As 收费项目 " & _
            "FROM 体检项目计价 a, " & _
                "体检项目清单 b, " & _
                "体检项目医嘱 c, " & _
                "收费项目目录 d, " & _
                "病人医嘱记录 e, " & _
                "病人信息 f " & _
            "Where b.登记id = [1] " & _
             "And d.ID = a.收费细目ID " & _
             "And C.清单id=b.ID " & _
             "And c.临时标记=1 " & _
             "And c.医嘱id In (e.id,e.相关id) " & _
             "And ((e.诊疗类别='E' And a.计价性质=2) Or (e.诊疗类别='C' And a.计价性质<>2) Or (e.诊疗类别='D' And a.计价性质<>2 And e.相关id Is Null)) " & _
             "And c.病人id=f.病人id " & _
             "And b.ID = a.清单id) x,部门表 y Where x.执行科室id=y.ID " & _
            "Order By x.收费细目id"
    
    Set rsCharge = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", lng登记id)
    If rsCharge.BOF = False Then
        Do While Not rsCharge.EOF
            If Val(rsCharge("计价性质").Value) > 0 And zlCommFun.NVL(rsCharge("No").Value) <> "" Then
                
                int单据 = 0
                lng类别id = 0
                If rsCharge("收费类别").Value = "4" Then
                    int跟踪在用 = 0
                    
                    strSQL = "Select 跟踪在用 From 材料特性 Where 材料ID=[1]"
                    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsCharge("收费细目id").Value))
                    If rs.BOF = False Then int跟踪在用 = zlCommFun.NVL(rs("跟踪在用").Value, 0)
                    If int跟踪在用 = 1 Then
                        int单据 = IIf(zlCommFun.NVL(rsCharge("结算途径").Value, 1) = 1, 41, 42)
                    End If
                ElseIf InStr("567", rsCharge("收费类别").Value) > 0 Then
                    int单据 = IIf(zlCommFun.NVL(rsCharge("结算途径").Value, 1) = 1, 9, 8)
                End If
                
                If int单据 > 0 Then
                    strSQL = "Select 类别id From 药品单据性质 Where 单据=[1]"
                    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", int单据)
                    If rs.BOF = False Then lng类别id = zlCommFun.NVL(rs("类别id").Value, 0)
                    If lng类别id = 0 Then
                        If rsCharge("收费类别").Value = "4" Then
                            ShowSimpleMsg "不能确定材料处方单据的入出类别,请先到入出类别管理中设置！"
                        Else
                            ShowSimpleMsg "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！"
                        End If
                        Exit Function
                    End If
                End If
                
                '检查药品或材料库存
                If InStr("4567", rsCharge("收费类别").Value) > 0 Then
                    
                    db库存数量 = GetStorage(Val(rsCharge("收费细目id").Value), Val(rsCharge("执行科室id").Value))
                    If Val(rsCharge("收费数量").Value) > db库存数量 Then
                    
                        int库存检查 = 0
                        If rsCharge("收费类别").Value = "4" Then
                            strSQL = "Select 检查方式 From 材料出库检查 Where 库房ID=[1]"
                        Else
                            strSQL = "Select 检查方式 From 药品出库检查 Where 库房ID=[1]"
                        End If
                        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsCharge("执行科室id").Value))
                        If rs.BOF = False Then int库存检查 = zlCommFun.NVL(rs("检查方式").Value, 0)
                        
                        '0-不检查;1-检查,不足提醒;2-检查，不足禁
                        Select Case int库存检查
                        Case 0
                            
                        Case 1
                            Set obj = frmWait.mfrmMain
                            
                            If Not (obj Is Nothing) Then Unload frmWait
                            If PromptStorageWarn(Val(rsCharge("收费数量").Value), db库存数量, rsCharge("收费项目").Value, rsCharge("执行科室").Value, rsCharge("计算单位").Value, int库存检查, 2) <> 0 Then
                                If Not (obj Is Nothing) Then Call frmWait.OpenWait(obj, "请稍等...")
                                Exit Function
                            End If
                            If Not (obj Is Nothing) Then Call frmWait.OpenWait(obj, "请稍等...")
                        Case 2
                            Set obj = frmWait.mfrmMain
                            If Not (obj Is Nothing) Then Unload frmWait
                            Call PromptStorageWarn(Val(rsCharge("收费数量").Value), db库存数量, rsCharge("收费项目").Value, rsCharge("执行科室").Value, rsCharge("计算单位").Value, int库存检查, 2)
                            If Not (obj Is Nothing) Then Call frmWait.OpenWait(obj, "请稍等...")
                            Exit Function
                        End Select
                        
                    End If
                    
                End If
            
                strSQL = "Select y.现价,Decode(x.总价,0,0,null,0,round(y.现价/x.总价,2)) As 比例,y.收入项目id,z.收据费目 " & _
                            "From ( " & _
                            "Select a.收费细目id,Sum(a.现价) As 总价 " & _
                            "From 收费价目 a " & _
                            "Where a.执行日期 <= SYSDATE " & _
                                "and (a.终止日期 IS NULL OR a.终止日期>SysDate) " & _
                                "and a.收费细目id=[1] " & _
                            "Group By a.收费细目id " & _
                            ") x, " & _
                            "收费价目 y, " & _
                            "收入项目 z " & _
                            "Where Y.执行日期 <= SYSDATE " & _
                              "and (y.终止日期 IS NULL OR y.终止日期>SysDate) " & _
                              "and y.收入项目id=z.id " & _
                              "and y.收费细目id=[1]"
                              
                Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsCharge("收费细目id").Value))
                If rsPrice.BOF = False Then
                    lngCount = 0
                    dbSum = 0
                    Do While Not rsPrice.EOF
                        lngCount = lngCount + 1
                        
'                        db应收金额 = rsPrice("现价").Value * rsCharge("收费数量").Value
                        db应收金额 = rsCharge("应收金额").Value
                        db实收金额 = rsPrice("比例").Value * rsCharge("实收金额").Value
                        
                        If lngCount = rsPrice.RecordCount Then
                            db实收金额 = rsCharge("实收金额").Value - dbSum
                        Else
                            dbSum = dbSum + db实收金额
                        End If
                        
                        If zlCommFun.NVL(rsCharge("费别").Value) <> "" Then
                            strSQL = "Select Round(Round([3], 5) * 实收比率 / 100, [4]) As 实收金额 From 费别明细 Where 收入项目id = [1] And 费别 = [2] And (Round([3], 5) Between 应收段首值 And 应收段尾值)"
                            Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", Val(rsPrice("收入项目id").Value), CStr(zlCommFun.NVL(rsCharge("费别").Value)), db实收金额, ParamInfo.费用金额小数位数)
                            If rs.BOF = False Then
                                db实收金额 = zlCommFun.NVL(rs("实收金额").Value, db实收金额)
                            End If
                            
                        End If
                        
                        If ParamInfo.费用金额小数位数 > 0 Then
                            db应收金额 = Format(db应收金额, "0." & String(ParamInfo.费用金额小数位数, "0"))
                            db实收金额 = Format(db实收金额, "0." & String(ParamInfo.费用金额小数位数, "0"))
                        End If
                        
                        If zlCommFun.NVL(rsCharge("结算途径").Value, 1) = 1 Then
                            
                            strSQL = "Select Nvl(Max(序号),0)+1 As 序号 From 病人费用记录 Where No=[1] And 记录性质=[2]"
                            Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", CStr(rsCharge("No").Value), 2)
                            If rs.BOF = False Then int序号 = rs("序号").Value
        
                            strSQL = "zl_门诊记帐记录_Insert('" & rsCharge("No").Value & "'," & int序号 & "," & _
                                                            rsCharge("病人id").Value & "," & ZVal(rsCharge("门诊号").Value) & "," & _
                                                            "'" & rsCharge("姓名").Value & "','" & rsCharge("性别").Value & "'," & _
                                                            "'" & rsCharge("年龄").Value & "','" & rsCharge("费别").Value & "'," & _
                                                            "Null,0," & _
                                                            rsCharge("开嘱科室id").Value & "," & rsCharge("开嘱科室id").Value & "," & _
                                                            rsCharge("开嘱科室id").Value & ",'" & UserInfo.姓名 & "'," & _
                                                            "Null," & rsCharge("收费细目id").Value & "," & _
                                                            "'" & rsCharge("收费类别").Value & "','" & rsCharge("计算单位").Value & "'," & _
                                                            "1," & rsCharge("收费数量").Value & "," & _
                                                            "0," & rsCharge("执行科室id").Value & "," & _
                                                            "Null," & rsPrice("收入项目ID").Value & "," & _
                                                            "'" & rsPrice("收据费目").Value & "'," & rsPrice("现价").Value & "," & _
                                                            db应收金额 & "," & db实收金额 & "," & _
                                                            "To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "Null,0," & _
                                                            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                                            ZVal(lng类别id) & ",Null,'" & rsCharge("医嘱内容").Value & "'," & rsCharge("医嘱ID").Value & ",Null,Null,Null,1,0,4)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "mdlMedical")
                        Else
                            strSQL = "Select Nvl(Max(序号),0)+1 As 序号 From 病人费用记录 Where No=[1] And 记录性质=[2]"
                            Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", CStr(rsCharge("No").Value), 1)
                            If rs.BOF = False Then int序号 = rs("序号").Value
                            
                            strSQL = "zl_门诊划价记录_Insert('" & rsCharge("No").Value & "'," & int序号 & "," & _
                                                            rsCharge("病人id").Value & ",Null," & ZVal(rsCharge("门诊号").Value) & ",Null," & _
                                                            "'" & rsCharge("姓名").Value & "','" & rsCharge("性别").Value & "'," & _
                                                            "'" & rsCharge("年龄").Value & "','" & rsCharge("费别").Value & "'," & _
                                                            "Null," & _
                                                            rsCharge("开嘱科室id").Value & "," & rsCharge("开嘱科室id").Value & "," & _
                                                            rsCharge("开嘱科室id").Value & ",'" & UserInfo.姓名 & "'," & _
                                                            "Null," & rsCharge("收费细目id").Value & "," & _
                                                            "'" & rsCharge("收费类别").Value & "','" & rsCharge("计算单位").Value & "',Null," & _
                                                            "1," & rsCharge("收费数量").Value & "," & _
                                                            "0," & rsCharge("执行科室id").Value & "," & _
                                                            "Null," & rsPrice("收入项目ID").Value & "," & _
                                                            "'" & rsPrice("收据费目").Value & "'," & rsPrice("现价").Value & "," & _
                                                            db应收金额 & "," & db实收金额 & "," & _
                                                            "To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            "'医嘱发送','" & UserInfo.姓名 & "'," & _
                                                            ZVal(lng类别id) & ",'" & rsCharge("医嘱内容").Value & "'," & rsCharge("医嘱ID").Value & ",Null,Null,Null,1,0,4)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "mdlMedical")
                        End If
                        
                        rsPrice.MoveNext
                    Loop
                    
                End If
            End If
            rsCharge.MoveNext
        Loop
    End If
    
    MakeMedicalCharge = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function DataMove(ByVal strRec As String, Optional ByVal bytMode As Byte = 1) As Boolean
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    DataMove = False
    
    Select Case bytMode
    Case 1
        strSQL = "Select 1 From H体检登记记录 Where ID=[1]"
        strRec = Val(strRec)
    Case 2
        strSQL = "Select 1 From H体检人员档案 Where ID=[1]"
        strRec = Val(strRec)
    Case 3
        strSQL = "Select 1 From H体检登记记录 Where 体检号=[1]"
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlMedical", strRec)
    If rs.BOF = False Then
        DataMove = True
    End If
    
errHand:

End Function

Public Function DeleteMedicalItems(ByRef strSQL() As String, ByVal rs As ADODB.Recordset, ByVal str体检号 As String, ByVal lng登记id As Long, Optional ByVal lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            '作废此体检项目所产生的医嘱

            If lng病人id > 0 Then
                strSQL(ReDimArray(strSQL)) = "ZL_体检项目清单_DELETE(" & lng登记id & ",NULL," & Val(rs("清单id").Value) & "," & lng病人id & ")"
            Else
                strSQL(ReDimArray(strSQL)) = "ZL_体检项目清单_DELETE(" & lng登记id & ",'" & rs("组别").Value & "'," & Val(rs("清单id").Value) & ",0)"
            End If
            
            rs.MoveNext
        Loop
    End If
    
    DeleteMedicalItems = True
    
End Function

Public Function InsertMedicalItems(ByRef strSQL() As String, ByVal rs As ADODB.Recordset, ByVal lng登记id As Long, Optional ByVal lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  新加入体检项目
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngLoop As Long
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            strTmp = ""
            varRow = Split(rs("计费明细").Value, ";")
            For lngLoop = 0 To UBound(varRow)
                
                varCol = Split(varRow(lngLoop), ":")
                
                If strTmp <> "" Then strTmp = strTmp & ";"
                strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                
            Next
            
            '将此体检项目产生为医嘱
            If lng病人id > 0 Then
                
                strTmp = "ZL_体检项目清单_INSERT(" & lng登记id & "," & _
                                                    "NULL," & _
                                                    rs("ID").Value & ",'" & _
                                                    rs("体检类型").Value & "'," & _
                                                    Val(rs("基本价格").Value) & "," & _
                                                    Val(rs("体检价格").Value) & "," & _
                                                    Val(rs("执行科室id").Value) & "," & _
                                                    IIf(rs("采集方式id") = "", "NULL", rs("采集方式id")) & "," & _
                                                    IIf(rs("采集科室id") = "", "NULL", rs("采集科室id")) & ",'" & _
                                                    zlCommFun.NVL(rs("检验标本").Value) & "','" & _
                                                    rs("检查部位").Value & "','" & _
                                                    rs("检查部位id").Value & "'," & lng病人id & "," & IIf(rs("结算方式").Value = "记帐", "1", "2") & ",'" & strTmp & "')"
        
            Else
            
                strTmp = "ZL_体检项目清单_INSERT(" & lng登记id & ",'" & _
                                            rs("组别").Value & "'," & _
                                            rs("ID").Value & ",'" & _
                                            rs("体检类型").Value & "'," & _
                                            Val(rs("基本价格").Value) & "," & _
                                            Val(rs("体检价格").Value) & "," & _
                                            rs("执行科室id").Value & "," & _
                                            IIf(rs("采集方式id") = "", "NULL", rs("采集方式id")) & "," & _
                                            IIf(rs("采集科室id") = "", "NULL", rs("采集科室id")) & ",'" & _
                                            rs("检验标本").Value & "','" & _
                                            rs("检查部位").Value & "','" & _
                                            rs("检查部位id").Value & "',NULL," & IIf(rs("结算方式").Value = "记帐", "1", "2") & ",'" & strTmp & "')"
            End If
            
            strSQL(ReDimArray(strSQL)) = strTmp
            
            rs.MoveNext
        Loop
    End If
    
    InsertMedicalItems = True
    
End Function


Public Function DeleteItem(ByRef rsSQL As ADODB.Recordset, ByVal rs As ADODB.Recordset, ByVal str体检号 As String, ByVal lng登记id As Long, Optional ByVal lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            '作废此体检项目所产生的医嘱

            If lng病人id > 0 Then
                strTmp = "ZL_体检项目清单_DELETE(" & lng登记id & ",NULL," & Val(rs("清单id").Value) & "," & lng病人id & ")"
            Else
                strTmp = "ZL_体检项目清单_DELETE(" & lng登记id & ",'" & rs("组别").Value & "'," & Val(rs("清单id").Value) & ",0)"
            End If
            
            Call SQLRecordAdd(rsSQL, strTmp)
            
            rs.MoveNext
        Loop
    End If
    
    DeleteItem = True
    
End Function

Public Function NewItem(ByRef rsSQL As ADODB.Recordset, ByVal rs As ADODB.Recordset, ByVal lng登记id As Long, Optional ByVal lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  新加入体检项目
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngLoop As Long
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            strTmp = ""
            varRow = Split(rs("计费明细").Value, ";")
            For lngLoop = 0 To UBound(varRow)
                
                varCol = Split(varRow(lngLoop), ":")
                
                If strTmp <> "" Then strTmp = strTmp & ";"
                strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                
            Next
            
            '将此体检项目产生为医嘱
            If lng病人id > 0 Then
                
                strTmp = "ZL_体检项目清单_INSERT(" & lng登记id & "," & _
                                                    "NULL," & _
                                                    rs("ID").Value & ",'" & _
                                                    rs("体检类型").Value & "'," & _
                                                    Val(rs("基本价格").Value) & "," & _
                                                    Val(rs("体检价格").Value) & "," & _
                                                    Val(rs("执行科室id").Value) & "," & _
                                                    IIf(rs("采集方式id") = "", "NULL", rs("采集方式id")) & "," & _
                                                    IIf(rs("采集科室id") = "", "NULL", rs("采集科室id")) & ",'" & _
                                                    zlCommFun.NVL(rs("检验标本").Value) & "','" & _
                                                    rs("检查部位").Value & "','" & _
                                                    rs("检查部位id").Value & "'," & lng病人id & "," & IIf(rs("结算方式").Value = "记帐", "1", "2") & ",'" & strTmp & "')"
        
            Else
            
                strTmp = "ZL_体检项目清单_INSERT(" & lng登记id & ",'" & _
                                            rs("组别").Value & "'," & _
                                            rs("ID").Value & ",'" & _
                                            rs("体检类型").Value & "'," & _
                                            Val(rs("基本价格").Value) & "," & _
                                            Val(rs("体检价格").Value) & "," & _
                                            rs("执行科室id").Value & "," & _
                                            IIf(rs("采集方式id") = "", "NULL", rs("采集方式id")) & "," & _
                                            IIf(rs("采集科室id") = "", "NULL", rs("采集科室id")) & ",'" & _
                                            rs("检验标本").Value & "','" & _
                                            rs("检查部位").Value & "','" & _
                                            rs("检查部位id").Value & "',NULL," & IIf(rs("结算方式").Value = "记帐", "1", "2") & ",'" & strTmp & "')"
            End If
            
            Call SQLRecordAdd(rsSQL, strTmp)
            
            rs.MoveNext
        Loop
    End If
    
    NewItem = True
    
End Function

Public Function OutPutQuestBill(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngPatientKey As Long, ByVal strDeptID As String, ByVal strSample As String, _
                                Optional ByVal blnVerfiy As Boolean, Optional ByVal blnCheck As Boolean, Optional ByVal bytMode As Byte = 1) As Boolean
                                
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    
    On Error GoTo errHand
    
    strSQL = "Select * From (Select Nvl(c.样本条码,'0') As 样本条码,a.执行科室id,Decode(a.检验标本,Null,1,2) As 检验标本 " & _
                "From 体检项目清单 a,体检项目医嘱 b,病人医嘱发送 c " & _
                "Where b.清单id=a.ID And a.登记id=[1] And c.医嘱id(+)=b.医嘱id And b.病人id+0=[2] And (c.样本条码 Is Null Or (c.样本条码 Is Not Null And Instr([3],''''||a.检验标本||'''')>0)) Group By Decode(a.检验标本,Null,1,2),a.执行科室id,Nvl(c.样本条码,'0')) Order by 检验标本,执行科室id,样本条码"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, frmMain.Caption, lngKey, lngPatientKey, strSample)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            If InStr(strDeptID, "," & zlCommFun.NVL(rs("执行科室id").Value, 0) & ",") > 0 Then
                If zlCommFun.NVL(rs("样本条码").Value, "0") <> "0" Then
    
                    If blnVerfiy Then
                    
                        '检查是否有需要调用两次的
                        gstrSQL = "Select  Nvl(Max(C.排列顺序),1) As 调用次数 " & _
                                    "From 体检项目清单 A,诊疗项目目录 B,体检项目医嘱 E,体检项目排列 C,病人医嘱发送 d " & _
                                    "Where E.清单ID = A.ID " & _
                                        "AND B.ID=A.诊疗项目id And d.医嘱id(+)=e.医嘱id " & _
                                        "AND B.类别='C' AND C.诊疗项目id=B.ID AND C.排列性质=2 AND C.排列顺序>1 " & _
                                        "AND A.登记ID=[1] " & _
                                        "AND E.病人ID+0=[2] " & _
                                        "AND Nvl(d.样本条码,'0')=[3] " & _
                                        "AND A.执行科室ID+0=[4] " & _
                                        "AND Instr([5],''''||A.检验标本||'''')>0 "
                                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, frmMain.Caption, lngKey, lngPatientKey, zlCommFun.NVL(rs("样本条码").Value, "0"), Val(zlCommFun.NVL(rs("执行科室id").Value)), strSample)
                                        
                        For intCount = 1 To rsTmp("调用次数").Value
                            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_4", frmMain, "登记id=" & lngKey, "病人id=" & lngPatientKey, "样本条码=" & zlCommFun.NVL(rs("样本条码").Value, "0"), "执行科室id=" & zlCommFun.NVL(rs("执行科室id").Value, 0), "检验标本=" & strSample, bytMode)
                        Next
                        
                    End If
                    
                ElseIf blnCheck Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_5", frmMain, "登记id=" & lngKey, "病人id=" & lngPatientKey, "执行科室id=" & zlCommFun.NVL(rs("执行科室id").Value, 0), bytMode)
                End If
            End If
            rs.MoveNext
        Loop
    End If
            
    OutPutQuestBill = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function WriteFile(ByVal strFile As String, ByVal strText As String) As Boolean

    '******************************************************************************************************************
    '功能：写信息入指定文件
    '参数：文件名
    '返回：信息内容
    '******************************************************************************************************************
    
    Dim fso As New FileSystemObject
    Dim objTxt As TextStream
    
    On Error GoTo errHand
    
    Set objTxt = fso.OpenTextFile(strFile, ForAppending, True)
    objTxt.WriteLine strText
    
errHand:
    
End Function

