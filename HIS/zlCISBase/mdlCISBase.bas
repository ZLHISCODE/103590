Attribute VB_Name = "mdlCISBase"
Option Explicit
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrProductName As String
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long
Public gblnCancel As Boolean                '记录界面中的取消按钮是否被点击了

Public gstrDBOwner As String                '当前系统所有者
Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称
Public gstrItemName As String

Public gstrUnitName As String               '用户单位名称
Public gfrmMain As Object

Public gstrMatchMethod As String            '匹配方式:0表示双向匹配

Public gstrSql As String
Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号
Public gblnOK As Boolean


Public glngPreHWnd As Long '用于支持鼠标滚轮功能

Public gobjKernel As New clsCISKernel       '临床核心部件
Public gobjLogisticPlatform As Object       '物流平台接口
Public gstrPriceClass As String         '价格等级

Public gobjRIS As Object                    '新网RIS接口对象
Public Enum RISBaseItemOper                 '新网RIS基础数据操作类型：1-新增；2-修改；3-删除
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '新网RIS基础数据类型：1：诊疗项目目录，2：诊疗项目部位
    ClinicItem = 1
    ClinicItemPart = 2
End Enum

Public gblnKSSStrict As Boolean             '是否启用抗菌药物严格控制
Public gblnIncomeItem As Boolean            '记录收入项目是否设置

Public Type type_user_Digits
    dig_成本价 As Double
    dig_零售价 As Double
    dig_数量 As Double
    dig_金额 As Double
End Type
Public gtype_MaxDigits As type_user_Digits  '用来记录最大精度

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    用药级别 As Long
End Type
Public UserInfo As TYPE_USER_INFO
Public Const gstrLisHelp As String = "zl9LisWork"               'LIS调用帮助时使用的部件名
Public glngTXTProc As Long '保存默认的消息函数的地址
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const GCST_INVALIDCHAR = "'"             '对于输入的无效字符

'支持滑轮的常量
Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = -4

Public Const GWL_STYLE = (-16)
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

'私有、公共模块参数
Public Enum 参数_药品目录管理_公共
    P1_西成药收入项目 = 1
    P2_中成药收入项目 = 2
    P3_中草药收入项目 = 3
    P4_应用范围 = 4
    P5_时价药品按批次调价 = 5
End Enum
Public grsPriceGrade As ADODB.Recordset
 
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Function Select部门选择器(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str工作性质 As String = "", _
    Optional bln操作员 As Boolean = False, _
    Optional strSql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '功能:部门选择器
    '参数:objCtl-指定控件
    '     strSearch-要搜索的条件
    '     str工作性质-工作性质:如"V,W,K"
    '     bln操作员-是否加操作员限制
    '     strSQL-直接根据SQL获取数据(但部门表的别名一定要是A)
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim strPa As String
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strTittle = "部门选择器"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    strPa = zlDatabase.GetPara(44, glngSys, 0): strPa = IIf(strPa = "", "11", strPa)
    
    If strSql <> "" Then
    
        gstrSql = strSql
    Else
        gstrSql = "" & _
        "   Select distinct a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
        "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
    
        If str工作性质 = "" And bln操作员 = False Then
            gstrSql = gstrSql & vbCrLf & _
            "   From 部门表 a" & _
            "   Where 1=1"
        Else
            gstrSql = gstrSql & vbCrLf & _
            "   From 部门表 a, 部门性质分类 b,部门性质说明 c" & _
            "   Where c.工作性质 = b.名称" & IIf(str工作性质 = "", "(+)", " and B.编码 in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) ") & _
            "         AND a.id = c.部门id " & _
            IIf(bln操作员 = False, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        End If
        gstrSql = gstrSql & vbCrLf & _
            "   and  (a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) And (a.站点=[4] or a.站点 is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.编码 like upper([3]) or a.简码 like upper([3]) or a.名称 like [3] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
            If Mid(strPa, 1, 1) = "1" Then strFind = " And (A.编码 Like Upper([3]))"
        ElseIf zlStr.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式" ))
            If Mid(strPa, 2, 1) = "1" Then strFind = " And  (a.简码 Like Upper([3]))"
        ElseIf zlStr.IsCharChinese(strSearch) Then  '全汉字
            strFind = " And a.名称 Like [3] "
        End If
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strSql = "" Then
        gstrSql = gstrSql & _
        "   Start With A.上级id Is Null Connect By Prior A.ID = A.上级id "
    Else
        gstrSql = gstrSql & vbCrLf & strFind & vbCrLf & " Order by A.编码"
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strSql = "" Then
        '分上下级
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSql, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSql, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.ID, str工作性质, strKey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "没有满足条件的部门,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!ID) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            MsgBox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
        objCtl.Tag = Val(rsTemp!ID)
    End If
    zlCommFun.PressKey vbKeyTab
    Select部门选择器 = True
End Function

Public Function CheckPriceAdjust(ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long, Optional ByVal bln忽略零差价管理属性 As Boolean = False) As Boolean
    '零差价管理模式时，判断价格是否满足零差价管理要（成本价和售价一致）
    '定价药品：售价是固定的，比较所有药房的成本价，如果存在不一致的就不能销售出库
    '时价药品：比较药房库存记录的零售价和成本价，如果存在不一致的就不能销售出库
    '无库存时：成本价取药品规格的成本价
    '参数：lng药品id-药品规格ID，为0则检查所有药品；lng库房id-对应的库房ID，为0则检查所有库房；lng批次-对应的批次，如果传入-1则不关联批次
    '      bln忽略零差价管理属性：true-忽略零差价属性(用于批量修改属性时，界面修改但未实际保存)
    '返回：True-正常；false-有不满足零差价管理要求的药品
    '
    Dim rsData As ADODB.Recordset
    Dim str条件 As String
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(zlDatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjust = True: Exit Function
    
    '检查有无库存
    If lng药品ID > 0 Then
        If lng库房ID > 0 Then
            gstrSql = "Select 1 from 药品库存 Where 性质=1 and 药品id=[1] and 库房id=[2] " & _
                " And Not (批次 = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0)"
            
            If lng批次 > 0 Then
                gstrSql = gstrSql & " and Nvl(批次,0)=[3] "
            End If
        Else
            gstrSql = "Select 1 from 药品库存 Where 性质=1 and 药品id=[1] " & _
                " And Not (批次 = 0 And 可用数量 < 0 And 实际数量 = 0 And 实际金额 = 0 And 实际差价 = 0)"
        End If
        Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "CheckPriceAdjust", lng药品ID, lng库房ID, lng批次)
        
        If rsData.EOF Then
            '无库存时，从收费价目取售价，从药品规格取成本价
            gstrSql = "Select a.成本价, b.现价 As 售价 " & _
                " From 药品规格 A, 收费价目 B " & _
                " Where a.药品id = b.收费细目id And (Sysdate Between b.执行日期 And b.终止日期) " & IIf(bln忽略零差价管理属性 = False, " And Nvl(a.是否零差价管理, 0) = 1 ", "") & _
                " And b.现价 <> a.成本价 And a.药品id = [1] " & GetPriceClassString("B")
            Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "CheckPriceAdjust", lng药品ID)
            
            If rsData.EOF Then
                '没找到表示价格一致
                CheckPriceAdjust = True
            Else
                '找到表示价格不一致
                CheckPriceAdjust = False
            End If
            
            Exit Function
        End If
    End If
    
    If lng药品ID > 0 Then
        str条件 = IIf(str条件 = "", "", str条件) & " and a.药品id=[1] "
    End If
    
    If lng库房ID > 0 Then
        str条件 = IIf(str条件 = "", "", str条件) & " and d.库房id=[2] "
    End If
    
    If lng批次 >= 0 Then
        str条件 = IIf(str条件 = "", "", str条件) & " and nvl(d.批次,0)=[3] "
    End If
    
    If bln忽略零差价管理属性 = False Then
        str条件 = IIf(str条件 = "", "", str条件) & " And Nvl(a.是否零差价管理, 0) = 1 "
    End If
    
    gstrSql = "Select 药品id, 通用名, 规格, 0 As 库房id, '' As 库房, 生产商, '' As 批号, 批次, 单位, 药库包装, 售价, Sum(成本价 * 实际数量) / Sum(实际数量) As 成本价, 是否时价" & vbNewLine & _
        " From (Select a.药品id, '['|| c.编码 || ']'|| c.名称||decode(c.产地,null,null,'('||c.产地||')') ||c.规格 As 通用名, c.规格, c.产地 As 生产商, Null As 批次, a.药库单位 As 单位, a.药库包装, b.现价 As 售价," & vbNewLine & _
        "              nvl(d.平均成本价,a.成本价) As 成本价, 0 As 是否时价, d.实际数量" & vbNewLine & _
        "       From 药品规格 A, 收费价目 B, 收费项目目录 C, 药品库存 D" & vbNewLine & _
        "       Where a.药品id = b.收费细目id And a.药品id = c.Id And a.药品id = d.药品id And d.性质 = 1 And (Sysdate Between b.执行日期 And b.终止日期) And" & vbNewLine & _
        "             (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.是否变价 = 0  And" & vbNewLine & _
        "             b.现价 <> nvl(d.平均成本价,a.成本价) " & str条件 & GetPriceClassString("B") & vbNewLine & _
        "  And Not (D.批次 = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0))" & vbNewLine & _
        " Group By 药品id, 通用名, 规格, 生产商, 批次, 单位, 药库包装, 售价, 是否时价 " & vbNewLine & _
        " Having Sum(实际数量) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.药品id, '['|| c.编码 || ']'|| c.名称||decode(c.产地,null,null,'('||c.产地||')') ||c.规格 As 通用名, c.规格, d.库房id, e.名称 As 库房, d.上次产地 As 生产商, d.上次批号 As 批号, d.批次," & vbNewLine & _
        "       a.药库单位 As 单位, a.药库包装, d.零售价 As 售价, nvl(d.平均成本价,a.成本价) As 成本价, 1 As 是否时价" & vbNewLine & _
        " From 药品规格 A, 收费项目目录 C, 药品库存 D, 部门表 E" & vbNewLine & _
        " Where a.药品id = c.Id And a.药品id = d.药品id And d.库房id = e.Id And d.性质 = 1 And c.是否变价 = 1 And" & vbNewLine & _
        "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And nvl(d.零售价,0) <> nvl(d.平均成本价,a.成本价)" & vbNewLine & _
        " " & str条件 & "" & vbNewLine & _
        "  And Not (D.批次 = 0 And D.可用数量 < 0 And D.实际数量 = 0 And D.实际金额 = 0 And D.实际差价 = 0) " & vbNewLine & _
        " Order By 通用名,库房id,批号"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "CheckPriceAdjust", lng药品ID, lng库房ID, lng批次)
    
    '没找到不满足零差价管理要求的记录，返回true
    If rsData.EOF Then CheckPriceAdjust = True: Exit Function
    
    CheckPriceAdjust = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'功能：初始化新网接口部件
'参数：blnMsg－创建失败时是否提示
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
    End If
End Sub
Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Sub GetMaxDigit()
    '用来取药品的各种最大精度
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSql = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "最大精度")
    If rsTemp.RecordCount = 0 Then
        gtype_MaxDigits.dig_成本价 = 7
        gtype_MaxDigits.dig_金额 = 2
        gtype_MaxDigits.dig_零售价 = 7
        gtype_MaxDigits.dig_数量 = 7
    Else
        gtype_MaxDigits.dig_成本价 = rsTemp.Fields(1).NumericScale
        gtype_MaxDigits.dig_金额 = rsTemp.Fields(0).NumericScale
        gtype_MaxDigits.dig_零售价 = rsTemp.Fields(2).NumericScale
        gtype_MaxDigits.dig_数量 = rsTemp.Fields(3).NumericScale
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

'取药品金额、价格和数量的小数位数
Public Function GetDigit(ByVal int类别 As Integer, ByVal int内容 As Integer, Optional ByVal int单位 As Integer) As Integer
    'int类别：1-药品;2-卫材
    'int内容：1-成本价;2-零售价;3-数量;4-金额
    'int单位：如果是取金额位数，可以不输入该参数
    '         药品单位:1-售价;2-门诊;3-住院;4-药库;
    '         卫材单位:1-散装;2-包装
    '返回：最小2，最大为数据库最大小数位数
    
    Dim rsTmp As ADODB.Recordset
    Dim intMax金额 As Integer
    Dim intMax成本价 As Integer
    Dim intMax零售价 As Integer
    Dim intMax数量 As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSql = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, "取药品精度")
    
    intMax金额 = rs.Fields(0).NumericScale
    intMax成本价 = rs.Fields(1).NumericScale
    intMax零售价 = rs.Fields(2).NumericScale
    intMax数量 = rs.Fields(3).NumericScale
    
    gstrSql = "Select Nvl(精度, 0) 精度 From 药品卫材精度 Where 类别 = [1] And 内容 = [2] And 单位 = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "取药品" & Choose(int内容, "成本价", "零售价", "数量") & "小数位数", int类别, int内容, int单位)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!精度
    End If
    
    If GetDigit = 0 Then
        '如果没有设置精度，则取数据库允许的最大位数
        GetDigit = Choose(int内容, intMax成本价, intMax零售价, intMax数量)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int内容, intMax成本价, intMax零售价, intMax数量, intMax金额)
End Function


Public Function GetUserInfo() As Boolean
    '功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        gstrUserName = UserInfo.姓名
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MoveSpecialChar(ByVal strInputString As String, Optional ByVal blnMoveSpace As Boolean = True) As String
    '1 去除一般字符: " '_%?"，把_%?转换为对应的全角字符
    '2 去除特殊字符:退格、制表、换行、回车
    '3 blnMoveSpace，是否去掉字符中的空格，Ture-去掉空格；注意头尾空格默认去掉
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intAsc As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '允许转换的字符
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "？"
                Case "%"
                    strTmp = strTmp & "％"
                Case "_"
                    strTmp = strTmp & "＿"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intAsc = Asc(Mid(strText, n, 1))
        Select Case intAsc
            Case 8, 9, 10, 13
            Case 32
                '空格处理
                If blnMoveSpace = False Then
                    strTmp = strTmp & Mid(strText, n, 1)
                End If
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function

Public Function zlClinicCodeRepeat(strInputCode As String, Optional lngSelfID As Long) As Boolean
    '----------------------------------
    '功能：检查诊疗项目编码的是否与现有编码重复，重复则给出提示
    '入参：strInputCode-输入的编码；lngSelfID-自己的ID号，当修改时，需要将自身除开才能判断
    '出参：重复返回True；否则反馈Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.名称||' ['||I.编码||']'||I.名称 as 名称" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.类别=K.编码 and I.编码=[1] " & _
            "       and I.ID<>[2]"
    err = 0: On Error GoTo ErrHand
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", strInputCode, lngSelfID)
        
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "该项目与“" & !名称 & "”编码重复！", vbExclamation, gstrSysName
            zlClinicCodeRepeat = True
        Else
            zlClinicCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlClinicCodeRepeat = True
End Function


Public Function zlExistItem(ByVal strTbleName As String, ByVal strField As String, ByVal varValues As Variant, _
                            ByVal strItemName As String) As Boolean
    
    '----------------------------------
    '功能：检查项目是否存在,用于并发操作时的检查
    '入参：strTableName 表名 ,strField 字段名 , ,lngItemID,字段的值,strItemName 提示时显示的项目名称
    '出参：存在返回True；否则反馈Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    err = 0: On Error GoTo ErrHand
    strSql = "Select " & strField & " From " & strTbleName & " Where " & strField & "=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", varValues)
    If rsTmp.RecordCount > 0 Then
        zlExistItem = True
    Else
         MsgBox "“" & strItemName & "”已经被其他操作员删除！", vbExclamation, gstrSysName
        zlExistItem = False
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlExistItem = False
End Function

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

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intType As Integer
    Dim curDate As Date
    
    On Error GoTo errH
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    strSql = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!编号规则)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSql = Format(CDate(Format(rsTmp!日期, "YYYY-MM-dd")) - CDate(Format(rsTmp!日期, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & strSql & Format(Right(strNo, 4), "0000")
    Else
        '按年编号
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DelInvalidChar(ByVal strchar As String, Optional ByVal strInvalidChar As String) As String
    '删除非法字符
    'strChar: 要处理的字符
    'strInvalidChar：非法字符串，如果为空，则为~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,否则按传入的字符处理
    Dim strBit As String, i As Integer, strWord As String
    strWord = "~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>"
    If strInvalidChar <> "" Then strWord = strInvalidChar
    If Len(strchar) > 0 Then
        For i = 1 To Len(strchar)
            strBit = Mid$(strchar, i, 1)
            If InStr(strWord, strBit) <= 0 Then
                DelInvalidChar = DelInvalidChar & strBit
            End If
        Next
    End If
End Function

Public Function CheckKSSPrivilege() As Boolean
'功能：检查系统是否存在抗菌药物授权的人员，并且设置当前操作员的用药级别到UserInfo对象
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    UserInfo.用药级别 = 0
    
    On Error GoTo errH
    strSql = "Select 级别 From 人员抗菌药物权限 Where 记录状态=1 and 人员ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        UserInfo.用药级别 = Val("" & rsTmp!级别)
        CheckKSSPrivilege = True
    Else
        strSql = "Select 1 From 人员抗菌药物权限 Where 记录状态=1 and Rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel")
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub GetPriceClass()
    '根据登录站点获取药品的价格等级
    Dim rsData As ADODB.Recordset
    
    If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
        gstrSql = " Select a.价格等级 " & _
            " From 收费价格等级应用 A, 收费价格等级 B " & _
            " Where a.价格等级 = b.名称 And a.性质 = 0 And b.是否适用药品 = 1 And a.站点 = [1] And Nvl(b.撤档时间, Sysdate + 1) > Sysdate "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "GetPriceClass", gstrNodeNo)
        
        If rsData.RecordCount > 0 Then gstrPriceClass = rsData!价格等级
    End If
End Sub

Public Function GetPriceClassString(strTableName As String) As String
    '根据传入表的别名返回价格等级的条件串
    GetPriceClassString = " And " & IIf(strTableName = "", "价格等级 Is Null ", strTableName & ".价格等级 Is Null ")
    
End Function

Public Function zlGetrsPriceGrade(ByRef rsOutPriceGrade As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取价格等级记录集
    '入参:
    '出参:rsOutPriceGrade-返回价格等级，未启用或获取失赃物时，返回Nothing
    '返回:如果获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-06-30 14:08:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    If Not grsPriceGrade Is Nothing Then
        If grsPriceGrade.State = 1 Then Set rsOutPriceGrade = grsPriceGrade: zlGetrsPriceGrade = True: Exit Function
    End If
    '检查是否启用，如果启用
    strSql = "" & _
    "   Select 编码,名称 From 收费价格等级 A where nvl(撤档时间,sysdate+1)>sysdate Order by 编码"
    Set grsPriceGrade = zlDatabase.OpenSQLRecord(strSql, "获取价格等级")
    Set rsOutPriceGrade = grsPriceGrade
    zlGetrsPriceGrade = rsOutPriceGrade.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function FmgFlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'支持frmDoctorManage窗体滚轮的滚动
    On Error GoTo errH
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
            Case -7864320  '向下滚
                If frmDoctorManage.vscBar.Value <> frmDoctorManage.vscBar.Max Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageDown
                End If
            Case 7864320   '向上滚
                If frmDoctorManage.vscBar.Value <> 0 Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageUp
                End If
        End Select
    End Select
    FmgFlexScroll = CallWindowProc(glngPreHWnd, hWnd, wMsg, wParam, lParam)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowSpecChar(frmParent As Object) As String
'功能：以模态窗体运行特殊字符程序
'参数：frmParent=调用父窗体
'返回：选择的特殊字符串；取消操作返回空
    Dim frmNew As frmSpecChar
    Set frmNew = New frmSpecChar
    frmNew.Show 1, frmParent
    If gblnOK Then ShowSpecChar = frmNew.mstrChar
End Function

Public Sub ArrayIcons(objLvw As ListView, Optional intBegin As Integer = 1, Optional blnShow As Boolean)
'功能：根据第一个图标的位置重新排列所有图标
    Dim i As Integer, t As Long
    Dim r As RECT

    Call GetClientRect(objLvw.hWnd, r)
    
    If blnShow Then
        If objLvw.ListItems(intBegin).Top < 30 Then
           objLvw.ListItems(intBegin).Top = 30
        ElseIf objLvw.ListItems(intBegin).Top + objLvw.ListItems(intBegin).Height > (r.Bottom - r.Top) * Screen.TwipsPerPixelY Then
            objLvw.ListItems(intBegin).Top = (r.Bottom - r.Top) * Screen.TwipsPerPixelY - objLvw.ListItems(intBegin).Height
        End If
    End If
    
    '下面的图标
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            'Item的Width包含文字部分,Left仅指图标
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t + .Height
        End With
    Next
    
    '上面的图标
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To 1 Step -1
        With objLvw.ListItems(i)
            'Item的Width包含文字部分,Left仅指图标
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t - .Height
        End With
    Next
End Sub
