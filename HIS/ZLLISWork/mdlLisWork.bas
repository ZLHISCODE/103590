Attribute VB_Name = "mdlLisWork"
Option Explicit

Public gobjEmrInterface As Object           '新版病历申请附项读取部件
Public gobjpublicExpenses As Object

Public Enum COLOR
    
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    橙色 = &H40C0&
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
End Enum
'    报警背景色 = &HFF&
'    报警前景色 = &H8000000F
'    超标背景色 = &H80FF&
'    超标前景色 = &H80000008
'    默认前景色 = &H80000008


'取得当前形所有列的SQL：select Wmsys.Wm_Concat(Column_Name) From User_Tab_Columns Where Table_Name = Upper('病人余额')
Public Const gConst_病人信息_列名 As String = "a.病人id,a.门诊号,a.住院号,a.就诊卡号,a.卡验证码,a.费别,a.医疗付款方式,a.姓名,a.性别,a.年龄,a.出生日期," & _
                                              "a.出生地点,a.身份证号,a.身份,a.职业,a.民族,a.国籍,a.区域,a.学历,a.婚姻状况,a.家庭地址,a.家庭电话,a.家庭地址邮编," & _
                                              "a.联系人关系,a.联系人地址,a.联系人电话,a.合同单位ID,a.工作单位,a.单位电话,a.单位邮编,a.单位开户行,a.单位帐号," & _
                                              "a.担保人,a.担保性质,a.就诊时间,a.就诊状态,a.就诊诊室,a.住院次数,a.当前科室ID,a.当前病区ID,a.入院时间,a.出院时间," & _
                                              "a.IC卡号,a.健康号,a.险类,a.登记时间,a.停用时间,a.当前床号,a.医保号,a.查询密码,a.在院,a.其他证件,a.监护人,a.锁定,a.主页id"

                                              
Public Const gConst_病人余额_列名 As String = " a.病人ID,a.性质,a.预交余额,a.费用余额 "

Public Const gConst_检验仪器_列名 As String = "a.波长,a.振板频率,a.振板时间,a.进板方式,a.空白形式,a.对数质控图,a.发送时指定杯号,a.质控水平数,a.上次质控日,a.QC码," & _
                                              "a. 试剂来源,a.校准物来源,a.ID,a.编码,a.名称,a.简码,a.连接计算机,a.通讯程序名,a.通讯端口,a.波特率,a.数据,a.停止位," & _
                                              "a.校验位,a. 仪器类型,a.仪器标志色,a.使用小组ID,a.质控标本号,a.备注,a.微生物,a.转换日期,a.转换仪器ID,a.质控周期,a.周期单位"
                                              
Public Const gConst_病人医嘱发送_列名 As String = "a.接收批次,a.完成时间,a.送检人,a.条码打印,a.重采标本,a.标本送出时间,a.完成人,a.医嘱ID,a.发送号,a.记录性质,a.NO," & _
                                                  "a. 记录序号,a.发送数次,a.发送人,a.发送时间,a.首次时间,a.末次时间,a.执行状态,a.执行部门ID,a.计费状态,a.执行间," & _
                                                  "a.执行过程,a. 采样人,a.采样时间,a.样本条码,a.报告ID,a.结果阳性,a.报到时间,a.执行说明,a.接收人,a.接收时间,a.安排时间"
                                                  
Public Const gConst_检验标本记录_列名 As String = "a.主页ID,a.检验项目,a.操作类型,a.接收人,a.接收时间,a.标识号,a.床号,a.病人科室,a.杯号," & _
                                                  "a.初审人,a. 初审时间,a.一级报告,a.二级报告,a.三级报告,a.审核未通过,a.年龄数字,a.年龄单位,a.紧急," & _
                                                  "a.挂号单,a.门诊号,a. 住院号,a.出生日期,a.ID,a.医嘱ID,a.标本序号,a.采样时间,a.采样人,a.标本类型,a.核收人," & _
                                                  "a.核收时间,a.样本状态,a.检验人,a. 检验时间,a.审核人,a.审核时间,a.合并报告号,a.打印次数,a.申请类型,a.仪器ID," & _
                                                  "a.样本条码,a.报告结果,a.备注,a.未通过审核原因,a. 申请时间,a.标本形态,a.是否质控品,a.执行科室ID,a.微生物标本," & _
                                                  "a.NO,a.是否传送,a.标本类别,a.检验备注,a.申请人,a.申请科室ID,a. 病人来源,a.病人ID,a.婴儿,a.姓名,a.性别,a.年龄,a.合并ID"
                                                  
Public Const gconst_病人医嘱记录_列名 As String = "a.可否分零,a.屏蔽打印,a.检查方法,a.执行标记,a.送检人,a.上次打印时间,a.体检号,a.审核标记,a.搞要,a.执行性质," & _
                                                  "a.紧急标志,a.开始执行时间,a.执行终止时间,a.上次执行时间,a.开嘱科室ID,a.开嘱医生,a.开嘱时间,a.校对护士,a.校对时间," & _
                                                  "a.停嘱医生,a.停嘱时间,a.确认停嘱时间,a.申请ID,a.是否上传,a.审查结果,a.首次保存时间,a.摘要,a.ID,a.相关ID,a.前提ID," & _
                                                  "a.病人来源,a.病人ID,a.主页ID,a.挂号单,a.婴儿,a.姓名,a.性别,a.年龄,a.病人科室ID,a.序号,a.医嘱状态,a.医嘱期效," & _
                                                  "a.诊疗类别,a.诊疗项目ID,a.标本部位,a.收费细目ID,a.天数,a.单次用量,a.总给予量,a.医嘱内容,a.医生嘱托,a.执行科室ID," & _
                                                  "a.皮试结果,a.执行频次,a.频率次数,a.频率间隔,a.间隔单位,a.执行时间方案,a.计价特性"


Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrProductName As String            '产品简称，例如：中联
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long
Public gcolPrivs As Collection              '记录内部模块的权限
Public gintSelectFocus As Integer           '由于控件焦点有问题，人为选择正确窗体的焦点
                                            '1=Dkp_ID_List;2=frmLabRequest;3=frmLisStationWrite;4=frmLisStationWrite2(1)
                                            '5=frmLisStationWrite2(2)

'医保变量
Public gclsInsure As New clsInsure
Public gblnInsure As Boolean '是否连接医保
Public gintInsure As Integer

Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称
Public gobjKernel As New clsCISKernel       '临床核心部件
Public gobjRichEPR As New cRichEPR          '病历核心部件
Public gbytCardNOLen As Long                '就诊卡长度

Public gobjEmr As Object                    '病历部件


Public gstrUnitName As String               '用户单位名称
Public gfrmMain As Object

Public gstrSql As String
Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号
Public gblnOK As Boolean
Public gLabcboDept As Object

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO


'HIS系统参数

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
    support出院无实际交易 = 29 '出院接口中是否要与接口商进行交易
End Enum

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
End Enum

'电子签名
Public gintCA As Integer '电子签名认证中心
Public gstrESign As String '电子签名控制场合
Public gobjESign As Object '电子签名接口部件

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const LONG_MAX = 2147483647 'Long型最大值
Public Const CuvetteNumberLen = 12 '试管条码长度

Public glngTXTProc As Long '保存默认的消息函数的地址

Public grsDuty As ADODB.Recordset '存放医生职务
Public grsSysPars As ADODB.Recordset
Public gbln手术授权管理 As Boolean  '是否启用手术按医师授权管理


Public gblnManualPH As Boolean '手工使用批号作为标本号
Public gintNumberPH As Integer '每批最大标本数
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip
Private mlngImageID As Long  '避免重新读取相同图片
Public gbln执行后审核 As Boolean    '执行后自动审核划价单
Public mobjLisInsideComm As Object                                      'LIS内部接口
Public mobjZLIHISPlugIn As Object                                       'ZLHIS通用插件接口
Private mintWarn As Integer                                             '-1=要显示,0=缺省为否,1-缺省为是

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    Set rsTmp = zlDatabase.GetUserInfo
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
    On Error GoTo ErrHand
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "生成字符串的简码")
    With rsTmp
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

ErrHand:
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
    On Error GoTo errH
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    X = objPoint.X * 15 + objBill.CellLeft
    Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal blnMerge As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '   blnMerge：是否合并相同ID的行。如果合并，则其他不同的列值以“;”分隔
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngloop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngloop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        lngCurrRow = -1
        If blnMerge Then lngCurrRow = FindGridLine(objMsf, CStr(zlCommFun.Nvl(rsData("ID"))))
        If lngCurrRow = -1 Then
            lngRow = lngRow + 1
            lngCurrRow = lngRow
        End If
        
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
        
        On Error GoTo ErrHand
        
        For lngloop = 0 To objMsf.Cols - 1
            
            If Trim(objMsf.TextMatrix(0, lngloop)) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngloop)
                                        
                On Error GoTo ErrHand
                
                strOldValue = objMsf.TextMatrix(lngCurrRow, lngloop)
                If strMask <> "" Then
                    strNewValue = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop))), strMask)
                Else
                    strNewValue = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop)))
                End If
                objMsf.TextMatrix(lngCurrRow, lngloop) = IIf(Trim(strOldValue) = "", strNewValue, _
                     strOldValue & IIf(InStr(";" & strOldValue & ";", ";" & strNewValue & ";") > 0, "", ";" & strNewValue))
            End If
            
        Next
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function FindGridLine(ByRef objMsf As Object, ByVal strSeekID As String) As Long
    '-------------------------------------------------------------------------------------------------------------
    '功能:查找RowData等于strSeekID的行
    '参数:
    '返回:行号或-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    FindGridLine = -1
    For i = 1 To objMsf.Rows - 1
        If objMsf.RowData(i) = strSeekID Then Exit For
    Next
    If i <= objMsf.Rows - 1 Then FindGridLine = i
End Function

Public Function FillListData(ByRef objLvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '-------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem
    Dim lngloop As Long
    
    Dim blnForeColor As Boolean
    
    On Error Resume Next
    
    blnForeColor = (rs("前景色").Name = "前景色")
    
    On Error GoTo ErrHand
    
    LockWindowUpdate objLvw.hWnd
    
    
    Do While Not rs.EOF
        
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("名称").Value, rs("图标").Value, rs("图标").Value)
        For lngloop = 2 To objLvw.ColumnHeaders.Count
            objItem.SubItems(lngloop - 1) = zlCommFun.Nvl(rs(objLvw.ColumnHeaders(lngloop).Text).Value)
        Next
        
        If blnForeColor Then
            objItem.ForeColor = Val(rs("前景色").Value)
            For lngloop = 2 To objLvw.ColumnHeaders.Count
                objItem.ListSubItems(lngloop - 1).ForeColor = objItem.ForeColor
            Next
        End If
                        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillListData = True
    
    Exit Function
ErrHand:
    LockWindowUpdate 0
    If ErrCenter = 1 Then Resume
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
        If InStr("0123456789<>", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.-<>+Ee", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    Dim lngloop As Long
    
    Select Case bytMode
    Case 1
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 99
        For lngloop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngloop, 1)) = 0 Then
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
    
    Set rs = zlDatabase.OpenSQLRecord("SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", "mdlCISBase")
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
    objVsf.Rows = 1
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

Public Function CheckIsAllowAuditing(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：检验是否允许审核,即是否满足审核的条件
    '--------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
        
    strSQL = "SELECT ROWNUM AS 序号,A.核收时间 AS 标本时间,A.标本序号, D.中文名 AS 检验项目, A.检验结果 AS 本次结果,B.检验结果 AS 上次结果 " & _
                 "FROM (SELECT A.核收时间,A.标本序号, B.检验项目id, B.检验结果, A.检验时间 " & _
                         "FROM 检验标本记录 A, 检验普通结果 B " & _
                        "WHERE A.ID = B.检验标本ID AND A.报告结果 = B.记录类型 AND A.ID = [1]) A, " & _
                      "(SELECT C.检验项目id, C.检验结果, A.检验时间 " & _
                         "FROM 检验标本记录 A,病人医嘱记录 B,检验普通结果 C, " & _
                              "(SELECT B.病人ID, B.主页ID " & _
                                 "FROM 检验标本记录 A, 病人医嘱记录 B " & _
                                "WHERE A.医嘱ID + 0 = B.ID AND A.ID = [1] ) D " & _
                        "WHERE (C.检验项目id,A.检验时间) IN (SELECT D.检验项目id,MAX(A.检验时间) " & _
                                           "FROM 检验标本记录 A,病人医嘱记录 B,检验普通结果 D," & _
                                                "(SELECT B.病人ID, B.主页ID, A.检验时间 " & _
                                                   "FROM 检验标本记录 A, 病人医嘱记录 B,检验普通结果 C " & _
                                                  "WHERE C.检验标本ID=A.ID AND A.报告结果=C.记录类型 AND A.医嘱ID + 0 = B.ID AND A.ID = [1] ) C " & _
                                          "WHERE A.检验时间 < C.检验时间 AND A.医嘱ID = B.ID AND D.检验项目id=C.检验项目id AND A.报告结果=D.记录类型 AND D.检验标本ID=A.ID AND " & _
                                                "B.病人ID = C.病人ID AND NVL(B.主页ID,0) = NVL(C.主页ID,0) GROUP BY D.检验项目id) AND " & _
                              "A.医嘱ID = B.ID AND B.病人ID = D.病人ID AND " & _
                              "NVL(B.主页ID,0) = NVL(D.主页ID,0) AND C.检验标本ID = A.ID AND " & _
                              "A.报告结果 = C.记录类型) B, " & _
                      "检验项目 C,诊治所见项目 D " & _
                "WHERE A.检验项目id = B.检验项目id(+) AND C.结果类型 = 1 AND " & _
                      "C.结果异常条件 IS NOT NULL AND C.诊治项目id = D.ID AND " & _
                      "C.诊治项目id = A.检验项目id AND " & _
                      "(A.检验时间 - B.检验时间) <=TO_NUMBER(SUBSTR(C.结果异常条件, 1, INSTR(C.结果异常条件, ';') - 1)) AND " & _
                      "ABS(TO_NUMBER(A.检验结果) - TO_NUMBER(B.检验结果)) >=TO_NUMBER(SUBSTR(C.结果异常条件,INSTR(C.结果异常条件, ';') + 1,LENGTH(C.结果异常条件) - INSTR(C.结果异常条件, ';')))"
                      
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngKey)
    
    CheckIsAllowAuditing = (rs.BOF = True)
    If rs.BOF = False Then
        CheckIsAllowAuditing = frmLisStationError.ShowError(frmMain, rs)
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal BeginDate As String) As String
    '-----------------------------------------------------------------------------------------
    '功能:获取特殊时间
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    Dim dateNow As Date
    
    If BeginDate = "" Then
        dateNow = zlDatabase.Currentdate
    Else
        dateNow = BeginDate
    End If
    
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(dateNow, "YYYY-MM-DD")))
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 2, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 8 - intDay, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dateNow, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(dateNow, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(dateNow, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "指定开始日期"
        If bytFlag = 1 Then
            GetDateTime = zlDatabase.GetPara("历次检验范围指定开始日期", 100, 1208, Format(dateNow - 30, "yyyy-mm-dd 00:00:00"))
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "不重复"
        If bytFlag = 1 Then
            GetDateTime = "2000-01-01 00:00:00"
        Else
            GetDateTime = "3000-12-31 23:59:59"
        End If
    Case "自定义"
        GetDateTime = "自定义"
    End Select
    
End Function

Public Sub ApplyResultColor(vsf As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim lngColor As Long, lngForeColor As Long
    Dim lngReferenceLow As Long                             '参考低颜色
    Dim lngReferenceHigh As Long                            '参考高颜色
    Dim lngReferenceExigency As Long                        '参考警示颜色
    
    '读取颜色
    lngReferenceLow = Val(zlDatabase.GetPara("参考颜色_偏低", 100, 1208, 0))
    If lngReferenceLow = 0 Then lngReferenceLow = 8454143
    lngReferenceHigh = Val(zlDatabase.GetPara("参考颜色_偏高", 100, 1208, 0))
    If lngReferenceHigh = 0 Then lngReferenceHigh = 8438015
    lngReferenceExigency = Val(zlDatabase.GetPara("参考颜色_警示", 100, 1208, 0))
    If lngReferenceExigency = 0 Then lngReferenceExigency = 16576
    
    Select Case bytMode
        Case 0, 1
            lngColor = &H80000005
            lngForeColor = COLOR.默认前景色
        Case 5, 6 '异常低、高
            lngColor = lngReferenceExigency
            lngForeColor = COLOR.报警前景色
        Case 2
            lngColor = lngReferenceLow
            lngForeColor = COLOR.超标前景色
        Case Else
            lngColor = lngReferenceHigh
            lngForeColor = COLOR.超标前景色
    End Select
    
    vsf.Cell(flexcpBackColor, lngRow, lngCol, lngRow, lngCol) = lngColor
    vsf.Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = lngForeColor
    
    
End Sub

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
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
End Sub

Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = objVsf.BackColorSel
    End If
    
End Sub

Public Function GetReportCode(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能;
    '--------------------------------------------------------------------------------------------------------
    Dim rsPaitentType As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng医嘱ID = 0 And lng发送号 = 0 Then Exit Function
    
'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
                       "A.NO," & _
                       "A.记录性质 " & _
                "FROM 病人医嘱发送 A,病历文件列表 C,病人医嘱记录 D,病历单据应用 E " & _
                "Where E.病历文件id = C.ID " & _
                        "AND D.诊疗项目ID=E.诊疗项目ID " & _
                      "AND A.医嘱ID=D.ID AND E.应用场合=Decode(D.病人来源,2,2,4,4,1) " & _
                      " AND D.相关id= [1] "
    strSQL = "select b.病人性质 from 病人医嘱记录 A,病案主页 B  where a.病人id=b.病人id and a. 相关id = [1]"
    Set rsPaitentType = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lng医嘱ID)
    
    If rsPaitentType.RecordCount > 0 Then
        strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.编号, '00000')) || '-2' As 报表编号, A.NO, A.记录性质, F.ID, F.编码" & vbNewLine & _
                "From 病人医嘱发送 A, 病历文件列表 C, 病人医嘱记录 D, 病历单据应用 E, 诊疗项目目录 F,病案主页 G" & vbNewLine & _
                "Where E.病历文件id = C.ID And D.诊疗项目id = E.诊疗项目id And D.诊疗项目id = F.ID And A.医嘱id = D.ID and d.病人id=g.病人id And" & vbNewLine & _
                "      E.应用场合 = Decode(D.病人来源, 2, Decode(g.病人性质,1,1,2), 4, 4, 1) And D.相关id = [1] " & vbNewLine & _
                "Order By F.编码 "
    Else
        
        strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.编号, '00000')) || '-2' As 报表编号, A.NO, A.记录性质, F.ID, F.编码" & vbNewLine & _
                "From 病人医嘱发送 A, 病历文件列表 C, 病人医嘱记录 D, 病历单据应用 E, 诊疗项目目录 F" & vbNewLine & _
                "Where E.病历文件id = C.ID And D.诊疗项目id = E.诊疗项目id And D.诊疗项目id = F.ID And A.医嘱id = D.ID And" & vbNewLine & _
                "      E.应用场合 = Decode(D.病人来源, 2, 2, 4, 4, 1) And D.相关id = [1] " & vbNewLine & _
                "Order By F.编码 "
    End If
                          
    If DataMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If

'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
'                       "A.NO," & _
'                       "A.记录性质 " & _
'                "FROM 病历单据应用 A,病历文件目录 C,病人医嘱记录 D,病人医嘱发送 B " & _
'                "Where A.病历文件id = C.ID " & _
'                      "AND A.诊疗项目id=D.诊疗项目ID " & _
'                      "AND B.病人ID=D.病人ID " & _
'                      "AND NVL(B.主页ID,0)=NVL(D.主页ID,0) " & _
'                      "AND B.文件id=C.ID " & _
'                      "AND D.相关id=" & lng医嘱id & " " & _
'                      "AND A.发送号=" & lng发送号

    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lng医嘱ID, lng发送号)
                      
    
    If rs.BOF = False Then
        strCode = zlCommFun.Nvl(rs("报表编号"))
        strNO = zlCommFun.Nvl(rs("NO"))
        bytMode = zlCommFun.Nvl(rs("记录性质"), 1)
    End If
    
    GetReportCode = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckChargeState(ByVal lngKey As Long, Optional ByVal blnOrder As Boolean = True, Optional ByVal DataMoved As Boolean = False) As Boolean
    '检验收费状态
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim strSQLbak As String
    Dim intPatientType As Integer               '病人来源
    On Error GoTo errH
    
    CheckChargeState = False
    
    strSQL = "select 病人来源 from 检验标本记录 where id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "检验查费用", lngKey)
    If rs.EOF = True Then Exit Function
    intPatientType = rs("病人来源")
    
    If blnOrder Then
        strSQL = _
            "select NVL(A.记录状态,0) As 记录状态 " & _
                  "from 住院费用记录 A, " & _
                  "( " & _
                       "select No from 病人医嘱发送 where 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE [1] In (ID,相关id))  " & _
                       "Union " & _
                       "select No from 病人医嘱附费 where 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE [1] In (ID,相关id)) " & _
                  ") B " & _
                "Where A.NO = B.NO "
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
        End If
    Else
        strSQL = _
            "select NVL(A.记录状态,0) As 记录状态 " & _
                  "from 住院费用记录 A, " & _
                  "( " & _
                       "select No,记录性质 from 病人医嘱发送 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id In (A.ID,A.相关id) and A.诊疗类别 = 'C' ) " & _
                       "Union " & _
                       "select No,记录性质 from 病人医嘱附费 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id In (A.ID,A.相关id) and A.诊疗类别 = 'C' ) " & _
                  ") B " & _
                "Where A.NO = B.NO and mod(a.记录性质,10) = b.记录性质 "
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
        End If
    End If
    
    strSQL = strSQL & " Order by 记录状态 "
    If DataMoved Then
        strSQL = Replace(strSQL, "住院费用记录", "H住院费用记录")
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "检验标本记录", "H检验标本记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱附费", "H病人医嘱附费")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

    If rs.BOF Then Exit Function
    If rs("记录状态").Value = 0 Then Exit Function
    
    CheckChargeState = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadAdvicePrice(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal str费别 As String, Optional ByVal DataMoved As Boolean = False) As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能：读取指定医嘱的计价关系到临时记录集
    '说明：要计算的项目应该不是叮嘱,院外执行,无需计费
    '----------------------------------------------------------------------------------------------------
    Dim rsAdvice As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
        
    On Error GoTo errH
            
    '读取要计算主费用的医嘱记录(包含附加手术,检查部位；手术麻醉单独)
    strSQL = _
        " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID," & _
        " Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
        " Where B.相关ID= [1] " & _
        " And A.医嘱ID=B.ID And A.发送号+0=" & lng发送号 & _
        " And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
        " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,A.发送数次"
    strSQL = strSQL & " Union ALL " & _
        " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID," & _
        " Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
        " Where B.ID= [1] " & _
        " And A.医嘱ID=B.ID And A.发送号+0=" & lng发送号 & _
        " And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
        " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,A.发送数次" & _
        " Order by 序号"
    If DataMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱执行", "H病人医嘱执行")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
            
    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lng医嘱ID)
    
    For i = 1 To rsAdvice.RecordCount

        strSQL = _
            " Select 1 " & _
            " From 诊疗收费关系 A,收费价目 B,收费项目目录 C,收入项目 D" & _
            " Where A.诊疗项目ID= [1] " & _
            " And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And B.收入项目ID=D.ID" & _
            " And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
            " And Nvl(A.固有对照,0)=1 And Nvl(C.是否变价,0)=0"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", Val(rsAdvice!诊疗项目ID))
        
        If rsTmp.RecordCount > 0 Then
            LoadAdvicePrice = True
            Exit Function
        End If
        
        rsAdvice.MoveNext
    Next
        
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetColNumber(objVsf As Object, ByVal strCaption As String) As Long
    
    Dim lngloop As Long
    
    GetColNumber = -1
    
    For lngloop = 0 To objVsf.Cols - 1
        If objVsf.TextMatrix(0, lngloop) = strCaption Then
            GetColNumber = lngloop
            Exit Function
        End If
    Next
    
End Function

Public Sub VsfCellFormat(objVsf As Object, ByVal lngCol As Long, ByVal strFormat As String, Optional ByVal iType As Integer = -1, Optional ByVal iTypeCol As Integer = -1)
    'iType：需格式化的数据类型
    '  0：数字、1：字符、2：日期、3：逻辑、-1：不限（缺省）
    'iTypeCol：数据类型的存储字段序号
    Dim lngloop As Long
    On Error GoTo errH
    For lngloop = 1 To objVsf.Rows - 1
        If iType = -1 Then
            objVsf.TextMatrix(lngloop, lngCol) = Format(objVsf.TextMatrix(lngloop, lngCol), strFormat)
        Else
            If iTypeCol = -1 Then
                If iType = 0 And IsNumeric("-" & objVsf.TextMatrix(lngloop, lngCol)) Then objVsf.TextMatrix(lngloop, lngCol) = Format(objVsf.TextMatrix(lngloop, lngCol), strFormat)
            Else
                If iType = Val(objVsf.TextMatrix(lngloop, iTypeCol)) Then objVsf.TextMatrix(lngloop, lngCol) = Format(objVsf.TextMatrix(lngloop, lngCol), strFormat)
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DrawLine(pic As PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    On Error GoTo errH
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Public Function CreateVsf(ByRef objVsf As Object, ByVal strVsf As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '
    '-------------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim varArray As Variant
    Dim varItem As Variant
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    objVsf.Cols = 0
    
    varArray = Split(strVsf, ";")
    For lngloop = 0 To UBound(varArray)
        varItem = Split(varArray(lngloop), ",")
                
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
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngloop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errH
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngloop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngloop) = False Then
            lngLastRow = lngloop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngloop = 1 To objLineX.UBound
        objLineX(lngloop).Visible = False
    Next
    
    For lngloop = 1 To objLineY.UBound
        objLineY(lngloop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngloop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngloop Then Load objLineY(lngloop)

        With objLineY(lngloop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngloop) - 15
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
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function



Public Function ShowGrdFilterDialog(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 6000, _
                                    Optional ByVal lngCY As Long = 3000, _
                                    Optional ByVal blnFilter As Boolean = True, _
                                    Optional ByVal blnPrompt As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;显示文本输入选择列表(只用于表格控件)
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    
    On Error GoTo ErrHand

    If InStr(objVsf.EditText, "'") > 0 Then Exit Function
        
    Call ClientToScreen(objVsf.hWnd, objPoint)
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
            
    '执行查询
    Set rs = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If rs.BOF Then
        If blnPrompt Then MsgBox "没有找到相匹配的结果！", , gstrSysName
        Exit Function                            '没有结果，直接返回
    End If
            
    If rs.RecordCount = 1 And blnFilter Then GoTo Over                    '因为是输入查找，如果只有一条，则直接返回
    If frmSelectList.ShowSelect(frmParent, rs, strLvw, lngX, lngY, lngCX, lngCY, strSavePath, strDescrible, , , objVsf.CellHeight) Then GoTo Over
    
    Exit Function
    
Over:
    
    Set rsResult = rs
    
    ShowGrdFilterDialog = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowGrdSelectDialog(ByVal frmParent As Object, _
                                    ByVal objVsf As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal strSQL As String, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开树型+列表结构,应用于表格控件
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngX As Long
    Dim lngY As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
        
    If Trim(strSQL) = "" Then Exit Function
    
    On Error GoTo ErrHand
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption)
    If rs.BOF Then
        MsgBox "没有可选择的数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ClientToScreen(objVsf.hWnd, objPoint)
    
    lngX = objPoint.X * Screen.TwipsPerPixelX + objVsf.CellLeft
    lngY = objPoint.Y * Screen.TwipsPerPixelY + objVsf.CellTop + objVsf.CellHeight
    
    
    If frmSelectExplorer.ShowSelect(frmParent, rs, lngX, lngY, lngCX, lngCY, objVsf.CellHeight, strSavePath, strLvw, strDescrible) Then
                        
        Set rsResult = rs
        ShowGrdSelectDialog = True
        
    End If
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
    
End Function

Public Sub LocationVsf(objVsf As Object, ByVal lngRow As Long, ByVal lngCol As Long)
    
    On Error Resume Next
    
    objVsf.Row = lngRow
    objVsf.Col = lngCol
    objVsf.ShowCell objVsf.Row, objVsf.Col
    objVsf.SetFocus
End Sub

Public Function CheckNumeric(ByVal strText As String, ByVal lngLength As Long, Optional ByVal lngDecLength As Long = 0, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:检测字符串的数值有效性
    '--------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    
    Dim str整数部份 As String
    Dim str小数部份 As String
    On Error GoTo errH
    If lngDecLength = 0 Then
        '整数
        Select Case bytMode
        Case 1      '正整数
            str整数部份 = strText
        Case 2      '负整数
            If Left(strText, 1) <> "-" And strText <> "0" Then
                CheckNumeric = "应为负数或者零！"
                Exit Function
            End If
            str整数部份 = Mid(strText, 2)
            
        Case 3      '正负整数
            If Left(strText, 1) = "-" Then str整数部份 = Mid(strText, 2)
        End Select
    Else
        '小数
        Select Case bytMode
        Case 1      '正小数
            If Len(strText) > lngLength + 1 Then
                CheckNumeric = "长度超过了" & lngLength & "位！"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '有小数部份
                str整数部份 = Left(strText, InStr(strText, ".") - 1)
                str小数部份 = Mid(strText, InStr(strText, ".") + 1)
            Else
                str整数部份 = strText
            End If
            
        Case 2      '负小数
            If Len(strText) > lngLength + 2 Then
                CheckNumeric = "长度超过了" & lngLength & "位！"
                Exit Function
            End If
            
            If Left(strText, 1) <> "-" Then
                CheckNumeric = "不是负数！"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '有小数部份
                str整数部份 = Mid(strText, 2, InStr(strText, ".") - 2)
                str小数部份 = Mid(strText, InStr(strText, ".") + 1)
            Else
                str整数部份 = Mid(strText, 2)
            End If
            
        Case 3      '正负小数
            If Left(strText, 1) = "-" Then
                If Len(strText) > lngLength + 2 Then
                    CheckNumeric = "长度超过了" & lngLength & "位！"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '有小数部份
                    str整数部份 = Mid(strText, 2, InStr(strText, ".") - 2)
                    str小数部份 = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str整数部份 = Mid(strText, 2)
                End If
            Else
                If Len(strText) > lngLength + 1 Then
                    CheckNumeric = "长度超过了" & lngLength & "位！"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '有小数部份
                    str整数部份 = Mid(strText, 1, InStr(strText, ".") - 1)
                    str小数部份 = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str整数部份 = strText
                End If
                
            End If
        End Select
    End If
    
    If Len(str整数部份) > (lngLength - lngDecLength) Then
        If lngDecLength = 0 Then
            CheckNumeric = "长度超过了" & (lngLength - lngDecLength) & "位！"
        Else
            CheckNumeric = "整数部份长度超过了" & (lngLength - lngDecLength) & "位！"
        End If
        Exit Function
    End If
    
    If Len(str小数部份) > lngDecLength Then
        CheckNumeric = "小数部份长度超过了" & lngDecLength & "位！"
        Exit Function
    End If
    
    For lngloop = 1 To Len(str整数部份)
        If Mid(str整数部份, lngloop, 1) < "0" Or Mid(str整数部份, lngloop, 1) > "9" Then
            CheckNumeric = "应为数字型！"
            Exit Function
        End If
    Next
    
    For lngloop = 1 To Len(str小数部份)
        If Mid(str小数部份, lngloop, 1) < "0" Or Mid(str小数部份, lngloop, 1) > "9" Then
            CheckNumeric = "应为数字型！"
            Exit Function
        End If
    Next
    
    
    CheckNumeric = ""
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MovedByDate(ByVal vDate As Date) As Boolean
'功能：判断指定日期之前的是否可能已经执行了数据转出
'参数：vDate=时间点或时间段的开始时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 上次日期 From zlDataMove Where 系统=[1] And 组号=1 And 上次日期 is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '上次日期没有时点,"<"判断与转出过程中一致
        If vDate < rsTmp!上次日期 Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetConnectDevs() As String
'功能：获取本机连接的检验仪器，以;分隔
'赵彤宇
    Dim aPorts As Variant, i As Integer, PortIndex As Integer
    Dim lngDeviceID As Long
    
    GetConnectDevs = ""
    On Error Resume Next
    aPorts = GetAllSettings("ZLSOFT", "公共模块\ZlLISSrv")
    If Not IsEmpty(aPorts) Then
        For i = 0 To UBound(aPorts)
            PortIndex = Val(Mid(aPorts(i, 0), 4)) - 1
            lngDeviceID = Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
            If lngDeviceID > 0 Then
                GetConnectDevs = GetConnectDevs & ";" & lngDeviceID
            End If
        Next
        If Len(GetConnectDevs) > 0 Then GetConnectDevs = Mid(GetConnectDevs, 2)
    End If
End Function

Public Function FindComboItem(objCombox As Object, ByVal lngFind As Long) As Integer
    Dim i As Integer
    
    For i = 0 To objCombox.ListCount - 1
        If objCombox.ItemData(i) = lngFind Then Exit For
    Next
    If i > objCombox.ListCount - 1 Then i = -1
    
    FindComboItem = i
End Function

'---以下为直接申请添加

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    On Error GoTo errH
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    On Error GoTo errH
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'功能：去掉TextBox的默认右键菜单
    If Msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
    End If
End Function

Public Function CheckOneDuty(ByVal str医嘱 As String, ByVal str职务 As String, ByVal str医生 As String, ByVal bln医保 As Boolean) As String
'功能：检查当前指定药品处方职务是否符合
'参数：str医嘱=药品医嘱提示内容
'      str职务=药品处方职务
'      str医生=开嘱医生
'      bln医保=是否公费或医保病人
'      grsDuty=记录医生职务缓存
'返回：职务不满足的提示信息，如果满足则返回空。
    Const STR_职务 = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim int职务A As Integer, int职务B As Integer
    
    If Len(str职务) <> 2 Or str医生 = "" Then Exit Function
    
    '取药品处方职务
    If bln医保 Then
        int职务B = Val(Right(str职务, 1))
    Else
        int职务B = Val(Left(str职务, 1))
    End If
    If int职务B = 0 Then Exit Function '不限制
    
    '取医生职务
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "医生", adVarChar, 50
        grsDuty.Fields.Append "职务", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.filter = "医生='" & str医生 & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select 姓名,Nvl(聘任技术职务,0) as 职务 From 人员表 Where 姓名='" & str医生 & "'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork")
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!医生 = rsTmp!姓名
            grsDuty!职务 = rsTmp!职务
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        int职务A = grsDuty!职务
    End If
        
    '检查职务要求
    If int职务A = 0 Then
        '医生未设置职务的情况
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIf(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """未设置职务。"
    ElseIf int职务B < int职务A Then
        '数值越小职务越高
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIf(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """的职务为""" & Split(STR_职务, ",")(int职务A - 1) & """。"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetSysParVal(Optional ByVal int参数号 As Integer = -9999, Optional ByVal strDefault As String) As String
'功能：获取指定系统参数的值
'参数：int参数号=为-9999时，初始化参数集
'      strDefault=如果没有值或为空的缺省值
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If Not grsSysPars Is Nothing Then
        If grsSysPars.State = 1 Then blnDo = False
    End If
    
    GetSysParVal = zlDatabase.GetPara(int参数号, glngSys)
    
'    If blnDo Then
'        strSQL = "Select 参数号,参数名,参数值 From 系统参数表"
'        Set grsSysPars = New ADODB.Recordset
'        Call zldatabase.OpenRecordset(grsSysPars, strSQL, "GetSysParVal")
'    End If
'
'    If int参数号 <> -9999 Then
'        grsSysPars.Filter = "参数号=" & int参数号
'        If Not grsSysPars.EOF Then
'            GetSysParVal = Nvl(grsSysPars!参数值, strDefault)
'        Else
'            GetSysParVal = strDefault
'        End If
'    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValidPH(ByVal SerialNO As String, ByRef ErrMessage As String) As String
'功能：判断标本批号是否合法，并返回格式化的批号或错误信息
    Dim i As Integer, intPh As Integer, intNumber As Integer
    Dim strTmp As String, blnError As Boolean
    
    ErrMessage = "": blnError = False
    For i = 1 To Len(SerialNO)
        If InStr("0123456789", Mid(SerialNO, i, 1)) > 0 Then
            strTmp = strTmp & Mid(SerialNO, i, 1)
        ElseIf Mid(SerialNO, i, 1) = "-" Then
            If intPh = 0 Then
                '是批号
                If Val(strTmp) > 9999 Then
                    blnError = True
                Else
                    intPh = Val(strTmp)
                    strTmp = ""
                End If
            Else
                blnError = True
            End If
        Else
            blnError = True
        End If
        
        If blnError Then Exit For
    Next
    
    If Not blnError Then
        If intPh = 0 Then
            blnError = True
        Else
            If Val(strTmp) = 0 Or Val(strTmp) > gintNumberPH Then
                blnError = True
            Else
                intNumber = Val(strTmp)
            End If
        End If
    End If
    If blnError Then
        ErrMessage = "标本批次号格式为：XXX-XXXX。" & vbCrLf & _
            "批号范围1～9999、批内编号范围1～" & gintNumberPH
        ValidPH = ""
    Else
        ValidPH = Format(intPh, "0000") & "-" & Format(intNumber, "0000")
    End If
End Function

Public Function TransSampleNO(ByVal varSampleNO As Variant) As String
    On Error Resume Next
    
    If InStr(varSampleNO, "-") = 0 Then
        TransSampleNO = varSampleNO
    Else
        TransSampleNO = (Split(varSampleNO, "-")(0) - 1) * 10000 + Split(varSampleNO, "-")(1)
    End If
End Function

Public Function TransSampleNO_PH(ByVal varSampleNO As Variant, ByVal lngDeviceID As Long) As String
    On Error Resume Next
    Dim lngTmp As Long
    
    If lngDeviceID <> -1 Or Not gblnManualPH Or InStr(varSampleNO, "-") > 0 Then
        TransSampleNO_PH = CStr(varSampleNO)
    Else
        lngTmp = Val(varSampleNO)
        TransSampleNO_PH = Format(((lngTmp \ 10000) + 1), "0000") & "-" & Format((lngTmp Mod 10000), "0000")
    End If
End Function

Public Function GetSampleNOStr(StartNO As String, EndNO As String, Optional ByRef strErr As String) As String
    '功能   返回开始和结束中间的标本字串
    Dim strNO As String
    Dim intRow As Integer
    Dim strTemp As String

    On Error GoTo errH

    If StartNO = "" And EndNO = "" Then Exit Function

    If StartNO = "" Then
        StartNO = EndNO
    End If

    If EndNO = "" Then
       EndNO = StartNO
    End If

    GetSampleNOStr = StartNO
    strNO = StartNO
    Do Until strNO = EndNO
        intRow = intRow + 1
        If intRow > 1000 Then
            strErr = "输入标本段超过了1000，请重新输入，并检查标本段是否正确！" ' & vbCrLf & _
                    '"正确的标本段如有字母，字母应是一致的如：S1 到 S50  不能写成 S1 到  D50"
            GetSampleNOStr = ""
            Exit Function
        End If
        
        strNO = IncStr(strNO)
        strTemp = strTemp & "," & strNO
        If Len(strTemp) > 3900 Then
            GetSampleNOStr = GetSampleNOStr & strTemp & ";"
            strTemp = ""
        End If
    Loop
    
    GetSampleNOStr = GetSampleNOStr & strTemp
    
    Exit Function
errH:
    strErr = "出错函数(GetSampleNOStr),出错信息:" & Err.Number & " " & Err.Description
End Function

Public Function IncStr(ByVal strVal As String, Optional intUpDown As Integer, Optional ByRef strErr As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'参数：strVal=要加1的字符串
'      intUpDown = 0 加1 =1 减1
    Dim strValuse As String
    Dim intAdd As Integer
    Dim intUp As Integer
    Dim strValue As String
    Dim strValueOne As String
    Dim strHead As String
    Dim i  As Integer
    
    On Error GoTo errH
    
    strVal = UCase(strVal)

    For i = Len(strVal) To 1 Step -1
        strValueOne = Mid(strVal, i, 1)
        If Asc(strValueOne) >= Asc("0") And Asc(strValueOne) <= Asc("9") Then
        Else
            '不是数字
            strHead = Mid$(strVal, 1, i)
            strVal = Mid$(strVal, i + 1)
            Exit For
        End If
    Next
    
    strVal = UCase(strVal)
    
    If intUpDown = 0 Then
        '加1
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = 1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp < 10 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "0" & strValue
                    intUp = 1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp <= Asc("Z") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "A" & strValue
                    intUp = 1
                End If
            End If
        Next
        If intUp = 1 Then
            If IsNumeric(strValueOne) Then
                strValue = "1" & strValue
            Else
                strValue = "A" & strValue
            End If
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    Else
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = -1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp >= 0 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
'                    If intAdd = 0 Then
                        strValue = "9" & strValue
'                    End If
                    intUp = -1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp >= Asc("A") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    If intAdd = 0 Then
                        strValue = "Z" & strValue
                    End If
                    intUp = -1
                End If
            End If
        Next
        If intUp = 1 Then
            strValue = -1
        End If
        If Mid(strValue, 1, 1) = "0" Or Mid(strValue, 1, 1) = "A" Then
            strValue = Mid(strValue, 2)
            If strValue = "" Then strValue = 1
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    End If
    Exit Function
errH:
    strErr = "出错函数(IncStr),出错信息:" & Err.Number & " " & Err.Description
End Function

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim RptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand
    For Each rptRow In rptList.Rows
        If rptRow.Childs.Count > 0 Then rptRow.Expanded = True
    Next
    If rptList.Rows.Count < 1 Then zlReportToVSFlexGrid = False: Exit Function
        
    With vfgList
        .Clear
        .Rows = 1: .FixedRows = 1: .RowHeight(.Rows - 1) = 280
        .Cols = 0
        .MergeCells = flexMergeFree
        
        '标题行复制
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = rptCol.Caption
                .ColData(.Cols - 1) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(.Cols - 1) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(.Cols - 1) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(.Cols - 1) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, .Cols - 1, .FixedRows - 1) = flexAlignCenterCenter
                If rptCol.Width < 20 * IIf(rptList.GroupsOrder.Count = 0, 1, rptList.GroupsOrder.Count) Then
                    .ColWidth(.Cols - 1) = 0
                Else
                    .ColWidth(.Cols - 1) = rptCol.Width * Screen.TwipsPerPixelX
                End If
            End If
        Next
        
        '数据行复制
        Dim intTiers As Integer, rptParent As ReportRow, rptChild As ReportRow
        For Each rptRow In rptList.Rows
            .Rows = .Rows + 1: .RowHeight(.Rows - 1) = 280
            If rptRow.GroupRow Then
                intTiers = 0
                Set rptParent = rptRow
                Do While Not (rptParent.ParentRow Is Nothing)
                    intTiers = intTiers + 1
                    Set rptParent = rptParent.ParentRow
                Loop
                Set rptChild = rptRow.Childs(0)
                Do While rptChild.GroupRow
                    Set rptChild = rptChild.Childs(0)
                Loop
                .MergeRow(.Rows - 1) = True
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "　") & rptList.GroupsOrder(intTiers).Caption & ": "
                    .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & rptChild.Record(rptList.GroupsOrder(intTiers).ItemIndex).Value
                Next
            Else
                For lngCol = 0 To .Cols - 1
                    If rptList.Columns(.ColData(lngCol)).TreeColumn Then
                        intTiers = 0
                        Set rptParent = rptRow
                        Do While Not (rptParent.ParentRow Is Nothing)
                            intTiers = intTiers + 1
                            Set rptParent = rptParent.ParentRow
                        Loop
                        .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "　") & rptRow.Record(.ColData(lngCol)).Value
                    Else
                        .TextMatrix(.Rows - 1, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                    End If
                    .Cell(flexcpAlignment, .Rows - 1, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Public Function FillGrid_UQ(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '   blnMerge：是否合并相同ID的行。如果合并，则其他不同的列值以“;”分隔
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngloop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngloop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngloop) = ""
        Next
        lngRow = 0
    Else
        '预先有一空行
        lngRow = objMsf.Rows - 2
    End If
    
    Do While Not rsData.EOF
        lngCurrRow = FindGridLine(objMsf, CStr(zlCommFun.Nvl(rsData("ID"))))
        If lngCurrRow = -1 Then
            lngRow = lngRow + 1
            lngCurrRow = lngRow
        
            If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
            
            On Error Resume Next
            objMsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
            
            On Error GoTo ErrHand
            
            For lngloop = 0 To objMsf.Cols - 1
                
                If Trim(objMsf.TextMatrix(0, lngloop)) <> "" Then
                    If objMsf.TextMatrix(0, lngloop) = "#" Then
                        objMsf.TextMatrix(lngCurrRow, lngloop) = lngCurrRow
                    Else
                    
                        On Error Resume Next
                        
                        strMask = ""
                        strMask = MaskArray(lngloop)
                                                
                        On Error GoTo ErrHand
                        
                        If strMask <> "" Then
                            strNewValue = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop))), strMask)
                        Else
                            strNewValue = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngloop)))
                        End If
                        objMsf.TextMatrix(lngCurrRow, lngloop) = strNewValue
                    End If
                End If
                
            Next
        End If
        
        rsData.MoveNext
    Loop
    
    FillGrid_UQ = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetLabItems(objParent As Object, Optional ByVal strType As String = "", Optional ByVal strCode As String = "", Optional ByVal lngExeDept As Long, Optional objContainer As Object = Nothing) As String
'选择检验项目(不含微生物项目)，可复选
'strType：检验类型
'lngExeDept：执行科室
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim sglX As Single, sglY As Single
    Dim objMain As Object
    On Error GoTo errH
    GetLabItems = ""
    If objContainer Is Nothing Then
        Set objMain = objParent.Container
    Else
        Set objMain = objContainer
    End If
    
'    strSQL = "Select 0 As 末级, to_Char(Rownum) As ID, '' As 上级id, 编码, 名称 " & _
'        "From 诊疗检验类型 A " & IIf(Len(strType) = 0, "", "Where A.名称 = [1] ") & _
'        "Union All " & _
'        "Select 1 As 末级, to_Char(B.ID) As ID, to_Char(A.ID) As 上级id, B.编码, B.名称 " & _
'        "From (Select Rownum As ID, 名称 From 诊疗检验类型) A, 诊疗项目目录 B " & _
'        "Where B.类别 = 'C' And A.名称 = B.操作类型 " & IIf(Len(strType) = 0, "", "And B.操作类型 = [1]")
    
    If Len(strCode) = 0 Then
        strSQL = "Select 0 As 选择, B.ID, B.编码, B.名称 " & _
            "From 诊疗项目目录 B Where B.类别 = 'C' " & IIf(Len(strType) = 0, "", "And B.操作类型 = [1]")
    Else
        strSQL = "Select Distinct 0 As 选择, B.ID, B.编码, B.名称 " & _
            "From 诊疗项目目录 B,诊疗项目别名 C,检验报告项目 D,检验项目 E " & _
            "Where B.ID=C.诊疗项目ID And B.ID=D.诊疗项目ID " & _
            "And D.报告项目ID=E.诊治项目ID And D.细菌ID Is Null And B.类别 = 'C' And E.项目类别<>2 " & _
            IIf(Len(strType) = 0, "", "And B.操作类型 = [1] ") & _
            "And (B.编码 Like [2] Or C.简码 Like [2] Or (Nvl(B.组合项目,0)=0 And Upper(E.缩写) Like [2]))"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检验项目", strType, UCase(strCode) & "%")
    If rsTmp.EOF Then Exit Function
    Call CalcPosition(sglX, sglY, objParent)
    If rsTmp.RecordCount = 1 Then
        Call frmSelectMuli.ShowSelect(objMain, rsTmp, "编码,1200,0,1;名称,3000,0,1", sglX, sglY, 5000, 3000, strTitle:="检验项目")
        frmSelectMuli.ReturnSelect
        If rsTmp.RecordCount > 0 Then
            GetLabItems = rsTmp("ID")
        End If
        Exit Function
    End If
    
    If frmSelectMuli.ShowSelect(objMain, rsTmp, "编码,1200,0,1;名称,3000,0,1", sglX, sglY, 5000, 3000, strTitle:="检验项目") Then
        Do While Not rsTmp.EOF
            GetLabItems = GetLabItems & "," & rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        If Len(GetLabItems) > 0 Then GetLabItems = Mid(GetLabItems, 2)
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub RenumVsf(objVsf As Object, intNumCol As Integer)
    Dim lngRow As Long
    Dim intLoop As Integer
    Dim dblTmp As Double
    Const intColCount As Integer = 27
    Dim GetColCount As Integer
    
    On Error Resume Next
    
    If objVsf.Cols <= intColCount Then
        GetColCount = 0
    Else
        dblTmp = objVsf.Cols / intColCount
        If InStr(dblTmp, ".") > 0 Then
            GetColCount = Mid(dblTmp, 1, InStr(dblTmp, ".") - 1)
        Else
            GetColCount = dblTmp
        End If
    End If

    For intLoop = 0 To GetColCount
        For lngRow = 1 To objVsf.Rows - 1
            objVsf.TextMatrix(lngRow, intNumCol) = lngRow
            objVsf.Cell(flexcpData, lngRow, intLoop * intColCount, lngRow, intLoop * intColCount) = ""
        Next
    Next
End Sub
Public Function VerifyAuditingRule(lngSampleID As Long, Optional strErrMessage As String, Optional ByVal iLoadProg As Integer = 1) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       审核时检验审核规则
    '参数                       lngSampleID 标本ID; strErrMessage 返回1时的错误提示。iLoadProg :调用程序，1-审核调用 2-批量审核调用
    '返回                       0 正常 1 有结果超出警示值
    '
    '结果标志 3-↑、2-↓、1-正常、4-异常、5-↓↓、6-↑↑
    '
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim int病人id As Integer '
    Dim strTmp As String
    On Error GoTo errH
    
    '处理超出警示值的结果
    strSQL = " select 结果标志 from 检验标本记录 a , 检验普通结果 b " & _
             " Where a.ID = b.检验标本id and a.id = [1] and (b.结果标志 = 5 Or b.结果标志 = 6)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngSampleID)
    If rsTmp.EOF = False Then
        VerifyAuditingRule = 1: strErrMessage = "  结果超过警示值！"
    End If
    '-- 德阳修改，检验结果全部为空，则提示。
    strSQL = "Select Count(B.ID) - Sum(Decode(Trim(b.检验结果), Null, 1, 0)) As 结果" & vbNewLine & _
             "From 检验标本记录 a , 检验普通结果 B Where a.id = b.检验标本ID and  a.id = [1] and a.微生物标本 is null "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngSampleID)
    Do Until rsTmp.EOF
        If Nvl(rsTmp("结果")) <> "" Then
            If Val("" & rsTmp!结果) <= 0 Then
               VerifyAuditingRule = 1: strErrMessage = strErrMessage & "  结果全部为空！"
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    int病人id = zlDatabase.GetPara("历史病人识别", 100, 1208, 0)
    
    'If VerifyAuditingRule <> 1 And strErrMessage = "" Then
        strSQL = "Select Zl_检验审核规则_Check(" & lngSampleID & "," & int病人id & "," & iLoadProg & ") as 审核结果 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName)
        If rsTmp.RecordCount <= 0 Then
            VerifyAuditingRule = 1
            strErrMessage = strErrMessage & "  计算过程调用错误! "
            Exit Function
        End If

        If Mid(rsTmp.Fields(0).Value, 1, 2) = "1|" Then
            strTmp = "1|"
        Else
            strTmp = ""
        End If
        strErrMessage = strErrMessage & "" & Mid(rsTmp.Fields(0).Value, 15)
        strErrMessage = strTmp & strErrMessage
    'End If
    strSQL = "Zl_检验标本记录_审核未通过(" & lngSampleID & ",'" & strErrMessage & "')"
    zlDatabase.ExecuteProcedure strSQL, "审核规则"
    If strErrMessage <> "" Then
        VerifyAuditingRule = 1
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSql = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSql, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

ErrHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function
'################################################################################################################
'## 功能：  替换要素
'## 参数：  BasicName     :要素名称
'## 返回：  要素内容
'################################################################################################################
Public Function ReplaceBasic(BasicName As String, lngPatientID As Long, lngPatientPage As Integer, intPatientType As Integer, lngAdvice As Long) As String
    Dim rsTmp As New ADODB.Recordset
    gstrSql = " Select Zl_Replace_Element_Value('" & BasicName & "'," & lngPatientID & "," & IIf(lngPatientPage = 0, "Null", lngPatientPage) & _
               "," & intPatientType & "," & lngAdvice & ") From dual "
    zlDatabase.OpenRecordset rsTmp, gstrSql, gstrSysName
    ReplaceBasic = Nvl(rsTmp(0))
    
End Function
'################################################################################################################
'## 功能：  产生保存指定的文件到指定表记录BLOB字段的SQL语句
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##         arySql()    :在该数据的基础上扩展增加保存的SQL语句；不指定时，取当前路径产生文件名
'##
'## 返回：  成功返回True，失败返回False
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByRef arySql() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo ErrHand
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        arySql(lngUBound + lngCount + 1) = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
    Next
    Close lngFileNum
    zlBlobSql = True
    Exit Function

ErrHand:
    Close lngFileNum
    zlBlobSql = False
End Function

Public Function GetHexStr(strChr As String) As String
    '功能                   得到汉字的16进制的字串
    '参数:strChr            传入字串
    '返回                   Rtf格式的字串
    Dim lngloop As Long
    Dim strTmp As String
    Dim strHeight As String
    Dim strLow As String
    
    For lngloop = 1 To Len(strChr)
         If Asc(Mid(strChr, lngloop, 1)) < 0 Then
            strTmp = Hex(Asc(Mid(strChr, lngloop, 1)))
            If Len(strTmp) = 3 Then strTmp = "0" & strTmp
            strLow = Mid(strTmp, 1, 2)
            If Len(strLow) = 0 Then
                strLow = "0" & strLow
            End If
            strHeight = Mid(strTmp, 3)
            If Len(strHeight) = 0 Then
                strHeight = "0" & strHeight
            End If
            GetHexStr = GetHexStr & "\'" & strLow & "\'" & strHeight
         Else
            strTmp = Hex(Asc(Mid(strChr, lngloop, 1)))
            If Len(strTmp) = 0 Then
                strTmp = "0" & strTmp
            End If
            GetHexStr = GetHexStr & "\'" & strTmp
         End If
    Next
    If Mid(GetHexStr, 1, 1) = "\" Then
        GetHexStr = Mid(GetHexStr, 2)
    End If
    If Trim(GetHexStr) = "" Then GetHexStr = " "
End Function
Public Function GetSouceElement(RtfTxt As RichTextBox, lngElement As Long) As String
    '功能                   通过要素名称得到要素字串(用替换)
    '参数: lngElement       要素Number
    '返回                   要替换的要素字串
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStart As String
    Dim strEnd As String
    
    strStart = "ES(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    strEnd = "EE(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    
    lngStart = InStr(RtfTxt.Text, strStart)
    lngEnd = InStr(lngStart, RtfTxt.Text, strEnd)
    
    GetSouceElement = Mid(RtfTxt.Text, lngStart, lngEnd - lngStart + 16)
    
End Function

Public Function GetReplaceElement(RtfTxt As RichTextBox, lngElement As Long, strElementReplace As String) As String
    '功能                   生成要替换的要素
    '参数 lngElement        要素Number
    '     strreplace        替换要素字串
    '返回                   替换的整个串
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strReplace As String
    Dim strNewChr As String
    
    strStart = "ES(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    strEnd = "EE(" & Format(lngElement, Replace(Space(8), " ", "0")) & ",0,0)"
    
    lngStart = InStr(RtfTxt.Text, strStart)
    lngEnd = InStr(lngStart, RtfTxt.Text, strEnd)
    GetReplaceElement = Mid(RtfTxt.Text, lngStart, lngEnd - lngStart + 16)
    
    lngStart = InStr(GetReplaceElement, "{")
    lngEnd = InStr(GetReplaceElement, "}")
    strReplace = Mid(GetReplaceElement, lngStart, lngEnd - lngStart + 1)
    
    GetReplaceElement = Replace(GetReplaceElement, strReplace, GetHexStr(strElementReplace))
    
    GetReplaceElement = Replace(GetReplaceElement, "\highlight2", "\highlight0")
    GetReplaceElement = Replace(GetReplaceElement, "\ulwave", "\ulnone")
        
End Function

Public Sub InstrtVerifyResult(RtfTxt As RichTextBox, lngSyllabus As Long, strSyllabusReplace As String)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       把检验结果增加到文件中
    '参数   lngSyllabus         提纲编号
    '       strSyllabusReplace
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStart As String
    Dim strEnd As String
    
    strEnd = "OE(" & Format(lngSyllabus, Replace(Space(8), " ", "0")) & ",0,0)"
    lngStart = InStr(RtfTxt.Text, strEnd)
    If lngStart = 0 Then
        strEnd = "OE(" & Format(lngSyllabus, Replace(Space(8), " ", "0")) & ",1,0)"
        lngStart = InStr(RtfTxt.Text, strEnd)
        If lngStart = 0 Then Exit Sub
    End If
    lngStart = InStr(lngStart, RtfTxt.Text, "\par") + 4
    
    RtfTxt.Text = Mid(RtfTxt.Text, 1, lngStart) & strSyllabusReplace & _
                        Mid(RtfTxt.Text, lngStart)
End Sub
Public Sub AuditingReport(RtfTxt As RichTextBox, lngSampleID As Long, intPatientType As Integer, lngPatientID As Long, intBaby As Integer, lngApplyDept As Long, _
                          lngAdviceID As Long, intRepotrCount As Integer, lngPatientPage As Integer)
    '功能           生成检验报告项目
    '参数           intPatientType              病人来源
    '               lngPatientID                病人ID
    '               intBaby
    '               lngApplyDept
    '               lngAdviceID
    '               intRepotrCount
    '               lngPatientPage
    Dim rsTmp As New ADODB.Recordset
    Dim rsVerify As New ADODB.Recordset         '检验指标
    Dim strZipFile As String                    '解压文件临时路径
    Dim strFilePath As String                   '临时RTF文件路径
    Dim lngNewCaseHistory As Double               '新的电子病历记录ID
    Dim astrSQL() As String                     '数组SQL字串
    Dim lngSQLCount As Integer                  'SQL字串数组长度
    Dim intLoop As Integer                      '临时循环变量
    Dim lngResult                               '显示检验结果ID
    Dim lngNextID As Double                       '得到下一个ID
    Dim strSampleID As String                   '标本ID
    Dim strLine As String                       '一行数据
    Dim strRtfTxt As String                     '从文件读到数据
    Dim strSouce As String                      '源要素字串
    Dim strReplace As String                    '替换的要素字串
    Dim lngUPID As Double                         '上级ID
    Dim lngFileID As Double                       '文件ID
    Dim blBeginTrans As Boolean                 '是否开始事务
    Dim strRtf() As String                      '保存Rtf字串数据
    
    On Error GoTo errH
    
    If intPatientType = 1 Then
        gstrSql = "Select Count(Id) As 主页ID From 病人挂号记录 Where 记录状态 =1 and 记录性质 =1 and  病人ID  = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngPatientID)
        If rsTmp.EOF = False Then
            lngPatientPage = Val(Nvl(rsTmp("主页ID")))
        End If
    Else
        If intPatientType <> 2 Then
            lngPatientPage = 0
        End If
    End If
    
    gstrSql = "Select 病历文件ID From 病人医嘱记录 a , 病历单据应用 b , 病历文件列表 c" & vbNewLine & _
              " Where a.诊疗项目id = b.诊疗项目id And b.病历文件id = c.Id" & vbNewLine & _
              "      And a.相关Id = [1] And b.应用场合 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngAdviceID, IIf(intPatientType = 3, 1, intPatientType))
    If rsTmp.EOF = True Then
        MsgBox "没有找到对应的病历文件ID!", vbInformation, gstrSysName
        Exit Sub
    End If
    lngFileID = rsTmp("病历文件ID")
    
    '处理电子病历内容
    gstrSql = "Select Id,文件ID,nvl(父ID,0) as 父ID,对象序号,对象类型,对象标记,保留对象,对象属性,内容行次,内容文本,是否换行,预制提纲ID" & vbNewLine & _
              "       复用提纲,使用时机,诊治要素ID,替换域,要素名称,要素类型,要素长度,要素小数,要素单位,要素表示,输入形态,要素值域" & vbNewLine & _
              " From 病历文件结构 Where 文件id = [1]" & vbNewLine & _
              "  Start With 父id  Is Null " & _
              "  Connect By Prior Id = 父id "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngFileID)
    
    '第一次循环处理要素
    rsTmp.filter = "对象类型 = 1 and 内容文本 = '检验结果'"
    If rsTmp.EOF = False Then
        lngResult = rsTmp("ID")
        rsTmp.filter = "父ID <> " & lngResult
    Else
        rsTmp.filter = ""
    End If
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有找到对应的电子病历内容", vbInformation, gstrSysName
        Exit Sub
    End If
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    intLoop = 0
    Do While Not rsTmp.EOF
    
        If rsTmp("ID") <> lngResult Then
            intLoop = intLoop + 1
            lngSQLCount = lngSQLCount + 1
            ReDim Preserve astrSQL(1 To lngSQLCount)
            
            lngNextID = GetNextNextId("电子病历内容")
            If rsTmp("父ID") = 0 Then
                lngUPID = lngNextID
            End If
            
            astrSQL(lngSQLCount) = "Zl_电子病历内容_Update(" & lngNextID & "," & lngFileID & ",1,0," & IIf(rsTmp("父ID") = 0, "Null", lngUPID) & "," & intLoop & "," & _
                                    rsTmp("对象类型") & "," & Nvl(rsTmp("对象标记"), "Null") & "," & Nvl(rsTmp("保留对象"), "Null") & ",'" & _
                                    Nvl(rsTmp("对象属性"), "") & "'," & Nvl(rsTmp("内容行次"), "Null") & ","
                                    
            If Nvl(rsTmp("对象类型"), 0) = 4 And Nvl(rsTmp("替换域"), 0) = 1 Then
                astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & _
                                       ReplaceBasic(rsTmp("要素名称"), lngPatientID, lngPatientPage, intPatientType, lngAdviceID) & _
                                       "'," & Nvl(rsTmp("是否换行"), "Null") & ")"
            Else
                astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & rsTmp("内容文本") & "'," & Nvl(rsTmp("是否换行"), "Null") & ")"
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    '处理检验提纲下的内容
    If lngResult > 0 Then
        rsTmp.filter = "id = " & lngResult
        rsTmp.MoveFirst
        intLoop = intLoop + 1
        lngSQLCount = lngSQLCount + 1
        ReDim Preserve astrSQL(1 To lngSQLCount)
        lngNextID = GetNextNextId("电子病历内容")
        If rsTmp("父ID") = 0 Then
            lngUPID = lngNextID
        End If
        '写入检验提纲
        astrSQL(lngSQLCount) = "Zl_电子病历内容_Update(" & lngNextID & "," & lngFileID & ",1,0," & IIf(rsTmp("父ID") = 0, "Null", lngUPID) & "," & intLoop & "," & _
                                Nvl(rsTmp("对象类型"), "Null") & "," & Nvl(rsTmp("对象标记"), "Null") & "," & Nvl(rsTmp("保留对象"), "Null") & ",'" & _
                                Nvl(rsTmp("对象属性"), "") & "'," & Nvl(rsTmp("内容行次"), "Null") & ",'" & Nvl(rsTmp("内容文本")) & "'," & _
                                Nvl(rsTmp("是否换行"), "Null") & ")"
                                
        '写入检验指标
        gstrSql = "Select 0 As 排列序号, '' As  检验项目编码,'   ' || rpad('检验项目',32) as 检验项目," & vbNewLine & _
                    "       rpad('本次结果', 10) As 本次结果," & vbNewLine & _
                    "       lpad('标志', 8) As 标志," & vbNewLine & _
                    "       lpad('单位',10) As 单位, lpad('参考',15) As 参考  From dual" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select /*+ RULE */ 排列序号, 检验项目编码,'   ' || rpad(检验项目,32) as 检验项目," & vbNewLine & _
                    "       rpad(Decode(本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**',Null,' ', 本次结果),10,' ') As 本次结果," & vbNewLine & _
                    "       lpad(decode(标志,null,' ',标志),8,' ') as 标志," & vbNewLine & _
                    "       lpad(decode(单位,null,' ',单位),10,' ' ) as 单位, lpad(decode(参考,Null,' ',参考),15,' ') as 参考　" & vbNewLine & _
                    "From (Select " & vbNewLine & _
                    "        A.检验项目id, A.排列序号, A.检验项目 As 检验项目编码, Decode(A.排列序号, Null, 0, 1) As 固定项目, C.ID," & vbNewLine & _
                    "        C.中文名 || Decode(D.缩写, Null, '', '(' || D.缩写 || ')') As 检验项目, B.原始结果, '' As 上次结果, '' As Cv," & vbNewLine & _
                    "        B.检验结果 As 本次结果, D.计算公式, D.结果类型," & vbNewLine & _
                    "        Decode(B.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
                    "        Nvl(E.仪器id, -1) As 仪器id, Nvl(E.标本类别, 0) As 标本类别, E.核收时间, E.标本序号," & vbNewLine & _
                    "        Decode(E.仪器id, Null," & vbNewLine & _
                    "                To_Char(Trunc(E.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(E.标本序号, 10000), '0000'), E.标本序号) As 标本号显示," & vbNewLine & _
                    "        E.检验备注, F.姓名, F.性别, F.年龄, F.门诊号, F.住院号, F.当前床号, 0 As 主页id, D.结果范围," & vbNewLine & _
                    "        Nvl(G.小数位数, 2) As 小数, D.警戒上限, D.警戒下限, D.单位," & vbNewLine & _
                    "        Trim(Replace(Replace(' ' || Zlgetreference(C.ID, E.标本类型, Decode(E.性别, '男', 1, '女', 2, 0), F.出生日期," & vbNewLine & _
                    "                                                    E.仪器id, E.年龄,E.申请科室id), ' .', '0.'), '～.', '～0.')) As 参考 " & vbNewLine
                    
        gstrSql = gstrSql & "　From (Select A.检验项目id, Min(Decode(E.报告项目id, A.检验项目id, F.编码, 99999)) As 检验项目," & vbNewLine & _
                            "              Max(Decode(E.报告项目id, A.检验项目id, E.排列序号, F.编码)) As 排列序号" & vbNewLine & _
                            "       From 检验普通结果 A, 检验报告项目 E, 诊疗项目目录 F," & vbNewLine & _
                            "            (Select Distinct C.诊疗项目id" & vbNewLine & _
                            "              From 检验项目分布 B, 病人医嘱记录 C" & vbNewLine & _
                            "              Where B.标本id = [1] And B.医嘱id = C.相关id) D" & vbNewLine & _
                            "       Where E.诊疗项目id = D.诊疗项目id And A.检验项目id = E.报告项目id(+) And E.诊疗项目id = F.ID(+) And" & vbNewLine & _
                            "             A.检验标本id = [1]" & vbNewLine & _
                            "       Group By A.检验项目id" & vbNewLine & _
                            "       Order By Min(Decode(E.报告项目id, A.检验项目id, F.编码, 99999))," & vbNewLine & _
                            "                Max(Decode(E.报告项目id, A.检验项目id, E.排列序号, F.编码))) A, 检验普通结果 B, 诊治所见项目 C," & vbNewLine & _
                            "     检验项目 D, 检验标本记录 E, 病人信息 F, 检验仪器项目 G　" & vbNewLine & _
                            "Where B.检验项目id = A.检验项目id(+) And B.检验标本id = [1] And B.检验项目id = C.ID And C.ID = D.诊治项目id And" & vbNewLine & _
                            "      B.检验标本id = E.ID And E.病人id = F.病人id(+) And B.检验项目id = G.项目id(+) And B.记录类型 = [2] And" & vbNewLine & _
                            "      (G.仪器id = E.仪器id + 0 Or G.仪器id Is Null Or E.仪器id Is Null)　" & vbNewLine & _
                            "Order By 检验项目编码, 排列序号) A"

                            
                                    
        Set rsVerify = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngSampleID, intRepotrCount)
        
        Do While Not rsVerify.EOF
            intLoop = intLoop + 1
            lngSQLCount = lngSQLCount + 1
            ReDim Preserve astrSQL(1 To lngSQLCount)
            lngNextID = GetNextNextId("电子病历内容")
            
            strLine = Nvl(rsVerify("检验项目")) & Nvl(rsVerify("本次结果")) & Nvl(rsVerify("单位")) & Nvl(rsVerify("标志")) & Nvl(rsVerify("参考"))
            '写入检验提纲
            astrSQL(lngSQLCount) = "Zl_电子病历内容_Update(" & lngNextID & "," & lngFileID & ",1,0," & lngUPID & "," & intLoop & ",2" & _
                                    ",Null" & "," & "Null,0,Null,'" & strLine & "',1)"
            rsVerify.MoveNext
        Loop
        
        '写入提纲下的内容
        rsTmp.filter = " 父id = " & lngResult
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                intLoop = intLoop + 1
                lngSQLCount = lngSQLCount + 1
                ReDim Preserve astrSQL(1 To lngSQLCount)
                
                astrSQL(lngSQLCount) = "Zl_电子病历内容_Update(" & lngNextID & "," & lngFileID & ",1,0," & lngUPID & "," & intLoop & "," & _
                                Nvl(rsTmp("对象类型"), "Null") & "," & Nvl(rsTmp("对象标记"), "Null") & "," & Nvl(rsTmp("保留对象"), "Null") & ",'" & _
                                Nvl(rsTmp("对象属性"), "") & "'," & Nvl(rsTmp("内容行次"), "Null") & ","
                                
                If Nvl(rsTmp("对象类型"), 0) = 4 And Nvl(rsTmp("替换域"), 0) = 1 Then
                    astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & _
                                           ReplaceBasic(rsTmp("要素名称"), lngPatientID, lngPatientPage, intPatientType, lngAdviceID) & _
                                           "'," & Nvl(rsTmp("是否换行"), "Null") & ")"
                Else
                    astrSQL(lngSQLCount) = astrSQL(lngSQLCount) & "'" & rsTmp("内容文本") & "'," & Nvl(rsTmp("是否换行"), "Null") & ")"
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    '得到病历文件
    strZipFile = zlBlobRead(1, lngFileID)
    If gobjFSO.FileExists(strZipFile) Then
        strFilePath = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strFilePath) = False Then
            MsgBox "没有找到病历文件!", vbInformation, gstrSysName
            Exit Sub
        End If
        Kill strZipFile
    End If
    
    '清空文档
    RtfTxt.Text = ""
    '读入Rtf文件
    Open strFilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, strLine
        RtfTxt.Text = RtfTxt.Text & IIf(RtfTxt.Text <> "", vbCrLf, "") & strLine
    Loop
    Close #1
    
    '查找并替换要素
    rsTmp.filter = "对象类型 = 4 and 替换域 = 1 "
    Do While Not rsTmp.EOF
        strSouce = GetSouceElement(RtfTxt, rsTmp("对象标记"))
        strReplace = GetReplaceElement(RtfTxt, rsTmp("对象标记"), ReplaceBasic(rsTmp("要素名称"), lngPatientID, lngPatientPage, intPatientType, lngAdviceID))
        RtfTxt.Text = Replace(RtfTxt.Text, strSouce, strReplace)
        rsTmp.MoveNext
    Loop
    
    '写入检验结果
    If lngResult > 0 Then
        strReplace = ""
        If rsVerify.RecordCount > 0 Then rsVerify.MoveFirst
        Do While Not rsVerify.EOF
            strLine = Nvl(rsVerify("检验项目")) & Nvl(rsVerify("本次结果")) & Nvl(rsVerify("单位")) & Nvl(rsVerify("标志")) & Nvl(rsVerify("参考"))
            strReplace = strReplace & "\" & GetHexStr(strLine) & "\par "
            rsVerify.MoveNext
        Loop
        
        rsTmp.filter = "对象类型 = 1 and 内容文本 = '检验结果'"
        rsTmp.MoveFirst
'        strRePlace = Mid(strRePlace, 5)
        InstrtVerifyResult RtfTxt, rsTmp("对象标记"), strReplace
    End If
    
    '保存RTF文件
    strRtf = Split(RtfTxt.Text, vbCrLf)
    If UBound(strRtf) < 0 Then Exit Sub
    Open strFilePath For Output As #1
    For intLoop = 0 To UBound(strRtf)
        Print #1, strRtf(intLoop)     ' 将文本数据写入文件。
    Next
    Close #1
    
'    strRtf = Split(Me.RtfTxt.Text, vbCrLf)
'    If UBound(strRtf) < 0 Then Exit Sub
'    Open "c:\10.rtf" For Output As #1
'    For intLoop = 0 To UBound(strRtf)
'        Print #1, strRtf(intLoop)     ' 将文本数据写入文件。
'    Next
'    Close #1
    
    lngNewCaseHistory = GetNextNextId("电子病历记录")
    strZipFile = zlFileZip(strFilePath)
    If gobjFSO.FileExists(strZipFile) Then
        zlBlobSql 5, lngNewCaseHistory, strZipFile, astrSQL
    End If
    Kill strFilePath
    Kill strZipFile
        
    '--电子病历记录
    lngSQLCount = UBound(astrSQL) + 1
    ReDim Preserve astrSQL(1 To lngSQLCount)
    astrSQL(lngSQLCount) = "Zl_电子病历记录_Update(" & lngNewCaseHistory & "," & intPatientType & "," & lngPatientID & "," & lngPatientPage & "," & _
                 intBaby & "," & lngApplyDept & "," & lngFileID & "," & lngAdviceID & ")"
                 
    blBeginTrans = True
    gcnOracle.BeginTrans
    For intLoop = 1 To UBound(astrSQL)
        zlDatabase.ExecuteProcedure Replace(astrSQL(intLoop), "Call", ""), gstrSysName
'        Debug.Print aStrSQL(intLoop)
    Next
    gcnOracle.CommitTrans
    Exit Sub
errH:
    If blBeginTrans = True Then gcnOracle.RollbackTrans: blBeginTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetNextNextId(strTable As String) As Double
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
    '参数：
    '   strTable：表名称
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strtab As String
    
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!
    '31730
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "门诊费用记录" Or strtab = "住院费用记录" Then strtab = "病人费用记录"
    
    strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    Call SQLTest(App.ProductName, "mdlCommon", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    GetNextNextId = rsTmp.Fields(0).Value

End Function

Public Function CheckExesState(lngKey As Long) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能:      检查住院病人出院后是否还有划价单需要进行审核
    '参数       标本ID
    '返回       有划价单未审核 = Fasle 没有则 = True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    CheckExesState = True
    
    '81号系统不生效时不检查
'    If zlDatabase.GetPara(81, 100) <> 1 Then Exit Function
        
    '当前病人是否已出院或预出院
    gstrSql = "select d.no" & vbNewLine & _
            "from (select distinct d.医嘱id" & vbNewLine & _
            "       from 检验标本记录 a, 病人信息 b, 病案主页 c, 检验项目分布 d" & vbNewLine & _
            "       where a.病人id = b.病人id and a.病人id = c.病人id and a.主页id = c.主页id and" & vbNewLine & _
            "             a.id = [1] and a.病人来源 = 2 and (b.出院时间 is not null or c.状态 = 3) and" & vbNewLine & _
            "             a.id = d.标本id) a, 病人医嘱记录 b, 病人医嘱发送 c, 住院费用记录 d" & vbNewLine & _
            "where a.医嘱id in (b.相关id, b.id) and b.id = c.医嘱id and c.记录性质 = d.记录性质 and" & vbNewLine & _
            "      c.no = d.no and d.记录性质 = 2 and d.记录状态 = 0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "检验技师工作站-费用状态检查", lngKey)
    
    CheckExesState = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
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
Public Function Between(X, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function
Public Sub SendSample(WinsockC As Winsock, ByVal strIP As String, ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, _
                    Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0)
    With WinsockC
        .SendData "SendSample," & strIP & "," & lngDeviceID & "," & strSampleDate & "," & strSampleNO & "," & _
                    Replace(strAdviceIDs, ",", ";") & "," & blnUndo & "," & iType
    End With
End Sub

Public Sub GetResultFromFile(WinsockC As Winsock, ByVal strIP As String, ByVal strFile As String, ByVal lngDeviceID As Long, _
            ByVal strSampleNO As String, ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))
    With WinsockC
        .SendData "ResultFromFile," & strIP & "," & strFile & "," & lngDeviceID & "," & strSampleNO & "," & dtStart & "," & _
                  dtEnd
    End With
End Sub

Public Sub Open_LIS_Report(ByVal frmParent As Object, ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal lng病人ID As Long, ByVal lng标本ID As Long, ByVal blnCurrMoved As Boolean, ByVal blnPrint As Boolean)
    '调用带图形的LIS报表
    '生成图形供自定义报表调用
'    mfrmLabMainImage.zlRefresh mlngKey, True
    Dim strChart(0 To 8) As String
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer
    On Error GoTo ErrHandle
    strSQL = "select id from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lng标本ID)
    intLoop = 0
    Do Until rsTmp.EOF
        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
        Debug.Print strChart(intLoop)
        Call LoadImageData(App.path, rsTmp("ID"))
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                        "病人ID=" & lng病人ID, "标本ID=" & lng标本ID, _
                        "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                        "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                        "图形9=" & strChart(8), IIf(blnPrint, 2, 1))
    End If

    '删除图形文件
    For intLoop = 0 To 8
        If strChart(intLoop) <> "" Then
            If Dir(strChart(intLoop)) <> "" Then Kill strChart(intLoop)
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadImageData(ByVal strPath As String, ByVal lngID As Long) As Boolean
        '从数据库读取图形数据，绘制后保存到指定的路径下。
        '入参：
        '   strPath 路径
        '   lngID   检验图像结果的ID
        '--如果有的话, 删除原来的临时图形文件
        Static objImg As Object
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '是否图片格式
        Dim lngFileNum As Long, lngCount As Long, lngBound As Long
        Dim aryChunk() As Byte, strFile As String
        Dim intLayOut As Integer
        Dim objPic As New frmChartPic
        Dim killFile As String
    
        Dim blnFtp As Boolean       'FTP是否可用
        Static strFtpPara As String       '保存FTP参数
        Dim strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String
        Dim strDownOk As String, strFtpPath   As String, strLocalFile As String
        Dim objStream As TextStream
    
        On Error GoTo ErrHandle
    
100     If Dir(strPath & "\" & lngID & ".cht") <> "" Then
102         LoadImageData = True
            Exit Function
        End If
    
        'FTP连接检查，有效则可以按FTP方式取图片
104     blnFtp = False
106     If strFtpPara = "" Then
108         strFtpPara = zlDatabase.GetPara("FTP设置", glngSys, 1208, "")
        End If
110     If UBound(Split(strFtpPara, ";")) >= 3 Then
112        strFtpUser = Split(strFtpPara, ";")(0)
114        strFtpPass = Split(strFtpPara, ";")(1)
116        strFtpIP = Split(strFtpPara, ";")(2)
118        strFtpDir = Split(strFtpPara, ";")(3)
120        If TestFTP(strFtpUser, strFtpPass, strFtpIP, strFtpDir) = "" Then
122             blnFtp = True
           End If
        End If
    
124     mlngImageID = lngID
    
126     lngCount = 0
128     strFile = ""
   
130     strSQL = "select 标本id,图像类型,图像位置 from 检验图像结果 where id = [1] "
132     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngID)
    
134     If rsTmp.EOF = True Then
            Exit Function
        End If
    
136     If objImg Is Nothing Then Set objImg = CreateObject("zlLisDev.clsDrawGraph")
    
138     Do Until rsTmp.EOF
140         strImageType = Trim("" & rsTmp("图像类型"))
142         strFtpPath = Trim("" & rsTmp!图像位置)
144         If InStr(strFtpPath, ";") <= 0 Or Not blnFtp Then
                '- 图像存在数据库中，按原来的方式处理
146             gstrSql = "select Zl_FUN_Get检验图像([1],[2],[3]) from dual "
148             Set rsImage = zlDatabase.OpenSQLRecord(gstrSql, "LoadImgData", CLng(rsTmp("标本id")), CStr(Nvl(rsTmp("图像类型"))), CInt("0"))
150             strTmp = Nvl(rsImage(0))
152             strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
            
154             If strImageData <> "" Then
156                 intLoop = 0
                
158                 If Val(Mid(strImageData, 1, 3)) >= 100 And Val(Mid(strImageData, 1, 3)) <= 227 And Mid(strImageData, 4, 1) = ";" Then
                
160                     blnPic = True
162                     If Mid(strImageData, 1, 3) >= 100 And Mid(strImageData, 1, 3) <= 107 Then
164                         strFile = App.path & "\zlLisPic" & lngID & ".bmp"
166                     ElseIf Mid(strImageData, 1, 3) >= 110 And Mid(strImageData, 1, 3) <= 117 Then
168                         strFile = App.path & "\zlLisPic" & lngID & ".jpg"
170                     ElseIf Mid(strImageData, 1, 3) >= 120 And Mid(strImageData, 1, 3) <= 127 Then
172                         strFile = App.path & "\zlLisPic" & lngID & ".gif"
174                     ElseIf Mid(strImageData, 1, 3) >= 200 And Mid(strImageData, 1, 3) <= 227 Then
176                         If gobjFSO.FolderExists(App.path & "\ZLLIS_ZIP") = False Then
178                             gobjFSO.CreateFolder App.path & "\ZLLIS_ZIP"
                            End If
180                         If gobjFSO.FolderExists(App.path & "\ZLLIS_ZIP\" & lngID) = False Then
182                             gobjFSO.CreateFolder App.path & "\ZLLIS_ZIP\" & lngID
                            End If
184                         strFile = App.path & "\ZLLIS_ZIP\" & lngID & "\ZLISPIC.ZIP"
                        End If
                    
                    
186                     intLayOut = Val(Mid(strImageData, 1, 3))
188                     strImageData = Mid(strImageData, 5)
190                     lngFileNum = FreeFile
192                     lngCount = 0
    
194                     If Dir(strFile) <> "" Then Kill strFile
196                     Open strFile For Binary As lngFileNum
198                     ReDim aryChunk(Len(strImageData) / 2 - 1) As Byte
200                     For lngBound = LBound(aryChunk) To UBound(aryChunk)
202                         aryChunk(lngBound) = CByte("&H" & Mid(strImageData, lngBound * 2 + 1, 2))
                        Next
                    
204                     Put lngFileNum, , aryChunk()
                    
                    End If
                    '-------保存为图片文件
206                 Do While strTmp <> ""
208                     intLoop = intLoop + 1
210                     gstrSql = "select Zl_FUN_Get检验图像([1],[2],[3]) from dual "
212                     Set rsImage = zlDatabase.OpenSQLRecord(gstrSql, "LoadImgData", CLng(rsTmp("标本id")), CStr(Nvl(rsTmp("图像类型"))), intLoop)
                    
214                     strTmp = Nvl(rsImage(0))
    
216                     If blnPic Then
                            '
218                         If strTmp <> "" Then
220                             ReDim aryChunk(Len(strTmp) / 2 - 1) As Byte
222                             For lngBound = LBound(aryChunk) To UBound(aryChunk)
224                                 aryChunk(lngBound) = CByte("&H" & Mid(strTmp, lngBound * 2 + 1, 2))
                                Next
                            
226                             Put lngFileNum, , aryChunk()
                            End If
                        Else
                            '图形数据
228                         strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                        End If
                    Loop
                
230                 If blnPic Then
232                     strImageData = intLayOut & ";" & strFile
234                     Close lngFileNum
                    End If
                End If
            Else
                '图像存在FTP中，从FTP中取数据
                '图像位置的数据格式为：图像格式;FTP文件路径
            
236             intLayOut = Val(Split(strFtpPath, ";")(0))
238             strFtpPath = Trim(Split(strFtpPath, ";")(1))
240             strImageData = ""
242             If intLayOut >= 100 And intLayOut <= 227 Then
                    ' 图片文件，直接下载到本地
244                 strLocalFile = strPath & "\" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
246                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
248                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
250                 If strDownOk = "" Then
252                     strImageData = intLayOut & ";" & strLocalFile
                    End If
                Else
                    ' 图形数据，需要从下载的文本文件中读取数据
254                 strLocalFile = strPath & "\" & lngID & "_" & strImageType & ".txt"
256                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
258                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
260                 If strDownOk = "" Then
262                     Set objStream = gobjFSO.OpenTextFile(strLocalFile, ForReading)
264                     Do Until objStream.AtEndOfLine
266                         strImageData = strImageData & objStream.ReadLine
                        Loop
268                     objStream.Close
270                     Set objStream = Nothing
272                     strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
274                     strImageData = intLayOut & ";" & strImageData
                    End If
276                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
                End If
            End If
        
278         If Len(strImageData) <> 0 Then
280             If Not objImg Is Nothing Then
282                 LoadImageData = objImg.DrawImg(strImageType, strImageData, strPath & "\" & lngID & ".cht")
                End If
            End If
        
284         strTmp = "": strImageData = ""
286         rsTmp.MoveNext
        Loop
        Exit Function
ErrHandle:
        WriteLog "mdlLisWork", "LoadImagedata", CStr(Erl()) & "行，" & Err.Description
288     If ErrCenter() = 1 Then
290         Resume
        End If
End Function

Public Function ReadVerifyData(lngID As Long, intRule As Integer) As String
    ''''''''''''''''''''''''''''''''''''''''''''''
    '功能       生成电子签名所用的字串
    '参数       lngID=标本ID
    '           intRule=生成字串规则
    '返回       生成好的字串
    '''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim strData As String
    Dim strBase As String
    
    On Error GoTo errH
    If intRule = 1 Then
        '得到基本信息
        gstrSql = "Select ID, 医嘱id, 标本序号, 采样人, 采样时间, 标本类型, 核收人, 核收时间, 检验人, 检验时间, 审核人, 审核时间, 申请类型," & vbNewLine & _
                "       仪器id, 样本条码, 报告结果, 备注, 申请时间, 标本形态, 执行科室id," & vbNewLine & _
                "       微生物标本, NO, 标本类别, 检验备注, 申请人, 申请科室id, 病人来源, 病人id, 婴儿, 姓名, 性别, 年龄数字, 年龄单位," & vbNewLine & _
                "       紧急, 挂号单, 门诊号, 住院号, 出生日期, 主页id, 检验项目, 操作类型," & vbNewLine & _
                "       接收人, 接收时间, 标识号, 床号, 病人科室,  杯号" & vbNewLine & _
                "From 检验标本记录 A" & vbNewLine & _
                "Where A.ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "读取电子签名", lngID)
        If rsTmp.EOF = True Then Exit Function
        For intLoop = 0 To rsTmp.Fields.Count - 1
            strBase = strBase & "," & rsTmp(intLoop)
        Next
        strBase = Mid(strBase, 2)
        '得到结果
        gstrSql = "Select 检验标本id, 检验项目id, 检验结果, 结果标志, 结果参考, 修改者, 修改时间, 记录类型, 原始结果, 原始记录时间, 记录者," & vbNewLine & _
                "       是否检验, 修改原因, 细菌id, 仪器id, 培养描述, 诊疗项目id, 排列序号, Od, Cutoff, Sco, 酶标板id, 弃用结果" & vbNewLine & _
                "From 检验普通结果 " & vbNewLine & _
                "where 检验标本ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "读取电子签名", lngID)
        Do Until rsTmp.EOF
            strData = strData & "|"
            For intLoop = 0 To rsTmp.Fields.Count - 1
                strData = strData & "," & rsTmp(intLoop)
            Next
            rsTmp.MoveNext
        Loop
        strData = Mid(strData, 3)
        ReadVerifyData = strBase & ";" & strData
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Signature(lngID As Long, Optional strAuditingMan As String, Optional strType As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           对LIS报告单签名
    '参数           lngID=检验标本ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSource As String                                 '取得电子病历签名字串
    Dim lng证书ID As Long                                   '证书ID
    Dim strSign As String                                   '签名后生成的字串
    Dim strTimeStamp As String                              '时间戳
    Dim strTimeStampCode As String                          '时间戳信息
    Dim rsTmp As New ADODB.Recordset
    Dim intSaveInfoSign As Integer                          '核收登记保存时签名 1=保存
    Dim intSaveReprotSign As Integer                        '报告单保存时签名 0=保存
    Dim strSQL As String
    
    
    '检查当前科室是否使用签名
    strSQL = "select 执行科室ID from 检验标本记录 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检验签名", lngID)
    
    strSQL = "select Zl_Fun_Getsignpar(6," & rsTmp("执行科室ID") & ") as tag from dual "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检验签名")
    
    If rsTmp("tag") = 0 Then
        '没有启用的科室直接处理为签名成功
        Signature = True
        Exit Function
    End If
    
    
    intSaveInfoSign = zlDatabase.GetPara("核收登记保存时签名", 100, 1208, 1)
    intSaveReprotSign = zlDatabase.GetPara("报告单保存时签名", 100, 1208, 1)
    
    If strType = "核收" And intSaveInfoSign = 0 Then
        Signature = True
        Exit Function
    End If
    
    If strType = "报告" And intSaveReprotSign = 0 Then
        Signature = True
        Exit Function
    End If
    
    On Error GoTo errH
    '电子签名
    If Not gobjESign Is Nothing Then
        If Not gobjESign.CheckCertificate(IIf(strAuditingMan <> "", strAuditingMan, gstrDBUser)) Then Exit Function
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then
            strSource = ReadVerifyData(lngID, 1)
            If strSource = "" Then
                MsgBox "不能读取要签名的检验报告单。", vbInformation, gstrSysName
                Exit Function
            End If
            strSign = gobjESign.Signature(strSource, IIf(strAuditingMan <> "", strAuditingMan, gstrDBUser), lng证书ID, strTimeStamp, , strTimeStampCode)
            If strSign = "" Then Exit Function
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            strTimeStampCode = IIf(strTimeStampCode = "", "NULL", "'" & strTimeStampCode & "'")
            gstrSql = "Select A.姓名 From 人员表 A, 上机人员表 B Where A.Id = B.人员id and b.用户名 = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "检验签名", IIf(strAuditingMan <> "", strAuditingMan, gstrDBUser))
            If rsTmp.EOF = False Then
                gstrSql = "zl_检验签名记录_Insert(" & lngID & ",1,'" & Replace(strSign, "'", "''") & _
                             "'," & lng证书ID & "," & strTimeStampCode & "," & strTimeStamp & ",'" & rsTmp("姓名") & "')"
                zlDatabase.ExecuteProcedure gstrSql, "检验签名"
            End If
        End If
    End If
    Signature = True
    Exit Function
errH:
    Err.Raise Err.Number, "检验签名"
End Function

Public Function VerifySignature(lngID As Long) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           验证签名
    '参数           lngID = 检验标本ID
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    VerifySignature = gobjESign.VerifySignature(ReadVerifyData(lngID, 1), lngID, 4)
        
    
End Function
Public Function GetAdviceMoney(ByVal str组ID As String, ByVal str医嘱ID As String, ByVal str发送号 As String, _
    str类别 As String, str类别名 As String, Optional ByVal bln单独执行 As Boolean, Optional ByVal strItemType As String) As Currency
'功能：根据指定的医嘱ID串，获取医嘱对应未审核的记帐费用合计
'参数：str组ID,str医嘱ID,str发送号="ID1,ID2,..."
'      bln单独执行=检验项目单独执行，这时只有一个医嘱ID
'      strItemType=是否使用诊疗类别进行限制，主要用于区分采集
'返回：str类别,str类别名=用于报警提示
'说明：当系统参数为执行后审核费用时才返回。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, curMoney As Currency
    Dim strSQLbak As String
    Dim intPatientType As Integer                                   '病人来源
    
    str类别 = "": str类别名 = ""
    
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */ 病人来源 From 病人医嘱记录 Where ID In (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str组ID)
    If rsTmp.EOF = True Then Exit Function
    intPatientType = rsTmp("病人来源")
    
    If bln单独执行 Then
        strSQL = _
            " Select B.编码,B.名称,Sum(A.实收金额) as 金额" & _
            " From 住院费用记录 A,收费项目类别 B" & _
            " Where A.医嘱序号 + 0 = [2] And (mod(A.记录性质,10), A.NO) In" & _
            "      (Select 记录性质, NO From 病人医嘱附费 Where 医嘱id = [2] And 发送号 + 0 = [3]" & _
            "       Union All" & _
            "       Select 记录性质, NO From 病人医嘱发送 Where 医嘱id = [2] And 发送号 + 0 = [3])" & _
            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码" & _
            " Group by B.编码,B.名称"
            
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
        End If
    Else
'        strSQL = _
'            " Select B.编码,B.名称,Sum(A.实收金额) as 金额 From 住院费用记录 A,收费项目类别 B,病人医嘱记录 C" & _
'            " Where A.医嘱序号 + 0 In" & _
'            "      (Select ID From 病人医嘱记录" & _
'            "       Where ID In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B)" & _
'            "       Union All" & _
'            "       Select ID From 病人医嘱记录" & _
'            "       Where 相关id In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B))" & _
'            "  And (mod(A.记录性质,10) , A.NO) In" & _
'            "      (Select 记录性质, NO From 病人医嘱附费" & _
'            "       Where 医嘱id In" & _
'                "      (Select ID From 病人医嘱记录" & _
'                "       Where ID In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B)" & _
'                "       Union All" & _
'                "       Select ID From 病人医嘱记录" & _
'                "       Where 相关id In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B))" & _
'            "         And 发送号 + 0 In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) B)" & _
'            "       Union All" & _
'            "       Select 记录性质, NO From 病人医嘱发送" & _
'            "       Where 医嘱id In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B)" & _
'            "         And 发送号 + 0 In (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) B))" & _
'            "  And A.记帐费用 = 1 And A.记录状态 = 0 And A.收费类别=B.编码 and a.医嘱序号 = c.id  "
        strSQL = "Select *" & vbNewLine & _
                "From (With T1 As (Select /*+cardinality(b,10)*/ b.Column_Value As ID" & vbNewLine & _
                "                  From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B" & vbNewLine & _
                "                  Union All" & vbNewLine & _
                "                  Select /*+cardinality(b,10)*/ a.Id" & vbNewLine & _
                "                  From 病人医嘱记录 A, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B" & vbNewLine & _
                "                  Where a.相关id = b.Column_Value)" & vbNewLine & _
                "       Select b.编码, b.名称, Sum(a.实收金额) As 金额" & vbNewLine & _
                "       From 住院费用记录 A, 收费项目类别 B, 病人医嘱记录 C, T1" & vbNewLine & _
                "       Where a.医嘱序号 = T1.Id And (Mod(a.记录性质, 10), a.No) In" & vbNewLine & _
                "             (Select /*+cardinality(b, 10)*/ a.记录性质, a.No" & vbNewLine & _
                "              From 病人医嘱附费 A, Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) B, T1" & vbNewLine & _
                "              Where a.医嘱id = T1.Id And a.发送号 = b.Column_Value" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select /*+cardinality(a, 10) cardinality(b, 10)*/ a.记录性质, a.No" & vbNewLine & _
                "              From 病人医嘱发送 A, Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B," & vbNewLine & _
                "                   Table(Cast(f_Num2list([3]) As Zltools.t_Numlist)) C" & vbNewLine & _
                "              Where a.医嘱id = b.Column_Value And a.发送号 = c.Column_Value) And a.记帐费用 = 1 And a.记录状态 = 0 And" & vbNewLine & _
                "             a.收费类别 = b.编码 And a.医嘱序号 = c.Id "
        strSQL = strSQL & _
            "" & IIf(strItemType <> "", " And c.诊疗类别 = [5] ", "") & _
            " Group by B.编码, B.名称)"
        If intPatientType <> 2 Then
            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
        End If
    End If
'    strSQLbak = strSQL
'    strSQLbak = Replace$(strSQLbak, "住院费用记录", "门诊费用记录")
'    strSQL = strSQL & " union all " & strSQLbak
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str组ID, str医嘱ID, str发送号, glngSys, strItemType)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!金额, 0)
        str类别 = str类别 & rsTmp!编码
        str类别名 = str类别名 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    str类别名 = Mid(str类别名, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str姓名 As String, ByVal cur剩余款额 As Currency, _
    ByVal cur当日金额 As Currency, ByVal cur记帐金额 As Currency, ByVal cur担保金额 As Currency, _
    ByVal str收费类别 As String, ByVal str类别名称 As String, str已报类别 As String, _
    intWarn As Integer, Optional ByVal bln划价 As Boolean) As Integer
'功能:对病人记帐进行报警提示
'参数:rsWarn=包含报警参数设置的记录集(该病人病区,并区分好了医保)
'     str收费类别=当前要检查的类别,用于分类报警
'     str类别名称=类别名称,用于提示
'     bln划价=生成划价费用时的报警，类似具有强制记帐权限时的处理
'     intWarn=是否显示询问性的提示,-1=要显示,0=缺省为否,1-缺省为是
'返回:str已报类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
'     intWarn=本次询问性提示中的选择结果,0=为否,1-为是
'     0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
    Dim bln已报警 As Boolean, byt标志 As Byte
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str担保 As String, i As Long
    
    BillingWarn = 0
    If mintWarn = 0 Then mintWarn = intWarn
    
    '报警参数检查:NULL是没有设置,0是设置了的
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str收费类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str收费类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str收费类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    '示例："-" 或 ",ABC,567,DEF"
    '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
    bln已报警 = InStr(str已报类别, str收费类别) > 0 Or str已报类别 Like "-*"
    
    If bln已报警 Then '当intWarn = -1时,也可强行再报警
        If byt标志 = 2 Then
            If str已报类别 Like "-*" Then
                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
            Else
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str收费类别) > 0 Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        'Exit For '取消说明见住院记帐模块
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
    str担保 = IIf(cur担保金额 = 0, "", "(含担保额:" & Format(cur担保金额, "0.00") & ")")
    cur剩余款额 = cur剩余款额 + cur担保金额 - cur记帐金额
    cur当日金额 = cur当日金额 + cur记帐金额
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then mintWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then mintWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If mintWarn = 0 Then
                                BillingWarn = 2
                            ElseIf mintWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & " 低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur剩余款额 < 0 Then
                        byt方式 = 2
                        If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                If vMsg = vbIgnore Then mintWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str类别名称 & IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
                                If vMsg = vbIgnore Then mintWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf cur剩余款额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then mintWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then mintWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If mintWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf mintWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If mintWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                                If vMsg = vbIgnore Then mintWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur剩余款额 < 0 Then
                            byt方式 = 2
                            If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                                If mintWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                    If vMsg = vbIgnore Then mintWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If mintWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str类别名称 & IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & "已经耗尽。", frmParent, True)
                                    If vMsg = vbIgnore Then mintWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then mintWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then mintWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If mintWarn = 0 Then
                                BillingWarn = 2
                            ElseIf mintWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";强制记帐;") > 0 Or bln划价) Then
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If mintWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln划价, "", "强制记帐") & "提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日金额, gstrDec) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", frmParent, True)
                            If vMsg = vbIgnore Then mintWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志3
            End If
        End If
    End If
End Function
Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng病人ID As Long, _
    ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal cur金额 As Currency, ByVal str类别 As String, ByVal str类别名 As String) As Boolean
'功能：当执行完成有自动审核的费用时，对病人费用进行记帐报警。
'参数：str类别="CDE..."，报警金额涉及到的收费类别
'      str类别名="检查,检验,..."，对应的类别名用于提示
    Dim rspati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur当日 As Currency
    
    On Error GoTo errH
    
    If lng主页ID <> 0 Then
        '住院病人报警
        strSQL = _
            " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
            " Union ALL" & _
            " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
        strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
        
        strSQL = "Select A.姓名,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人,C.剩余款," & _
            " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
            " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
            " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
            " And A.病人ID=[1] And B.主页ID=[2]"
        Set rspati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID, lng主页ID)
        zlDatabase.Currentdate
    Else
        '其他按门诊报警
        strSQL = "Select 病人ID,预交余额,费用余额 From 病人余额 Where 性质=1 And 病人ID=[1]"
        strSQL = "Select A.姓名,zl_PatiWarnScheme(A.病人ID) as 适用病人,A.担保额," & _
            " Nvl(B.预交余额,0)-Nvl(B.费用余额,0)+Nvl(E.帐户余额,0) as 剩余款" & _
            " From 病人信息 A,(" & strSQL & ") B,医保病人关联表 D,医保病人档案 E" & _
            " Where A.病人ID=B.病人ID(+) And A.病人id = D.病人id(+) And A.险类=D.险类(+)" & _
            " And D.险类=E.险类(+) And D.中心=E.中心(+) And D.医保号=E.医保号(+) And D.标志(+)=1" & _
            " And A.病人ID=[1]"
        Set rspati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病人ID)
    End If
    
    intWarn = -1 '记帐报警时缺省要提示
    '执行报警:门诊病人病区ID=0
    strSQL = "Select Nvl(报警方法,1) as 报警方法,报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线 Where Nvl(病区ID,0)=[1] And 适用病人=[2]"
    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng病区ID, CStr(Nvl(rspati!适用病人)))
    If Not rsWarn.EOF Then
        If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(lng病人ID)
        str类别名 = Mid(str类别名, 2)
        For i = 1 To Len(str类别)
            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rspati!姓名), Nvl(rspati!剩余款, 0), cur当日, cur金额, Nvl(rspati!担保额, 0), Mid(str类别, i, 1), Split(str类别名, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Chk划价费用(Objfrm As Object, str医嘱ID组 As String, lng标本ID As Long, Optional strItemType As String) As Boolean
    '功能 检验划价单记帐时是报费禁止还是提醒
    '参数 str医嘱ID组=如有医嘱时直接传入医嘱组用","分隔
    '     lng标本ID=按标本ID查出需要的医嘱ID
    '     是否权制诊疗项目的类别
    '     上面两个参数只有一个起做用
    Dim curMoney As Currency, str类别 As String, str类别名 As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strIDs As String
    Dim int病人来源 As Integer
    Dim int主页ID As Integer
    Dim lng病区ID As Long
    Dim str医嘱ID As String
    Dim str发送号 As String
    Dim lng病人ID As Long
    
    On Error GoTo errH
    
    If lng标本ID <> 0 Then
        strSQL = "Select Distinct Decode(B.医嘱id, Null, A.ID, B.医嘱id) As 医嘱id" & vbNewLine & _
                "From 检验标本记录 A, 检验项目分布 B" & vbNewLine & _
                "Where A.ID = B.标本id And A.ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Chk划价费用", lng标本ID)
        Do While rsTmp.EOF
            strIDs = strIDs & "," & rsTmp("医嘱ID")
            rsTmp.MoveNext
        Loop
        strIDs = Mid(strIDs, 2)
    Else
        strIDs = str医嘱ID组
    End If
    
    strIDs = Replace(Replace(strIDs, ";", ","), "|", ",")
    
    strSQL = "Select /*+ rule */ a.病人ID,A.主页id, A.病人来源, B.医嘱id, B.发送号, C.当前病区id" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C" & vbNewLine & _
            "Where A.ID = B.医嘱id And A.病人id = C.病人id And" & vbNewLine & _
            "      A.ID In (Select * From Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist)))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Chk划价费用", strIDs)
    If rsTmp.EOF = True Then Exit Function
    int病人来源 = Nvl(rsTmp("病人来源"))
    int主页ID = Nvl(rsTmp("主页ID"), 0)
    lng病区ID = Nvl(rsTmp("当前病区id"), 0)
    lng病人ID = Nvl(rsTmp("病人ID"))
    Do While Not rsTmp.EOF
        str医嘱ID = str医嘱ID & "," & rsTmp("医嘱ID")
        str发送号 = str发送号 & "," & rsTmp("发送号")
        rsTmp.MoveNext
    Loop
    curMoney = GetAdviceMoney(str医嘱ID, str医嘱ID, str发送号, str类别, str类别名, False, strItemType)
    
    If Not FinishBillingWarn(Objfrm, gstrPrivs, lng病人ID, int主页ID, lng病区ID, curMoney, str类别, str类别名) Then Exit Function
    
    Chk划价费用 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function
Public Function GetUser科室IDs(Optional ByVal bln病区 As Boolean) As String
'功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
'参数：是否取所属病区下的科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    '没有强制限制临床,可能医技科室用
    strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1]"
    If bln病区 Then
        strSQL = strSQL & " Union" & _
            " Select Distinct B.科室ID From 部门人员 A,病区科室对应 B" & _
            " Where A.部门ID=B.病区ID And A.人员ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    GetUser科室IDs = Mid(GetUser科室IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'-----以下是 FTP 相关函数
Private Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP类
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.path, 1) <> "\", App.path & "\", App.path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP测试" & strPath) > 0 Then
            TestFTP = "在FTP上不能创建目录！"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "上传文件失败"
            Else
                FtpNet.FuncFtpDisConnect '先断开，再删除，不然删不掉
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP不能连接！"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP测试" & strPath) > 0 Then
                    TestFTP = "在FTP上不能删除目录"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "不能连接FTP！"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Private Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '从FTP服务器下载文件。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpFile :FTP上的文件。
        'strFile    :本地文件全路径。
        '返回：空串表示成功，否则为错误提示。
        Dim objFtp As New clsFtp, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "请指定要下载的文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "请指定下载的文件保存到何处！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "要下载的文件已存在！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "不能连接服务器！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFile = "不能进入指定的目录，可能是权限不足或服务器上无此目录！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFtpDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "下载失败，可能是权限不足或服务器上无此文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "行，" & Err.Description
End Function

Private Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '按本地文件名上传文件到FTP服务器。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpPath :FTP上的目录，无目录会自动创建。
        'strFile    :本地文件全路径。
        'strNewFileName: 传到FTP上后的文件名，为空则按本地文件名保存
        '返回：空串表示成功，否则为错误提示。
    
        Dim objFtp As New clsFtp, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "文件" & strLocaFile & "不存在!"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        If strNewFileName = "" Then
112         strFileName = Split(strLocaFile, "\")(UBound(Split(strLocaFile, "\")))
        Else
            strFileName = strNewFileName
        End If
114     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
116     If lngReturn <> 0 Then
            '检查目录是否存在
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "创建目录失败！可能是权限不足！"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "上传文件失败，可能是权限不足！"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "不能连接服务器！"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "行，" & Err.Description
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

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function


Public Sub WriterBarCodeToLIS(rsBarcode As ADODB.Recordset, intMode As Integer, Optional ByVal intContinue As Integer)
    '功能   把条码写入LIS申请单
    Dim strErr As String
    If rsBarcode.RecordCount > 0 Then
        If Not mobjLisInsideComm Is Nothing Then
            rsBarcode.MoveFirst
            Do Until rsBarcode.EOF
                If intMode = 3 Then
                    If mobjLisInsideComm.SampleBarcodeWrite(rsBarcode("医嘱ID串"), rsBarcode("样本条码"), UserInfo.姓名, strErr, intContinue) = False Then
                        MsgBox "写入条码到LIS申请单出错!" & vbCrLf & strErr
                    End If
                Else
                    If mobjLisInsideComm.SampleBarcodeWrite(rsBarcode("医嘱ID串"), "", "", strErr, intContinue) = False Then
                        MsgBox "写入条码到LIS申请单出错!" & vbCrLf & strErr
                    End If
                End If
                rsBarcode.MoveNext
            Loop
        End If
    End If
End Sub
Public Sub WriterSampleSendDateToLIS(strAdvice As String, intType As Integer, ByVal strUser As String)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能   把送检时间写入LIS申请单中
    '参数   strAdvice = 医嘱串　逗号分隔
    '       intType = 0 送检  1 = 取消送检
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.SampleSendInfo(strAdvice, intType, strUser, strErr) = False Then
            MsgBox "写入送检时间到LIS申请单出错!" & vbCrLf & strErr
        End If
    End If
End Sub

Public Sub SaveFlexState(objThis As Object, strForm As String)
    Dim strWidth As String, strText As String, i As Integer
        
    On Error Resume Next
    
    strWidth = "": strText = ""
    For i = 0 To objThis.Cols - 1
        strWidth = strWidth & "," & objThis.Body.ColWidth(i)
        If UCase(TypeName(objThis)) = UCase("BillEdit") Then
            If objThis.msfObj.FixedRows = 1 Then strText = strText & "," & objThis.TextMatrix(0, i)
        Else
            If objThis.FixedRows = 1 Then strText = strText & "," & objThis.TextMatrix(0, i)
        End If
    Next
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "宽度", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "名称", Mid(strText, 2)
    
    If TypeName(objThis) = "VSFlexGrid" Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "冻结", objThis.FrozenCols
    End If
End Sub
Public Function RestoreFlexState(objThis As Object, strForm As String) As Boolean
    Dim strWidth As String, strText As String
    Dim arrText As Variant, i As Integer
        
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        RestoreFlexState = True: Exit Function
    End If
    
    
        
    strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "宽度", "")
    If UBound(Split(strWidth, ",")) >= objThis.Cols - 1 Then
        For i = 0 To objThis.Cols - 1
            objThis.Body.ColWidth(i) = Split(strWidth, ",")(i)
        Next
        RestoreFlexState = True
    End If
    
    
End Function

Public Function CheckDocEmpower(ByVal lng诊疗项目ID As Long, ByVal strAppend As String) As Boolean
'功能：检查操作员是否具有手术项目的执行权
'参数：strAppend=当前申请附项的填写情况串,格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    Dim lngID As Long
    Dim strDoc As String
    
    On Error GoTo errH
    strSQL = "select A.ID from 诊治所见项目 A,诊治所见分类 B where a.分类id=b.id and b.编码='06' and A.中文名='主刀医生'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDocEmpower")
    If rsTmp.RecordCount > 0 Then
        lngID = rsTmp!ID
        arrItem = Split(strAppend, "<Split1>")
        For i = 0 To UBound(arrItem)
            arrSub = Split(arrItem(i), "<Split2>")
            If Val(arrSub(2)) = lngID Then
                If Trim(arrSub(3)) <> "" Then
                    strDoc = Trim(arrSub(3))
                End If
                Exit For
            End If
        Next
    End If
    If strDoc = "" Then strDoc = UserInfo.姓名
    strSQL = "Select Count(*) as 权限 From 人员手术权限 A,人员表 B Where A.人员id = B.ID And B.姓名=[1] And A.诊疗项目id = [2] And A.记录性质 = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, lng诊疗项目ID)
    CheckDocEmpower = Val(rsTmp!权限 & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetOrderInspectInfo(ByVal lng病人ID As Long, ByVal strCondition As String) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等
    
    If gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = DynamicCreate("zl9EmrInterface.ClsEmrInterface", "新版病历")
    End If
    If Not gobjEmrInterface Is Nothing Then
        GetOrderInspectInfo = gobjEmrInterface.GetOrderInspectInfo(lng病人ID, strCondition)
    End If
    
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function BlnIsNumber(ByVal strCode As String) As Boolean
    '数字，及条码判断
     If IsNumeric(strCode) And Len(strCode) >= 12 And InStr("*-+./", Mid(strCode, 1, 1)) = 0 Then
        BlnIsNumber = True
     Else
        BlnIsNumber = False
     End If
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "") As String
    '功能:年龄合法性检查
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        strSQL = "select Zl_Age_Check([1],[2]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge, CDate(strBirthDay))
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    End If
    CheckAge = Nvl(rsTemp.Fields(0).Value)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ModifyApplyToLIS(strAdvices As String, intType As Integer)
    '功能   把签收信息写入LIS
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.ModifyApplyItemStateYJ(strAdvices, intType) = False Then
            
        End If
    End If
End Sub

Public Function GetAdvicePrice(ByVal lngPatientID As Long, ByVal lngPageID As Long) As String
    Dim strYPJG As String   '药品价格等级
    Dim strWCJG As String   '卫材价格等级
    Dim strPTXM As String   '普通项目价格等级
    
    On Error GoTo ErrHand
    
    If gobjpublicExpenses Is Nothing Then
        Set gobjpublicExpenses = CreateObject("zlPublicExpense.clsPublicExpense")
        Call gobjpublicExpenses.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    End If
    If gobjpublicExpenses.zlGetPriceGrade(gstrNodeNo, lngPatientID, lngPageID, "", strYPJG, strWCJG, strPTXM) = True Then
        If strPTXM <> "" Then
            GetAdvicePrice = " = '" & strPTXM & "'"
        Else
            GetAdvicePrice = " is null "
        End If
    Else
        GetAdvicePrice = " is null "
    End If
    
    Exit Function
ErrHand:
    MsgBox "出错函数(GetAdvicePrice),错误信息:" & Err.Number & " " & Err.Description, vbInformation, "提示"
    Err.Clear
        
End Function
