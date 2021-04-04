Attribute VB_Name = "mdlCISAudit"
Option Explicit

'######################################################################################################################
'常量定义

Public Const strSplitCmb = "—"
Public Const SPI_GETWORKAREA = 48
Public Enum DbType
    T_AnsiString = 0
    T_Binary = 1
    T_Byte = 2
    T_Boolean = 3
    T_Currency = 4
    T_Date = 5
    T_DateTime = 6
    T_Decimal = 7
    T_Double = 8
    T_Guid = 9
    T_Int16 = 10
    T_Int32 = 11
    T_Int64 = 12
    T_Object = 13
    T_SByte = 14
    T_Single = 15
    T_String = 16
    T_Time = 17
    T_UInt16 = 18
    T_UInt32 = 19
    T_UInt64 = 20
    T_VarNumeric = 21
    T_AnsiStringFixedLength = 22
    T_StringFixedLength = 23
    T_xml = 25
    T_DateTime2 = 26
    T_DateTimeOffset = 27
End Enum
Public Enum COLOR
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E0E0
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
    拖动色 = &HFFE0D9
    公共模块色 = &HC00000
    
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    
End Enum

Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

'Public Const 指标数据类型 = "1,文本;2,数值;3,日期;4,逻辑"

'----------------------------------------------------------------------------------------------------------------------
'类型定义

'用户信息
Public Type USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    模块权限 As String
    单位名称 As String
    部门名称 As String
    数据库用户 As String
End Type

'系统参数信息
Public Type SYSPARAM_INFO
    费用金额小数位数 As String
    收费诊疗项目匹配 As String
    结帐票据号长度 As Integer
    收费票据号长度 As Integer
    就诊卡号码长度 As Integer
    就诊卡字母前缀 As String
    就诊卡密文显示 As Boolean
    项目输入匹配方式 As Integer '0-双向;1-从左
    系统号 As Long
    系统名称 As String
    产品名称 As String
    模块号 As Long
    所有者 As String
    收费票种 As Integer
    结帐票种 As Integer
    结帐票号严格控制 As Boolean
    收费票号严格控制 As Boolean
    连接HIS报告 As Byte
    启用RIS As Boolean
End Type

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'----------------------------------------------------------------------------------------------------------------------
'全局变量申明
Public gcnOracle As ADODB.Connection                    '公共数据库连接，特别注意：不能设置为新的实例
Public gobjFSO As New Scripting.FileSystemObject        'FSO对象
Public ParamInfo As SYSPARAM_INFO
Public glngUserId As Long                   '当前用户id
Public UserInfo As USER_INFO
Public gblnInsure As Boolean
Public gstrSQL As String
Public gblnShowInTaskBar As Boolean
Public gfrmMain As Object
Public glngTXTProc As Long                              '保存默认的消息函数的地址
Public glngShareUseID As Long
Public gobjKernel As New clsCISKernel       '临床核心部件
Public gobjRichEPR As New cRichEPR          '病历核心部件
Public gobjPath As New clsCISPath           '临床路径部件
Public gobjEmr As Object                    '新版电子病历
Public gobjJob As Object                    '临床工作部件ZL9CISJOB
Public gobjXWHIS As Object     '新网接口部件zl9XWInterface.clsHISInner
Public gobjPlugIn As Object    '插件对象
Public gobjLIS As Object     'Lis部件
'病案评分
Public gstrPrivs As String
Public gstrDeptName As String
Public glngDeptId As Long
Public gstrDBUser As String
Public glngSys As Long
Public glngModul As Long
Public gstrSysName As String
Public gstrUserName As String
Public OldWindowProc As Long  ' Original window proc

'PDF打印
Public gstrInputSeverName As String
Public gstrInputUser As String
Public gstrInputPwd As String
    
'公共图标定义
Public Const Icon_History = 1000
Public Const Icon_Charge = 1001
Public Const Icon_Item = 1002
Public Const Icon_Report = 1003
Public Const Icon_Archives = 1004
Public Const Icon_Package = 1005
Public Const Icon_WaitPerson = 1006
Public Const Icon_NowPerson = 1007
Public Const Icon_OverPerson = 1008

Public gclsPackage As New clsPackage
Public gstrMatchMethod      As String '输入法区配方式
Public glngHIS共享号 As Long

'Private mclsUnzip As New clsUnZip

'----------------------------------------------------------------------------------------------------------------------
'模块变量申明




'######################################################################################################################
'过程清单

Public Function GetUserInfo() As Boolean
    '******************************************************************************************************************
    '功能:获取登陆用户信息
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    
    UserInfo.用户名 = UserInfo.数据库用户
    UserInfo.姓名 = UserInfo.数据库用户
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.部门名称 = IIf(IsNull(rsTmp!部门名), "", rsTmp!部门名)
        GetUserInfo = True
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '打印设置
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '打印数据,预览数据,输出到Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        If Not SearchPrintData(objPrnVsf, frmPubResource.msfPrint) Then
            MsgBox "你打印的网络不存在数据，请重新检视！", vbInformation, ParamInfo.系统名称
            Exit Function
        End If
        
        '调用打印部件处理
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("打印时间:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '帮助主题
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.系统号) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strTmp As String
        
    On Error GoTo errHand
    
    '票据号长度
    '------------------------------------------------------------------------------------------------------------------
    
    strTmp = zlDatabase.GetPara(20, ParamInfo.系统号)

    If strTmp <> "" Then
        If UBound(Split(strTmp, "|")) >= 4 Then ParamInfo.就诊卡号码长度 = Val(Split(strTmp, "|")(4))
    End If
    
    InitSysPara = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function

Public Function AdjustCodePostion(ByVal frmMain As Object, ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    objTxt.Top = objTxtParent.Top + 45
    objTxt.Left = objTxtParent.Left + frmMain.TextWidth(objTxtParent.Text) + 60
    objTxt.Width = objTxtParent.Width - frmMain.TextWidth(objTxtParent.Text) - 120
    objTxt.BackColor = objTxtParent.BackColor
    
    AdjustCodePostion = True
    
End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '功能： 将指定的信息保存在注册表中
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strKeyValue-键值
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        Call SaveSetting("ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue)
        
    Case 私有模块

        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 私有全局

        Call SaveSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strKeyValue)
        
    Case 公共模块

        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 公共全局
        
        Call SaveSetting("ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '功能： 将指定的注册信息读取出来
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strDefKeyValue-缺省键值
    '返回： strKeyValue-键值
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        strValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, strDefKeyValue)
        
    Case 私有模块

        strValue = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 私有全局

        strValue = GetSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共模块

        strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共全局
        
        strValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    GetPara = zlDatabase.GetPara(varPara, ParamInfo.系统号, lngModual, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************

    On Error GoTo ErrH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, ParamInfo.系统号, lngModual, blnSetup)

    Exit Function
    
ErrH:

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
        strSQL = "Select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "Select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        zlDatabase.OpenSQLRecord strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function GetApplyMode(ByVal StrText As String) As Byte
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    If CheckStrType(StrText, 1) And Left(ParamInfo.收费诊疗项目匹配, 1) = 1 Then
        '是全数字，按编码查找
            
        GetApplyMode = 1
        
    ElseIf CheckStrType(StrText, 2) And Left(ParamInfo.收费诊疗项目匹配, 2) = 1 Then
        '是全字母，按简码查找
        
        GetApplyMode = 2
    Else
        GetApplyMode = 3
    End If
End Function

Public Function VsfInputIsCard(ByRef vsfInput As Object, ByVal KeyAscii As Integer, ByVal lngSys As Long) As Boolean
    '******************************************************************************************************************
    '功能：判断指定文本框中当前输入是否在刷卡(是否达到卡号长度，在调用程序中判断),并根据系统参数处理是否密文显示
    '参数：KeyAscii=在KeyPress事件中调用的参数
    '******************************************************************************************************************
    
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, StrText As String
        
    '刷卡时含有特殊符号的要取消
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Function
    
    '处理当前键入后显示的内容(还未显示出来)
    StrText = vsfInput.EditText
    If vsfInput.EditSelLength = Len(vsfInput.EditText) Then StrText = ""
    If KeyAscii = 8 Then
        If StrText <> "" Then StrText = Mid(StrText, 1, Len(StrText) - 1)
    Else
        StrText = StrText & Chr(KeyAscii)
    End If
        
    '判断是否在刷卡
    If IsNumeric(StrText) And IsNumeric(Left(StrText, 1)) Then  '姓名输入框如果输的是全数字，认为是刷卡
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = Timer
        If vsfInput.EditText = "" Or StrText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(StrText), "0.000") < 0.04 Then blnCard = True   '用一台笔记本测试，一般在0.014左右
        End If
    End If
    
'    '刷卡时卡号是否密文显示
'    If blnCard Then
'        vsfInput.PasswordChar = IIf(gobjComLib.zlDatabase.GetPara(12, lngSys) = "0", "", "*")
'    Else
'        vsfInput.PasswordChar = ""
'    End If
    
    VsfInputIsCard = blnCard
End Function



Public Sub WaitOpen(ByVal frmParent As Object, ByVal strTitle As String)
    frmPubWait.OpenWait frmParent, strTitle
End Sub

Public Sub WaitClose()
    frmPubWait.CloseWait
End Sub

Public Sub WaitInfo(ByVal strInfo As String)
    frmPubWait.WaitInfo = strInfo
End Sub

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub

Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If objPrintVsf.Cols = 0 Then Exit Function
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    objPrintVsf.Cols = 0
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                lngPrintCol = lngPrintCol + 1
                
                objPrintVsf.Cols = lngPrintCol + 1
                
                objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
                objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
                If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                    objPrintVsf.ColAlignment(lngPrintCol) = 4
                Else
                    objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
                End If
            End If
        End If
    Next
    
    If objPrintVsf.Cols = 0 Then Exit Function
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
                If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                    lngPrintCol = lngPrintCol + 1
                    
                    If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "√", "")
                    Else
                        strFormat = objVsf.ColFormat(lngCol)
                        If strFormat = "" Then
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                        Else
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                        End If
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
    SearchPrintData = True
End Function


Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnOneReturn As Boolean = False) As Byte
    '******************************************************************************************************************
    '功能：打开树型+列表结构,应用于表格控件
    '参数：
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '返回：0:取消选择;1:选择;2:无数据返回
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As PointAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If blnOneReturn Then
        If rsData.RecordCount = 1 Then
            Set rsResult = rsData
            ShowPubSelect = 1
            Exit Function
        End If
    End If
    
    Call ClientToScreen(obj.hWnd, objPoint)
    
    Select Case TypeName(obj)
    Case "TextBox", "CommandButton"
    
        lngX = objPoint.x * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = obj.Height + objPoint.y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        lngObjHeight = obj.Height
        
    Case Else
        lngX = objPoint.x * Screen.TwipsPerPixelX + obj.CellLeft
        lngY = objPoint.y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
        lngObjHeight = obj.CellHeight
    End Select
    
    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, strFilterControl)
                                
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AnalyseAge(strOld As String, ByRef strAgeNumber As String, ByRef strAgeUnit As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '功能:将数据库中保存的年龄按估计的格式加载到界面
    
    Dim strTmp As Long
    
    If strOld = "岁" Then Exit Function
    
    If InStr(strOld, "岁") > 0 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "岁"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "月"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "天"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf IsNumeric(strOld) Then
        strAgeNumber = strOld
        strAgeUnit = "岁"
    Else
        strAgeNumber = strOld
        strAgeUnit = ""
    End If
    
    AnalyseAge = True
    
End Function


Public Function GetImageList(Optional ByVal intIconSize As Integer = 16) As ImageList
    Set GetImageList = frmPubResource.GetImageList(intIconSize)
End Function

Public Function CreateHelpMenu(cbsMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
        
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & ParamInfo.产品名称)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, ParamInfo.产品名称 & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.产品名称 & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With
    
    CreateHelpMenu = True
    
End Function
Public Function Get新版护理(ByVal lng病人ID, ByVal lng主页ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
     On Error GoTo errHand
    strSQL = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在老板数据", lng病人ID, lng主页ID)
    If rsTemp.RecordCount > 0 Then
        Get新版护理 = False
    Else
        Get新版护理 = True
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


Public Function Have部门性质(ByVal lng科室ID As Long, ByVal str性质 As String) As Boolean
'功能：检查指定科室是否具有指定工作性质
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    On Error GoTo ErrH
    
    strSQL = "Select 部门ID From 部门性质说明 Where 部门ID=[1] And 工作性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng科室ID, str性质)
    Have部门性质 = Not rsTmp.EOF
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetlngID(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
'功能：检查当前病人所在的科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select 入院科室ID,出院科室ID From 病案主页  where 病人ID =[1] and 主页ID =[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng主页ID)
    If rsTmp.RecordCount = 1 Then
        If Val(rsTmp!出院科室ID) = 0 Then
            GetlngID = Val(rsTmp!入院科室ID)
        Else
            GetlngID = rsTmp!出院科室ID
        End If
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemField(ByVal strTable As String, ByVal lngID As Long, Optional ByVal strField As String) As Variant
'功能：获取指定表指定字段信息
'说明：未处理NULL值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrH
    
    If strField = "" Then
        strSQL = "Select * From " & strTable & " Where ID=[1]"
    Else
        strSQL = "Select " & strField & " From " & strTable & " Where ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
    If Not rsTmp.EOF Then
        If strField = "" Then
            Set GetItemField = rsTmp
        Else
            GetItemField = rsTmp.Fields(strField).Value
        End If
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function CheckAuditSql_IN(strSQL As String, Optional blnMsg As Boolean = False, Optional intSource As Integer = 0) As Boolean
'=功能： 检测SQL主句功能 通过 函数 判断正误
'intSource 数据源 =0 表明是在标准版内执行，=1表明在EMR库执行
'返回：正常通过返回True
    Dim rsTemp          As ADODB.Recordset
    Dim zlCheck         As New clsCheck
    Dim strReturn       As String, strEMRSQL As String, strParm As String
    On Error GoTo ErrH
    If strSQL = "" Then CheckAuditSql_IN = True: Exit Function
    If strSQL = "" Then Exit Function
    strSQL = UCase(strSQL)
    If intSource = 1 Then
        'EMR库的查询不能被转成大写执行，有些情况可能需要区分大小写,转成大写只用于检查是否存在更新删除语句
        strEMRSQL = Replace(strSQL, "[MID]", "Hextoraw(:mid)")
        strEMRSQL = Replace(strEMRSQL, "[ALIDIN]", "Hextoraw(:alidin)")
    Else
        strSQL = Replace(strSQL, "[病人ID]", "-1")
        strSQL = Replace(strSQL, "[主页ID]", "-1")
        strSQL = "Select * From (" & strSQL & ") "
    End If
    
    If InStr(1, strSQL, "INSERT") = 1 Then
        zlCheck.Msg_OK "语法检测失败！不能写入数据！", vbCritical
        Exit Function
    ElseIf InStr(1, strSQL, "UPDATE") = 1 Then
        zlCheck.Msg_OK "语法检测失败！不能更新数据！", vbCritical
        Exit Function
    ElseIf InStr(1, strSQL, "DELETE") = 1 Then
        zlCheck.Msg_OK "语法检测失败！不能删除数据！", vbCritical
        Exit Function
    ElseIf InStr(1, strSQL, ";") > 0 Then
        zlCheck.Msg_OK "语法检测失败！不能使用“;”！", vbCritical
        Exit Function
    End If
    
    If intSource = 1 Then
        strParm = IIf(InStr(strEMRSQL, ":mid") = 0, "", "A^" & DbType.T_String & "^mid")
        If InStr(strEMRSQL, ":alidin") > 0 Then
            If InStr(strEMRSQL, ":mid") > 0 Then
                strParm = strParm & "|"
            End If
            strParm = strParm & "A^" & DbType.T_String & "^alidin"
        End If
        strReturn = gobjEmr.OpenSQLRecordset(strEMRSQL, strParm, rsTemp)
        If (strReturn <> "" And InStr(strReturn, "ORA-01403") = 0) Or rsTemp Is Nothing Then
            zlCheck.Msg_OK "审查依据检测失败:" & vbCrLf & "【" & strReturn & "】" & vbCrLf & "请重新录入审查依据检测语句！", vbExclamation
            Exit Function
        End If
    Else
        strSQL = "Select * From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSQL, "'", "''") & "') from dual", "mdlCISAudit")
        If rsTemp Is Nothing Then
            zlCheck.Msg_OK "语法检测失败！", vbCritical
            Exit Function
        Else
            If InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]:ORA-01403") > 0 Then
                '没找到任何数据
            ElseIf InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]") > 0 Then
                zlCheck.Msg_OK "审查依据检测失败:" & vbCrLf & "【" & Mid(rsTemp.Fields(0), 23) & "】" & vbCrLf & "请重新录入审查依据检测语句！", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If blnMsg Then zlCheck.Msg_OK "语法检测检测成功！"
    CheckAuditSql_IN = True
    Set zlCheck = Nothing
    Exit Function
ErrH:
    zlCheck.Msg_OK "语法检测检测失败！" & vbCrLf & Err.Description, vbCritical
    '判断状态 无需进错误处理中心
    Err.Clear
    Set zlCheck = Nothing
End Function

'==============================================================================
'=功能： 检测SQL主句功能
'==============================================================================
Public Function CheckAuditSql_OUT(strSQL As String, Optional lng病人ID As Long = -1, Optional lng主页ID As Long = -1) As String
    On Error GoTo ErrH
    strSQL = UCase(strSQL)
    strSQL = Replace(strSQL, "[病人ID]", CStr(lng病人ID))
    strSQL = Replace(strSQL, "[主页ID]", CStr(lng主页ID))
    
    CheckAuditSql_OUT = strSQL
        
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And Nvl(附加床位, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 1 And Nvl(附加床位, 0) = 0) B"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人入院ID", lngPatiID, lngPageID)
    If rsTemp Is Nothing Then Exit Function
    If NVL(rsTemp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTemp!ID
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEMR_MID_ALIDIN(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strMid As String, ByRef strAlidin As String) As Boolean
    On Error GoTo errHandle
    Dim strReturn As String, strExtend_Tag As String, rsTemp As New ADODB.Recordset
    strExtend_Tag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtend_Tag = "" Then Exit Function
    gstrSQL = "Select Rawtohex(ID) As ID, Rawtohex(Master_Id) As Master_Id From Bz_Act_Log Where Extend_Tag = :extendtag"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtend_Tag & "^" & DbType.T_String & "^extendtag", rsTemp)
    If strReturn <> "" Then Exit Function
    strMid = rsTemp!Master_id
    strAlidin = rsTemp!ID
    
    GetEMR_MID_ALIDIN = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'==============================================================================
'=功能： 审查项目 适用对象 与 病历种类对照
'==============================================================================
Public Function AuditFileTran(strUsed As String, intSource As Integer) As String
    Dim strType     As String
    On Error GoTo ErrH
    Select Case strUsed
        Case "2" '住院病历
            strType = IIf(intSource = 0, "2", "02")
        Case "3" '护理病历
            strType = IIf(intSource = 0, "4", "03")
        Case "4" '护理记录
            strType = "3"
        Case "6" '医嘱报告
            strType = "7"
        Case "7" '疾病证明
            strType = IIf(intSource = 0, "5", "04")
        Case "8" '知情文件
            strType = IIf(intSource = 0, "6", "05")
    End Select
    AuditFileTran = strType
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Function GetPPFS() As String
    On Error GoTo ErrH
    '输入匹配
    GetPPFS = ""
    If Val(zlDatabase.GetPara("输入匹配")) = 0 Then
        GetPPFS = "%"
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=名称:去掉字符串中的单引号("'")(ConvertString)
'=入口参数:
'=1).sStr          类型:String
'=出口参数:空
'=功能:去掉字符串(sStr)中的单引号
'=日期:2004-12-11
'=编程:欧阳
'=说明:在SQL语句中不能带单引号
'==================================================================================================
Function ConvertString(ByVal sStr As String) As String
On Error GoTo ErrH
    ConvertString = Replace(sStr, "'", "")
    ConvertString = Replace(ConvertString, "—", "")
    ConvertString = Replace(ConvertString, "&", "")
    Exit Function
ErrH:
    Err.Clear
    ConvertString = ""
End Function

'/******************************************/
'=函数:BigNote
'=功能:放大备注字段编辑框,并返回编辑后的文本
' 参数:mStr   已经库存的字符串
'      mTitle 编辑窗口的标题
'=编程:朱红军
'=日期:2002-04-03
'/******************************************/
Function Big_Note(mStr As String, mTitle As String, Optional bReadOnly As Boolean, Optional bSqlCheck As Boolean = False, Optional SqlSource As Integer = 0) As String
On Error GoTo ErrH
    With FrmNoteBox
        .intSource = SqlSource
        .SqlCheck = bSqlCheck
        .StrText = mStr
        .StrTile = mTitle
        .ReadOnly = bReadOnly
        .Show vbModal
        DoEvents
        Big_Note = IIf(bSqlCheck, .StrText, ConvertString(.StrText))
    End With
    Set FrmNoteBox = Nothing
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function RecordEprPrintInfo(ByVal bytMode As Byte, ByVal strRecordKey As String, ByVal lngNo As Long, Optional ByVal lngPatientKey As Long, Optional ByVal lngPatientPageKey As Long) As Boolean
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    If lngNo = 0 Then
        lngNo = 1
        strSQL = "Select Nvl(Max(打印次数),0)+1 As 打印次数 From 病案打印记录 Where 病人id=[1] And 主页id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", lngPatientKey, lngPatientPageKey)
        If rsTmp.BOF = False Then
            lngNo = rsTmp("打印次数").Value
        End If
    End If
    
    Select Case bytMode
    Case 1
        strSQL = "Select 病人id,主页id,病历名称 From　电子病历记录 a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_病案打印记录_Insert(" & Val(rs("病人id").Value) & "," & Val(rs("主页id").Value) & "," & lngNo & ",'" & rs("病历名称").Value & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    Case 2
        strSQL = "Zl_病案打印记录_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & strRecordKey & "','" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
    Case 3
        strSQL = "Select 名称 From　病历文件列表 a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_病案打印记录_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & rs("名称").Value & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    End Select
    
    RecordEprPrintInfo = True
    
End Function

'检测长度是否超过长度(字节数)
Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

Public Sub SetVsFlexGridChangeHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid, lngNo As Long)
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNo = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               '为了支持zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(.FixedCols + i) = False
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
'                    .ColData
                    '为了支持zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  '为了支持zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
                End If
            End If
            .ColData(i) = Val(Split(arrHead(i), ",")(3)) '将标提作为列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

Public Function zl_VsGrid_SaveToPara(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, _
ByVal lngMoudel As Long, ByVal strParaName As String, Optional ByVal bln私有 As Boolean = True, _
    Optional ByVal bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到参数表
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     lngMoudel-模块号
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------

    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If bln强制恢复保存 = False Then
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    End If

    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    zlDatabase.SetPara strParaName, strCol, glngSys, lngMoudel ', bln私有
    zl_VsGrid_SaveToPara = True
End Function

Public Function zl_VsGrid_FromParaRestore(ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal lngMoudle As Long, _
    ByVal strParaName As String, Optional bln私有 As Boolean = True, _
    Optional ByVal bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从参数表中恢复网格的宽度
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     lngMoudle-模块号
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------

    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String

    If bln强制恢复保存 = False Then
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    End If

    strParaValue = zlDatabase.GetPara(strParaName, glngSys, lngMoudle, "")
    If strParaValue = "" Then Exit Function
    
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...

    Err = 0: On Error GoTo errHand:

    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_VsGrid_FromParaRestore = True
    Exit Function
errHand:
End Function

Public Function GetCustomWhere(Optional ByVal lng病种id As Long, _
                                    Optional ByVal str性别 As String, _
                                    Optional ByVal lng科室ID As Long, _
                                    Optional ByVal int开始年龄 As Integer, _
                                    Optional ByVal int结束年龄 As Integer, _
                                    Optional ByVal str婚姻状况 As String, _
                                    Optional ByVal str住院号 As String, _
                                    Optional ByVal str病案号 As String) As String
    '******************************************************************************************************************
    '功能：组合病案借阅查询的基本条件
    '参数：
    '返回：返回记查询条件
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim strSQL As String
    
    'A 病人信息 B病案主页
    If lng病种id > 0 Then strSQL = " And (Y1.病人id,Y1.主页id) In (Select 病人id,主页id From 病人诊断记录 Where 疾病id=" & lng病种id & ")"
    If str性别 <> "" Then strSQL = strSQL & " And X.性别='" & str性别 & "'"
    If lng科室ID > 0 Then strSQL = strSQL & " And Y1.出院科室id=" & lng科室ID
    If str婚姻状况 <> "" Then strSQL = strSQL & " And Y1.婚姻状况='" & str婚姻状况 & "'"
    If Val(str住院号) > 0 Then strSQL = strSQL & " And Y1.住院号='" & str住院号 & "'"
    If str病案号 <> "" Then strSQL = strSQL & " And Z.病案号='" & str病案号 & "'"
    
    If int开始年龄 <> 0 Or int结束年龄 <> 0 Then
'        strSQL = "Select * From (" & strSQL & ") Where 年龄 Between [5] And [6]"
         strSQL = strSQL & " And X.年龄 Between '" & int开始年龄 & "' And '" & int结束年龄 & "'"
    End If
    
    GetCustomWhere = strSQL
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Where撤档时间(Optional strAlias As String) As String
    If strAlias = "" Then
        Where撤档时间 = " (撤档时间=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null) "
    Else
        Where撤档时间 = " (" & strAlias & ".撤档时间=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".撤档时间 is null) "
    End If
End Function

Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取站点条件限制
    '入参:blnAnd-是否加入 And 语句
    '出参:str别名-加入别名
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-03-02 17:27:54
    '-----------------------------------------------------------------------------------------------------------
    Dim strWhere As String
    Dim strAlia As String
    
'    If gbln存在站点控制 = False Then
        '不作控制
        zl_获取站点限制 = "": Exit Function
'    End If
    
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
     zl_获取站点限制 = strWhere
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng主例 As Long = -1, Optional lng尾列 As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '功能:移动单元格的列
    '入参:blnEdit-当前正处于编辑状态,允许新增行
    '     lng主例-主列,如果<0,则主列为0列,否则为指定的列
    '     lng尾列-尾列,如果<0,则主列为.cols-1,否则为指定的列
    '出参:lngRow-如果存在插入行,则返回被插入的行号,否则返回-1
    '返回:
    '编制:刘兴洪
    '日期:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    
    Err = 0: On Error GoTo errHand:
    
    'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
    If lng主例 <> -1 Then
        lngCol = lng主例
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lng尾列 < 0, vsGrid.Cols - 1, lng尾列)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
errHand:
End Sub

Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:插入行
    '参数:vsGrid-插入行的网格格件
    '     lngRow-当前行
    '     blnBefor-在lngrow之间或之后.true:之间,false-之后
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo errHand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
        Else
            .AddItem "", lngRow + 1
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT

    Err = 0: On Error GoTo errHand:

    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
errHand:
End Function

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '-------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

'排序处理
Public Sub zl_VsGridBeforeSort(ByVal vsGrid As VSFlexGrid, ByRef Col As Long, ByRef Order As Integer, Optional strSpaceRowNotCheckCol As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '功能:处理排序(排序时,不包含空白行)
    '入参:strSpaceRowNotCheckCol-不检查空行中的哪些列(列1,列2...)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-07-25 11:38:23
    '-----------------------------------------------------------------------------------------------------------
    Dim lngStartRow As Long, lngEndRow As Long, lngStartCol As Long, lngEndCol As Long
    Dim lngRow As Long, lngCol As Long
    Dim blnAllowSelect As Boolean, blnAllowBigSel As Boolean
    Dim lngOldBackColor As Long
    
    If vsGrid.ExplorerBar > &H1000& Then Exit Sub
    '保存当前的选择区域
    vsGrid.GetSelection lngStartRow, lngStartCol, lngEndRow, lngEndCol
    vsGrid.Redraw = flexRDNone
    blnAllowBigSel = vsGrid.AllowBigSelection: blnAllowSelect = vsGrid.AllowSelection
    
    '不排序空白行
    With vsGrid
        For lngRow = .Rows - 1 To .FixedRows Step -1
            For lngCol = 0 To .Cols - 1
               If InStr(1, "," & strSpaceRowNotCheckCol & ",", "," & lngCol & ",") > 0 Then
               Else
                    If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then GoTo GoNext:
               End If
            Next
        Next
GoNext:
        If lngRow > .FixedRows Then
            
             .Select .FixedRows, Col, lngRow, Col
            .Sort = Order
        End If
        ' 恢复以前选择的区域
        .Select lngStartRow, lngStartCol, lngEndRow, lngEndCol
            
        .Redraw = flexRDDirect
    End With
    Order = 0
End Sub


Public Function zl_vsGrid_Para_Save(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制保存 As Boolean = False, Optional blnHaveParaPrivs As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '功能:保存vsFlex的宽度到注册表
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnSaveToDataBase = False Then
        zl_vsGrid_Para_Save = True
        If bln强制保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    If blnSaveToDataBase Then
        zlDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
    Else
        Call SaveRegInFor(g私有模块, strCaption, strKey, strCol)
    End If
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional bln强制恢复保存 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:从数据库中恢复网格的宽度等信息
    '参数:vsGrid-对应的网络控件
    '     strCaption-窗体名
    '     strKey-主建
    '     blnSaveToDataBase-是否是往数据库中保存参数(如果是往数据库中保存,则强制保存为true,否则根据是否使用个性化风格来确定)
    '     bln强制恢复保存-决定是否将保存注册表的参数值,进行强制恢复
    '返回:恢复成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        '只有在本地注册表中才会处理个性化设置
        zl_vsGrid_Para_Restore = True
        If bln强制恢复保存 = False Then
            If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g私有模块, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:保存格式:列主键,列宽,列隐藏|列主键,列宽,列隐藏|...
    Err = 0: On Error GoTo errHand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
errHand:
End Function
 


'*********************************************************************************************************************
Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
errHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
errHand:
End Sub

Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '功能:获取bool列的值
    '返回:是该单元格为true,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/28
    '------------------------------------------------------------------------------
    Dim strTemp As String
    Err = 0: On Error GoTo errHand:
    With vsGrid
        strTemp = .TextMatrix(lngRow, lngCol)
    End With
    If UCase(strTemp) = UCase("True") Then
        GetVsGridBoolColVal = True: Exit Function
    End If
    GetVsGridBoolColVal = Val(strTemp) <> 0
    Exit Function
errHand:
End Function

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function


'################################################################################################################
'## 功能：  将病案审查方案文件导出到XML文档中
'##
'## 参数：  tvwThis     :   RTB编辑器控件
'##         strFileName :   XML文件名（全路径）
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ExportToXMLFile(ByRef tvwThis As TreeView, ByVal strFileName As String) As Boolean
    Dim i As Long, j As Long, k As Long
    Dim oDoc As DOMDocument             'xml文档
    Dim oRoot  As IXMLDOMElement        '根节点
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '节点
    Dim oSubNode3 As IXMLDOMNode        '节点
 
    Dim strPath As String
    Dim strSolutionID As String         '方案ID
    Dim strSQL As String                'SQL
    Dim rsTree As ADODB.Recordset       '分类记录集
    
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strSolutionID = Replace(tvwThis.SelectedItem.Key, "Root", "")
    
    'XML文档
    Set oDoc = New DOMDocument
    '注释
    oDoc.appendChild oDoc.createComment(gstrSysName & "  " & _
        "操作员:" & gstrUserName & "，部门:" & gstrDeptName & "，时间:" & _
        Format(Now(), "YYYY年MM月DD日"))
    '根节点
    Set oRoot = oDoc.createElement("SpotCheck")
    Set oDoc.documentElement = oRoot    '设置为根节点
    Call oRoot.setAttribute("SolutionName", tvwThis.SelectedItem.Text)
    Call oRoot.setAttribute("SolutionID", Replace(tvwThis.SelectedItem.Key, "Root", ""))
    
    strSQL = "SELECT /*+ rule */ id,上级ID,方案ID,编码,名称 FROM 病案审查分类 Where 方案ID=[1] START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID"
    Set rsTree = zlDatabase.OpenSQLRecord(strSQL, "病案审查分类", strSolutionID)
    rsTree.Sort = "编码"
    
    '病案审查分类
    Set oNode = CreateNode(1, oRoot, "Classify", NODE_ELEMENT, "")
    rsTree.MoveFirst
    Do Until rsTree.EOF
      '添加子节点
       Set oSubNode1 = CreateNode(2, oNode, "Class", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "ID", , zlCommFun.NVL(rsTree!ID, 0)
            CreateNode 3, oSubNode1, "方案ID", , zlCommFun.NVL(rsTree!方案ID)
            CreateNode 3, oSubNode1, "上级ID", , zlCommFun.NVL(rsTree!上级ID)
            CreateNode 3, oSubNode1, "编码", , zlCommFun.NVL(rsTree!编码)
            CreateNode 3, oSubNode1, "名称", , zlCommFun.NVL(rsTree!名称)
        rsTree.MoveNext
    Loop
    
    
    strSQL = "Select /*+ rule */ a.id,a.分类id,a.编码,a.名称,a.简码,a.分值,a.分制,a.说明,a.审查依据,a.适用对象,a.文件ID,a.适用环节" & vbNewLine & _
            "  From 病案审查目录 a, 病案审查分类 b,病案审查方案 C" & vbNewLine & _
            " Where a.分类id = b.ID And b.方案id = C.id And C.id =[1]"
    Set rsTree = zlDatabase.OpenSQLRecord(strSQL, "病案审查分类", strSolutionID)
    rsTree.Sort = "分类id"
    
    '病案审查目录
    Set oNode = CreateNode(1, oRoot, "Catalogue", NODE_ELEMENT, "")
    rsTree.MoveFirst
    Do Until rsTree.EOF
      '添加子节点
        Set oSubNode1 = CreateNode(2, oNode, "Catalog", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "ID", , zlCommFun.NVL(rsTree!ID, 0)
            CreateNode 3, oSubNode1, "分类ID", , zlCommFun.NVL(rsTree!分类id)
            CreateNode 3, oSubNode1, "编码", , zlCommFun.NVL(rsTree!编码)
            CreateNode 3, oSubNode1, "名称", , zlCommFun.NVL(rsTree!名称)
            CreateNode 3, oSubNode1, "简码", , zlCommFun.NVL(rsTree!简码)
            CreateNode 3, oSubNode1, "分值", , zlCommFun.NVL(rsTree!分值)
            CreateNode 3, oSubNode1, "分制", , zlCommFun.NVL(rsTree!分制)
            CreateNode 3, oSubNode1, "说明", , zlCommFun.NVL(rsTree!说明)
            CreateNode 3, oSubNode1, "审查依据", , zlCommFun.NVL(rsTree!审查依据)
            CreateNode 3, oSubNode1, "适用对象", , zlCommFun.NVL(rsTree!适用对象)
            CreateNode 3, oSubNode1, "文件ID", , zlCommFun.NVL(rsTree!文件ID)
            CreateNode 3, oSubNode1, "适用环节", , zlCommFun.NVL(rsTree!适用环节)
         rsTree.MoveNext
    Loop
 
    '版本信息
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = oDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call oDoc.insertBefore(pi, oDoc.childNodes(0))
    '直接保存成文件即可
    oDoc.Save strFileName
    
    Set oDoc = Nothing
    ExportToXMLFile = True
    Exit Function
LL:
    ExportToXMLFile = False
End Function


'################################################################################################################
'## 功能：  从XML文件导入病案审查方案
'##
'## 参数：  tvwThis     :   RTB编辑器控件
'##         strFileName :   XML文件名（全路径）
'##         blnPrompt   :   是否提示导入覆盖，默认为True
'##         blnForUndoRedo : 是否用于Undo/Redo，默认为False
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ImportFromXMLFile(ByRef tvwThis As TreeView, _
    ByVal strFileName As String, _
    Optional blnPrompt As Boolean = True, _
    Optional blnForUndoRedo As Boolean = False) As Boolean
    
    Dim i As Long, j As Long, k As Long, lngSelStart As Long, lngSelEnd As Long
    
    Dim lKey As Long
    Dim oDoc As DOMDocument             'xml文档
    Dim oRoot  As IXMLDOMElement        '根节点
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    
    Dim cllTemp As New Collection
    Dim strCurSolutionID As String      '当前方案ID
    Dim strSolutionName As String       '原来方案名称
    Dim strSolutionID As String         '原来方案ID
    
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCode As String
    Dim strPath As String
    Dim strID As String
    Dim lngTypeID As Long
    Dim lngTypePrivID As Long
    Dim strTypeCode As String
    Dim strTypeName As String
    Dim lngProjectID As String
    Dim strParentCode As String
    
    
    Dim lngItemID  As Long
    Dim intItemTypeID As Integer
    Dim strItemCode As String
    Dim strItemName As String
    Dim strItemMnemonicCode As String
    Dim strItemDescription As String
    Dim strItemAudit As String
    Dim intItemUsed As Integer
    Dim intItemFileID As Integer
    Dim strItemLink As String
    Dim strPalValue As String
    Dim strNumValue As String
    
    On Error GoTo ErrH
    
    strCurSolutionID = Replace(tvwThis.SelectedItem.Key, "Root", "")
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    
    Set oDoc = New DOMDocument
    oDoc.Load strFileName
    '如果不包含任何元素，则退出
    If oDoc.documentElement Is Nothing Then
        Exit Function
    End If
    If blnPrompt Then
        If MsgBox("注意：导入文件后原有内容将不可恢复，是否继续覆盖当前文件？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then
            Exit Function
        End If
    End If
    '读取文件结构
    Set oRoot = oDoc.selectSingleNode("SpotCheck")       'oRoot置为根节点
    If oRoot Is Nothing Then MsgBox "该XML文件不是正确的病案审查方案文件！", vbInformation, gstrSysName: Exit Function
    
    '获取基础信息
    On Error Resume Next
    strSolutionName = oRoot.getAttributeNode("SolutionName").Text
    strSolutionID = Val(oRoot.getAttributeNode("SolutionID").Text)
    Screen.MousePointer = vbHourglass
    
    
    '读取病案审查分类 Classify:
    Set oNode = oRoot.selectSingleNode("Classify")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "ID", 0)
        If lKey > 0 Then
            If Val(GetNodeValue(oSubNode1, "上级ID", 0)) = 0 Then
                strID = "-1"
            Else
                strID = GetCllValue(cllTemp, Val(GetNodeValue(oSubNode1, "上级ID", 0)))
            End If
            
            strSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病案审查分类", strID)
            If rsTemp.RecordCount = 1 Then
                    strParentCode = "" & rsTemp!编码
            Else
                    strParentCode = ""
            End If
            
            If Val(GetNodeValue(oSubNode1, "上级ID", 0)) = 0 Then
               lngTypePrivID = 0
               strSQL = "select max(编码) as 编码 from 病案审查分类 a Where a.上级ID is null"
               Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病案审查分类", strID)
            Else
               lngTypePrivID = strID
               strSQL = "select max(编码) as 编码 from 病案审查分类 a Where a.上级ID = [1]"
               Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病案审查分类", strID)
            End If
            
            strCode = ""
            If rsTemp.RecordCount = 1 Then
                strCode = rsTemp!编码
                strCode = IncStr(strCode)
            End If
            If strCode = "" Then
                strTypeCode = strParentCode & "01"
            Else
                strTypeCode = strCode
            End If
            strTypeName = GetNodeValue(oSubNode1, "名称", "")
            lngProjectID = Val(strCurSolutionID)
            lngTypeID = zlDatabase.GetNextId("病案审查分类")
            
            AddArray cllTemp, lngTypeID & ";" & lKey
            
            strSQL = "Zl_病案审查分类_Insert (" + CStr(lngTypeID) & "," & IIf(lngTypePrivID = 0, "NULL", CStr(lngTypePrivID)) + "," + "'" + strTypeCode + "'" + "," + "'" + strTypeName + "'," & CStr(0) & "," & lngProjectID & ")"
            zlDatabase.ExecuteProcedure strSQL, "病案审查分类"
        End If
    Next


    '病案审查目录
    Set oNode = oRoot.selectSingleNode("Catalogue")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "ID", 0)
        If lKey > 0 Then
            strID = GetCllValue(cllTemp, Val(GetNodeValue(oSubNode1, "分类ID", 0)))
            
            lngItemID = zlDatabase.GetNextId("病案审查目录")
            intItemTypeID = Val(strID)
    
            strItemName = GetNodeValue(oSubNode1, "名称", "")
            strItemMnemonicCode = GetNodeValue(oSubNode1, "简码", "")
            strItemDescription = GetNodeValue(oSubNode1, "说明", "")
            strItemAudit = GetNodeValue(oSubNode1, "审查依据", "")
            intItemUsed = Val(GetNodeValue(oSubNode1, "适用对象", 0))
            intItemFileID = Val(GetNodeValue(oSubNode1, "文件ID", 0))
            strItemLink = GetNodeValue(oSubNode1, "适用环节", "")
            strPalValue = GetNodeValue(oSubNode1, "分制", "")
            strNumValue = GetNodeValue(oSubNode1, "分值", "")
            
            strSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病案审查分类", strID)
   
            If rsTemp.RecordCount = 1 Then
                 strSQL = "select nvl(Max(编码),0) from 病案审查目录 where 分类id=[1] and 编码 like [2] || '%'"
                 strItemCode = IncStr(zlDatabase.OpenSQLRecord(strSQL, "病案审查目录", intItemTypeID, rsTemp!编码).Fields(0))
                
                 If strItemCode = "1" Then
                     strItemCode = rsTemp!编码 & Format(strItemCode, "0000")
                 End If
                 strItemCode = InsertNewCode(strItemCode)
            
                strSQL = "Zl_病案审查目录_Insert (" + CStr(lngItemID) + "," + IIf(intItemTypeID = 0, "NULL", CStr(intItemTypeID)) + "," + "'" + strItemCode + "'" + "," + "'" + strItemName + "','" & strItemMnemonicCode & "','" & strItemDescription & "','" & strItemAudit & "'," & intItemUsed & ",'" & intItemFileID & "','" & strItemLink & "'," & Val(strPalValue) & " ," & Val(strNumValue) & ")"
                zlDatabase.ExecuteProcedure strSQL, "病案审查目录"
            End If
        End If
    Next

    Screen.MousePointer = 0
    ImportFromXMLFile = True
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  获取一个节点的值
'##
'## 参数：  CurNode         :   当前节点对象
'##         SubNodeName     :   子节点名称
'##         DefaultValue    :   默认值
'################################################################################################################
Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, _
    ByVal SubNodeName As String, _
    Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    Dim NodeTMP As IXMLDOMNode
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
    
    If InStr(GetNodeValue, vbCr) > 0 And InStr(GetNodeValue, vbLf) = 0 Then '只有回车符无换行符
        GetNodeValue = Replace(GetNodeValue, vbCr, vbCrLf)
    ElseIf InStr(GetNodeValue, vbLf) > 0 And InStr(GetNodeValue, vbCr) = 0 Then '只有换行符无回车符
        GetNodeValue = Replace(GetNodeValue, vbLf, vbCrLf)
    End If
End Function

'################################################################################################################
'## 功能：  创建一个XML节点并赋值
'##
'## 参数：  TabNumber   :   缩进层次数（表示有多少个Tab制表符，便于阅读）
'##         Parent      :   父节点
'##         Node_Type   :   节点类型（目前支持 NODE_ELEMENT 、NODE_CDATA_SECTION 、NODE_COMMENT 、NODE_ATTRIBUTE等）
'##         Node_Name   :   节点名称
'##         Node_Value  :   节点文本
'################################################################################################################
Private Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal node_name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "")
    Dim New_Node As IXMLDOMNode
    
    '字符缩进值设置（不影响数据），只影响阅读美观度
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '创建文本节点
    '创建新节点
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, node_name, "")
    '设置文本值
    New_Node.Text = Node_Value
    '添加到父节点
    Parent.appendChild New_Node
    '添加末尾回车（不影响数据），只影响阅读美观度
    'Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf)   '创建文本节点
    Set CreateNode = New_Node
End Function


''''################################################################################################################
''''## 功能：  清除所有对象的ID号
''''################################################################################################################
'''Public Sub ClearAllIDs()
'''    Dim i As Long, j As Long
'''    For i = 1 To Me.Compends.count
'''        Me.Compends(i).ID = 0
'''    Next
'''    For i = 1 To Me.Pictures.count
'''        Me.Pictures(i).ID = 0
'''    Next
'''    For i = 1 To Me.elements.count
'''        Me.elements(i).ID = 0
'''    Next
'''    For i = 1 To Me.Signs.count
'''        Me.Signs(i).ID = 0
'''    Next
'''    For i = 1 To Me.Diagnosises.count
'''        Me.Diagnosises(i).ID = 0
'''    Next
'''    For i = 1 To Me.Tables.count
'''        Me.Tables(i).ID = 0
'''        For j = 1 To Me.Tables(i).Cells.count
'''            Me.Tables(i).Cells(j).ID = 0
'''        Next
'''        For j = 1 To Me.Tables(i).elements.count
'''            Me.Tables(i).elements(j).ID = 0
'''        Next
'''        For j = 1 To Me.Tables(i).Pictures.count
'''            Me.Tables(i).Pictures(j).ID = 0
'''        Next
'''    Next
'''End Sub

Private Function GetCllValue(ByVal cllTmp As Collection, ByVal strKey As String) As String
'获取审查分类中对应的新ID
    Dim lngNum As Long
    If cllTmp.count > 0 Then
        For lngNum = 1 To cllTmp.count
            If InStrRev(cllTmp.Item(lngNum), ";" & strKey) > 0 Then
                GetCllValue = Replace(cllTmp.Item(lngNum), ";" & strKey, "")
                Exit Function
            End If
        Next
    End If
    
End Function

'========================================================================================
'=新增 或 复制时检测产生的编码是否已存在
'========================================================================================
Private Function InsertNewCode(strInCode) As String
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "select 1 from 病案审查目录 where 编码 = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病案审查目录", strInCode)
    If rsTemp.RecordCount = 0 Then InsertNewCode = strInCode: Exit Function
    strInCode = IncStr(strInCode)
    InsertNewCode = InsertNewCode(strInCode)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowPubSelectTest(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnOneReturn As Boolean = False) As Byte

    On Error GoTo errHand

       ShowPubSelectTest = ShowPubSelect(frmParent, obj, bytStyle, strLvw, strSavePath, strDescrible, rsData, rsResult, lngCX, lngCY, blnMuliSel, strInitKey, strFilterControl, blnOneReturn)

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'功能：判断 RIS接口部件(zl9XWInterface.clsHISInner) 是否存在，并启用
'参数：blnMsg－创建失败时是否提示

    If Not ParamInfo.启用RIS Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

Public Function InitObjLis(Optional ByVal blnMsg As Boolean) As Boolean
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = GetObject(, "zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
    
        On Error Resume Next
        If gobjLIS Is Nothing Then Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
        
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If blnMsg Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
                Exit Function
            End If
        End If
    End If
    InitObjLis = True
End Function

Public Function CreateCISJOB() As Boolean
'功能：判断 临床工作站部件(ZL9CISJOB.CLSCISJob) 是否存在

    If Not gobjJob Is Nothing Then CreateCISJOB = True: Exit Function
    On Error Resume Next
    Set gobjJob = GetObject(, "ZL9CLSCISJOB.CLSCISJOB")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjJob Is Nothing Then Set gobjJob = DynamicCreate("ZL9CISJOB.CLSCISJob", "临床工作站", False)
    Err.Clear: On Error GoTo 0
    If gobjJob Is Nothing Then Exit Function
    CreateCISJOB = True
End Function
