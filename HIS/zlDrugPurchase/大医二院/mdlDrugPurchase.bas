Attribute VB_Name = "mdlDrugPurchase"
Option Explicit

Public Const GSTR_MESSAGE = "提示信息"

Public gstrUser As String, gstrUserNameNew As String
Public glngUserID As Long, glngDeptID As Long
Public gbyt效期 As Byte

Public gcnOutside As New ADODB.Connection           '外部数据库连接
Public gcnOracle As ADODB.Connection
Public gblnSetupFinish As Boolean
Public gobjComLib As Object

Public Const GSTR_SYSNAME = "采购数据交互接口"
Public Const GSTR_REGEDIT_PATH = "公共模块\DrugPurchaseDBServer"
Public Const MSTR_SERVER = "localhost"
Public Const MSTR_DBNAME = "GuoYaoDB"
Public Const MSTR_USER = "sa"
Public Const MSTR_PASSWORD = ""

Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Enum enm_Pop_File
     File = 1
     FilePrintSet = 181
     FilePreview = 102
     FilePrint = 103
     FileExit = 191
     Edit = 3
     EditAdd = 3212
     EditDel = 3213
     EditEdit = 23
     EditIgnore = 3214
     EditProcess = 3104
     EditCurrChoose = 301
     EditCurrCancel = 302
     EditChooChoose = 303
     EditChooCancel = 304
     EditAllChoose = 305
     EditAllCancel = 306
     View = 4
     ViewRefresh = 791
     ViewFindTitle = 411
     ViewFindEdit = 412
     ViewFindButton = 413
     ViewTools = 420
     ViewToolsButton = 421
     ViewToolsLabel = 422
     ViewToolsIcon = 423
     ViewStatebar = 430
     Import = 5
     ImportTitle = 51
     ImportControl = 52
     Help = 9
     HelpHelp = 901
     HelpWeb = 902
     HelpWebhome = 9021
     HelpWebBBS = 9022
     HelpWebFeelback = 9023
     HelpAbout = 903
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Function MSSQLServerOpen(ByVal strServerName As String, ByVal strDBName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的MS SQL Server 数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    If Len(Trim(strUserName)) = 0 Then
        MSSQLServerOpen = False
        MsgBox "请设置外联数据库信息！", vbInformation, GSTR_MESSAGE
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .ConnectionTimeout = 5
        .Open "Driver={SQL Server};Server=" & strServerName & ";Database=" & strDBName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Or Err.Number = -2147467259 Then
                MsgBox "中间数据库连接失败！", vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            MSSQLServerOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    MSSQLServerOpen = True
    Exit Function
    
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    MSSQLServerOpen = False
    Err = 0
End Function


Public Function OraDataOpenTest(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Then
                MsgBox Mid(strError, InStr(1, strError, "[SQL Server]"), Len(strError)), vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            OraDataOpenTest = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpenTest = True
    Exit Function
    
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    OraDataOpenTest = False
    Err = 0
End Function

Public Function StringEnDeCodecn(strSource As String, MA) As String
'该函数只对中西文起到加密作用
'参数为：源文件，密码
    On Error GoTo ErrEnDeCode
    Dim X As Single, i As Integer
    Dim CHARNUM As Long, RANDOMINTEGER As Integer
    Dim SINGLECHAR As String * 1
    Dim strTmp As String
    
    If MA < 0 Then
        MA = MA * (-1)
    End If
    
    X = Rnd(-MA)
    For i = 1 To Len(strSource) Step 1                 '取单字节内容
        SINGLECHAR = Mid(strSource, i, 1)
        CHARNUM = Asc(SINGLECHAR)
g:
        RANDOMINTEGER = Int(127 * Rnd)
        If RANDOMINTEGER < 30 Or RANDOMINTEGER > 100 Then GoTo g
        CHARNUM = CHARNUM Xor RANDOMINTEGER
        strTmp = strTmp & Chr(CHARNUM)
    Next i
    StringEnDeCodecn = strTmp
    Exit Function

ErrEnDeCode:
    StringEnDeCodecn = ""
    MsgBox Err.Number & "\" & Err.Description
End Function

Public Function GetUserNameInfo() As Boolean
'获取用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    
    With rsTmp
        If Not .EOF Then
            glngUserID = IIf(IsNull(!Id), 0, !Id)
            glngDeptID = IIf(IsNull(!部门id), 0, !部门id)
            gstrUserNameNew = IIf(IsNull(!姓名), "", !姓名) '当前用户姓名
            GetUserNameInfo = True
        Else
            glngUserID = 0
            glngDeptID = 0
            gstrUserNameNew = "" '当前用户姓名
        End If
    End With
    rsTmp.Close

    strSQL = "Select 参数号, 参数值, 缺省值 From Zlparameters Where 系统 = [1] And Nvl(私有, 0) = 0 And 模块 Is Null and 参数号=[2] "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "取系统参数", 100, 149)
    With rsTmp
        If Not .EOF Then
            gbyt效期 = IIf(IsNull(rsTmp!参数值), rsTmp!缺省值, rsTmp!参数值)
        Else
            gbyt效期 = 0
        End If
    End With
    
End Function

Public Sub SelText(ByVal ctlVal As Control)
    If TypeOf ctlVal Is TextBox Then
        ctlVal.SelStart = 0
        ctlVal.SelLength = Len(ctlVal.Text)
    End If
End Sub

Public Sub InitCommandBars(ByVal cmbVal As CommandBars)
    cmbVal.VisualTheme = xtpThemeOffice2003
    With cmbVal.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cmbVal.EnableCustomization False
    cmbVal.Icons = frmPublic.imgPublic.Icons
End Sub

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

Public Sub ProviderSelecter(frmParam As Form, ByVal objParam As Object, ByVal blnClick As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strParam As String
    Dim blnCancel As Boolean
    Dim vRect As RECT

    vRect = GetControlRect(objParam.hwnd)
    If blnClick = False Then
        strParam = "%" & UCase(Trim(objParam.Text)) & "%"
        strSQL = "SELECT id, 编码, 简码, 名称  " _
                & "FROM 供应商 a " _
                & "Where TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                & "  and (substr(类型,1,1)=1 Or Nvl(末级,0)=0)" _
                & "  and (a.编码 like [1] or a.简码 like [1] or a.名称 like [1]) " _
                & "order by a.编码"
    Else
        strSQL = "SELECT id, 编码, 简码, 名称 " _
                & "FROM 供应商 a " _
                & "Where TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " _
                & "  and (substr(类型,1,1)=1 Or Nvl(末级,0)=0)" _
                & "order by a.编码"
    End If
    
    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(frmParam, strSQL, 0, "供应商", False, "", "" _
              , False, False, True, vRect.Left, vRect.Top, objParam.Height, blnCancel, False, False, strParam)

    If Not rsTmp Is Nothing Then
        objParam.Text = rsTmp!名称
        rsTmp.Close
    End If
    
End Sub


Public Sub InitVSF(ByVal vsfVal As VSFlexGrid, blnVal As Boolean)
'切换显示VSF的标题栏
    Dim strCols As String
    Dim arrCols As Variant
    Dim i As Single
    
    With vsfVal
        .Rows = 1
        .ColWidth(0) = 130 * 2                               '第一列宽
        .ColWidth(1) = 130
        .FixedCols = 2                                       '固定前二列
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns               '运行时可调整Columns宽度
        .AllowSelection = True                               '多单元选择控制开关
        .SelectionMode = flexSelectionListBox                '多单元选择控制
        .ExplorerBar = flexExSortShow
        .BackColorSel = &HC0E0FF
        .BackColorAlternate = .BackColor      '&H80000003
        '.BackColorBkg = vbWhite
    End With
    
    If blnVal Then
        '发票导入
        strCols = "||选择,choose,440|H_供应商ID,providerid,800|供应商,provider,1500|H_药品ID,id,650|计划单号,plan_code,1000" & _
                  "|药品名称,name,2000|药品规格,spec,1500|发票数量,ivqty,850,r|PDA验收数量,pdaqty,1100,r|已验收数量,chkqty,1000,r" & _
                  "|验收数量,qty,850,r|药库单位,unit,800|验收人,Accepter,800|批发价,price,1000,r" & _
                  "|发票金额,iamount,1200,r|发票号,invoice,1000|发票日期,idate,1000|生产商,producer,1000" & _
                  "|批号,lot_no,1000|生产日期,pdate,1000|效期,avail_date,1000|H_DetailID,detail_id,0" & _
                  "|H_已导入,imported,600|消息,mess,2000"
    Else
        '计划导出
        strCols = "||选择,choose,440|H_计划ID,planid,0|计划单号,planno,1000|序号,xh,500|H_供应商ID,providerid,800|供应商,provider,1500" & _
                  "|H_药品ID,id,650|药品名称,name,2000|药品规格,spec,1500|计划数量,qty,850,r|药库单位,unit,800|单价,price,1000,r" & _
                  "|生产商,producer,1000|H_药库ID,wh_id,600|药库,wh,1000|H_药房ID,dh_id,0|药房,dh,1000|编制日期,edate,1000" & _
                  "|审核日期,cdate,1000|H_已导入,imported,600|备注,remark,3000|消息,mess,2000"
    End If
    arrCols = Split(strCols, "|")
    With vsfVal
        .Clear
        .Cols = UBound(arrCols) + 1
        For i = LBound(arrCols) To UBound(arrCols)
            If arrCols(i) = "" Then
                .TextMatrix(0, i) = ""
            Else
                .TextMatrix(0, i) = Split(arrCols(i), ",")(0)
                .ColKey(i) = Split(arrCols(i), ",")(1)
                .ColWidth(i) = Split(arrCols(i), ",")(2)
                'H_为隐藏列
                If Mid(Split(arrCols(i), ",")(0), 1, 2) = "H_" Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                    If UBound(Split(arrCols(i), ",")) > 2 Then
                        .ColAlignment(i) = flexAlignRightCenter
                    End If
                End If
            End If
        Next
        .ColDataType(.ColIndex("choose")) = flexDTBoolean    '设置为Check控件
    End With
    
End Sub

Public Function CheckProvider(ByVal lngProviderID As Long) As String
'审核供应商ID
    Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord("select 名称 from 供应商 where 撤档时间>to_date('2999/12/31','yyyy/mm/dd') And id=[1]", "审核供应商ID", lngProviderID)
'    Set rsTmp = zlDatabase.OpenSQLRecord("Select (Select 名称 From 供应商 Where 撤档时间=to_date('3000/1/1','yyyy/mm/dd') And Id=[1]) 名称, " & _
'                                         "(Select count(1) From 药品中标单位 Where 单位Id=[1] And 药品ID=[2] and 撤档时间=to_date('3000/1/1','yyyy/mm/dd')) 是否中标 " & _
'                                         "from dual ", "审核供应商ID", intProviderID, lngDrugID)
    If rsTmp.RecordCount = 1 Then
'        CheckProvider = rsTmp!名称 & "|" & rsTmp!是否中标
        CheckProvider = rsTmp!名称
    End If
    rsTmp.Close
End Function

Public Sub DataLoading(ByVal vsfVal As VSFlexGrid, ByVal rsVal As ADODB.Recordset, ByVal bytTab As Byte, Optional ByVal bytMarked As Byte = 0)
    Dim i As Integer, j As Integer
    Dim strName As String, strSpec As String, strUnit As String, strProvider As String
    Dim blnGet As Boolean
    Dim dblCost As Double

    On Error GoTo errHandle
    With vsfVal
        .Rows = 1
        .Rows = rsVal.RecordCount + 1
        If rsVal.RecordCount > 0 Then rsVal.MoveFirst
        For i = 1 To rsVal.RecordCount
            strName = "": strSpec = "": strUnit = ""
            
            Err = 0: On Error Resume Next
            blnGet = GetMedicalInfo(IIf(IsNull(rsVal!药品id), -1, rsVal!药品id), strName, strSpec, strUnit)
            If Err <> 0 Then
                .TextMatrix(i, .ColIndex("mess")) = "“药品ID”" & Err.Description & "[外部数据库]。"
            End If
            Err = 0: On Error GoTo errHandle
            
            .TextMatrix(i, 1) = i   '序号
            '采购单导出
            If bytTab = 0 Then
                .TextMatrix(i, .ColIndex("planid")) = IIf(IsNull(rsVal!Id), "", rsVal!Id)
                .TextMatrix(i, .ColIndex("planno")) = IIf(IsNull(rsVal!no), "", rsVal!no)
                .TextMatrix(i, .ColIndex("xh")) = IIf(IsNull(rsVal!序号), "", rsVal!序号)
                .TextMatrix(i, .ColIndex("providerid")) = IIf(IsNull(rsVal!供应商id), "", rsVal!供应商id)
                .TextMatrix(i, .ColIndex("provider")) = IIf(IsNull(rsVal!上次供应商), "", rsVal!上次供应商)
                .TextMatrix(i, .ColIndex("producer")) = IIf(IsNull(rsVal!上次生产商), "", rsVal!上次生产商)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rsVal!药品id), "", rsVal!药品id)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rsVal!名称), "", rsVal!名称)
                .TextMatrix(i, .ColIndex("spec")) = IIf(IsNull(rsVal!规格), "", rsVal!规格)
                .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(rsVal!药库单位), "", rsVal!药库单位)
                .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(rsVal!计划数量), "0", rsVal!计划数量)
                '.ColFormat(.ColIndex("qty")) = "#0"
                .TextMatrix(i, .ColIndex("price")) = IIf(IsNull(rsVal!单价), "0", rsVal!单价)
                .ColFormat(.ColIndex("price")) = "#0.0000"
                .TextMatrix(i, .ColIndex("wh_id")) = IIf(IsNull(rsVal!药库id), "", rsVal!药库id)
                .TextMatrix(i, .ColIndex("wh")) = IIf(IsNull(rsVal!药库), "", rsVal!药库)
                .TextMatrix(i, .ColIndex("dh_id")) = IIf(IsNull(rsVal!药房id), "", rsVal!药房id)
                .TextMatrix(i, .ColIndex("dh")) = IIf(IsNull(rsVal!药房), "", rsVal!药房)
                .TextMatrix(i, .ColIndex("edate")) = IIf(IsNull(rsVal!编制日期), "", rsVal!编制日期)
                .ColFormat(.ColIndex("edate")) = "yyyy-mm-dd"
                .TextMatrix(i, .ColIndex("cdate")) = IIf(IsNull(rsVal!审核日期), "", rsVal!审核日期)
                .ColFormat(.ColIndex("cdate")) = "yyyy-mm-dd"
                
                If rsVal!是否上传 = 1 Then
                    .TextMatrix(i, .ColIndex("choose")) = 0
                    .TextMatrix(i, .ColIndex("imported")) = "1,0"
                    .Cell(flexcpForeColor, i, 3, i, .ColIndex("mess")) = vbBlue
                ElseIf .TextMatrix(i, .ColIndex("qty")) > 0 Then
                    .TextMatrix(i, .ColIndex("choose")) = 1
                    .TextMatrix(i, .ColIndex("imported")) = "1,1"
                Else
                    .TextMatrix(i, .ColIndex("choose")) = 0
                    .TextMatrix(i, .ColIndex("imported")) = "1,0"
                    .Cell(flexcpForeColor, i, 3, i, .ColIndex("mess")) = vbRed
                End If
            '入库单导入
            Else
                .TextMatrix(i, .ColIndex("providerid")) = IIf(IsNull(rsVal!供应商id), "-1", rsVal!供应商id)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rsVal!药品id), "-1", rsVal!药品id)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(strName), "", strName)
                .TextMatrix(i, .ColIndex("spec")) = IIf(IsNull(strSpec), "", strSpec)
                .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(strUnit), "", strUnit)
                .TextMatrix(i, .ColIndex("ivqty")) = IIf(IsNull(rsVal!发票数量), "0", rsVal!发票数量)
                .ColFormat(.ColIndex("ivqty")) = "#0.000"
                .TextMatrix(i, .ColIndex("pdaqty")) = IIf(IsNull(rsVal!PDA验收数量), "0", rsVal!PDA验收数量)
                .ColFormat(.ColIndex("pdaqty")) = "#0.000"
                .TextMatrix(i, .ColIndex("chkqty")) = IIf(IsNull(rsVal!已验收数量), "0", rsVal!已验收数量)
                .ColFormat(.ColIndex("chkqty")) = "#0.000"
                '读已标记的数据
                If bytMarked = 1 Then
                    .TextMatrix(i, .ColIndex("qty")) = 0 'IIf(IsNull(rsVal!验收数量), "0", rsVal!验收数量)
                Else
                    'If Val(.TextMatrix(i, .ColIndex("pdaqty"))) <= 0 Then
                    '    .TextMatrix(i, .ColIndex("qty")) = 0
                    'Else
                        .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(rsVal!PDA验收数量), "0", rsVal!PDA验收数量)
                    'End If
                End If
                .ColFormat(.ColIndex("qty")) = "#0.000"
                .ColDataType(.ColIndex("qty")) = flexDTLong
                '.TextMatrix(i, .ColIndex("price")) = IIf(IsNull(rsVal!批发价), "0", rsVal!批发价)
                
                '药库单位的成本价
                Err.Clear: On Error Resume Next
                dblCost = GetCostPrice(IIf(IsNull(rsVal!药品id), "-1", rsVal!药品id))
                If Err <> 0 Then
                    .TextMatrix(i, .ColIndex("mess")) = "“药品ID”" & Err.Description & "[外部数据库]。"
                    dblCost = 0
                End If
                .TextMatrix(i, .ColIndex("price")) = dblCost
                Err = 0: On Error GoTo errHandle
                
                .ColFormat(.ColIndex("price")) = "#0.0000"
                .TextMatrix(i, .ColIndex("producer")) = IIf(IsNull(rsVal!生产商), "", rsVal!生产商)
                .TextMatrix(i, .ColIndex("lot_no")) = IIf(IsNull(rsVal!批号), "", rsVal!批号)
                .ColDataType(.ColIndex("lot_no")) = flexDTString
                .TextMatrix(i, .ColIndex("pdate")) = IIf(IsNull(rsVal!生产日期), "", rsVal!生产日期)
                .ColFormat(.ColIndex("pdate")) = "yyyy-mm-dd"
                .TextMatrix(i, .ColIndex("avail_date")) = IIf(IsNull(rsVal!效期), "", rsVal!效期)
                .TextMatrix(i, .ColIndex("invoice")) = IIf(IsNull(rsVal!发票号), "", rsVal!发票号)
                .ColDataType(.ColIndex("invoice")) = flexDTString
                .TextMatrix(i, .ColIndex("idate")) = IIf(IsNull(rsVal!发票日期), "", rsVal!发票日期)
                .ColFormat(.ColIndex("idate")) = "yyyy-mm-dd"
                '.TextMatrix(i, .ColIndex("iamount")) = IIf(IsNull(rsVal!发票金额), "0", rsVal!发票金额)
                .TextMatrix(i, .ColIndex("iamount")) = dblCost * IIf(IsNull(rsVal!发票数量), "0", rsVal!发票数量)
                .ColFormat(.ColIndex("iamount")) = "#0.0000"
                .TextMatrix(i, .ColIndex("detail_id")) = IIf(IsNull(rsVal!detail_id), "0", rsVal!detail_id)
                .TextMatrix(i, .ColIndex("plan_code")) = IIf(IsNull(rsVal!计划单号), "", rsVal!计划单号)
                .TextMatrix(i, .ColIndex("Accepter")) = IIf(IsNull(rsVal!验收人), "", rsVal!验收人)
                
                '检查供应商ID
                If Trim(.TextMatrix(i, .ColIndex("providerid"))) = "" Then
                    .TextMatrix(i, .ColIndex("mess")) = "“供应商ID”未填写[外部数据库]。"
                    strProvider = ""
                Else
                    Err = 0: On Error Resume Next
                    strProvider = CheckProvider(Val(.TextMatrix(i, .ColIndex("providerid"))))
                    If Err <> 0 Then
                        .TextMatrix(i, .ColIndex("mess")) = "“供应商ID”" & Err.Description & "[外部数据库]。"
                        strProvider = ""
                    End If
                    Err = 0: On Error GoTo errHandle
                End If
                .TextMatrix(i, .ColIndex("provider")) = strProvider
                
                'If .TextMatrix(i, .ColIndex("providerid")) = "" Or .TextMatrix(i, .ColIndex("providerid")) = "-1" Then
                If strProvider = "" Then
                    '为不可修改提供信息
                    .TextMatrix(i, .ColIndex("provider")) = "供应商ID无"
                    .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
'                ElseIf Mid(strProvider, InStr(strProvider, "|") + 1, Len(strProvider)) = "0" Or Mid(strProvider, InStr(strProvider, "|") + 1, Len(strProvider)) = "" Then
'                    .TextMatrix(i, .ColIndex("provider")) = "未设置中标单位"
'                    .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                Else
                    'Choose可点击修改
                    If Len(Trim(strName)) = 0 Then
                        .TextMatrix(i, .ColIndex("provider")) = "药品ID无/与HIS不对应"
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                    ElseIf bytMarked = 1 And Val(.TextMatrix(i, .ColIndex("pdaqty"))) <> 0 Then
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                    ElseIf Val(.TextMatrix(i, .ColIndex("ivqty"))) <= 0 Then 'Or Val(.TextMatrix(i, .ColIndex("qty"))) <= 0 Then
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                    ElseIf Val(.TextMatrix(i, .ColIndex("ivqty"))) > Val(.TextMatrix(i, .ColIndex("qty"))) And Val(.TextMatrix(i, .ColIndex("qty"))) > 0 Then
                        '发票数量大于验收数量
                        If bytMarked = 0 And Val(.TextMatrix(i, .ColIndex("qty"))) > 0 Then
                            .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",1"
                        Else
                            .TextMatrix(i, .ColIndex("provider")) = "发票数量大于验收数量"
                            .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",1"
                        End If
                    Else
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",1"
                    End If
                End If
                
                If Mid(.TextMatrix(i, .ColIndex("imported")), 3, 1) = "1" Then
                    '能Check选
                    If Val(.TextMatrix(i, .ColIndex("qty"))) > 0 Then
                        .TextMatrix(i, .ColIndex("choose")) = IIf(Len(Trim(strName)) = 0, 0, 1)
                    End If
                Else
                    '不能Check选
                    .TextMatrix(i, .ColIndex("choose")) = 0
                    .Cell(flexcpForeColor, i, 3, i, .Cols - 1) = vbRed
                End If
                If Left(.TextMatrix(i, .ColIndex("imported")), 1) = "1" Then
                    .Cell(flexcpForeColor, i, 3, i, .ColIndex("mess")) = vbBlue
                    .TextMatrix(i, .ColIndex("choose")) = 0
                End If
                
            End If
            rsVal.MoveNext
            .ColWidth(1) = IIf(.Rows > 0, Len(Trim(Str(.Rows))) * 130 + 70, 200)
        Next
        
    End With
    Exit Sub

errHandle:
    MsgBox "控件装载数据时异常！", vbInformation, GSTR_MESSAGE
End Sub

Public Function GetMedicalInfo(ByVal intID As Long, strName As String, strSpec As String, strUnit As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "select b.名称,b.规格,a.药库单位 from 药品规格 a, 收费项目目录 b where a.药品id=[1] and b.id=[1]  and a.药品id=b.id and rownum=1 "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "", intID)
    On Error GoTo ErrHand
    If rsTmp.RecordCount = 1 Then
        strName = IIf(IsNull(rsTmp!名称), "", rsTmp!名称)
        strSpec = IIf(IsNull(rsTmp!规格), "", rsTmp!规格)
        strUnit = IIf(IsNull(rsTmp!药库单位), "", rsTmp!药库单位)
    End If
    rsTmp.Close
    GetMedicalInfo = True
    Exit Function
ErrHand:
    GetMedicalInfo = False
End Function

Public Sub RefreshTVWProvider(ByVal tvwVal As TreeView, ByVal vsfVal As VSFlexGrid)
    Dim i As Long, j As Long
    Dim blnFind As Boolean
    Dim nodTmp As Node
    Dim rsTmp As New ADODB.Recordset
    
    With rsTmp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "ID", adInteger, 18, adFldIsNullable
        .Fields.Append "Name", adVarChar, 50, adFldIsNullable
        .Open
    End With
    With tvwVal
        .Nodes.Clear
        .Nodes.Add , , "Root", "全部"
        .Nodes(1).Checked = True
        .Nodes(1).Expanded = True
    End With
    '保存到RecordSet中
    With vsfVal
        For i = 1 To .Rows - 1
            blnFind = False
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                If rsTmp!Id = Val(.TextMatrix(i, .ColIndex("providerid"))) Then
                    blnFind = True
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If blnFind = False And .TextMatrix(i, .ColIndex("imported")) <> "0,0" Then
                rsTmp.AddNew
                rsTmp!Id = Val(.TextMatrix(i, .ColIndex("providerid")))
                rsTmp!Name = .TextMatrix(i, .ColIndex("provider"))
            End If
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("imported")) = "0,0" Then
                rsTmp.AddNew
                rsTmp!Id = -1
                rsTmp!Name = "错误记录"
                Exit For
            End If
        Next
    End With
    '排序
    With rsTmp
        .Sort = "Name"
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Set nodTmp = tvwVal.Nodes.Add("Root", tvwChild, "K" & !Id, !Name)
            nodTmp.Tag = !Id
            nodTmp.Checked = True
            .MoveNext
        Loop
    End With
End Sub

Public Function CheckRecord(ByVal vsfVal As VSFlexGrid) As Boolean
    Dim i As Integer
    With vsfVal
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False And Val(.TextMatrix(i, .ColIndex("choose"))) <> 0 Then
                CheckRecord = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function GetCostPrice(ByVal lngID As Long) As Double
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    strTmp = "Select nvl(成本价,0) * 药库包装 成本价 From 药品规格 Where 药品id=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strTmp, "取药品规格的成本价", lngID)
    If Not rsTmp.EOF Then
        GetCostPrice = rsTmp!成本价
    End If
    rsTmp.Close
End Function
