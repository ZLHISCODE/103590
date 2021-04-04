VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmQCAddSample1 
   Caption         =   "质控标本"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   Icon            =   "frmQCAddSample1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11235
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt样本号 
      Height          =   300
      Left            =   6900
      TabIndex        =   4
      Top             =   570
      Width           =   1770
   End
   Begin VB.ComboBox cbo质控品 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   3000
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgSample 
      Height          =   4500
      Left            =   270
      TabIndex        =   0
      Top             =   1410
      Width           =   10785
      _cx             =   19024
      _cy             =   7937
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483632
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.ComboBox cbo仪器 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   3200
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   300
      Left            =   4950
      TabIndex        =   3
      Top             =   600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   98566147
      CurrentDate     =   39590
      MaxDate         =   401769
      MinDate         =   2
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   180
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQCAddSample1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private marrData() As String  '保存原始数据
Private mstrPriv As String

Private Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
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
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 350
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

Private Sub initCbsThis(cbsMain As CommandBars)
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)  '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")  '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")

        'Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "放弃(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&P)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有

    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

    '查找项特殊处理
    '-----------------------------------------------------
'    主菜单右侧的查找 按就诊卡号查找，支持刷卡
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Dept, "仪器")
        objControl.ID = conMenu_View_Dept
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Dept + 1, "")
        objCustom.Handle = cbo仪器.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
        Set objControl = .Add(xtpControlLabel, conMenu_View_FindType, "质控品")
        objControl.ID = conMenu_View_FindType
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = cbo质控品.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "放弃")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
        
        Set objControl = .Add(xtpControlLabel, conMenu_EditPopup + 1, "日期")
        objControl.ID = conMenu_EditPopup + 1
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_EditPopup + 2, "")
        objCustom.Handle = dtpStart.hWnd
        objCustom.Flags = xtpFlagRightAlign
        

        
        Set objControl = .Add(xtpControlLabel, conMenu_EditPopup + 7, "样本号")
        objControl.ID = conMenu_EditPopup + 7
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_EditPopup + 8, "")
        objCustom.Handle = txt样本号.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings

        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印

        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet         '打印设置
        .AddHiddenCommand conMenu_File_Excel            '输出到Excel
    End With

    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
'    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)

End Sub
Private Sub reSetHead()
    '初始化vfgSample控件标题
    Dim strHead As String, lng次数 As Long
    Dim i As Integer
    
    lng次数 = 9
    For i = 1 To lng次数
        strHead = strHead & ";第" & i & "次,900,7"
    Next
    
    strHead = "检验项目,1200,1" & strHead & ";项目ID,0,1;类型,0,1;序列,0,1"
    Call SetVsFlexGridHead(strHead, vfgSample)

End Sub

Private Sub RefreshData()
    Dim lng质控ID As Long, int检验次数 As Integer
    Dim dateStart As Date, dateEnd As Date
    Dim i As Integer
    
    Dim strsql As String
    Dim rsTmp As ADODB.Recordset
    dateStart = Format(dtpStart.Value, "yyyy-MM-dd")
    dateEnd = dateStart + 1
    
    Call reSetHead
    ReDim marrData(vfgSample.Rows, vfgSample.Cols)
    
    If cbo质控品.ListIndex < 0 Then Exit Sub
    
    lng质控ID = cbo质控品.ItemData(cbo质控品.ListIndex)
    If lng质控ID <= 0 Then Exit Sub

    '------------- 读数据
    Dim intRow As Integer, intFindRow As Integer
    
    On Error GoTo ErrHandle
    strsql = "Select 标本号 From 检验质控品 where id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng质控ID)
    Do Until rsTmp.EOF
        If "" & rsTmp!标本号 <> "" Then
            txt样本号 = "" & rsTmp!标本号
        End If
        rsTmp.MoveNext
    Loop
    
    '--- 加空白入项目
        
    
    strsql = "Select A.质控品id, A.项目id, A.取值序列, A.序列值, E.结果类型, F.编码, F.中文名, E.缩写" & vbNewLine & _
            "From 检验质控品项目 A, 检验项目 E, 诊治所见项目 F" & vbNewLine & _
            "Where A.项目id = E.诊治项目id And A.项目id = F.ID And A.质控品id = [1]" & vbNewLine & _
            "Order By F.编码"

    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng质控ID)
    With vfgSample
        Do Until rsTmp.EOF
           
            .TextMatrix(.Rows - 1, .ColIndex("项目ID")) = "" & rsTmp!项目id
            .TextMatrix(.Rows - 1, .ColIndex("检验项目")) = "" & rsTmp!中文名 & " " & rsTmp!缩写
            .TextMatrix(.Rows - 1, .ColIndex("类型")) = "" & rsTmp!结果类型
            .TextMatrix(.Rows - 1, .ColIndex("序列")) = "" & rsTmp!取值序列
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If Trim(.TextMatrix(.Rows - 1, 0)) = "" Then .Rows = .Rows - 1
        
        ReDim marrData(vfgSample.Rows, vfgSample.Cols)
    End With
    
    '---- 取具体数据
    strsql = "Select A.*, E.结果类型, F.编码, F.中文名, E.缩写, D.检验结果, T.标记" & vbNewLine & _
            "From (Select A.质控品id, A.项目id, B.标本id, B.检验时间, A.取值序列, A.序列值, B.测试次数" & vbNewLine & _
            "       From 检验质控品项目 A, 检验质控记录 B" & vbNewLine & _
            "       Where A.质控品id = B.质控品id(+) And A.质控品id = [1] And" & vbNewLine & _
            "             B.检验时间(+) Between [2] And [3] And B.测试次数(+) between 1 and 9) A," & vbNewLine & _
            "     检验普通结果 D, 检验项目 E, 诊治所见项目 F,检验质控报告 T" & vbNewLine & _
            "Where D.ID=T.结果ID(+) And A.标本id = D.检验标本id And A.项目id = D.检验项目id And A.项目id = E.诊治项目id And A.项目id = F.ID" & vbNewLine & _
            "Order By A.检验时间, F.编码"
    dateEnd = dateEnd + 1
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng质控ID, CDate(dateStart), CDate(dateEnd))
    
    With vfgSample
        'intRow = .FixedRows
        Do Until rsTmp.EOF
            intFindRow = .FindRow("" & rsTmp!项目id, .FixedRows, .ColIndex("项目ID"))
            If intFindRow > 0 Then
                intRow = intFindRow
            Else
                intRow = .Rows - 1
            End If
            .TextMatrix(intRow, .ColIndex("项目ID")) = "" & rsTmp!项目id
            .TextMatrix(intRow, .ColIndex("检验项目")) = "" & rsTmp!中文名 & " " & rsTmp!缩写

            
            For i = 1 To .Cols - 1
                If Val("" & rsTmp!测试次数) = Val(Mid(.TextMatrix(0, i), 2)) Then
                    .TextMatrix(intRow, i) = "" & rsTmp!检验结果
                    marrData(intRow, i) = "" & rsTmp!检验结果 & "|" & rsTmp!标本ID
                    If Val("" & rsTmp!标记) = 2 Then
                        .Cell(flexcpForeColor, intRow, i) = vbRed
                    ElseIf Val("" & rsTmp!标记) = 0 Then
                        .Cell(flexcpForeColor, intRow, i) = .ForeColor
                    Else
                        .Cell(flexcpForeColor, intRow, i) = vbMagenta
                    End If
                                        
                    Exit For
                End If
            Next
            If Not (intFindRow = intRow And intFindRow > 0) Then
                intRow = intRow + 1
                .Rows = .Rows + 1
            End If
            rsTmp.MoveNext
        Loop
        If Trim(.TextMatrix(.Rows - 1, 0)) = "" Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize .ColIndex("检验项目")
        '.FrozenCols = 1
        .AllowUserFreezing = flexFreezeColumns
        
        .Editable = flexEDNone
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub

Private Sub Load仪器()
    Dim strsql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If InStr(1, mstrPriv, "所有科室") > 0 Then
        strsql = " Select Distinct  a.id,a.编码 , a.名称  From 检验仪器 a ,部门表 b,检验质控品 c " & _
                  "Where a.使用小组ID = b.ID and a.id = c.仪器id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName)
        
    Else
        strsql = " Select Distinct a.id,a.编码 , a.名称  From 部门人员 D,检验仪器 a ,部门表 b , 检验质控品 c " & _
                  " Where a.使用小组ID = b.ID and a.使用小组id=D.部门id and D.人员id = [1]  " & _
                  " and a.id = c.仪器Id "
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, UserInfo.ID)
    End If
    
    cbo仪器.Clear
    Do Until rsTemp.EOF
        cbo仪器.AddItem "" & rsTemp!编码 & " " & rsTemp!名称
        cbo仪器.ItemData(cbo仪器.NewIndex) = rsTemp!ID
        rsTemp.MoveNext
    Loop
    If cbo仪器.ListCount > 0 Then cbo仪器.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SaveData()
    '保存数据
    Dim intRow As Integer, intCol As Integer
    Dim strData As String, strOLDdata As String
    Dim lng项目ID As Long, str日期 As String
    Dim lng标本ID As Long, lng次数 As Long
    Dim bln提示标本号 As Boolean
    Dim strNewItem As String '保存新增检验项目
    bln提示标本号 = False
    
    For intCol = vfgSample.ColIndex("检验项目") + 1 To vfgSample.ColIndex("项目ID") - 1
        strNewItem = ""
        str日期 = Format(dtpStart, "yyyy-MM-dd")
        lng标本ID = 0
        lng次数 = Val(Mid(vfgSample.TextMatrix(0, intCol), 2))
        For intRow = 1 To vfgSample.Rows - 1
            strData = vfgSample.TextMatrix(intRow, intCol)
            strOLDdata = marrData(intRow, intCol)
            
            If InStr(strOLDdata, "|") > 0 Then
                lng标本ID = Split(strOLDdata, "|")(1)
            End If
            
            If strOLDdata <> "" Then
                If strData <> Split(strOLDdata, "|")(0) Then
                    '要保存
                    If InStr(strOLDdata, "|") > 0 Then
                        '有原始记录
                        
                        lng项目ID = Val(vfgSample.TextMatrix(intRow, vfgSample.ColIndex("项目ID")))
                        strNewItem = strNewItem & "|" & lng项目ID & "^" & strData
                    Else
                        If vfgSample.TextMatrix(intRow, intCol - 1) = "" Then
                            MsgBox "第" & intRow & "行的数据不连续，请处理!", vbQuestion, Me.Caption
                            Exit Sub
                        End If
                        '新增
                        lng项目ID = Val(vfgSample.TextMatrix(intRow, vfgSample.ColIndex("项目ID")))
                        strNewItem = strNewItem & "|" & lng项目ID & "^" & strData
                        If Val(txt样本号) = 0 Then bln提示标本号 = True
                    End If
                End If
            Else

                If strData <> "" Then
                    '新增
                    If vfgSample.TextMatrix(intRow, intCol - 1) = "" Then
                        MsgBox "第" & intRow & "行的数据不连续，请处理!", vbQuestion, Me.Caption
                        Exit Sub
                    End If
                    lng项目ID = Val(vfgSample.TextMatrix(intRow, vfgSample.ColIndex("项目ID")))
                    strNewItem = strNewItem & "|" & lng项目ID & "^" & strData
                    If Val(txt样本号) = 0 Then bln提示标本号 = True
                End If
            End If
        Next
        If bln提示标本号 Then
            MsgBox "请填写标本号！", vbInformation, Me.Caption
            Exit Sub
        End If
        If strNewItem <> "" Then
            strNewItem = Mid(strNewItem, 2)
            lng标本ID = Edit_Sample(lng标本ID, lng次数, str日期, strNewItem)
        End If
    Next
    
    Call RefreshData
End Sub

Private Function Edit_Sample(ByVal lngID_in As Long, ByVal lng次数 As Long, _
                        ByVal str日期 As String, ByVal strItemRecords As String) As Long
    '增加质控标本
    Dim lngID As Long       '标本id
    Dim lngDeviceID As Long '仪器id
    Dim strSampleNO As String '标本号
    Dim lngQCID As Long '质控品ID
    
    Dim blnTrans As Boolean '是否开始事务
    On Error GoTo ErrHandle
    
    If lngID_in = 0 Then
        lngID = zlDatabase.GetNextId("检验标本记录")
    Else
        lngID = lngID_in
    End If
    
    strSampleNO = Val(txt样本号) + lng次数 - 1
    lngDeviceID = cbo仪器.ItemData(cbo仪器.ListIndex)
    lngQCID = cbo质控品.ItemData(cbo质控品.ListIndex)
    
'    gcnOracle.BeginTrans
'    blnTrans = True
    If lngID_in = 0 Then
        gstrSql = "ZL_检验标本记录_INSERT(" & lngID & ",NULL,'" & _
            strSampleNO & "',NULL,NULL," & lngDeviceID & ",NULL," & _
            "To_Date('" & str日期 & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
            "To_Date('" & str日期 & "','yyyy-mm-dd hh24:mi:ss'),'" & UserInfo.姓名 & "'," & _
            "Null,To_Date('" & str日期 & "','yyyy-mm-dd hh24:mi:ss'),'" & gstrUserName & "','0',Null,0,Null)"
        zlDatabase.ExecuteProcedure gstrSql, "插入检验临时记录"
    End If
    
    gstrSql = "ZL_检验普通结果_BATCHUPDATE(" & lngID & "," & _
        lngDeviceID & ",Null,Null,Null,'" & strItemRecords & "')"
    zlDatabase.ExecuteProcedure gstrSql, "检验结果报告"
    gstrSql = "Zl_重新计算结果_Cale(" & lngID & ")"
    zlDatabase.ExecuteProcedure gstrSql, "检验结果报告"
    
    gstrSql = "ZL_检验质控记录_EDIT(1," & lngID & "," & lngQCID & ")"
    zlDatabase.ExecuteProcedure gstrSql, "保存为质控品"
    
'    gcnOracle.CommitTrans
    Edit_Sample = lngID
    blnTrans = False
    Exit Function
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'------ 以下是控件过程

Private Sub cbo检验次数_Click()

    Call RefreshData

End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Me.vfgSample.Editable = flexEDKbdMouse
    Case conMenu_Edit_Untread
        Call RefreshData
    Case conMenu_Edit_Save
        Call SaveData
    Case conMenu_View_Refresh
        Call RefreshData
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Control.Enabled = Not (Me.vfgSample.Editable = flexEDKbdMouse)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (Me.vfgSample.Editable = flexEDKbdMouse)
    End Select
End Sub

Private Sub dtpStart_Change()
    Call RefreshData
End Sub

Private Sub Form_Load()
    
    Call initCbsThis(cbsThis)
    
    '设日期及检验次数
    dtpStart = Now
    Call reSetHead

    
    Call Load仪器
    ReDim marrData(vfgSample.Rows, vfgSample.Cols)
End Sub

Private Sub Form_Resize()
    Call cbsThis_Resize
End Sub

Private Sub cbo质控品_Click()
    Call RefreshData
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With vfgSample
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbo仪器_Click()
    Dim strsql As String
    Dim rsTmp As ADODB.Recordset
    Dim lng仪器id As Long
    Dim dateStart As Date
    Dim dateEnd As Date
    
    On Error GoTo ErrHandle
    If cbo仪器.ListIndex < 0 Then Exit Sub
    
    lng仪器id = cbo仪器.ItemData(cbo仪器.ListIndex)
    dateStart = Format(dtpStart.Value, "yyyy-MM-dd")
    dateEnd = dateStart + 1
    strsql = "Select ID,名称,批号,浓度,水平 From 检验质控品 Where [2] between 开始日期 and 结束日期 and [3] between 开始日期 and　结束日期 and 仪器ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng仪器id, dateStart, dateEnd)
    cbo质控品.Clear
    Do Until rsTmp.EOF
        cbo质控品.AddItem "" & rsTmp!名称 & " " & rsTmp!批号 & " 水平:" & rsTmp!水平
        cbo质控品.ItemData(cbo质控品.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Loop
    If cbo质控品.ListCount > 0 Then cbo质控品.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowMe(ByVal strPrivate As String, ByVal frmMain As Form)
    mstrPriv = strPrivate
    
    Me.Show vbModal, frmMain
End Sub

Private Sub vfgSample_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strLists As String, strValue As String
    Dim lngCount As Long
    
    If Col = 0 Then Exit Sub
    If Trim(Me.vfgSample.TextMatrix(Row, Col)) = "" Then Exit Sub
    
    strLists = Trim(Me.vfgSample.TextMatrix(Row, vfgSample.ColIndex("序列")))
    strValue = Trim(Me.vfgSample.TextMatrix(Row, Col))
    
    If strLists = "" Then
        If InStr(strValue, "E+") > 0 And Val(strValue) > 0 Then
            Me.vfgSample.TextMatrix(Row, Col) = strValue
        Else
            Me.vfgSample.TextMatrix(Row, Col) = Format(Val(strValue), "0.00")
        End If
        
        Exit Sub
    End If
    For lngCount = 0 To UBound(Split(strLists, ";"))
        If vfgSample = Split(strLists, ";")(lngCount) Then Exit Sub
    Next
    Me.vfgSample.TextMatrix(Row, Col) = ""
    
    strValue = "该项目为半定量项目，需符合取值序列(" & strLists & ")要求！"
    MsgBox strValue, vbInformation, gstrSysName
End Sub

Private Sub vfgSample_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub
