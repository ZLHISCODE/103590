VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMiningVessels 
   Caption         =   "采血管设置"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   Icon            =   "frmMiningVessels.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   11700
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   4950
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   635
      SimpleText      =   $"frmMiningVessels.frx":000C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMiningVessels.frx":0053
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   3600
      Left            =   45
      TabIndex        =   0
      Top             =   645
      Width           =   7860
      _cx             =   13864
      _cy             =   6350
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.TextBox txtEdit 
         Height          =   375
         Left            =   5865
         TabIndex        =   2
         Top             =   435
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin XtremeCommandBars.ImageManager imgICON 
      Left            =   915
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMiningVessels.frx":08E7
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   300
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMiningVessels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' API declares
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'公共部份菜单ID定义:*表示有图标
'*********************************************************************
Private Const mconMenu_FilePopup = 1 '文件
Private Const mconMenu_ManagePopup = 2 '管理
Private Const mconMenu_EditPopup = 3 '编辑
Private Const mconMenu_ReportPopup = 4 '报表
Private Const mconMenu_ViewPopup = 7 '查看
Private Const mconMenu_ToolPopup = 8 '工具
Private Const mconMenu_HelpPopup = 9 '帮助

'文件菜单
Private Const mconMenu_File_Open = 100            '*打开(&O)…
Private Const mconMenu_File_PrintSet = 101        '*打印设置(&S)…
Private Const mconMenu_File_Preview = 102         '*预览(&V)
Private Const mconMenu_File_Print = 103           '*打印(&P)
Private Const mconMenu_File_Excel = 104           '输出到&Excel…

Private Const mconMenu_File_MedRecSetup = 1051        '打印设置(&S)
Private Const mconMenu_File_MedRecPreview = 1052      '打印预览(&P)

Private Const mconMenu_File_RowPrint = 121        '记录打印(&R)

Private Const mconMenu_File_Exit = 191            '*退出(&X)

'编辑

Private Const mconMenu_Manage_Append = 3001     '*增加(&Y)
Private Const mconMenu_Manage_Delete = 3004     '*删除(&D)
Private Const mconMenu_Manage_Modify = 3003       '*修改(&M)
Private Const mconMenu_Manage_ModifyNo = 228       '*修改编码(&M)
Private Const mconMenu_Manage_deleCos = 21205       '*删除材料

Private Const mconMenu_Manage_Stop = 3503     '*保存(&C)
Private Const mconMenu_Manage_Cancle = 3014   '取消

'查看菜单
Private Const mconMenu_View_ToolBar = 701              '工具栏(&T)
Private Const mconMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Private Const mconMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Private Const mconMenu_View_ToolBar_Size = 7013           '大图标(&B)
Private Const mconMenu_View_StatusBar = 702            '状态栏(&S)
Private Const mconMenu_View_Append = 703               '附加信息(&A)
Private Const mconMenu_View_Expend = 711               '展开/折叠组(&X)
Private Const mconMenu_View_Expend_CurCollapse = 7111     '折叠当前组(&C)
Private Const mconMenu_View_Expend_CurExpend = 7112       '展开当前组(&E)
Private Const mconMenu_View_Expend_AllCollapse = 7113     '折叠所有组(&L)
Private Const mconMenu_View_Expend_AllExpend = 7114       '展开所有组(&X)
Private Const mconMenu_View_Find = 721                 '*查找(&F)
Private Const mconMenu_View_FindNext = 722             '继续查找(&N)
Private Const mconMenu_View_FindType = 723             '查找方式(&Y)
Private Const mconMenu_View_Filter = 731               '*数据过滤(&I),子窗体的过滤功能
Private Const mconMenu_View_Notify = 732               '*医嘱提醒(&B)
Private Const mconMenu_View_Busy = 733                 '诊室忙(&M)
Private Const mconMenu_View_Hide = 741                 '*隐藏(&H)
Private Const mconMenu_View_Show = 742                 '*显示(&S)
Private Const mconMenu_View_Backward = 743             '*后退(&B)
Private Const mconMenu_View_Forward = 744              '*前进(&F)
Private Const mconMenu_View_Option = 781               '选项(&O)
Private Const mconMenu_View_Refresh = 791              '*刷新(&R)
Private Const mconMenu_View_Jump = 792                 '跳转(&J)

'帮助菜单
Private Const mconMenu_Help_Help = 901        '*帮助主题(&H)
Private Const mconMenu_Help_Web = 902         '&WEB上的中联
Private Const mconMenu_Help_Web_Home = 9021       '中联主页(&H)
Private Const mconMenu_Help_Web_Forum = 9023   '中联论坛(&F)
Private Const mconMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Private Const mconMenu_Help_About = 991       '关于(&A)…

'其它常量定义
'*********************************************************************
'CommandBar固有常量定义
Private Const mXTP_ID_WINDOW_LIST = 35000 '窗体列表
Private Const mXTP_ID_TOOLBARLIST = 59392 '工具栏列表
Private Const mID_INDICATOR_CAPS = 59137 '状态栏（大写）
Private Const mID_INDICATOR_NUM = 59138 '状态栏（数字）
Private Const mID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Private Const mFSHIFT = 4
Private Const mFCONTROL = 8
Private Const mFALT = 16

Private mlngModul As Long, mstrPrivs As String

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    Dim strNo As String, strSql As String, strNewNo As String
    
    Select Case Control.ID
    Case mconMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case mconMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case mconMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case mconMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case mconMenu_Help_Help '帮助
        Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case mconMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case mconMenu_File_Exit '退出
        Unload Me
        
    Case mconMenu_Manage_Modify  '修改
        vfgList.SelectionMode = flexSelectionFree
        vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_ModifyNo '修改编码
        If vfgList.Row > 0 Then
            
            strNo = Trim(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("Modify")))
            strNewNo = InputBox("将编码" & strNo & "改为：" & vbNewLine & "(注：修改编码后，检验项目中的管码将同步调整。)")
            If strNewNo <> "" And strNo <> "" Then
                With vfgList
                    For i = .FixedRows To .Rows - 1
                        If strNewNo = .TextMatrix(i, .ColIndex("Modify")) Then
                            MsgBox strNewNo & "与现有编码重复，请输入另一个编码！"
                            Exit Sub
                        End If
                    Next
                End With
            
                strSql = "Zl_采血管类型_Clear(1,'" & strNo & "','" & strNewNo & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "修改管码")
                Call LoadData
            End If
        End If
    Case mconMenu_Manage_deleCos
        With vfgList
            .TextMatrix(.Row, .ColIndex("对应材料")) = ""
            .TextMatrix(.Row, .ColIndex("材料ID")) = ""
        End With
    Case mconMenu_Manage_Stop '保存

        If SaveData = True Then
            vfgList.Editable = flexEDNone
            vfgList.SelectionMode = flexSelectionByRow
            Call LoadData
        End If
    Case mconMenu_Manage_Cancle  '取消
        vfgList.Editable = flexEDNone
        vfgList.SelectionMode = flexSelectionByRow
        Call LoadData
    Case mconMenu_Manage_Append '增加
        With vfgList
            .SelectionMode = flexSelectionFree
            .Editable = flexEDKbdMouse
            .Rows = .Rows + 1
            If .Rows > 1 Then
                If Val(.TextMatrix(.Rows - 2, .ColIndex("编码"))) <> 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("编码")) = Val(.TextMatrix(.Rows - 2, .ColIndex("编码"))) + 1
                End If
            End If
            .Cell(flexcpFloodColor, .Rows - 1, .ColIndex("颜色")) = vbWhite
            .Cell(flexcpFloodPercent, .Rows - 1, .ColIndex("颜色")) = 100
            .Select .Rows - 1, .ColIndex("名称")
            
        End With
    Case mconMenu_Manage_Delete '删除
    
        If vfgList.Row > 0 Then
            '删除前检查是否已使用，使用了则清空。
            If MsgBox("提示：删除该项目后，检验项目中对应的管码设置将清空，是否继续！", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                strNo = Trim(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("Modify")))
                If strNo <> "" Then
                    strSql = "Zl_采血管类型_Clear(1,'" & strNo & "',Null)"
                    Call zlDatabase.ExecuteProcedure(strSql, "清除管码")
                End If
                Call vfgList.RemoveItem(vfgList.Row)
            End If
        End If
    End Select

End Sub

Private Function SaveData() As Boolean
    '保存数据
    Dim i As Integer, iRow As Integer
    Dim strSql() As String, str编码 As String, str名称 As String
    Dim lngColor As Long '-214748363
    On Error GoTo errHandle
    
    
    With vfgList
        
        If .Rows > 1 Then
            ReDim strSql(.Rows - 1)
        End If
        For i = 1 To .Rows - 1
            str编码 = Replace(Trim(.TextMatrix(i, .ColIndex("编码"))), "'", "")
            If str编码 = "" Then
                MsgBox "编码不能为空！", vbInformation, gstrSysName
                Exit Function
            End If
            If IsNumeric(str编码) = False Then
                MsgBox "编码不能为字符！", vbInformation, gstrSysName
                Exit Function
            End If
            str名称 = Replace(Trim(.TextMatrix(i, .ColIndex("名称"))), "'", "")
            If str名称 = "" Then
                MsgBox "名称不能为空！", vbInformation, gstrSysName
                Exit Function
            End If
            
            '检查编码是否有重复
            For iRow = 1 To .Rows - 1
                If i <> iRow Then
                    If str编码 = Replace(Trim(.TextMatrix(iRow, .ColIndex("编码"))), "'", "") Then
                        MsgBox "编码[" & str编码 & "]重复，请调整！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
            
            lngColor = Val(.Cell(flexcpFloodColor, i, .ColIndex("颜色")))
            If lngColor = -214748363 Then lngColor = 0
            strSql(i) = "Zl_采血管类型_Update('" & str编码 & "','" & str名称 & "','" & Replace(Trim(.TextMatrix(i, .ColIndex("规格"))), "'", "") & "'," & _
                        "'" & Replace(Trim(.TextMatrix(i, .ColIndex("添加剂"))), "'", "") & "','" & Replace(Trim(.TextMatrix(i, .ColIndex("采血量"))), "'", "") & "'," & _
                        lngColor & "," & Val(.TextMatrix(i, .ColIndex("材料ID"))) & ")"
        Next
        
        If vfgList.Rows > 1 Then
            'gstrSql = "Zl_采血管类型_Clear"
            gcnOracle.BeginTrans
            'Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            For i = LBound(strSql) To UBound(strSql)
                If strSql(i) Like "Zl_采血管类型_Update*" Then
                    Call zlDatabase.ExecuteProcedure(strSql(i), Me.Caption)
                End If
            Next
            gcnOracle.CommitTrans
            SaveData = True
        End If
    End With
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.vfgList
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top - stbThis.Height
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case mconMenu_Manage_Stop, mconMenu_Manage_deleCos '保存
        Control.Enabled = vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_Cancle '取消
        Control.Enabled = vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_ModifyNo, mconMenu_Manage_Delete ', mconMenu_Manage_deleCos
        Control.Enabled = Not (vfgList.Editable = flexEDKbdMouse)
     
    End Select
End Sub

Private Sub Form_Load()

    '菜单工具栏
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
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
    cbsMain.Icons = imgICON.Icons
    Call initMenus
    Call LoadData
    
End Sub

Private Sub initVfgList()
    Dim strHead As String
    '1 左对齐 4 居中 7 右对齐
    strHead = "编码,600,4;名称,1500,1;规格,1500,1;添加剂,1800,1;采血量,1500,1;颜色,600,1;对应材料,2800,1;材料ID,0,1;Modify,0,1"
    Call SetVsFlexGridHead(strHead, vfgList)
    
End Sub

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
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

Private Sub LoadData()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    strSql = "Select A.编码,A.名称,A.规格,A.添加剂,A.采血量,A.颜色,A.材料ID,B.编码||' '||B.名称 as 对应材料 From 采血管类型 A,收费项目目录 B Where A.材料ID=B.ID(+) Order by A.编码"
    Dim lngColor As Long
    
    On Error GoTo errHandle
    With vfgList
        .Clear
        Call initVfgList
        '.ColComboList(.ColIndex("颜色")) = "..."
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        Do Until rsTmp.EOF
             
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = Nvl(rsTmp.Fields("编码"))
            .TextMatrix(.Rows - 1, .ColIndex("名称")) = Nvl(rsTmp.Fields("名称"))
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = Nvl(rsTmp.Fields("规格"))
            .TextMatrix(.Rows - 1, .ColIndex("添加剂")) = Nvl(rsTmp.Fields("添加剂"))
            .TextMatrix(.Rows - 1, .ColIndex("采血量")) = Nvl(rsTmp.Fields("采血量"))
            .TextMatrix(.Rows - 1, .ColIndex("Modify")) = Nvl(rsTmp.Fields("编码"))
            
            .Cell(flexcpFloodPercent, .Rows - 1, .ColIndex("颜色")) = 100
            lngColor = Val(Nvl(rsTmp.Fields("颜色")))
            If lngColor = 0 Then lngColor = -214748363
            .Cell(flexcpFloodColor, .Rows - 1, .ColIndex("颜色")) = lngColor
            
            .TextMatrix(.Rows - 1, .ColIndex("材料ID")) = Nvl(rsTmp.Fields("材料ID"))
            .TextMatrix(.Rows - 1, .ColIndex("对应材料")) = Trim(Nvl(rsTmp.Fields("对应材料")))
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        If .Rows > 0 Then
            .Rows = .Rows - 1
        End If
        '行选择
        .SelectionMode = flexSelectionByRow
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    vfgList.TextMatrix(Row, vfgList.ColIndex("Modify")) = "Update"
End Sub

Private Sub vfgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEdit As Boolean
    Call vfgList_BeforeEdit(NewRow, NewCol, blnEdit)
    If blnEdit Then
        vfgList.ComboList = ""
        'vsList.FocusRect = flexFocusLight
    Else
        'vsList.FocusRect = flexFocusSolid
        If NewCol = vfgList.ColIndex("对应材料") Then
            vfgList.ComboList = "..."
        Else
            vfgList.ComboList = ""
        End If
    End If
End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgList.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
    If Col = vfgList.ColIndex("颜色") Then
        Cancel = True
    ElseIf Col = vfgList.ColIndex("编码") Then
        If vfgList.TextMatrix(Row, vfgList.ColIndex("Modify")) <> "" Then
            '已使用的编码不能改
            Cancel = True
        End If
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_DblClick()

    Dim pt As POINTAPI
    
    With vfgList
        If .Editable = flexEDNone Then Exit Sub
        
        If .MouseCol = .ColIndex("颜色") Then
            pt.x = .ColPos(.MouseCol) \ Screen.TwipsPerPixelX
            pt.y = (.RowPos(.MouseRow) + .RowHeight(.MouseRow)) \ Screen.TwipsPerPixelY
            ClientToScreen .hWnd, pt
            
            frmSelColor.lblRow = .MouseRow
            frmSelColor.lblCol = .MouseCol
            frmSelColor.lngColor = .Cell(flexcpFloodColor, .MouseRow, .MouseCol)
            frmSelColor.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
            frmSelColor.Show vbModal, Me
        End If
    End With
End Sub

Private Sub initMenus()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = mconMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_Excel, "输出到&Excel…")

        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ManagePopup, "编辑(&E)", -1, False)
    objMenu.ID = mconMenu_ManagePopup
    With objMenu.CommandBar.Controls
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "保存(&S)")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "取消(&C)")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "增加(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "删除(&D)")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_ModifyNo, "修改编码(&N)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_deleCos, "删除材料"): objControl.BeginGroup = True
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = mconMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, mconMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, mconMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, mconMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, mconMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, mconMenu_View_StatusBar, "状态栏(&S)") '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = mconMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "帮助主题(&H)") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, mconMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, mconMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, mconMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, mconMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "预览") '固有
        Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "打印") '固有

        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "取消")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_ModifyNo, "修改编码"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_deleCos, "删除材料"): objControl.BeginGroup = True
        
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "增加"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "修改")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "删除")

        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出") '固有
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, mconMenu_File_Print '打印
        .Add 0, vbKeyF1, mconMenu_Help_Help '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet '打印设置
        .AddHiddenCommand mconMenu_File_Excel '输出到Excel
    End With

    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vfgListButtonClick(Row, Col)
End Sub

Private Sub vfgList_EnterCell()
    On Error GoTo errHandle
    With vfgList
    
        If .Col = .ColIndex("对应材料") And .Row > 0 Then
            If txtEdit.Tag = "False" Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
                txtEdit.Tag = "True"
            End If
        Else
            txtEdit.Tag = "False"
        End If
        
        Dim blnCancle As Boolean
        Call vfgList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
 With vfgList
    If (.Col = .ColIndex("对应材料")) And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    Else
        If .Col = .ColIndex("对应材料") And vfgList.ComboList = "..." Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                txtEdit.Text = .EditText
                Call vfgList_CellButtonClick(.Row, .Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End If
 End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vfgList
            If Col = .ColIndex("对应材料") Then
                txtEdit.Text = .EditText
                .EditText = ""
                Call vfgListButtonClick(Row, Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            
            ElseIf Col + 1 > .Cols - 3 Then
                If Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    
                    If Val(.TextMatrix(Row, .ColIndex("编码"))) <> 0 Then
                        .TextMatrix(Row + 1, .ColIndex("编码")) = Val(.TextMatrix(Row, .ColIndex("编码"))) + 1
                    End If
                End If
                .Cell(flexcpFloodColor, .Rows - 1, .ColIndex("颜色")) = vbWhite
                .Cell(flexcpFloodPercent, .Rows - 1, .ColIndex("颜色")) = 100
                .Select Row + 1, .ColIndex("名称")
                
            Else
                .Select Row, Col + 1
            End If
        End With
    End If
End Sub


Private Sub vfgList_LeaveCell()
    Dim blnCancle As Boolean
    
    With vfgList
        Call vfgList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            On Error Resume Next
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vfgList_RowColChange()
    On Error GoTo errHandle
    With vfgList
        If txtEdit.Tag = "True" Then
            txtEdit.Left = .CellLeft
            txtEdit.Top = .CellTop
            txtEdit.Height = .CellHeight - 12
            txtEdit.Width = .CellWidth - 12
        End If
    End With
    Exit Sub
errHandle:
    If err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vfgListButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql   As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean
    Dim i As Integer
    On Error GoTo errHandle
    
    If Col = vfgList.ColIndex("对应材料") Then
        '提取材料
        '--------------------------------------------------------------------------------------
            strInput = DelInvalidChar(UCase(Trim(txtEdit)))
            If InStr(strInput, " ") > 0 Then
                strInput = Trim(Split(strInput, " ")(0))
            End If
            If strInput = "" Then
                strSql = "Select A.ID, A.编码, A.名称, A.规格, A.计算单位, B.现价, A.费用类型," & vbNewLine & _
                        "       Decode(A.服务对象, 1, '门诊', 2, '住院', '门诊和住院') As 服务对象" & vbNewLine & _
                        "From (Select 现价, 收费细目id From 收费价目 Where (终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                        IIf(gstrPriceClass = "", " And 价格等级 Is Null ", " And 价格等级 = [1] ") & ") B," & vbNewLine & _
                        "     收费项目目录 A" & vbNewLine & _
                        "Where A.ID = B.收费细目id And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.类别 = '4'"
            Else
                strSql = "Select /*+ rule */" & vbNewLine & _
                    " A.ID, A.编码, A.名称, A.规格, A.计算单位, B.现价, A.费用类型," & vbNewLine & _
                    " Decode(A.服务对象, 1, '门诊', 2, '住院', '门诊和住院') As 服务对象" & vbNewLine & _
                    "From 收费项目别名 E," & vbNewLine & _
                    "     (Select 现价, 收费细目id From 收费价目 Where (终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                    IIf(gstrPriceClass = "", " And 价格等级 Is Null ", " And 价格等级 = [1] ") & ") B," & vbNewLine & _
                    "     收费项目目录 A" & vbNewLine & _
                    "Where A.ID = E.收费细目id And A.ID = B.收费细目id And E.码类 = 1 And" & vbNewLine & _
                    "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.类别 = '4' And" & vbNewLine & _
                    "      (E.简码 Like '%" & strInput & "%' Or A.名称 Like '%" & strInput & "%' Or A.编码 Like '%" & strInput & "%')"

            End If

            vRect = zlControl.GetControlRect(txtEdit.hWnd)
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "材料", False, "", "选择材料", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True, gstrPriceClass)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vfgList
                        .EditText = Trim(Nvl(rsTmp.Fields("编码")) & " " & Nvl(rsTmp.Fields("名称")))
                        .TextMatrix(.Row, .ColIndex("对应材料")) = Trim(Nvl(rsTmp.Fields("编码")) & " " & Nvl(rsTmp.Fields("名称")))
                        .TextMatrix(.Row, .ColIndex("材料ID")) = Nvl(rsTmp.Fields("ID"), "")
                    End With
                End If
                Set rsTmp = Nothing
            End If
            txtEdit = ""
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

