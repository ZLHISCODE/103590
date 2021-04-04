VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISAuditSafeKeep 
   Caption         =   "病案封存记录"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmCISAuditSafeKeep.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   11670
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   5835
      Index           =   2
      Left            =   45
      ScaleHeight     =   5835
      ScaleWidth      =   8100
      TabIndex        =   0
      Top             =   465
      Width           =   8100
      Begin VB.PictureBox picPane 
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   0
         Left            =   45
         ScaleHeight     =   840
         ScaleWidth      =   7860
         TabIndex        =   1
         Top             =   4890
         Width           =   7860
         Begin VB.TextBox txt理由 
            Height          =   300
            Left            =   1065
            TabIndex        =   9
            Top             =   30
            Width           =   6660
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "退出(&Q)"
            Height          =   350
            Left            =   6615
            TabIndex        =   7
            Top             =   345
            Width           =   1100
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "查询(&F)"
            Height          =   350
            Left            =   5475
            TabIndex        =   6
            Top             =   345
            Width           =   1100
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1065
            TabIndex        =   2
            Top             =   390
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   487194627
            CurrentDate     =   38083
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   3345
            TabIndex        =   3
            Top             =   390
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   487194627
            CurrentDate     =   38083
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "封存理由(&1)"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   75
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "封存时间(&2)"
            Height          =   180
            Index           =   8
            Left            =   0
            TabIndex        =   5
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   9
            Left            =   3180
            TabIndex        =   4
            Top             =   435
            Width           =   180
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1200
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
         Appearance      =   2
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmCISAuditSafeKeep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mstrDateFrom As String  '开始日期
Private mstrDateTo As String    '结束日期
Private mlngMoual As Long

Private mrsCondition    As ADODB.Recordset
Private mclsVsf(0)      As clsVsf

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object, ByVal lngMoual As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngMoual = lngMoual
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Or ExecuteCommand("刷新数据") = False Then Exit Function
    
End Function


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview

        Call RptPrint(2)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print

        Call RptPrint(1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel

        Call RptPrint(3)
        
    End Select
    
End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With vfgThis
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel               '预览,打印,输出到Excel
        
            Control.Enabled = (.Rows > .FixedRows + 1)
        
        End Select
        
    End With
    
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsThis)
    Set cbsThis.Icons = frmPubResource.imgApp.Icons
    cbsThis.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = True
    
    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "输出到&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    Call CreateHelpMenu(cbsThis)
    
    '命令的快键绑定:公共部份主界面已处理
    '------------------------------------------------------------------------------------------------------------------
    With cbsThis.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
    End With
    
End Function

Private Sub RptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = vfgThis
    objPrint.Title.Text = "病案封存记录清单"
    
    Set objPrint.Title.Font = vfgThis.Font

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    Me.vfgThis.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.vfgThis.Tag = ""
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
                
        Call InitCommandBar
        
        Set mclsVsf(0) = New clsVsf
        With mclsVsf(0)
                    Call .Initialize(Me.Controls, vfgThis, True, True, frmPubResource.GetImageList(16))
                    Call .ClearColumn
                    Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("病案状态值", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("数据转出", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("就诊卡号", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("床号", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
'                    Call .AppendColumn("封存时间", 0, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True, , , True)
                    
                    Call .AppendColumn("", 240, flexAlignCenterCenter, flexDTBoolean, , "[选择]", False)
                    Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[图标]", False)
                    Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[路径]", False)
                    Call .AppendColumn("姓名", 810, flexAlignLeftCenter, flexDTString, , , True)
    
                    Call .AppendColumn("住院号", 900, flexAlignLeftCenter, flexDTDecimal, , , True)
                    
                    Call .AppendColumn("年龄", 500, flexAlignLeftCenter, flexDTDecimal, , , True)
                    Call .AppendColumn("护理等级", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                    Call .AppendColumn("住院医师", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                    Call .AppendColumn("病人状态", 1080, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("出院科室", 1080, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("审查状态", 840, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("提交人", 810, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("提交时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                    Call .AppendColumn("接收人", 810, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("接收时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                    Call .AppendColumn("封存人", 990, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("封存时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                    Call .AppendColumn("封存理由", 2000, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("出院科室ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
              
                    .SysHidden(.ColIndex("ID")) = True
                    .SysHidden(.ColIndex("病人id")) = True
                    .SysHidden(.ColIndex("主页id")) = True
                    .SysHidden(.ColIndex("病案状态值")) = True
                    .SysHidden(.ColIndex("就诊卡号")) = True
                    .SysHidden(.ColIndex("床号")) = True
'                    .SysHidden(.ColIndex("封存时间")) = True
                    .SysHidden(.ColIndex("数据转出")) = True
                    .SysHidden(.ColIndex("出院科室ID")) = True
                    
                    Call .InitializeEdit(True, False, False)
                    Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
            .AppendRows = True
        End With
        DoEvents
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        dtp(0).Value = Format(DateAdd("d", -7, Now()), "YYYY-MM-DD 00:00:00")
        dtp(1).Value = Format(Now(), "YYYY-MM-DD 23:59:59")
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
    
        mclsVsf(0).ClearGrid
        
        Set rs = gclsPackage.GetAduitPatientSafeKeep(txt理由.Text, CStr(dtp(0).Value), CStr(dtp(1).Value))
        If rs.BOF = False Then
            Call mclsVsf(0).LoadDataSource(rs)
'            rs.MoveFirst
'            Do Until rs.EOF
'                With vfgThis
'                    .Rows = .Rows + 1
'                    .TextMatrix(.Rows - 2, .ColIndex("ID")) = rs!ID
'                    .TextMatrix(.Rows - 2, .ColIndex("病人id")) = rs!病人ID
'                    .TextMatrix(.Rows - 2, .ColIndex("主页id")) = rs!主页ID
'                    .TextMatrix(.Rows - 2, .ColIndex("病案状态值")) = zlCommFun.NVL(rs!病案状态值, 0)
'                    .TextMatrix(.Rows - 2, .ColIndex("数据转出")) = zlCommFun.NVL(rs!数据转出)
'                    .TextMatrix(.Rows - 2, .ColIndex("就诊卡号")) = zlCommFun.NVL(rs!就诊卡号)
'                    .TextMatrix(.Rows - 2, .ColIndex("床号")) = zlCommFun.NVL(rs!床号)
'                    .TextMatrix(.Rows - 2, .ColIndex("封存时间")) = zlCommFun.NVL(rs!封存时间)
'                    .TextMatrix(.Rows - 2, .ColIndex("姓名")) = zlCommFun.NVL(rs!姓名)
'                    .TextMatrix(.Rows - 2, .ColIndex("住院号")) = zlCommFun.NVL(rs!住院号)
'                    .TextMatrix(.Rows - 2, .ColIndex("年龄")) = zlCommFun.NVL(rs!年龄)
'                    .TextMatrix(.Rows - 2, .ColIndex("护理等级")) = zlCommFun.NVL(rs!护理等级)
'                    .TextMatrix(.Rows - 2, .ColIndex("住院医师")) = zlCommFun.NVL(rs!住院医师)
'                    .TextMatrix(.Rows - 2, .ColIndex("出院科室")) = zlCommFun.NVL(rs!出院科室)
'                    .TextMatrix(.Rows - 2, .ColIndex("审查状态")) = zlCommFun.NVL(rs!审查状态)
'
'                    .TextMatrix(.Rows - 2, .ColIndex("提交人")) = zlCommFun.NVL(rs!提交人)
'                    .TextMatrix(.Rows - 2, .ColIndex("提交时间")) = zlCommFun.NVL(rs!提交时间)
'                    .TextMatrix(.Rows - 2, .ColIndex("接收人")) = zlCommFun.NVL(rs!接收人)
'                    .TextMatrix(.Rows - 2, .ColIndex("接收时间")) = zlCommFun.NVL(rs!接收时间)
'                    .TextMatrix(.Rows - 2, .ColIndex("封存时间")) = zlCommFun.NVL(rs!封存时间)
'                    .TextMatrix(.Rows - 2, .ColIndex("封存理由")) = zlCommFun.NVL(rs!封存理由)
'                    .TextMatrix(.Rows - 2, .ColIndex("出院科室ID")) = zlCommFun.NVL(rs!出院科室ID)
'                End With
'                rs.MoveNext
'            Loop
'
'            mclsVsf(0).AppendRows = True
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:

End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case Else
    
        If Control.ID > 400 And Control.ID < 500 Then
           
        Else
             '与业务无关的功能，公共的功能
            Call CommandBarExecutePublic(Control, Me, vfgThis, "病案封存记录清单")
            
        End If
        
    
    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long
    Dim lngScaleTop  As Long
    Dim lngScaleRight  As Long
    Dim lngScaleBottom  As Long
    
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    picPane(2).Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call ExecuteCommand("刷新数据")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 2
        vfgThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height - picPane(0).Height
        picPane(0).Move 0, vfgThis.Top + vfgThis.Height, vfgThis.Width
        
        txt理由.Move txt理由.Left, txt理由.Top, picPane(0).Width - txt理由.Left - 30
        
        cmdCancel.Move picPane(0).Width - cmdCancel.Width - 30, cmdCancel.Top
        cmdOK.Move cmdCancel.Left - cmdOK.Width - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

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

Private Sub vfgThis_AfterMoveColumn(ByVal Col As Long, Position As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vfgThis_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
 mclsVsf(0).AppendRows = True
End Sub

