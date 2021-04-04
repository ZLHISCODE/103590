VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSetClinicCharge 
   Caption         =   "外科病区-费用对照"
   ClientHeight    =   7680
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetClinicCharge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTwo 
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   7320
      ScaleHeight     =   3540
      ScaleWidth      =   4920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2985
      Width           =   4920
      Begin VSFlex8Ctl.VSFlexGrid vsCharge1 
         Height          =   3030
         Left            =   210
         TabIndex        =   9
         Top             =   150
         Width           =   12240
         _cx             =   21590
         _cy             =   5345
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
         BackColorSel    =   -2147483637
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
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
   Begin VB.PictureBox picOne 
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   285
      ScaleHeight     =   3555
      ScaleWidth      =   7680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7680
      Begin VSFlex8Ctl.VSFlexGrid vsCharge 
         Height          =   3030
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12840
         _cx             =   22648
         _cy             =   5345
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
         BackColorSel    =   -2147483637
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
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
         Begin VB.TextBox txtEdit 
            Height          =   375
            Left            =   6390
            TabIndex        =   7
            Top             =   315
            Visible         =   0   'False
            Width           =   1125
         End
      End
   End
   Begin XtremeSuiteControls.TabControl TbcCharge 
      Height          =   3375
      Left            =   1800
      TabIndex        =   4
      Top             =   3930
      Width           =   7530
      _Version        =   589884
      _ExtentX        =   13282
      _ExtentY        =   5953
      _StockProps     =   64
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8730
      TabIndex        =   3
      Top             =   105
      Width           =   1590
   End
   Begin VB.Frame FraNs 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   -510
      MousePointer    =   7  'Size N S
      TabIndex        =   1
      Top             =   5235
      Width           =   8640
   End
   Begin VSFlex8Ctl.VSFlexGrid vsClinic 
      Height          =   2985
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   12840
      _cx             =   22648
      _cy             =   5265
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      BackColorSel    =   -2147483637
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7305
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSetClinicCharge.frx":29F2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17568
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   7575
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmSetClinicCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum COL_VsClinic
    ID = 0: 类别码: 类别: 编码: 名称: 标本部位: 计算单位: 操作类型: 执行频率: 计价性质: 计算方式: 计算规则: 服务对象: 站点: 简码
End Enum
Private Enum COL_VsCharge
    ID = 0: 序号: 项目名: 规格: 单位: 价格: 数量: 固定: 从项: 收费方式: 停用
End Enum
Private mintFindType As Integer '查找模式 1-简码 2-名称编码
Private mblnEditMode As Boolean '是否编辑模式

Private mlngDeptID As Long      '病区ID
Private mlngClinicID As Long    '当前诊疗项目ID
Private mlng计价性质 As Long    '当前诊疗项目的计价性质
Private mblnCopy As Boolean     '是否可以使用复制到其他病区的 功能
Private mstrTitle As String     '标题
Private mint场合 As Integer     '调用场和,1-门诊，2-住院
Private mblnModify As Boolean   '是否修改
Private mblnModifyPrivs As Boolean
Private mlngModul As Long       '调用的模块号

Public Sub ShowMe(ByVal lngDeptID As Long, ByVal lngMode As Long, frmMain As Form, int场合 As Integer, ByVal blnModify As Boolean)
    Dim strSql As String, rsTmp As ADODB.Recordset
    On Error GoTo errH
    If lngDeptID <= 0 Then Exit Sub
    mlngDeptID = lngDeptID
    mlngClinicID = 0
    mint场合 = int场合
    mblnModifyPrivs = blnModify
    
'    If InStr(frmMain.Caption, "门诊医生工作站 -") > 0 Then
'        mlngModul = 1260
'    ElseIf InStr(frmMain.Caption, "住院护士工作站 -") > 0 Then
'        mlngModul = 1262
'    ElseIf InStr(frmMain.Caption, "医技工作站 -") > 0 Then
'        mlngModul = 1263
'    End If
    mlngModul = glngModul
    
    If mint场合 = 2 Then
        strSql = "Select Distinct ID, 编码, 名称" & vbNewLine & _
        "From 部门表 A, 部门人员 B, 上机人员表 C, 部门性质说明 D" & vbNewLine & _
        "Where a.Id = b.部门id And b.人员id = c.人员id And a.Id = d.部门id And (d.服务对象 = 2 or d.服务对象=3) " & _
        "    And d.工作性质 in ('护理','检查','检验','手术','麻醉','治疗','营养') And A.ID<>[1] And c.用户名 = User"
    Else
        strSql = "Select Distinct ID, 编码, 名称" & vbNewLine & _
        "From 部门表 A, 部门人员 B, 上机人员表 C, 部门性质说明 D" & vbNewLine & _
        "Where a.Id = b.部门id And b.人员id = c.人员id And a.Id = d.部门id And (d.服务对象 = 1 Or d.服务对象 = 3) " & vbNewLine & _
        "    And Instr('临床,检验,检查,手术,麻醉,治疗,营养', d.工作性质) > 0 And a.Id <> [1] And c.用户名 = User"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    mblnCopy = Not rsTmp.EOF
    
    strSql = "Select Distinct ID, 编码, 名称 From 部门表 A Where A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    mstrTitle = "费用对照"
    If Not rsTmp.EOF Then mstrTitle = "" & rsTmp!名称 & "(" & rsTmp!编码 & ")" & "-" & IIf(mint场合 = 2, "住院费用对照", "门诊费用对照")
    Me.Show lngMode, frmMain

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    
    On Error GoTo ErrHandle
    Select Case Control.ID
 
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsThis.Count
            Me.cbsThis(i).Visible = Not Me.cbsThis(i).Visible
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsThis.Count
            For Each objControl In Me.cbsThis(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_Edit_Modify '进入编辑
        mblnModify = False
        Call CheckIsNoCharge '检查是否从未对照过费用，是则复制所有科室的对照。
        mblnEditMode = True
        vsClinic.Enabled = False
        txtFind.Enabled = vsClinic.Enabled
        With vsCharge
            .Editable = flexEDKbdMouse
            .SelectionMode = flexSelectionFree
            If TbcCharge.Selected.Index <> 0 Then TbcCharge.Item(0).Selected = True
            .SetFocus
            .Col = COL_VsCharge.项目名
        End With
        Select Case Trim("" & vsClinic.TextMatrix(vsClinic.Row, COL_VsClinic.计价性质))
            Case "正常计价": mlng计价性质 = 0
            Case "不计价": mlng计价性质 = 1
            Case "手工计价": mlng计价性质 = 2
        End Select
    Case conMenu_Edit_Save   '保存
        
        If SaveData(mlngDeptID) Then
            mblnEditMode = False
            vsClinic.Enabled = True
            txtFind.Enabled = vsClinic.Enabled
            vsCharge.Editable = flexEDNone
            vsCharge.SelectionMode = flexSelectionByRow
            Call vsChargeRef(mlngClinicID)
        End If
    Case conMenu_Edit_Untread '取消
        If mblnModify Then
            If MsgBox("是否放弃当前已调整的数据？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        mblnEditMode = False
        vsClinic.Enabled = True
        txtFind.Enabled = vsClinic.Enabled
        vsCharge.Editable = flexEDNone
        vsCharge.SelectionMode = flexSelectionByRow
        Call vsChargeRef(mlngClinicID)
    
    Case conMenu_Edit_NewItem '新增
        mblnModify = True
        With vsCharge
            If Val(.TextMatrix(.Rows - 1, COL_VsCharge.ID)) <> 0 Then
                .Rows = .Rows + 1
            ElseIf .Rows = .FixedRows Then
                .Rows = .Rows + 1
            End If
            .Row = .Rows - 1
        End With
    Case conMenu_Edit_Delete '删除
        '
        mblnModify = True
        Call DeleteCharge
        
    Case conMenu_View_FindType  '查找方式
        
        cbsThis.RecalcLayout
        txtFind.Text = ""
        txtFind.SetFocus
        
    Case conMenu_View_Find '查找
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus '有时需要定位一下
            If txtFind.Text <> "" Then
                Call ExecuteFind
            End If
        Else
            txtFind.SetFocus
        End If
    Case conMenu_View_FindNext '查找下一个
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call ExecuteFind(True)
        End If
    Case conMenu_Edit_Copy '应用于其他病区
        Call DeptCopy
    End Select
    Exit Sub

ErrHandle:
    Call ErrCenter
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    '右键弹出菜单
'    Select Case CommandBar.Parent.ID
'    Case conMenu_View_FindType '查找方式
'        With CommandBar.Controls
'            If .Count = 0 Then
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "简码(&1)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "编码或名称(&2)"
'            End If
'        End With
'    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    FraNs.Left = Me.ScaleLeft
    FraNs.Width = Me.ScaleWidth
    
    With vsClinic
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = FraNs.Top - .Top
    End With
    With TbcCharge
        .Left = lngLeft
        .Top = FraNs.Top + FraNs.Height
        .Width = lngRight - lngLeft
        If Me.stbThis.Visible Then
            .Height = Me.ScaleHeight - .Top - Me.stbThis.Height
        Else
            .Height = Me.ScaleHeight - .Top
        End If

    End With
    
'    With vsCharge
'        .Left = lngLeft
'        .Width = lngRight - lngLeft
'        .Top = FraNs.Top + FraNs.Height
'
'        If Me.stbThis.Visible Then
'            .Height = Me.ScaleHeight - .Top - Me.stbThis.Height
'        Else
'            .Height = Me.ScaleHeight - .Top
'        End If
'    End With


End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsThis.Count >= 2 Then
            Control.Checked = Me.cbsThis(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsThis.Count >= 2 Then
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case conMenu_Edit_Modify, conMenu_Edit_Copy
        Control.Enabled = Not mblnEditMode And mblnModifyPrivs
        If Control.ID = conMenu_Edit_Copy And Control.Enabled Then
            Control.Enabled = mblnCopy
        End If
    Case conMenu_Edit_Save, conMenu_Edit_Untread, conMenu_Edit_NewItem, conMenu_Edit_Delete
        Control.Enabled = mblnEditMode
'    Case conMenu_View_FindType '查找方式
'        If Control.Parent Is cbsThis.ActiveMenuBar Then
'            Control.Caption = "↓按" & IIf(mintFindType = 0, "简码", "编码或名称") & "查找"
'        End If
    End Select
End Sub

Private Sub Form_Activate()
    mlngClinicID = -1
    Me.Caption = mstrTitle
    Call vsClinic_RowColChange
End Sub

Private Sub Form_Load()
    
    '初始化界面
    
    Call initMenu
    Call initTbcCharge
    Call initVsClinic
    Call initVsCharge(vsCharge)
    Call initVsCharge(vsCharge1)
    '装入数据
    
    Call zlRefRecords
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEditMode Then Call cbsThis_Execute(Me.cbsThis.FindControl(, conMenu_Edit_Untread))
    If mblnEditMode Then
        Cancel = True
        Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub FraNs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        On Error Resume Next
        If vsClinic.Height + y < 1000 Or TbcCharge.Height - y < 1000 Then Exit Sub
        FraNs.Top = FraNs.Top + y
        vsClinic.Height = vsClinic.Height + y
        TbcCharge.Top = TbcCharge.Top + y
        TbcCharge.Height = TbcCharge.Height - y
    End If
End Sub

Private Sub initTbcCharge()
    With TbcCharge
        If .ItemCount <= 0 Then
            With .PaintManager
                .Appearance = xtpTabAppearanceExcel
                .ClientFrame = xtpTabFrameSingleLine
                
                .Position = xtpTabPositionTop  '选项卡在顶部
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
             
            .InsertItem(0, "当前" & IIf(mint场合 = 2, "病区", "科室"), picOne.hWnd, 0).Tag = "当前" & IIf(mint场合 = 2, "病区", "科室")
            .InsertItem(1, "全院通用", picTwo.hWnd, 0).Tag = "全院通用"
            .Item(0).Selected = True
        End If
    End With
End Sub

Private Sub initVsClinic()
    With vsClinic
        .Clear
        '配色
        .BackColorBkg = vbWindowBackground          '窗口背景*
        .BackColorSel = &HFFEBD7    'vbInactiveBorder            '非活动边框*
        .GridColor = vbActiveBorder                 '活动边框 *
        .SheetBorder = vbWindowBackground           '窗口背景*
        .ForeColorSel = vbWindowText                '窗口文本*
        '
        .FocusRect = flexFocusNone                  '无单元格焦点框
        
        '初始行，列
        .Cols = 15
        .Rows = 2
        
        .TextMatrix(0, COL_VsClinic.ID) = "ID": .ColWidth(COL_VsClinic.ID) = 0: .ColHidden(COL_VsClinic.ID) = True
        .TextMatrix(0, COL_VsClinic.类别码) = "类别码": .ColWidth(COL_VsClinic.类别码) = 0: .ColHidden(COL_VsClinic.类别码) = True
        .TextMatrix(0, COL_VsClinic.站点) = "站点": .ColWidth(COL_VsClinic.站点) = 0: .ColHidden(COL_VsClinic.站点) = True
        .TextMatrix(0, COL_VsClinic.站点) = "简码": .ColWidth(COL_VsClinic.简码) = 0: .ColHidden(COL_VsClinic.简码) = True
        .TextMatrix(0, COL_VsClinic.标本部位) = "标本部位": .ColWidth(COL_VsClinic.标本部位) = 0: .ColHidden(COL_VsClinic.标本部位) = True
        
        .TextMatrix(0, COL_VsClinic.类别) = "类别": .ColWidth(COL_VsClinic.类别) = 600
        .TextMatrix(0, COL_VsClinic.编码) = "编码": .ColWidth(COL_VsClinic.编码) = 1200
        .TextMatrix(0, COL_VsClinic.名称) = "诊疗项目名称": .ColWidth(COL_VsClinic.名称) = 3200
        
        .TextMatrix(0, COL_VsClinic.计算单位) = "计算单位": .ColWidth(COL_VsClinic.计算单位) = 900
        .TextMatrix(0, COL_VsClinic.操作类型) = "操作类型": .ColWidth(COL_VsClinic.操作类型) = 1200
        .TextMatrix(0, COL_VsClinic.执行频率) = "执行频率": .ColWidth(COL_VsClinic.执行频率) = 1200
        .TextMatrix(0, COL_VsClinic.计价性质) = "计价性质": .ColWidth(COL_VsClinic.计价性质) = 1200
        .TextMatrix(0, COL_VsClinic.计算方式) = "计算方式": .ColWidth(COL_VsClinic.计算方式) = 1200
        .TextMatrix(0, COL_VsClinic.计算规则) = "计算规则": .ColWidth(COL_VsClinic.计算规则) = 1200
        .TextMatrix(0, COL_VsClinic.服务对象) = "服务对象": .ColWidth(COL_VsClinic.服务对象) = 1200
        
        .Cell(flexcpAlignment, 0, .FixedCols, 0, .Cols - 1) = flexAlignCenterCenter
        
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
    End With

End Sub

Private Sub initVsCharge(ByRef objVsf As VSFlexGrid)
    With objVsf
        .Clear
        '配色
        .BackColorBkg = vbWindowBackground          '窗口背景*
        .BackColorSel = &HFFEBD7 'vbInactiveBorder            '非活动边框*
        .GridColor = vbActiveBorder                 '活动边框 *
        .SheetBorder = vbWindowBackground           '窗口背景*
        .ForeColorSel = vbWindowText                '窗口文本*
        '
        .FocusRect = flexFocusNone                  '无单元格焦点框
        
        '初始行，列
        .Cols = 11
        .Rows = 2
        .TextMatrix(0, COL_VsCharge.ID) = "ID": .ColWidth(COL_VsCharge.ID) = 0: .ColHidden(COL_VsCharge.ID) = True
        
        .TextMatrix(0, COL_VsCharge.序号) = "序号": .ColWidth(COL_VsCharge.序号) = 500
        .TextMatrix(0, COL_VsCharge.项目名) = "收费项目名称": .ColWidth(COL_VsCharge.项目名) = 3600
        .TextMatrix(0, COL_VsCharge.规格) = "规格": .ColWidth(COL_VsCharge.规格) = 2600
        .TextMatrix(0, COL_VsCharge.单位) = "单位": .ColWidth(COL_VsCharge.单位) = 900
        .TextMatrix(0, COL_VsCharge.价格) = "价格": .ColWidth(COL_VsCharge.价格) = 1000
        .TextMatrix(0, COL_VsCharge.数量) = "数量": .ColWidth(COL_VsCharge.数量) = 800
        .TextMatrix(0, COL_VsCharge.固定) = "固定": .ColWidth(COL_VsCharge.固定) = 500
        .TextMatrix(0, COL_VsCharge.从项) = "从项": .ColWidth(COL_VsCharge.从项) = 500
        .TextMatrix(0, COL_VsCharge.收费方式) = "收费方式": .ColWidth(COL_VsCharge.收费方式) = 1800
        .TextMatrix(0, COL_VsCharge.停用) = "停用": .ColWidth(COL_VsCharge.停用) = 500
        
        .Cell(flexcpAlignment, 0, .FixedCols, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(COL_VsCharge.固定) = flexAlignCenterCenter
        .ColAlignment(COL_VsCharge.从项) = flexAlignCenterCenter
        .AllowUserResizing = flexResizeColumns
        
        .ColComboList(COL_VsCharge.收费方式) = "0-正常收取|1-检验试管费用|2-一次发送只收取一次|3-当天只收取一次|4-当天未执行收取一次|5-当天只收取一次，排斥其他项目|6-当天未执行收取一次，排斥其他项目"
    End With
End Sub

Private Sub initMenu()
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup

    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objCustom As CommandBarControlCustom
 
    'Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With Me.cbsThis.Options
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
    cbsThis.EnableCustomization False
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    'Call Me.cbsThis.ActiveMenuBar.EnableDocking(xtpFlagAlignTop) '这句加了，就不能控制查找框的位置了
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "对照(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&U)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "应用于其他" & IIf(mint场合 = 2, "病区", "科室") & "(&C)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        'Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "查找(&Y)"): objPopup.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个")

        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '查找项特殊处理
    '-----------------------------------------------------
    
    '主菜单右侧的查找 按就诊卡号查找，支持刷卡
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlLabel, conMenu_View_FindType, "查找")
        cbrControl.ID = conMenu_View_FindType
        cbrControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.hWnd
        objCustom.Flags = xtpFlagRightAlign
'
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F3, conMenu_View_FindNext
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "对照")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "应用于其他" & IIf(mint场合 = 2, "病区", "科室")): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
End Sub

Private Sub picOne_Resize()
    With vsCharge
        .Top = picOne.ScaleTop
        .Width = picOne.ScaleWidth
        .Height = picOne.ScaleHeight
        .Left = picOne.ScaleLeft
    End With
End Sub

Private Sub picTwo_Resize()
    With vsCharge1
        .Top = picTwo.ScaleTop
        .Width = picTwo.ScaleWidth
        .Height = picTwo.ScaleHeight
        .Left = picTwo.ScaleLeft
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    '按回车
    If KeyAscii = vbKeyReturn Then
        ExecuteFind
    End If
End Sub

Private Sub vsCharge_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = COL_VsCharge.项目名 Or _
        Col = COL_VsCharge.数量 Or Col = COL_VsCharge.收费方式) Then
        Cancel = True
    End If
End Sub

Private Sub vsCharge_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnEditMode Then
        If NewCol = COL_VsCharge.项目名 Then
            vsCharge.ComboList = "..."
        Else
            vsCharge.ComboList = ""
        End If
    Else
        vsCharge.ComboList = ""
    End If
End Sub

Private Sub vsCharge_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '
    Call SelectCharge(vsCharge.Row, vsCharge.Col)
End Sub

Private Sub vsCharge_DblClick()
    With vsCharge
        If (.Col = COL_VsCharge.固定 Or .Col = COL_VsCharge.从项) And mblnEditMode = True Then
            mblnModify = True
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "√"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
End Sub

Private Sub vsCharge_EnterCell()
    With vsCharge
    
        If mblnEditMode Then
            If .Col = COL_VsCharge.项目名 Then
                .FocusRect = flexFocusHeavy
                If txtEdit.Tag = "False" Then
                    txtEdit.Left = .CellLeft
                    txtEdit.Top = .CellTop
                    txtEdit.Height = .CellHeight - 12
                    txtEdit.Width = .CellWidth - 12
                    txtEdit.Tag = "True"
                End If
            ElseIf .Col = COL_VsCharge.数量 _
                 Or .Col = COL_VsCharge.停用 _
                 Or .Col = COL_VsCharge.固定 _
                 Or .Col = COL_VsCharge.从项 _
                 Or .Col = COL_VsCharge.收费方式 _
                 Then
                .FocusRect = flexFocusHeavy
                txtEdit.Tag = "False"
            Else
                .FocusRect = flexFocusLight
                txtEdit.Tag = "False"
            End If
        Else
            .FocusRect = flexFocusNone
        End If
    End With
End Sub

Private Sub vsCharge_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Col = COL_VsCharge.项目名 Then
        vsCharge.ComboList = "..."
    Else
        vsCharge.ComboList = ""
    End If
End Sub

Private Sub vsCharge_KeyPress(KeyAscii As Integer)
    If Not mblnEditMode Then Exit Sub
    mblnModify = True
    With vsCharge
       If KeyAscii = vbKeyReturn Then
           KeyAscii = 0
           If .Col = COL_VsCharge.停用 And .Row = .Rows - 1 And Val(.TextMatrix(.Row, COL_VsCharge.ID)) > 0 Then
               .Rows = .Rows + 1
               .Select .Rows - 1, COL_VsCharge.项目名
           ElseIf .Col = COL_VsCharge.项目名 Then
               .Select .Row, COL_VsCharge.数量
           Else
               Call zlCommFun.PressKey(vbKeyRight)
           End If
       ElseIf .Col = COL_VsCharge.项目名 And .ComboList = "..." Then
           If KeyAscii = Asc("*") Then
               KeyAscii = 0
               txtEdit.Text = .EditText
               Call SelectCharge(.Row, .Col)
               txtEdit.Tag = False
               txtEdit.Visible = False
           Else
               .ComboList = "" '使按钮状态进入输入状态
           End If
    
       ElseIf (.Col = COL_VsCharge.固定 _
               Or .Col = COL_VsCharge.从项) _
           And Val(.TextMatrix(.Row, COL_VsCharge.ID)) > 0 _
           And KeyAscii = vbKeySpace Then
           
           If .TextMatrix(.Row, .Col) = "" Then
               .TextMatrix(.Row, .Col) = "√"
           Else
               .TextMatrix(.Row, .Col) = ""
           End If
       ElseIf KeyAscii = vbKeyDelete Then
           Call DeleteCharge
       End If
    End With
End Sub

Private Sub vsCharge_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsCharge
            If Col = COL_VsCharge.项目名 Then
                txtEdit.Text = .EditText
                .EditText = ""
                Call SelectCharge(Row, Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
                .TextMatrix(.Row, COL_VsCharge.数量) = 1
                .TextMatrix(.Row, COL_VsCharge.收费方式) = "0-正常收取"
                .Select .Row, COL_VsCharge.数量 - 1
            ElseIf Col = COL_VsCharge.数量 Or Col = COL_VsCharge.收费方式 Then
                Call zlCommFun.PressKey(vbKeyRight)
            End If
        End With
    End If
End Sub

Private Sub vsCharge_LeaveCell()
    With vsCharge
        On Error Resume Next
        If mblnEditMode Then
            .FocusRect = flexFocusLight
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        Else
            .FocusRect = flexFocusNone
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vsCharge_RowColChange()
    On Error GoTo ErrHandle
    
    With vsCharge
        If mblnEditMode Then
            If .Col = COL_VsCharge.项目名 Or .Col = COL_VsCharge.数量 Or .Col = COL_VsCharge.固定 Or .Col = COL_VsCharge.从项 Or .Col = COL_VsCharge.收费方式 Then
                '.FocusRect = flexFocusHeavy
                .FocusRect = flexFocusNone
                Call .CellBorder(vbBlue, 1, 1, 1.5, 1.5, 0, 0)
            Else
                .FocusRect = flexFocusLight
                Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
            End If
            If txtEdit.Tag = "True" Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
            End If
        Else
            .FocusRect = flexFocusNone
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
    Call SetVsfCharge
    
    Exit Sub
ErrHandle:
    If Err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteCharge()
    '删除收费项目
    Dim strName As String
    With vsCharge
        If .Row >= .FixedRows Then
            strName = .TextMatrix(.Row, COL_VsCharge.项目名)
            If MsgBox("是否删除“" & strName & "”？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call .RemoveItem(.Row)
                stbThis.Panels(2).Text = "“" & strName & "”已删除！"
            End If
        End If
    End With
End Sub

Private Sub vsCharge_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCharge
        If Col = COL_VsCharge.数量 Then
            If Not IsNumeric(.EditText) Then
                Cancel = True
                MsgBox "请输入数字！"
                Exit Sub
            ElseIf Not Val(.EditText) > 0 And Val(.EditText) <= 9999 Then
                Cancel = True
                MsgBox "请输入0-9999之间的数字！"
                Exit Sub
            End If
        ElseIf Col = COL_VsCharge.从项 Then
        ElseIf Col = COL_VsCharge.固定 Then
        ElseIf Col = COL_VsCharge.收费方式 Then
        ElseIf Col = COL_VsCharge.项目名 Then
        End If
        
    
    End With

End Sub

Private Sub vsClinic_RowColChange()
    With vsClinic
        If mlngClinicID = Val("" & .TextMatrix(.Row, COL_VsClinic.ID)) Then Exit Sub
        mlngClinicID = Val("" & .TextMatrix(.Row, COL_VsClinic.ID))
        Call vsChargeRef(mlngClinicID)
    End With
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    Dim iSubItemIndex As Integer, intCount As Integer, strTemp As String
    Dim rsTemp As ADODB.Recordset, intSelectRow As Integer, blnCharge As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim eTime As Single, sTime As Single
    Dim strDepts As String
    '---------------------------------------------
    '提取诊疗项目
    '只提取 诊疗项目目录.执行科室=1-病人所在科室 2-病人所在病区 和4-指定科室 中，指定科室为传入科室ID 的项目
    '                 并且 服务对象为 1-门诊，3-门诊住院(场合为1-门诊时) 或 2-住院,3-门诊住院(场合为2-住院时)
    '       并且项目大类为E-治疗 H-护理 I-膳食
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand

    If mlngModul = 1263 Then
        '医技工作站
        gstrSql = "Select /*+Rule */ Distinct A.*,b.诊疗项目id as 收费 From (" & vbNewLine
        strDepts = mlngDeptID
    Else
        '门诊医生工作站/住院护士工作站
        strDepts = mlngDeptID
        If mlngModul = 1262 Then
            '护士站允许调整病区对应科室执行的项目
            gstrSql = "Select 科室ID From 病区科室对应 Where 病区ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
            Do While Not rsTemp.EOF
                strDepts = strDepts & "," & rsTemp!科室ID
                rsTemp.MoveNext
            Loop
        End If
        gstrSql = "Select /*+Rule */ Distinct A.*,b.诊疗项目id as 收费 From (Select i.Id, i.编码, i.名称, i.标本部位, i.计算单位, i.类别 As 类别码, k.名称 As 类别, i.操作类型, i.执行频率, i.计算方式, i.计算规则," & vbNewLine & _
                "       Decode(i.服务对象, 1, '门诊', 2, '住院', 3, '门诊和住院', 4, '体检', '不直接应用于病人') As 服务对象," & vbNewLine & _
                "       Nvl(i.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间, i.站点, i.计价性质, zlSpellCode(i.名称) As 简码" & vbNewLine & _
                "From 诊疗项目目录 I, 诊疗项目类别 K" & vbNewLine & _
                "Where i.类别 = k.编码 And (Instr(',E,H,I,', i.类别) > 0 Or i.类别='D' and i.操作类型='其他') And (i.服务对象 = [2] Or i.服务对象 = 3) And" & vbNewLine & _
                "      instr([3],i.执行科室)>0 And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) " & vbNewLine & _
                "Union All" & vbNewLine
        
    End If
    gstrSql = gstrSql & _
            "Select i.Id, i.编码, i.名称, i.标本部位, i.计算单位, i.类别 As 类别码, k.名称 As 类别, i.操作类型, i.执行频率, i.计算方式, i.计算规则," & vbNewLine & _
            "       Decode(i.服务对象, 1, '门诊', 2, '住院', 3, '门诊和住院', 4, '体检', '不直接应用于病人') As 服务对象," & vbNewLine & _
            "       Nvl(i.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间, i.站点, i.计价性质, zlSpellCode(i.名称) As 简码" & vbNewLine & _
            "From 诊疗项目目录 I, 诊疗项目类别 K, 诊疗执行科室 J" & vbNewLine & _
            "Where i.类别 = k.编码 And (Instr(',E,H,I,', i.类别) > 0 Or i.类别='D' and i.操作类型='其他') And (i.服务对象 = [2] Or i.服务对象 = 3) And" & vbNewLine & _
            "      i.执行科室 = 4 And Nvl(j.病人来源,0) <> [2] And i.Id = j.诊疗项目id And j.执行科室id  In (Select Column_Value From Table(f_Num2list([1]))) And" & vbNewLine & _
            "      (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))) A,(Select 诊疗项目id From 诊疗收费关系 Where 病人来源 = [2] And 适用科室id  In (Select Column_Value From Table(f_Num2list([1]))) ) B" & vbNewLine & _
            "Where A.id=B.诊疗项目ID(+) Order By 类别, 操作类型, 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strDepts, mint场合, IIf(mlngModul = 1262, "12", mint场合))
    
    With vsClinic
        .Clear
        Call initVsClinic
        Do While Not rsTemp.EOF
            
            If Val(.TextMatrix(.Rows - 1, COL_VsClinic.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_VsClinic.ID) = "" & rsTemp!ID
            .TextMatrix(.Rows - 1, COL_VsClinic.类别) = "" & rsTemp!类别
            .TextMatrix(.Rows - 1, COL_VsClinic.编码) = "" & rsTemp!编码
            .TextMatrix(.Rows - 1, COL_VsClinic.名称) = "" & rsTemp!名称
            .TextMatrix(.Rows - 1, COL_VsClinic.标本部位) = "" & rsTemp!标本部位
            .TextMatrix(.Rows - 1, COL_VsClinic.计算单位) = "" & rsTemp!计算单位
            
            .TextMatrix(.Rows - 1, COL_VsClinic.执行频率) = "" & rsTemp!执行频率

            Select Case Val("" & rsTemp!执行频率)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.执行频率) = "可选频率"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.执行频率) = "一次性"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsClinic.执行频率) = "持续性"
            End Select
            Select Case Val("" & rsTemp!计算方式)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.计算方式) = "不确定"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.计算方式) = "计量"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsClinic.计算方式) = "计时"
            Case 3
                .TextMatrix(.Rows - 1, COL_VsClinic.计算方式) = "计次"
            End Select
            Select Case Val("" & rsTemp!计算规则)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.计算规则) = "正常计算"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.计算规则) = "取整计算"
            End Select
            .TextMatrix(.Rows - 1, COL_VsClinic.服务对象) = "" & rsTemp!服务对象
            .TextMatrix(.Rows - 1, COL_VsClinic.站点) = "" & rsTemp!站点
            
            Select Case rsTemp!类别码
            Case "E"
                intCount = Val("" & rsTemp!操作类型)
                strTemp = Switch(intCount = 0, "普通", _
                                intCount = 1, "过敏试验", _
                                intCount = 2, "给药方法(西药)", _
                                intCount = 3, "中药煎法", _
                                intCount = 4, "中药用(服)法", _
                                intCount = 5, "特殊治疗", _
                                intCount = 6, "采集方法", _
                                intCount = 7, "配血方法", _
                                intCount = 8, "输血途径", _
                                intCount = 9, "输血采集")
                .TextMatrix(.Rows - 1, COL_VsClinic.操作类型) = strTemp
            Case "H"
                If IIf(IsNull(rsTemp!操作类型), "0", rsTemp!操作类型) = "1" Then
                    .TextMatrix(.Rows - 1, COL_VsClinic.操作类型) = "护理等级"
                Else
                    .TextMatrix(.Rows - 1, COL_VsClinic.操作类型) = "护理常规"
                End If
            Case "Z"
                intCount = Val("" & rsTemp!操作类型)
                strTemp = Switch(intCount = 0, "普通", _
                                intCount = 1, "留观", _
                                intCount = 2, "住院", _
                                intCount = 3, "转科", _
                                intCount = 4, "术后", _
                                intCount = 5, "出院", _
                                intCount = 6, "转院", _
                                intCount = 7, "会诊", _
                                intCount = 8, "抢救", _
                                intCount = 9, "病重", _
                                intCount = 10, "病危", _
                                intCount = 11, "死亡", _
                                intCount = 12, "记录入出量")
                .TextMatrix(.Rows - 1, COL_VsClinic.操作类型) = strTemp
            Case Else
                .TextMatrix(.Rows - 1, COL_VsClinic.操作类型) = "" & rsTemp!操作类型
            End Select
            .TextMatrix(.Rows - 1, COL_VsClinic.类别码) = rsTemp!类别码
            
            Select Case Val("" & rsTemp!计价性质)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.计价性质) = "正常计价"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.计价性质) = "不计价"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsClinic.计价性质) = "手工计价"
            End Select
            
            .TextMatrix(.Rows - 1, COL_VsClinic.简码) = "" & rsTemp!简码
            .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
            
            '将停用项目显示为红色
            If Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HFF&
            End If
            
            '有对照，显示为兰色
            'gstrSql = "Select 诊疗项目ID From 诊疗收费关系 Where 病人来源 = 2 And 诊疗项目ID=[1] And 适用科室ID=[2]"
            'Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val("" & rsTemp!ID), mlngDeptID)
            If Val("" & rsTemp!收费) <> 0 Then
                .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = vbBlue
            End If

            If mlngClinicID <> 0 And mlngClinicID = rsTemp!ID Then
                intSelectRow = .Rows - 1
            End If
            
            rsTemp.MoveNext
        Loop
        
        If intSelectRow <> 0 Then
            .Select intSelectRow, COL_VsClinic.编码
        Else
            .Select .FixedRows, COL_VsClinic.编码
        End If
    End With
    
    Err = 0: On Error Resume Next
    If Val(vsClinic.TextMatrix(vsClinic.Rows - 1, COL_VsClinic.ID)) <> 0 Then
        Me.stbThis.Panels(2).Text = "该分类共有" & Me.vsClinic.Rows - 1 & "种项目"
    Else
        Call initVsCharge(vsCharge)
        Call initVsCharge(vsCharge1)
        Me.stbThis.Panels(2).Text = ""
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsChargeRef(ByVal lngClinicID As Long)
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim lngForeColor As Long, curTotal As Currency
    Dim strInfo As String
    On Error GoTo ErrHand
    
    Call initVsCharge(vsCharge)
    Call initVsCharge(vsCharge1)
    
    strInfo = ""
    Me.stbThis.Panels(2).Text = ""
    
    If lngClinicID <= 0 Then Exit Sub
    
    '---本病区的费用对照
    gstrSql = "select I.ID,R.检查部位,R.检查方法,R.费用性质,'['||I.编码||']'||I.名称 as 名称,I.规格,I.计算单位,decode(I.是否变价,1,'变价',to_char(P.价格)) As 价格," & _
            "       nvl(R.收费数量,0) as 数量,nvl(R.固有对照,0) as 固定,nvl(R.从属项目,0) as 从项," & _
            "Nvl(I.撤档时间,to_Date('3000-01-01','YYYY-MM-DD')) As 撤档时间,Nvl(R.收费方式,0) As 收费方式 " & _
            " from 诊疗收费关系 R,收费项目目录 I," & _
            "      (Select P.收费细目id,sum(P.现价) As 价格" & _
            "      From 收费价目 P " & _
            "      Where P.执行日期<=Sysdate And (P.终止日期 Is Null Or P.终止日期>=Sysdate)" & _
            IIf(gstrPriceClass = "", " And P.价格等级 Is Null ", " And P.价格等级 = [4] ") & _
            "      Group by P.收费细目id) P" & _
            " where R.收费项目ID=I.ID and I.ID=P.收费细目id(+) And R.病人来源 = [3] and R.诊疗项目ID=[1] And R.适用科室ID=[2]" & _
            " order by nvl(R.从属项目,0) ,R.ROWID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClinicID, mlngDeptID, mint场合, gstrPriceClass)
        
    With vsCharge
        curTotal = 0
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, COL_VsCharge.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_VsCharge.序号) = .Rows - 1
            .TextMatrix(.Rows - 1, COL_VsCharge.ID) = "" & rsTemp!ID
            .TextMatrix(.Rows - 1, COL_VsCharge.项目名) = "" & rsTemp!名称
            .TextMatrix(.Rows - 1, COL_VsCharge.规格) = "" & rsTemp!规格
            .TextMatrix(.Rows - 1, COL_VsCharge.单位) = "" & rsTemp!计算单位
            .TextMatrix(.Rows - 1, COL_VsCharge.价格) = FormatEx(Format("" & rsTemp!价格, "0.00"), 2)
            .TextMatrix(.Rows - 1, COL_VsCharge.数量) = FormatEx(Format("" & rsTemp!数量, "0.00000"), 5)
            .TextMatrix(.Rows - 1, COL_VsCharge.固定) = IIf(Val("" & rsTemp!固定) = 0, "", "√")
            .TextMatrix(.Rows - 1, COL_VsCharge.从项) = IIf(Val("" & rsTemp!从项) = 0, "", "√")
            .TextMatrix(.Rows - 1, COL_VsCharge.停用) = IIf(Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01", "停用", "")
            
            Select Case rsTemp!收费方式
            Case 0
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "0-正常收取"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "1-检验试管费用"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "2-一次发送只收取一次"
            Case 3
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "3-当天只收取一次"
            Case 4
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "4-当天未执行收取一次"
            Case 5
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "5-当天只收取一次，排斥其他项目"
            Case 6
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "6-当天未执行收取一次，排斥其他项目"
            Case Else
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "0-正常收取"
            End Select

            If Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                lngForeColor = &HFF&
            Else
                lngForeColor = &H0&
            End If
            .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = lngForeColor
            
            curTotal = curTotal + Val("" & rsTemp!价格) * Val("" & rsTemp!数量)
            rsTemp.MoveNext
        Loop
        If curTotal <> 0 Then strInfo = " " & IIf(mint场合 = 2, "病区", "科室") & "对照合计：" & FormatEx(Format(curTotal, "0.0000"), 4)

    End With
    '---- 所有科室的收费对照
    gstrSql = "select I.ID,R.检查部位,R.检查方法,R.费用性质,'['||I.编码||']'||I.名称 as 名称,I.规格,I.计算单位,decode(I.是否变价,1,'变价',to_char(P.价格)) As 价格," & _
            "       nvl(R.收费数量,0) as 数量,nvl(R.固有对照,0) as 固定,nvl(R.从属项目,0) as 从项," & _
            "Nvl(I.撤档时间,to_Date('3000-01-01','YYYY-MM-DD')) As 撤档时间,Nvl(R.收费方式,0) As 收费方式 " & _
            " from 诊疗收费关系 R,收费项目目录 I," & _
            "      (Select P.收费细目id,sum(P.现价) As 价格" & _
            "      From 收费价目 P " & _
            "      Where P.执行日期<=Sysdate And (P.终止日期 Is Null Or P.终止日期>=Sysdate)" & _
            IIf(gstrPriceClass = "", " And P.价格等级 Is Null ", " And P.价格等级 = [2] ") & _
            "      Group by P.收费细目id) P" & _
            " where R.收费项目ID=I.ID and I.ID=P.收费细目id(+) And R.病人来源 = 0 and R.诊疗项目ID=[1] And R.适用科室ID Is Null" & _
            " order by nvl(R.从属项目,0) ,R.ROWID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClinicID, gstrPriceClass)
    With vsCharge1
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, COL_VsCharge.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_VsCharge.序号) = .Rows - 1
            .TextMatrix(.Rows - 1, COL_VsCharge.ID) = "" & rsTemp!ID
            .TextMatrix(.Rows - 1, COL_VsCharge.项目名) = "" & rsTemp!名称
            .TextMatrix(.Rows - 1, COL_VsCharge.规格) = "" & rsTemp!规格
            .TextMatrix(.Rows - 1, COL_VsCharge.单位) = "" & rsTemp!计算单位
            .TextMatrix(.Rows - 1, COL_VsCharge.价格) = FormatEx(Format("" & rsTemp!价格, "0.00"), 2)
            .TextMatrix(.Rows - 1, COL_VsCharge.数量) = FormatEx(Format("" & rsTemp!数量, "0.00000"), 5)
            .TextMatrix(.Rows - 1, COL_VsCharge.固定) = IIf(Val("" & rsTemp!固定) = 0, "", "√")
            .TextMatrix(.Rows - 1, COL_VsCharge.从项) = IIf(Val("" & rsTemp!从项) = 0, "", "√")
            .TextMatrix(.Rows - 1, COL_VsCharge.停用) = IIf(Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01", "停用", "")
            
            Select Case rsTemp!收费方式
            Case 0
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "0-正常收取"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "1-检验试管费用"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "2-一次发送只收取一次"
            Case 3
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "3-当天只收取一次"
            Case 4
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "4-当天未执行收取一次"
            Case 5
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "5-当天只收取一次，排斥其他项目"
            Case 6
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "6-当天未执行收取一次，排斥其他项目"
            Case Else
                .TextMatrix(.Rows - 1, COL_VsCharge.收费方式) = "0-正常收取"
            End Select

            If Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                lngForeColor = &HFF&
            Else
                lngForeColor = &H0&
            End If
            .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = lngForeColor
            
            curTotal = curTotal + Val("" & rsTemp!价格) * Val("" & rsTemp!数量)
            rsTemp.MoveNext
        Loop
        If curTotal <> 0 Then strInfo = strInfo & " 所有科室对照合计：" & FormatEx(Format(curTotal, "0.0000"), 4)
    
    End With
    
    Call SetVsfCharge
    
    '-- 合计
    
    Me.stbThis.Panels(2).Text = "该分类共有" & Me.vsClinic.Rows - 1 & "种项目" & strInfo

    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveData(ByVal lngDeptID As Long, Optional blnShowInfo As Boolean = True) As Boolean
    '保存业务数据
    'lngDeptID : 病区ID
    'blnShowInfo:是否显示完成提示
    If mlngClinicID = 0 Then
        MsgBox "未正确指定诊疗项目！"
        Exit Function
    End If
    
    '校验从属：可以全部为主项(相当于不是套餐)，但如果存在从项，则只能有且必须有一个主项，且该主项必须为固定项目(不能删除)。
    Dim bln存在从项 As Boolean
    Dim int主项数 As Integer
    Dim int主项所在行 As Integer
    Dim intRows As Integer
    Dim rs As New ADODB.Recordset
    Dim intCount As Integer
    
    Err = 0: On Error GoTo ErrHand
    With vsCharge
        For intCount = .FixedRows To .Rows - 1
            If .TextMatrix(intCount, COL_VsCharge.从项) = "√" Then
                bln存在从项 = True
                Exit For
            End If
        Next
        If bln存在从项 Then
            For intCount = .FixedRows To .Rows - 1
                If .TextMatrix(intCount, COL_VsCharge.从项) <> "√" Then
                    int主项所在行 = intCount
                    int主项数 = int主项数 + 1
                    If int主项数 > 1 Then
                        If blnShowInfo Then MsgBox "提示：只能允许一个主项。"
                        Exit Function
                    End If
                End If
            Next
            
            If int主项数 = 1 Then
                If .TextMatrix(int主项所在行, COL_VsCharge.固定) <> "√" Then
                    If blnShowInfo Then MsgBox "提示：第" & int主项所在行 & "行是主项，必须为固定项目。"
                    Exit Function
                End If
            End If
            If int主项数 = 0 Then
                If blnShowInfo Then MsgBox "提示：必须要有一个主项。"
                Exit Function
            End If
        End If
    End With
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngChargeID As Long
    '检查主项的价格是否存在多个收入项目，如果有则提示，不能保存
    If bln存在从项 Then
        lngChargeID = Val(Me.vsCharge.TextMatrix(int主项所在行, COL_VsCharge.ID))
        gstrSql = "Select Id From 收费价目 Where 收费细目id=[1] And 执行日期 <= SYSDATE AND (终止日期 > SYSDATE OR 终止日期 IS NULL) " & _
                IIf(gstrPriceClass = "", " And 价格等级 Is Null ", " And 价格等级 = [2] ")
        
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngChargeID, gstrPriceClass)
        If rs.RecordCount > 1 Then
            If blnShowInfo Then MsgBox "提示：主项的价格存在多个收入项目，不能保存。"
            Exit Function
        End If
        rs.Close
    End If
    
    Dim strCharges As String   '保存收费细目ID,用于检查重复项目
    Dim strContent() As String '保存 zl_诊疗收费_UPDATE 要用的 收费内容 参数
    strCharges = "": ReDim strContent(0) As String
    
    With Me.vsCharge
        For intCount = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(intCount, COL_VsCharge.项目名)) <> "" And Val(.TextMatrix(intCount, COL_VsCharge.ID)) <> 0 Then
                If Not IsNumeric(Nvl(.TextMatrix(intCount, COL_VsCharge.数量), "X")) Then
                    If blnShowInfo Then MsgBox intCount & IIf(.TextMatrix(intCount, COL_VsCharge.数量) = "", "行的数量不能为空", "行需填写数字。")
                    Exit Function
                End If
                
                '可能是0.000等
                If Int(.TextMatrix(intCount, COL_VsCharge.数量)) = 0 And .TextMatrix(intCount, COL_VsCharge.固定) = "√" Then
                    .TextMatrix(intCount, COL_VsCharge.固定) = ""
                    MsgBox intCount & "行的数量为0,不能设为固定项,已自动更正."
                    Exit Function
                End If
            
                If InStr(1, strCharges & ";", ";" & Val(.TextMatrix(intCount, COL_VsCharge.ID)) & ";") > 0 Then
                    If blnShowInfo Then MsgBox intCount & "行收费项目与前面的收费项目有重复！"
                    Exit Function
                End If
                strCharges = strCharges & ";" & Val(.TextMatrix(intCount, COL_VsCharge.ID))
                If strContent(UBound(strContent)) <> "" Then ReDim Preserve strContent(UBound(strContent) + 1)
                '以"|"分隔的诊疗收费内容,每条记录按"诊疗项目ID^数量^固定^从项^性质^部位^检查方法^收费方式"组织
                strContent(UBound(strContent)) = Val(.TextMatrix(intCount, COL_VsCharge.ID)) & "^" & Val(.TextMatrix(intCount, COL_VsCharge.数量)) & "^" & IIf(Trim(.TextMatrix(intCount, COL_VsCharge.固定)) = "", 0, 1) & "^" & IIf(Trim(.TextMatrix(intCount, COL_VsCharge.从项)) = "", 0, 1) & "^0^^ " & Val(Mid(.TextMatrix(intCount, COL_VsCharge.收费方式), 1, 1))
            End If
        Next
    End With

    'If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    
    Dim lngCount As Long ' 总个数
    Dim lngLoop As Long, lngEndloop As Long
    Dim strItem As String, blnBeginTrans As Boolean, i As Integer
    
    lngCount = UBound(strContent)
    lngEndloop = 0
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For lngLoop = 0 To lngCount
        
        strItem = strItem & "|" & strContent(lngLoop)
        If i >= 40 Then
            strItem = Mid(strItem, 2)
            
            gstrSql = "zl_诊疗收费_UPDATE(" & mlngClinicID & "," & mlng计价性质 & ",'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & mint场合 & ")"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            i = 0: strItem = ""
            lngEndloop = lngEndloop + 1
        End If
        i = i + 1
    Next
    
    If Left(strItem, 1) = "|" Then
        strItem = Mid(strItem, 2)
        
        gstrSql = "zl_诊疗收费_UPDATE(" & mlngClinicID & "," & mlng计价性质 & ",'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & mint场合 & ")"
        Err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    If lngLoop = 0 Then '11303 不能全部删除对照的收费项目
        gstrSql = "zl_诊疗收费_UPDATE(" & mlngClinicID & "," & mlng计价性质 & ",'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & mint场合 & ")"
        Err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    SaveData = True
    If blnShowInfo Then MsgBox "收费对照保存成功！"
    
    Exit Function

ErrHand:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectCharge(ByVal Row As Long, ByVal Col As Long)
    '提取收费项目
    
    '只提取 收费项目目录.服务对象为 2，3 且在用的项目。
    '
    Dim rsTmp As New ADODB.Recordset
    Dim strSql   As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean
    Dim i As Integer
    On Error GoTo ErrHandle
    
    If Col = COL_VsCharge.项目名 Then
        '提取项目
        '--------------------------------------------------------------------------------------
            strInput = DelInvalidChar(UCase(Trim(txtEdit)))
            If InStr(strInput, " ") > 0 Then
                strInput = Trim(Split(strInput, " ")(0))
            End If
            strSql = "Select distinct i.* " & vbNewLine & _
                    "From (Select Distinct i.Id, i.编码, i.名称, i.规格, i.产地, i.计算单位," & vbNewLine & _
                    "                       Decode(Nvl(i.是否变价, 0), 0, LTrim(RTrim(To_Char(Nvl(d.现价, 0), '9999999990.0000')))," & vbNewLine & _
                    "                               Decode(Instr('4567', 类别), 0, LTrim(RTrim(To_Char(Nvl(d.缺省价格, 0), '9999999990.0000'))), '时价')) As 售价" & vbNewLine & _
                    "       From 收费项目目录 I, 收费价目 D" & vbNewLine & _
                    "       Where i.Id = d.收费细目id(+) And" & vbNewLine & _
                    "             (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And d.执行日期 <= Sysdate And" & vbNewLine & _
                    "             (d.终止日期 > Sysdate Or d.终止日期 Is Null) And (i.服务对象 = [3] Or i.服务对象 = 3) " & vbNewLine & _
                    IIf(gstrPriceClass = "", " And D.价格等级 Is Null ", " And D.价格等级 = [4] ") & vbNewLine & _
                    "      ) I, 收费项目别名 N" & vbNewLine & _
                    "Where i.Id = n.收费细目id And Rownum<2000 "

            If strInput <> "" Then
                strSql = strSql & " and (I.编码 like [1] " & _
                        "           or N.名称 like [2] " & _
                        "           or N.简码 like [2])"

            
            End If
            With vsCharge
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
            End With

            vRect = zlControl.GetControlRect(txtEdit.hWnd)
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "费用项目", False, "", "选择费用项目", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True, _
                                                 strInput & "%", gstrMatch & strInput & "%", mint场合, gstrPriceClass)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vsCharge
                        .EditText = "[" & Trim("" & rsTmp.Fields("编码") & "]" & rsTmp.Fields("名称"))
                        .TextMatrix(.Row, COL_VsCharge.项目名) = "[" & Trim("" & rsTmp.Fields("编码") & "]" & rsTmp.Fields("名称"))
                        .TextMatrix(.Row, COL_VsCharge.ID) = "" & rsTmp.Fields("ID")
                        .TextMatrix(.Row, COL_VsCharge.规格) = "" & rsTmp.Fields("规格")
                        .TextMatrix(.Row, COL_VsCharge.价格) = "" & rsTmp.Fields("售价")
                        .TextMatrix(.Row, COL_VsCharge.单位) = "" & rsTmp.Fields("计算单位")
                    End With
                End If
                Set rsTmp = Nothing
            End If
            txtEdit = ""
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub ExecuteFind(Optional ByVal blnNext As Boolean)
'功能：查找(下一个)病人
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim lngRow As Long, lngStart As Long
    Dim strFind As String, blnHave As Boolean
    strFind = IIf(gstrMatch = "%", "*", "") & DelInvalidChar(Trim(txtFind.Text))
            
    '开始查找行
    With vsClinic
        If blnNext Then
            If .Row + 1 <= .Rows - 1 Then
                lngStart = .Row + 1
            Else
                MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的项目。", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            lngStart = .FixedRows
        End If
        For lngRow = lngStart To .Rows - 1

            If UCase(.TextMatrix(lngRow, COL_VsClinic.名称)) Like UCase(strFind) & "*" _
               Or UCase(.TextMatrix(lngRow, COL_VsClinic.编码)) Like UCase(strFind) & "*" _
               Or UCase(.TextMatrix(lngRow, COL_VsClinic.简码)) Like UCase(strFind) & "*" _
            Then Exit For
        Next
        
        If lngRow <= .Rows - 1 Then
            '该行选中且显示在可见区域,并引发SelectionChanged事件
            .Select lngRow, COL_VsClinic.名称
            .ShowCell lngRow, COL_VsClinic.编码
        Else
            MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的项目。", vbInformation, gstrSysName
        End If
    End With


End Sub

Private Sub DeptCopy()
    '复制本科室的该项目对照到其他科室
    Dim strSql As String
    Dim rsDept As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInfo As String
    Dim varDept As Variant, strReturn  As String, strLine As String, i As Integer
    
    On Error GoTo ErrHandle
    If mint场合 = 2 Then
        strSql = "Select Distinct a.编码, a.名称, ID " & vbNewLine & _
            "From 部门表 A, 部门人员 B, 上机人员表 C, 部门性质说明 D" & vbNewLine & _
            "Where a.Id = b.部门id And b.人员id = c.人员id And a.Id = d.部门id And (d.服务对象 = 2 Or d.服务对象=3) And d.工作性质 = '护理' And A.ID<>[1] And c.用户名 = User"
    Else
        strSql = "Select Distinct a.编码, a.名称, ID " & vbNewLine & _
            "From 部门表 A, 部门人员 B, 上机人员表 C, 部门性质说明 D" & vbNewLine & _
            "Where a.Id = b.部门id And b.人员id = c.人员id And a.Id = d.部门id And (d.服务对象 = 1 Or d.服务对象=3) And Instr('检验,检查,手术,治疗,营养', d.工作性质) > 0 And A.ID<>[1] And c.用户名 = User"
    
    End If
    Set rsDept = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    strReturn = frmSelCur.ShowCurrSel(Me, rsDept, "编码,1200,0,2;名称,1800,0,2;ID,0,1,2", "选择" & IIf(mint场合 = 2, "病区", "科室"), True, , , 5000, True)
    If strReturn = "" Then Exit Sub
    varDept = Split(strReturn, "|")
    
    strInfo = ""
    For i = LBound(varDept) To UBound(varDept)
        '检验是否已设了对照，没有才能复制
        strLine = varDept(i)
        If UBound(Split(strLine, ",")) = 2 Then
            strSql = "Select 收费项目ID From 诊疗收费关系 Where 病人来源=[2] And 适用科室ID=[1] and 诊疗项目ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(Split(strLine, ",")(2)), mlngClinicID, mint场合)
            If rsTmp.EOF Then
               Call SaveData(CLng(Split(strLine, ",")(2)), False)
            Else
               strInfo = IIf(strInfo = "", "", vbNewLine) & "" & Split(strLine, ",")(0) & " " & Split(strLine, ",")(1) & " 该项目已经设定了费用！"
            End If
        End If
    Next
    If strInfo <> "" Then
        MsgBox strInfo
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CheckIsNoCharge()
    '如果从未对照过费用，是则复制所有科室的对照。
    Dim blnNoCharge As Boolean, intRow As Integer, intCol As Integer
    blnNoCharge = True
    With vsCharge
        For intRow = .FixedRows To .Rows - 1
             If Val(.TextMatrix(intRow, COL_VsCharge.ID)) <> 0 Then
                blnNoCharge = False
                Exit For
             End If
        Next
    End With
    If blnNoCharge Then
        If Me.vsCharge1.Rows < 2 Then
            Exit Sub
        ElseIf Me.vsCharge1.Rows = 2 Then
            If Me.vsCharge1.TextMatrix(1, COL_VsCharge.项目名) = "" Then Exit Sub
        End If
        If MsgBox("当前病区无收费项目，是否自动复制全院的收费项目？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        mblnModify = True
        With vsCharge1
            For intRow = .FixedRows To .Rows - 1
                If Val(.TextMatrix(intRow, COL_VsCharge.ID)) <> 0 Then
                   If Val(vsCharge.TextMatrix(vsCharge.Rows - 1, COL_VsCharge.ID)) <> 0 Then vsCharge.Rows = vsCharge.Rows + 1
                   For intCol = .FixedCols To .Cols - 1
                    vsCharge.TextMatrix(vsCharge.Rows - 1, intCol) = .TextMatrix(intRow, intCol)
                   Next
                End If
            Next
        End With
    End If
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.vsClinic.Rows <= 1 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    'If zlReportToVSFlexGrid(Me.vsfPrint, Me.vsClinic) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vsClinic
    objPrint.Title.Text = "目录"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub SetVsfCharge()
    With vsCharge
        .Cell(flexcpBackColor, 0, COL_VsCharge.序号, .Rows - 1, COL_VsCharge.序号) = &H8000000F
        .Cell(flexcpBackColor, 0, COL_VsCharge.规格, .Rows - 1, COL_VsCharge.价格) = &H8000000F
        .Cell(flexcpBackColor, 0, COL_VsCharge.停用, .Rows - 1, COL_VsCharge.停用) = &H8000000F
    End With
End Sub
