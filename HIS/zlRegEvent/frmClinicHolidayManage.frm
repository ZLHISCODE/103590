VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicHolidayManage 
   BorderStyle     =   0  'None
   Caption         =   "节假日管理"
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox pic调休情况 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   6780
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   3810
      Width           =   3195
      Begin VSFlex8Ctl.VSFlexGrid vsf调休情况 
         Height          =   1155
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   3015
         _cx             =   5318
         _cy             =   2037
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfWorkInfo 
         Height          =   1215
         Left            =   60
         TabIndex        =   7
         Top             =   2010
         Width           =   3015
         _cx             =   5318
         _cy             =   2143
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin XtremeSuiteControls.ShortcutCaption sccWorkInfo 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   1650
         Width           =   3105
         _Version        =   589884
         _ExtentX        =   5477
         _ExtentY        =   564
         _StockProps     =   6
         Caption         =   "上班信息"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   0
         GradientColorDark=   0
      End
      Begin XtremeSuiteControls.ShortcutCaption scc调休情况 
         Height          =   320
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3105
         _Version        =   589884
         _ExtentX        =   5477
         _ExtentY        =   564
         _StockProps     =   6
         Caption         =   "调休信息"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   0
         GradientColorDark=   0
      End
   End
   Begin zl9RegEvent.UserSelectPopup uspSelectYear 
      Height          =   315
      Left            =   390
      TabIndex        =   0
      Top             =   510
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      PopupWidth      =   1000
   End
   Begin zl9RegEvent.UserDatePicker dtpDay 
      Height          =   3045
      Left            =   180
      TabIndex        =   1
      Top             =   3900
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   5371
      HolidayStart    =   42379.5026851852
      TitleBackColor  =   -2147483626
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
      Height          =   2055
      Left            =   870
      TabIndex        =   5
      Top             =   1080
      Width           =   7905
      _cx             =   13944
      _cy             =   3625
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line LineX 
      BorderColor     =   &H8000000C&
      X1              =   6750
      X2              =   6750
      Y1              =   7530
      Y2              =   3720
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   180
      Top             =   1140
      Width           =   405
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   8625
      _Version        =   589884
      _ExtentX        =   15214
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "基础设置>节假日管理"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmClinicHolidayManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mGridHeadCol
    COl_节日 = 0
    COL_开始时间 = 1
    COL_结束时间 = 2
    COL_备注 = 3
    COL_允许预约 = 4
    COL_允许挂号 = 5
    
    COL_序号 = 0
    COL_原上班时间 = 1
    Col_调休时间 = 2
    
    COL_日期 = 0
    COL_挂号 = 1
    COL_预约 = 2
End Enum
Private mdatStart As Date, mdatEnd As Date, mvarWorks As Variant
Private mlngYear As Long

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo ErrHandler
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…", cbrControl.Index + 1)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加节假日(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改节假日(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除节假日(&D)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "刷新提醒(&B)", cbrControl.Index)
        cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加节假日", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改节假日", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除节假日", cbrControl.Index + 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnVisible As Boolean, blnEnabled As Boolean
    Dim dtTemp As Date
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "节假日设置")
    If vsfGrid.Row > 0 Then
        dtTemp = Val(uspSelectYear.SelectedKey) & "/" & vsfGrid.TextMatrix(vsfGrid.Row, COL_开始时间)
        blnEnabled = DateDiff("d", Now, dtTemp) > 0  '小于当前日期的不能删除、修改
    End If

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfGrid.Rows > vsfGrid.FixedRows
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strHolidayName As String
    Dim frmEdit As frmClinicHolidayEdit
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem
        Set frmEdit = New frmClinicHolidayEdit
        If frmEdit.ShowMe(Me, Fun_Add) Then
            RefrashData mlngYear '刷新数据
        End If
    Case conMenu_Edit_Modify
        If vsfGrid.Row < 1 Then Exit Sub
        
        strHolidayName = vsfGrid.TextMatrix(vsfGrid.Row, COl_节日)
        Set frmEdit = New frmClinicHolidayEdit
        If frmEdit.ShowMe(Me, Fun_Update, mlngYear, strHolidayName) Then
            RefrashData mlngYear '刷新数据
        End If
    Case conMenu_Edit_Delete
        If ExcuteDelete() Then
            RefrashData mlngYear '刷新数据
        End If
    Case conMenu_View_Refresh
        RefrashData mlngYear '刷新数据
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ExcuteDelete() As Boolean
    '功能:执行删除操作
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim strHolidayName As String
    
    On Error GoTo ErrHandler
    If vsfGrid.Row <= 0 Then Exit Function
    
    strHolidayName = vsfGrid.TextMatrix(vsfGrid.Row, COl_节日)
    
    If MsgBox("你确定要删除" & mlngYear & "年 " & strHolidayName & " 吗？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    '删除有效性检查
    strSQL = "Select 1" & vbNewLine & _
        "    From 临床出诊记录 A" & vbNewLine & _
        "    Where a.出诊日期 >= (Select 开始日期 From 法定假日表 Where 年份 = [1] And 节日名称 = [2] And 性质 = 0 And Rownum<2)" & vbNewLine & _
        "          And a.上班时段 Is Not Null And Nvl(a.是否发布, 0) = 1 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngYear, strHolidayName)
    If Not rsTemp.EOF Then
        MsgBox "当前节假日开始时间之后已有有效的出诊安排，不能删除！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_法定假日表_Delete(
    strSQL = "Zl_法定假日表_Delete("
    '年份_In     法定假日表.年份%Type,
    strSQL = strSQL & "" & mlngYear & ","
    '节日名称_In 法定假日表.节日名称%Type
    strSQL = strSQL & "'" & strHolidayName & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    ExcuteDelete = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGridHead()
    Dim strHead As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strHead = "节日,4,800|开始时间,4,1300|结束时间,4,1300|备注,1,4500|允许预约,1,0|允许挂号,1,0"
    With vsfGrid
        .Redraw = False
        .Rows = 1
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor '行交替色
        
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        Call RestoreFlexState(vsfGrid, App.ProductName & "\" & Me.Name)
        .Redraw = True
    End With

    strHead = "序号,4,500|原上班日期,4,1300|调休日期,4,1300"
    With vsf调休情况
        .Redraw = False
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor '行交替色
        .RowHeightMin = 300
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = True
    End With
    
    strHead = "日期,4,1300|允许挂号,4,1000|允许预约,4,1000"
    With vsfWorkInfo
        .Redraw = flexRDNone
        .FixedCols = 0: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
            If i > 0 Then
                .ColDataType(i) = flexDTBoolean
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDay_DayMetrics(Day As Date, Metrics As UserDatePickerDayMetrics)
    Dim dtTemp As Date, i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If CStr(Day) = "00:00:00" Then Exit Sub
    
    If Weekday(Day) = vbSunday Or Weekday(Day) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
    
    '标记休假日
    If DateDiff("d", Day, mdatStart) <= 0 And DateDiff("d", Day, mdatEnd) >= 0 Then
'        If HolidayIsWork(Day) = False Then
            Metrics.BackColor = &HC0E0FF
            Metrics.IsHoliday = True
'        End If
    End If
    
    '标记调班日
    For i = 1 To vsf调休情况.Rows - 1
        dtTemp = vsf调休情况.TextMatrix(i, Col_调休时间)
        If DateDiff("d", Day, dtTemp) = 0 Then
            Metrics.BackColor = &HFFFFC0
            Metrics.IsWorkFromHoliday = True
        End If
    Next
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function HolidayIsWork(ByVal Day As Date) As Boolean
    '检查节假日是否上班
    Dim i As Integer, j As Integer
    Dim var允许预约 As Variant, var允许挂号 As Variant
    
    Err = 0: On Error GoTo ErrHandler
    If vsfGrid.Row < 1 Then Exit Function
    
    Err = 0: On Error GoTo ErrHandler
    var允许预约 = Split(vsfGrid.TextMatrix(vsfGrid.Row, COL_允许预约), ";")
    var允许挂号 = Split(vsfGrid.TextMatrix(vsfGrid.Row, COL_允许挂号), ";")
    
    For j = 0 To UBound(var允许预约)
        If DateDiff("d", Day, var允许预约(j)) = 0 Then
            HolidayIsWork = True
            Exit Function
        End If
    Next
    
    For j = 0 To UBound(var允许挂号)
        If DateDiff("d", Day, var允许挂号(j)) = 0 Then
            HolidayIsWork = True
            Exit Function
        End If
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim varRow As Variant, varCol As Variant
    Dim i As Long, j As Long
    Err = 0: On Error GoTo ErrHandler
    Call InitGridHead
    scc调休情况.GradientColorDark = dtpDay.TitleBackColor
    scc调休情况.GradientColorLight = dtpDay.TitleBackColor
    
    mlngYear = Year(zlDatabase.Currentdate)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function RefrashData(Optional ByVal lngYear As Long) As Boolean
    Dim i As Long
    Dim lngMaxYear As Long, lngMinYear As Long
    Dim strSQL As String, rs年份 As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    lngMinYear = 3000: lngMaxYear = 1900
    
    strSQL = "Select Max(年份) As 年 From 法定假日表 Group By 年份"
    Set rs年份 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rs年份.EOF
        If lngMinYear > Val(Nvl(rs年份!年)) Then lngMinYear = Val(Nvl(rs年份!年))
        If lngMaxYear < Val(Nvl(rs年份!年)) Then lngMaxYear = Val(Nvl(rs年份!年))
        rs年份.MoveNext
    Loop
    
    If lngMinYear = 3000 Then lngMinYear = lngYear
    If lngMaxYear = 1900 Then lngMaxYear = lngYear
    If lngYear < lngMinYear Or lngYear > lngMaxYear Then
        lngYear = Year(zlDatabase.Currentdate)
    End If
    If lngYear < lngMinYear Then lngMinYear = lngYear
    If lngYear > lngMaxYear Then lngMaxYear = lngYear
    '为年份选择器加载数据
    uspSelectYear.Clear
    For i = lngMinYear To lngMaxYear
        uspSelectYear.AddItem i, i & "年"
    Next
    uspSelectYear.SelectedKey = lngYear '会触发uspSelectYear_ValueChanged事件
'    Call LoadData(lngYear)
    RefrashData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadData(Optional ByVal lngYear As Long) As Boolean
    Dim i As Long, j As Long, lngRow As Long
    Dim strSQL As String, rs节假日 As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    mdatStart = Empty: mdatEnd = Empty: mvarWorks = Empty
    vsf调休情况.Clear 1: vsf调休情况.Rows = 1
    vsfWorkInfo.Clear 1: vsfWorkInfo.Rows = 1
    dtpDay.RedrawControl
    
    strSQL = "Select 年份,节日名称,开始日期,终止日期,备注,允许预约日期,允许挂号日期 From 法定假日表" & vbNewLine & _
            " Where Nvl(性质,0)=0 And 年份=[1]" & vbNewLine & _
            " Order By 年份,开始日期"
    Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngYear)
    
    With rs节假日
        lngRow = vsfGrid.Row
        vsfGrid.Rows = .RecordCount + 1
        i = 1
        Do While Not .EOF
            uspSelectYear.SelectedKey = Nvl(!年份)
            vsfGrid.TextMatrix(i, COl_节日) = Nvl(!节日名称)
            vsfGrid.TextMatrix(i, COL_开始时间) = Format(Nvl(!开始日期), "mm-dd hh:mm")
            vsfGrid.Cell(flexcpData, i, COL_开始时间) = Nvl(!开始日期)
            vsfGrid.TextMatrix(i, COL_结束时间) = Format(Nvl(!终止日期), "mm-dd hh:mm")
            vsfGrid.Cell(flexcpData, i, COL_结束时间) = Nvl(!终止日期)
            vsfGrid.TextMatrix(i, COL_备注) = Nvl(!备注)
            vsfGrid.TextMatrix(i, COL_允许预约) = Nvl(!允许预约日期)
            vsfGrid.TextMatrix(i, COL_允许挂号) = Nvl(!允许挂号日期)
            i = i + 1
            .MoveNext
        Loop
    End With
    If vsfGrid.Rows > 1 Then
        vsfGrid.Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
        If lngRow = 0 Then
            vsfGrid.Row = 1
        ElseIf lngRow > vsfGrid.Rows - 1 Then
            vsfGrid.Row = vsfGrid.Rows - 1
        Else
            vsfGrid.Row = lngRow
        End If
    End If
    
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitDatePickerData(ByVal datStart As Date, ByVal datEnd As Date, Optional ByVal strWorks As String)
    '功能：根据节假日开始时间和结束时间，以及调班时间显示日历
    '参数：
    '   datStart - 开始时间
    '   datEnd - 结束时间
    '   varWorks - 调班(上班)时间，多个用"、"分隔
    Err = 0: On Error GoTo ErrHandler
    If datStart > datEnd Then '确定时间大小
        Dim datTemp As Date
        datTemp = datStart: datStart = datEnd: datEnd = datTemp
    End If
    mvarWorks = Empty
    If strWorks <> "" Then mvarWorks = Split(strWorks, "、")
    mdatStart = datStart: mdatEnd = datEnd
    
    dtpDay.HolidayStart = mdatStart '会触发RedrawControl()方法
'    dtpDay.RedrawControl
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    
    sccTitle.Move 10, 0, Me.ScaleWidth
    With uspSelectYear
        .Left = 10: .Top = sccTitle.Top + sccTitle.Height
        .Width = Me.ScaleWidth - .Left + 10
    End With
    
    With vsfGrid
        .Left = 0: .Top = uspSelectYear.Top + uspSelectYear.Height
        .Width = Me.ScaleWidth - .Left + 10
        .Height = Me.ScaleHeight * 7 / 12 - .Top - 10
    End With
    
    With dtpDay
        .Left = 10: .Top = vsfGrid.Top + vsfGrid.Height
        .Width = Me.ScaleWidth - 4000 - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
    LineX.X1 = dtpDay.Left + dtpDay.Width + 10: LineX.Y1 = dtpDay.Top
    LineX.X2 = dtpDay.Left + dtpDay.Width + 10: LineX.Y2 = dtpDay.Top + dtpDay.Height
    With pic调休情况
        .Left = dtpDay.Left + dtpDay.Width + 20: .Top = dtpDay.Top
        .Width = Me.ScaleWidth - .Left: .Height = dtpDay.Height
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(vsfGrid, App.ProductName & "\" & Me.Name)
End Sub

Private Sub pic调休情况_Resize()
    Err = 0: On Error Resume Next
    scc调休情况.Move 0, 0, pic调休情况.ScaleWidth
    With vsf调休情况
        .Left = 0: .Top = scc调休情况.Top + scc调休情况.Height
        .Width = pic调休情况.ScaleWidth + 20
    End With
    
    sccWorkInfo.Move 0, vsf调休情况.Top + vsf调休情况.Height, pic调休情况.ScaleWidth
    With vsfWorkInfo
        .Left = 0: .Top = sccWorkInfo.Top + sccWorkInfo.Height
        .Width = pic调休情况.ScaleWidth + 20
        .Height = pic调休情况.ScaleHeight - .Top + 10
    End With
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfGrid.Visible And vsfGrid.Enabled Then vsfGrid.SetFocus
End Sub

Private Sub uspSelectYear_ValueChanged(ByVal strKey As String, ByVal strValue As String)
    dtpDay.HolidayStart = strKey & "/02/01"
    mlngYear = Val(strKey)
    Call LoadData(mlngYear)
End Sub

Private Sub vsfGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    Dim strSQL As String, rs节假日 As ADODB.Recordset
    Dim strHolidayName As String
    
    If NewRow < 1 Or vsfGrid.Rows - 1 < NewRow Then Exit Sub
    strHolidayName = vsfGrid.TextMatrix(NewRow, COl_节日)
    With vsf调休情况
        strSQL = "Select 年份,节日名称,开始日期,终止日期,备注 From 法定假日表" & _
                " Where Nvl(性质,0)=1 And 年份=[1] And 节日名称=[2]"
        Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngYear, strHolidayName)
        .Rows = rs节假日.RecordCount + 1
        i = 1
        Do While Not rs节假日.EOF
            .TextMatrix(i, COL_序号) = i
            .TextMatrix(i, COL_原上班时间) = Format(Nvl(rs节假日!终止日期), "yyyy-mm-dd")
            .TextMatrix(i, Col_调休时间) = Format(Nvl(rs节假日!开始日期), "yyyy-mm-dd")
            .RowHeight(i) = .RowHeight(0)
            i = i + 1
            rs节假日.MoveNext
        Loop
    End With
    
    Call ShowDateRangeToGrid(vsfGrid.Cell(flexcpData, NewRow, COL_开始时间), vsfGrid.Cell(flexcpData, NewRow, COL_结束时间))
    Call LoadDateRegist(vsfGrid.TextMatrix(NewRow, COL_允许挂号), vsfGrid.TextMatrix(NewRow, COL_允许预约))
    
    If vsfGrid.TextMatrix(NewRow, COL_开始时间) <> "" Then
        InitDatePickerData vsfGrid.Cell(flexcpData, NewRow, COL_开始时间), _
              vsfGrid.Cell(flexcpData, NewRow, COL_结束时间)
    End If
End Sub

Private Sub vsfGrid_DblClick()
    Dim strHolidayName As String
    Dim dtTemp As Date
    Dim blnEnabled As Boolean
    Dim frmEdit As frmClinicHolidayEdit
    
    Err = 0: On Error GoTo ErrHandler
    If vsfGrid.Row <= 0 Then Exit Sub
    strHolidayName = vsfGrid.TextMatrix(vsfGrid.Row, COl_节日)
    If vsfGrid.Row > 0 Then
        dtTemp = mlngYear & "/" & vsfGrid.TextMatrix(vsfGrid.Row, COL_开始时间)
        blnEnabled = DateDiff("d", Now, dtTemp) > 0  '小于当前日期的不能删除
    End If
    
    Set frmEdit = New frmClinicHolidayEdit
    If zlStr.IsHavePrivs(mstrPrivs, "节假日设置") And blnEnabled Then
        If frmEdit.ShowMe(Me, Fun_Update, mlngYear, strHolidayName) Then Call RefrashData(mlngYear)   '刷新数据
    Else
        frmEdit.ShowMe Me, Fun_View, mlngYear, strHolidayName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlDataPrint(BytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If UserInfo.姓名 = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    objOut.Title.Text = mlngYear & "年节假日清单"
    Set objOut.Body = vsfGrid
    
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    If BytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, BytMode
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowDateRangeToGrid(ByVal dtStart As Date, dtEnd As Date)
    '显示日期到表格中
    Dim lngRow As Long, i As Integer
    Dim intCount As Integer
    
    Err = 0: On Error GoTo ErrHandler
    intCount = DateDiff("d", dtStart, dtEnd) '总天数
    With vsfWorkInfo
        .Clear 1
        .Rows = 1
        For i = 0 To intCount
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_日期) = Format(DateAdd("d", i, dtStart), "yyyy-mm-dd")
            .Cell(flexcpChecked, lngRow, COL_挂号, lngRow, COL_预约) = 2
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDateRegist(ByVal str允许挂号 As String, ByVal str允许预约 As String)
    '加载预约挂号情况
    Dim i As Integer, j As Integer
    Dim var允许预约 As Variant, var允许挂号 As Variant
    
    Err = 0: On Error GoTo ErrHandler
    var允许挂号 = Split(str允许挂号, ";")
    var允许预约 = Split(str允许预约, ";")
    With vsfWorkInfo
        For i = 1 To .Rows - 1
            For j = 0 To UBound(var允许挂号)
                If DateDiff("d", .TextMatrix(i, COL_日期), var允许挂号(j)) = 0 Then
                    .TextMatrix(i, COL_挂号) = 1
                End If
            Next
            For j = 0 To UBound(var允许预约)
                If DateDiff("d", .TextMatrix(i, COL_日期), var允许预约(j)) = 0 Then
                    .TextMatrix(i, COL_预约) = 1
                End If
            Next
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

