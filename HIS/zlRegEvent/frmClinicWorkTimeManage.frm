VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicWorkTimeManage 
   BorderStyle     =   0  'None
   Caption         =   "上班时间管理"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsfWorkTime 
      Height          =   2955
      Left            =   870
      TabIndex        =   0
      Top             =   1110
      Width           =   7035
      _cx             =   12409
      _cy             =   5212
      Appearance      =   0
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicWorkTimeManage.frx":0000
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
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   180
      Top             =   120
      Width           =   405
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   630
      TabIndex        =   1
      Top             =   540
      Width           =   7905
      _Version        =   589884
      _ExtentX        =   13944
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "基础设置>上班时间管理"
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
Attribute VB_Name = "frmClinicWorkTimeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mGridHead
    COL_站点 = 0
    COL_号类
    Col_时间段
    COL_上班时间
    COL_休息时间
    COL_出诊预留时间
    COL_缺省预约时间
    COL_提前挂号时间
End Enum

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
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加时间段(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改时间段(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除时间段(&D)")
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
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加时间段", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改时间段", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除时间段", cbrControl.index + 1)
        .Item(cbrControl.index + 1).BeginGroup = True
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
    Dim blnEnabled As Boolean, blnVisible As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnEnabled = Not vsfWorkTime.IsSubtotal(vsfWorkTime.Row) And vsfWorkTime.Rows > 1
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "时间段设置")

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfWorkTime.Rows > vsfWorkTime.FixedRows
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem '增加
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify '修改
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_Delete '删除
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim str站点 As String, str号类 As String, str时间段 As String
    Dim frmEdit As frmClinicWorkTimeEdit
    
    Err = 0: On Error GoTo ErrHandler
    
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem '增加上班时间
        Set frmEdit = New frmClinicWorkTimeEdit
        If frmEdit.ShowMe(Me, Fun_Add) Then Call LoadData: Set grsWorkTime = Nothing '重新取上班时段
    Case conMenu_Edit_Modify '调整上班时间
        With vsfWorkTime
            If .Row <= 0 Then Exit Sub
            If .IsSubtotal(.Row) Then Exit Sub
            
            str站点 = .Cell(flexcpData, .Row, COL_站点)
            str号类 = Trim(.TextMatrix(.Row, COL_号类))
            str时间段 = .TextMatrix(.Row, Col_时间段)
            Set frmEdit = New frmClinicWorkTimeEdit
            If frmEdit.ShowMe(Me, Fun_Update, str站点, str号类, str时间段) Then Call LoadData: Set grsWorkTime = Nothing '重新取上班时段
        End With
    Case conMenu_Edit_Delete '删除上班时间
        If ExcuteDelete() Then Call LoadData: Set grsWorkTime = Nothing '重新取上班时段
    Case conMenu_View_Refresh '刷新数据
        Call LoadData
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
    Dim str站点 As String, str号类 As String, str时间段 As String
    Dim blnUsed As Boolean
    On Error GoTo errHandle
    
    With vsfWorkTime
        If .Row <= 0 Then Exit Function
        If .IsSubtotal(.Row) Then Exit Function
        
        str站点 = .Cell(flexcpData, .Row, COL_站点)
        str号类 = Trim(.TextMatrix(.Row, COL_号类))
        str时间段 = .TextMatrix(.Row, Col_时间段)
    End With
    
    strSQL = "Select 1 From 临床出诊号源限制 Where 上班时段 = [1] And Rownum < 2" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 1 From 临床出诊限制 Where 上班时段 = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str时间段)
    If rsTemp.EOF Then
        If MsgBox("你确定要删除上班时间段（" & str时间段 & "）吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        blnUsed = True
        If MsgBox("注意：" & vbCrLf & _
                  "    上班时间段（" & str时间段 & "）可能已被使用，删除后你需要对所有使用了该上班时间段且启用了分时段的安排进行重新划分时段，否则，可能会导致预约挂号出错！" & vbCrLf & _
                  vbCrLf & _
                  "    你确定要删除上班时间段（" & str时间段 & "）吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    '删除有效性检查
    If CheckHaveUsed(str站点, str号类, str时间段) Then
        MsgBox "当前上班时间段已被使用，不能删除！", vbInformation, gstrSysName
        Exit Function
    End If
    'Zl_上班时段_Delete(
    strSQL = "Zl_上班时段_Delete("
    '站点_In   时间段.站点%Type,
    strSQL = strSQL & "'" & str站点 & "',"
    '号类_In   时间段.号类%Type,
    strSQL = strSQL & "'" & str号类 & "',"
    '时间段_In 时间段.时间段%Type
    strSQL = strSQL & "'" & str时间段 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If blnUsed Then
        MsgBox "注意：" & vbCrLf & _
               "    上班时间段（" & str时间段 & "）已被删除，请及时对所有使用了该上班时间段且启用了分时段的安排进行重新划分时段，否则，可能会导致预约挂号出错！", vbExclamation, gstrSysName
    End If
    
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckHaveUsed(ByVal str站点 As String, ByVal str号类 As String, ByVal str时间段 As String) As Boolean
    '检查当前上班时间段是否已被使用
    Dim strSQL As String, rs上班时间 As ADODB.Recordset
    Dim varTims As Variant, varRow As Variant
    
    Err = 0: On Error GoTo ErrHandler
    '检查原上班时段是否被使用，被使用的不能修改站点、号类、时间段
    '不能删除被使用的范围最广的那一个,被使用的时段只要有一个即可（不同站点，不同号类可能会有多个同名的时间段）
    '临床出诊号源限制
    strSQL = "Select 1" & vbNewLine & _
            " From (Select b.上班时段, c.站点, a.号类," & vbNewLine & _
            "              Row_Number() Over(Partition By b.上班时段 Order By b.上班时段, c.站点 Desc, a.号类 Desc) As 组号" & vbNewLine & _
            "        From 临床出诊号源 A, 临床出诊号源限制 B, 部门表 C" & vbNewLine & _
            "        Where a.Id = b.号源id And a.科室id = c.Id)" & vbNewLine & _
            " Where 组号 = 1 And Nvl(站点, '-') = Nvl([1], '-') And Nvl(号类, '-') = Nvl([2], '-') And 上班时段 = [3] And Rownum < 2"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str站点, str号类, str时间段)
    If Not rs上班时间 Is Nothing Then
        If Not rs上班时间.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '临床出诊限制(固定规则、模板)
    strSQL = "Select 1" & vbNewLine & _
            " From (Select a.上班时段, c.站点, b.号类," & vbNewLine & _
            "              Row_Number() Over(Partition By a.上班时段 Order By a.上班时段, c.站点 Desc, b.号类 Desc) As 组号" & vbNewLine & _
            "        From 临床出诊限制 A, 临床出诊安排 D, 临床出诊号源 B, 部门表 C" & vbNewLine & _
            "        Where a.安排id = d.Id And d.号源id = b.Id And b.科室id = c.Id)" & vbNewLine & _
            " Where 组号 = 1 And Nvl(站点, '-') = Nvl([1], '-') And Nvl(号类, '-') = Nvl([2], '-') And 上班时段 = [3] And Rownum < 2"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str站点, str号类, str时间段)
    If Not rs上班时间 Is Nothing Then
        If Not rs上班时间.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '临床出诊记录
    '不检查，因为该表太大，其次上班时段的信息都保存在了这个表中，没有找到上班时段时可由这个表的数据来提取
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
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    Call InitGrid
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitGrid()
    Dim strHead As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strHead = "站点,4,0|号类,4,1000|时间段,4,1000|上班时间,4,1500|休息时间,4,1500|出诊预留时间,4,1200|" & _
            "缺省预约时间,4,1200|提前挂号时间,4,1200"
    With vsfWorkTime
        .Redraw = False
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AutoSizeMode = flexAutoSizeRowHeight
        .RowHeightMin = 300
        .WordWrap = True
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        Call RestoreFlexState(vsfWorkTime, App.ProductName & "\" & Me.Name)
        .Redraw = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadData() As Boolean
    '加载上班时间段数据
    Dim i As Long, lngRow As Long, strSQL As String
    Dim rs上班时间 As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    strSQL = "Select a.时间段, a.号类, a.休息时段, a.开始时间, a.终止时间," & vbNewLine & _
            "        a.缺省时间, a.提前时间, a.出诊预留时间," & vbNewLine & _
            "        b.编号, b.名称 As 站点" & vbNewLine & _
            " From 时间段 A, Zlnodelist B" & vbNewLine & _
            " Where a.站点 = b.编号(+)" & vbNewLine & _
            " Order By Nvl(b.编号, -1), Nvl(a.号类, -1)"
    Set rs上班时间 = zlDatabase.OpenSQLRecord(strSQL, "获取上班时间段")
    With vsfWorkTime
        lngRow = .Row
        .Redraw = False
        .Clear 1
        .Subtotal flexSTClear
        .Rows = rs上班时间.RecordCount + 1
        i = 1
        Do While Not rs上班时间.EOF
            .TextMatrix(i, COL_站点) = Nvl(rs上班时间!站点, "全院")
            .Cell(flexcpData, i, COL_站点) = Nvl(rs上班时间!编号)
            .TextMatrix(i, COL_号类) = Nvl(rs上班时间!号类, " ")
            .TextMatrix(i, Col_时间段) = Nvl(rs上班时间!时间段)
            .TextMatrix(i, COL_上班时间) = Format(Nvl(rs上班时间!开始时间), "hh:mm") & "-" & Format(Nvl(rs上班时间!终止时间), "hh:mm")
            .TextMatrix(i, COL_休息时间) = FormatStr(Nvl(rs上班时间!休息时段))
            .TextMatrix(i, COL_出诊预留时间) = IIf(Val(Nvl(rs上班时间!出诊预留时间)) = 0, "", Nvl(rs上班时间!出诊预留时间))
            .TextMatrix(i, COL_缺省预约时间) = Format(Nvl(rs上班时间!缺省时间), "hh:mm")
            .TextMatrix(i, COL_提前挂号时间) = Format(Nvl(rs上班时间!提前时间), "hh:mm")
            .RowData(i) = IIf(i Mod 2 = 0, vbWindowBackground, G_AlternateColor) '用于设置行交替颜色
            i = i + 1
            rs上班时间.MoveNext
        Loop
        .AutoSize 0, .Cols - 1
        
        '设置行交替颜色
        For i = 1 To .Rows - 1
            .Cell(flexcpBackColor, i, Col_时间段, i, .Cols - 1) = .RowData(i)
        Next
        
        Call DataSplitGroup '分组显示
        If .Rows > 1 Then '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            If lngRow = 0 Then
                .Row = 1
            ElseIf lngRow > .Rows - 1 Then
                .Row = .Rows - 1
            Else
                .Row = lngRow
            End If
        End If
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FormatStr(ByVal strIn As String) As String
    '格式（开始时间1-终止时间1; 开始时间2-终止时间2;….）
    Dim varRow As Variant, varCol As Variant
    Dim i As Integer
    Dim strReturn As String
    
    If strIn = "" Then FormatStr = "": Exit Function
    varRow = Split(strIn, ";")
    For i = 0 To UBound(varRow)
        strReturn = strReturn & vbCrLf
        varCol = Split(varRow(i), "-")
        strReturn = strReturn & Format(varCol(0), "hh:mm") & "-" & Format(varCol(1), "hh:mm")
    Next
    If strReturn <> "" Then strReturn = Mid(strReturn, 3)
    FormatStr = strReturn
End Function

Private Sub DataSplitGroup()
    Dim i As Integer, j As Integer

    Err = 0: On Error GoTo ErrHandler
    With vsfWorkTime
        .OutlineBar = flexOutlineBarComplete '返回/设置显示目录树的线条
        .OutlineCol = COL_号类 '外面的线列
        .Outline COL_号类
        
        .Subtotal flexSTClear
        .Subtotal flexSTNone, COL_站点, , , , , True, "%s", , True
        .SubtotalPosition = flexSTAbove
        .MergeCells = flexMergeRestrictRows
        .MergeCol(COL_号类) = True

        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then '是否已小计
                .MergeRow(i) = True
                .IsCollapsed(i) = flexOutlineExpanded '是否展开状态
                .Cell(flexcpText, i, 1, i, .Cols - 1) = .Cell(flexcpTextDisplay, i, 0) 'Flexcptextdisplay 单元格格式化了的文本内容(只读)
                .RowHeight(i) = 300
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = flexAlignLeftCenter
            End If
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 0, 0, Me.ScaleWidth
    With vsfWorkTime
        .Left = 10: .Top = sccTitle.Top + sccTitle.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(vsfWorkTime, App.ProductName & "\" & Me.Name)
End Sub


Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfWorkTime.Visible And vsfWorkTime.Enabled Then vsfWorkTime.SetFocus
End Sub

Private Sub vsfWorkTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    '设置选择行颜色
    Call SetVsGridRowChangeBackColor(vsfWorkTime, OldRow, NewRow, OldCol, NewCol, _
        vsfWorkTime.BackColorSel, Col_时间段, vsfWorkTime.Cols - 1)
End Sub

Private Sub vsfWorkTime_DblClick()
    Dim blnUpdate As Boolean
    Dim str站点 As String, str号类 As String, str时间段 As String
    Dim frmEdit As frmClinicWorkTimeEdit
    
    Err = 0: On Error GoTo ErrHandler
    With vsfWorkTime
        If .Row < 1 Then Exit Sub
        If .IsSubtotal(.Row) Then Exit Sub
        
        str站点 = .Cell(flexcpData, .Row, COL_站点)
        str号类 = Trim(.TextMatrix(.Row, COL_号类))
        str时间段 = .TextMatrix(.Row, Col_时间段)
        
        Set frmEdit = New frmClinicWorkTimeEdit
        If zlStr.IsHavePrivs(mstrPrivs, "时间段设置") Then
            '修改
            If frmEdit.ShowMe(Me, Fun_Update, str站点, str号类, str时间段) Then
                Call LoadData '刷新数据
                Set grsWorkTime = Nothing '重新取上班时段
            End If
        Else
            '查看
            frmEdit.ShowMe Me, Fun_View, str站点, str号类, str时间段
        End If
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfWorkTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    objOut.Title.Text = "上班时间段清单"
    Set objOut.Body = vsfWorkTime
    
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
