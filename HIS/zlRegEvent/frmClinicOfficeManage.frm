VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmClinicOfficeManage 
   BorderStyle     =   0  'None
   Caption         =   "门诊诊室管理"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptData 
      Height          =   3105
      Left            =   570
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      _Version        =   589884
      _ExtentX        =   10610
      _ExtentY        =   5477
      _StockProps     =   0
      ShowGroupBox    =   -1  'True
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   6120
      MaxLength       =   100
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   150
      Top             =   300
      Width           =   405
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   510
      TabIndex        =   1
      Top             =   300
      Width           =   5895
      _Version        =   589884
      _ExtentX        =   10398
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "基础设置>门诊诊室设置"
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
Attribute VB_Name = "frmClinicOfficeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mRptHeadCol
    COL_ID = 0
    COL_站点
    COL_科室
    COL_编码
    COL_名称
    COL_简码
    COL_位置
    COL_闲忙标志
End Enum
Private mintFindType As Integer
Private mrsDoctorOffice As ADODB.Recordset

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
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
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
    cbrMenuBar.id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加诊室(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改诊室(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除诊室(&D)")
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
        If Val(Left(cbrControl.id, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.id, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加诊室", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改诊室", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除诊室", cbrControl.Index + 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "按科室过滤↓")
    objPopup.flags = xtpFlagRightAlign
    '被绑定的控件必须动态加载，因为工具栏一但被删除，被绑定的控件的句柄就会变成0
    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    If txtFind.UBound > 0 Then Unload txtFind(1)
    Load txtFind(1)
    objCustom.Handle = txtFind(1).Hwnd
    objCustom.flags = xtpFlagRightAlign
    
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
    Dim blnVisible As Boolean, blnEnable As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "门诊诊室设置")
    If rptData.SelectedRows.Count > 0 Then
        blnEnable = Not rptData.SelectedRows(0).GroupRow
    End If
      
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnable
    Case conMenu_View_FindType '查找方式
        Control.Caption = "按" & Decode(mintFindType, 0, "科室", 1, "诊室", 2, "站点", "科室") & "过滤↓"
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 9 '查找方式
        Control.Checked = Val(Right(Control.id, 2)) - 1 = mintFindType
    End Select
End Sub

Public Sub InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
        
    Select Case CommandBar.Parent.id
    Case conMenu_View_FindType
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "科室(&1)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "诊室(&2)"
                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "站点(&3)"
            End If
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frm As New frmClinicOfficeEdit, lngID As Long
    
    Err = 0: On Error GoTo ErrHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
        End If
    End If
    Select Case Control.id
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem
        Dim strNewItem As String
        If frm.ShowMe(Me, Fun_Add, , strNewItem) Then Call LoadData(, strNewItem)
    Case conMenu_Edit_Modify
        If frm.ShowMe(Me, Fun_Update, lngID) Then Call LoadData
    Case conMenu_Edit_Delete
        If ExcuteDelete() Then Call LoadData
    Case conMenu_View_Refresh
        Call GetRecords: Call ExecuteFilter
    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 3 '查找方式
        mintFindType = Val(Right(Control.id, 2)) - 1
        mcbsMain.RecalcLayout
        txtFind(1).Text = ""
        If txtFind(1).Visible And txtFind(1).Enabled Then txtFind(1).SetFocus
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecuteFilter()
    '过滤数据
    Dim strKey As String
    
    Err = 0: On Error GoTo ErrHandler
    Call zlControl.TxtSelAll(txtFind(1))
    
    If Not mrsDoctorOffice Is Nothing Then
        With mrsDoctorOffice
            If Trim(txtFind(1).Text) = "" Then
                .Filter = ""
            Else
                strKey = Replace(gstrLike, "%", "*") & UCase(txtFind(1).Text) & "*"
                Select Case mintFindType
                Case 0   '科室(简码)
                    .Filter = "科室 Like '" & strKey & "' Or 科室简码 Like '" & strKey & "'"
                Case 1   '诊室(简码)
                    .Filter = "名称 Like '" & strKey & "' Or 简码 Like '" & strKey & "'"
                Case 2   '站点
                    If Trim(txtFind(1).Text) = "全院" Then
                        .Filter = "站点名称=null"
                    Else
                        .Filter = "站点名称 Like '" & strKey & "'"
                    End If
                Case Else
                    .Filter = ""
                End Select
            End If
        End With
    End If
    If mintFindType = 8 Then mintFindType = 0 '清除
    Call LoadData(False)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ExcuteDelete() As Boolean
    '功能:执行删除操作
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim lngID As Long, str诊室 As String
    
    On Error GoTo ErrHandler
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function

    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    str诊室 = Trim(rptData.SelectedRows(0).Record(COL_名称).Value)
    
    If MsgBox("你确定要删除 " & str诊室 & " 吗？", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
    '检查，被使用的不能删除
    If CheckHaveUsed(lngID) Then
      MsgBox "当前诊室已被使用，不能删除！", vbInformation, gstrSysName: Exit Function
    End If

    'Zl_门诊诊室_Delete(
    strSQL = "Zl_门诊诊室_Delete("
    'Id_In 门诊诊室.Id%Type
    strSQL = strSQL & "" & lngID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    ExcuteDelete = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckHaveUsed(ByVal lng诊室ID As Long) As Boolean
    '检查当前诊室是否已被使用
    Dim strSQL As String, rs诊室 As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '检查原上班时段是否被使用，被使用的不能修改站点、号类、时间段
    '不能删除被使用的范围最广的那一个,被使用的时段只要有一个即可（不同站点，不同号类可能会有多个同名的时间段）
    '临床出诊号源诊室
    strSQL = "Select 1 From 临床出诊号源诊室 Where 诊室id = [1] And Rownum < 2"
    Set rs诊室 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng诊室ID)
    If Not rs诊室 Is Nothing Then
        If Not rs诊室.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '临床出诊诊室(固定规则、模板)
    strSQL = "Select 1 From 临床出诊诊室 Where 诊室id = [1] And Rownum < 2"
    Set rs诊室 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng诊室ID)
    If Not rs诊室 Is Nothing Then
        If Not rs诊室.EOF Then CheckHaveUsed = True: Exit Function
    End If
    
    '临床出诊诊室记录
    strSQL = "Select 1 From 临床出诊诊室记录 Where 诊室id = [1] And Rownum < 2"
    Set rs诊室 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng诊室ID)
    If Not rs诊室 Is Nothing Then
        If Not rs诊室.EOF Then CheckHaveUsed = True: Exit Function
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGrid()
    Dim i As Long
    Dim objCol As ReportColumn, lngIdx As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objItem As Field
    
    Err = 0: On Error GoTo ErrHandler
    With rptData
        .AutoColumnSizing = False '不使用自动列宽
        .AllowColumnRemove = False '不允许拖动删除诊室列
        .ShowGroupBox = True '显示分组框
        .ShowItemsInGroups = False '不显示已分组的列
        .MultipleSelection = False '不允许多行选择
'        .SetImageList Me.img16
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid '竖向表格线格式
            .HorizontalGridStyle = xtpGridSolid '横向表格线格式
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的内容..."
            .ShadeSortColor = .BackColor
            Set .CaptionFont = Me.Font
            Set .TextFont = Me.Font
            Set .PreviewTextFont = Me.Font
        End With
    End With

    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_站点, "站点", 50, True)
        Set objCol = .Add(COL_科室, "科室", 100, True)
        Set objCol = .Add(COL_编码, "编码", 60, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_名称, "名称", 100, True)
        Set objCol = .Add(COL_简码, "简码", 60, True)
        Set objCol = .Add(COL_位置, "位置", 100, True)
        Set objCol = .Add(COL_闲忙标志, "闲忙状态", 80, True): objCol.Alignment = xtpAlignmentCenter
        
        '动态加载用户扩展字段,113315
        lngIdx = COL_闲忙标志 + 1
        strSQL = "Select * From 门诊诊室 Where 1 = 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取门诊诊室表结构")
        For Each objItem In rsTemp.Fields
            If InStr(",ID,编码,名称,简码,位置,缺省标志,站点,", "," & UCase(objItem.Name) & ",") = 0 Then
                Set objCol = .Add(lngIdx, objItem.Name, 100, True): lngIdx = lngIdx + 1
                If objItem.Name Like "是否*" Or (objItem.Type = adNumeric And objItem.Precision = 1) Then
                    objCol.Alignment = xtpAlignmentCenter
                End If
            End If
        Next
    End With
    With rptData
        '将站点和科室分组
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_站点)
        .GroupsOrder.Add .Columns(COL_科室)
        .Columns(COL_站点).Visible = False
        .Columns(COL_科室).Visible = False
        
        '将站点+编码排序(升序)
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(COL_站点)
        .SortOrder.Add .Columns(COL_编码)
        .SortOrder(0).SortAscending = True
        .SortOrder(1).SortAscending = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetRecords()
    '读取记录
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = _
        "Select d.名称 As 站点名称, c.名称 As 科室, c.简码 As 科室简码, a.*" & vbNewLine & _
        " From 门诊诊室 A, 门诊诊室适用科室 B, 部门表 C, Zlnodelist D" & vbNewLine & _
        " Where a.Id = b.诊室id(+) And b.科室id = c.Id(+) And a.站点 = d.编号(+)" & vbNewLine & _
        "       And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))"
    Set mrsDoctorOffice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function LoadData(Optional ByVal blnReRead As Boolean = True, _
    Optional ByVal strNewItem As String) As Boolean
    '加载数据
    '入参：
    '   blnReRead 是否重新读取数据
    '   strNewItem 新增诊室名称，用于定位
    Dim i As Long, j As Long
    Dim lngSelectRow As Long
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objField As Field
    
    Err = 0: On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).Index
    rptData.Records.DeleteAll
    
    If mrsDoctorOffice Is Nothing Then
        Call GetRecords
    ElseIf mrsDoctorOffice.State <> adStateOpen Then
        Call GetRecords
    ElseIf blnReRead Then
        Call GetRecords
    End If
    
    Do While Not mrsDoctorOffice.EOF
        Set objRecord = rptData.Records.Add()
        With objRecord
            Set objItem = .AddItem(Val(Nvl(mrsDoctorOffice!id)))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!站点名称, "全院"))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!科室, "所有"))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!编码))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!名称))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!简码))
            Set objItem = .AddItem(Nvl(mrsDoctorOffice!位置))
            Set objItem = .AddItem(IIf(Val(Nvl(mrsDoctorOffice!缺省标志)) = 0, "闲", "忙"))
            
            '动态加载用户扩展字段,113315
            For Each objField In mrsDoctorOffice.Fields
                If InStr(",站点名称,科室,科室简码,ID,编码,名称,简码,位置,缺省标志,站点,", "," & UCase(objField.Name) & ",") = 0 Then
                    If objField.Name Like "是否*" Or (objField.Type = adNumeric And objField.Precision = 1) Then
                        Set objItem = .AddItem(IIf(Nvl(objField.Value) = "1", "√", ""))
                    ElseIf objField.Type = adDate Or objField.Type = adDBTimeStamp _
                        Or objField.Type = adDBDate Or objField.Type = adDBTime Then
                        Set objItem = .AddItem(Format(Nvl(objField.Value), "yyyy-mm-dd"))
                    Else
                        Set objItem = .AddItem(Nvl(objField.Value))
                    End If
                End If
            Next
        End With
        
        mrsDoctorOffice.MoveNext
    Loop

    Call rptData.Populate '发布数据以更新界面
    With rptData
        If .Rows.Count > 0 Then '该行选中且显示在可见区域
            If strNewItem <> "" Then
                For i = 0 To rptData.Rows.Count - 1
                    If Not rptData.Rows(i).GroupRow Then
                        If rptData.Rows(i).Record(COL_名称).Value = strNewItem Then
                            rptData.FocusedRow = rptData.Rows(i)
                            Exit For
                        End If
                    End If
                Next
            Else
                If lngSelectRow = 0 Then
                    .FocusedRow = .Rows(0)
                ElseIf lngSelectRow > .Rows.Count - 1 Then
                    .FocusedRow = .Rows(.Rows.Count - 1)
                Else
                    .FocusedRow = .Rows(lngSelectRow)
                End If
            End If
        End If
    End With
    
    Call SetReportControlBackColorAlternate(rptData)
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

Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then
        sccTitle.SetFocus
    ElseIf Not Me.ActiveControl Is txtFind(1) Then
        rptData.SetFocus
    End If
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo ErrHandler
    Call InitGrid
    RestoreWinState Me, App.ProductName
    
    Dim strFindType As String
    Call GetRegInFor(g私有模块, Me.Name, "FindType", strFindType)
    mintFindType = Val(strFindType)
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
    With rptData
        .Left = 10: .Top = sccTitle.Top + sccTitle.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call SaveRegInFor(g私有模块, Me.Name, "FindType", mintFindType)
    If Not mrsDoctorOffice Is Nothing Then Set mrsDoctorOffice = Nothing
End Sub

Private Sub rptData_ColumnOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
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

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim frm As New frmClinicOfficeEdit, lngID As Long
    
    Err = 0: On Error GoTo ErrHandler
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            lngID = rptData.SelectedRows(0).Record(COL_ID).Value
            If zlStr.IsHavePrivs(mstrPrivs, "门诊诊室设置") Then
                If frm.ShowMe(Me, Fun_Update, lngID) Then Call LoadData '刷新数据
            Else
                frm.ShowMe Me, Fun_View, lngID
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_SortOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub zlDataPrint(BytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    If UserInfo.姓名 = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    objOut.Title.Text = "门诊诊室清单"
    '将ReportControl转换为VSFlexGrid
    Set objOut.Body = GetVsfGridData(rptData, CStr(COL_ID))

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

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub

Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call ExecuteFilter
        If rptData.Visible Then rptData.SetFocus
    End If
End Sub

Private Sub txtFind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then
        '按了右键菜单快捷键，清除粘贴板内容
        If Clipboard.GetText <> "" Then Clipboard.Clear
    End If
End Sub

Private Sub txtFind_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtFind(Index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtFind(Index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtFind_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtFind(Index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
