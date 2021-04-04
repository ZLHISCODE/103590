VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmLabMainSampleUnion 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2775
      Left            =   570
      TabIndex        =   0
      Top             =   30
      Width           =   5745
      _Version        =   589884
      _ExtentX        =   10134
      _ExtentY        =   4895
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl RptComPare1 
      Height          =   2775
      Left            =   450
      TabIndex        =   1
      Top             =   3090
      Width           =   3075
      _Version        =   589884
      _ExtentX        =   5424
      _ExtentY        =   4895
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl RptComPare2 
      Height          =   2805
      Left            =   4350
      TabIndex        =   2
      Top             =   3090
      Width           =   3075
      _Version        =   589884
      _ExtentX        =   5424
      _ExtentY        =   4948
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   6930
      Top             =   510
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   7080
      Top             =   2070
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabMainSampleUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event StartEdit(Cancel As Boolean)
Private Enum mCol
    标本ID
    姓名
    性别
    年龄
    检验项目
    标识号
    仪器
    标本时间
    检验人
    合并状态
End Enum
Private Enum mRCol
    检验项目
    结果
    单位
    标志
    参考
End Enum

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intLoop As Integer
    Dim intNext As Integer
    Dim lngID As Long
    
    Select Case Control.ID
        Case conMenu_Edit_Insert                        '设置合并标本
            If Not Me.rptList.FocusedRow Is Nothing Then
                lngID = Me.rptList.FocusedRow.Record.Item(mCol.标本ID).Value
                For intLoop = 1 To Me.rptList.Records.Count
                    '重新写入合并或被合并标本
                    If Me.rptList.Records(intLoop - 1).Item(mCol.标本ID).Value = lngID Then
                        Me.rptList.Records(intLoop - 1).Item(mCol.合并状态).Value = "合并标本"
                    Else
                        Me.rptList.Records(intLoop - 1).Item(mCol.合并状态).Value = "被合并标本"
                    End If
                    '设置颜色
                    For intNext = 0 To Me.rptList.Columns.Count
                        If Me.rptList.Records(intLoop - 1).Item(mCol.合并状态).Value = "合并标本" Then
                            Me.rptList.Records(intLoop - 1).Item(intNext).ForeColor = vbBlue
                        Else
                            Me.rptList.Records(intLoop - 1).Item(intNext).ForeColor = vbRed
                        End If
                    Next
                Next
                Me.rptList.Populate
            End If
        Case conMenu_Manage_ThingDel                    '清空所有合并标本
            Me.rptList.Records.DeleteAll
            Me.rptList.Populate
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = Me.rptList.hWnd
        Case 2
            Item.Handle = Me.RptComPare1.hWnd
        Case 3
            Item.Handle = Me.RptComPare2.hWnd
    End Select
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    dkpMain.Options.DefaultPaneOptions = PaneNoCloseable
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(1, 200, 50, DockTopOf, Nothing)
    Pane1.Title = "合并列表"
    Pane1.Handle = Me.rptList.hWnd
    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(2, 200, 150, DockBottomOf, Nothing)
    Pane2.Title = "合并结果清单"
    Pane2.Handle = Me.RptComPare1.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set Pane3 = dkpMain.CreatePane(3, 200, 150, DockRightOf, Pane2)
    Pane3.Title = "被合并结果清单"
    Pane3.Handle = Me.RptComPare2.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With Me.rptList.Columns
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "双击左边的列表中的病人增加到合并列表中..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
        Set Column = .Add(mCol.标本ID, "标本ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.姓名, "姓名", 75, True)
        Set Column = .Add(mCol.性别, "性别", 40, True)
        Set Column = .Add(mCol.年龄, "年龄", 40, True)
        Set Column = .Add(mCol.仪器, "仪器", 100, True)
        Set Column = .Add(mCol.标识号, "标识号", 80, True)
        Set Column = .Add(mCol.标本时间, "标本时间", 75, False)
        Set Column = .Add(mCol.检验人, "检验人", 75, True)
        Set Column = .Add(mCol.合并状态, "合并状态", 75, True)
    End With
    
    With Me.RptComPare1.Columns
        RptComPare1.AllowColumnRemove = False
        RptComPare1.ShowItemsInGroups = False
        
        With RptComPare1.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "合并的检验项目为空..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mRCol.检验项目, "检验项目", 85, True)
        Set Column = .Add(mRCol.结果, "结果", 85, True)
        Set Column = .Add(mRCol.单位, "单位", 65, True)
        Set Column = .Add(mRCol.标志, "标志", 65, True)
        Set Column = .Add(mRCol.参考, "参考", 85, True)
    End With
    
    With Me.RptComPare2.Columns
        RptComPare2.AllowColumnRemove = False
        RptComPare2.ShowItemsInGroups = False
        
        With RptComPare2.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "被合并的检验项目为空..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mRCol.检验项目, "检验项目", 85, True)
        Set Column = .Add(mRCol.结果, "结果", 85, True)
        Set Column = .Add(mRCol.单位, "单位", 65, True)
        Set Column = .Add(mRCol.标志, "标志", 65, True)
        Set Column = .Add(mRCol.参考, "参考", 85, True)
    End With
End Sub

Private Sub Form_Resize()
'    With Me.rptList
'        .Left = 0
'        .Top = 0
'        .Width = Me.ScaleWidth
'        .Height = Me.ScaleHeight
'    End With
End Sub
Public Function zlRefresh(ByVal lngSampleID As Long, strName As String, strSex As String, strAge As String, ItemName As String, _
                        lngPatientID As String, strMachineName As String, SampleTime As String, strVerifyName As String) As Boolean
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim intRowIndex As Integer
    
    zlRefresh = True
    
    intRowIndex = CheckName
    
    If intRowIndex > 0 And strName <> "" Then
        If MsgBox("已存在一个合并标本,是否覆盖?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            
            zlRefresh = False
            Exit Function
        End If
        Me.rptList.Records.RemoveAt intRowIndex - 1
    End If
    
    For intLoop = 1 To Me.rptList.Records.Count
        If Me.rptList.Records(intLoop - 1).Item(mCol.标本ID).Value = lngSampleID Then
            MsgBox "您已加入了个同样的标本!", vbQuestion, gstrSysName
            zlRefresh = False
            Exit Function
        End If
    Next
    
    Set Record = Me.rptList.Records.Add
    
    For intLoop = 0 To Me.rptList.Columns.Count
        Record.AddItem ""
    Next
    
    Record.Item(mCol.标本ID).Value = lngSampleID
    Record.Item(mCol.姓名).Value = strName
    Record.Item(mCol.性别).Value = strSex
    Record.Item(mCol.年龄).Value = strAge
    Record.Item(mCol.检验项目).Value = ItemName
    Record.Item(mCol.标识号).Value = lngPatientID
    Record.Item(mCol.仪器).Value = strMachineName
    Record.Item(mCol.标本时间).Value = SampleTime
    Record.Item(mCol.检验人).Value = strVerifyName
    
    If strName <> "" Then
        '当第一个时默认为合并标本
        Record.Item(mCol.合并状态).Value = "合并标本"
        For intLoop = 1 To Me.rptList.Columns.Count
            Record.Item(intLoop).ForeColor = vbBlue
        Next
    Else
        Record.Item(mCol.合并状态).Value = "被合并标本"
        For intLoop = 1 To Me.rptList.Columns.Count
            Record.Item(intLoop).ForeColor = vbRed
        Next
    End If
    
    Me.rptList.Populate
    
    RefreshUnion
      
End Function

Private Sub Form_Unload(Cancel As Integer)
    Me.cbrthis.DeleteAll
    Me.dkpMain.DestroyAll
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    
    If Me.rptList.Records.Count = 0 Then Exit Sub

    Set cbrPopupBar = Me.cbrthis.Add("弹出菜单", xtpBarPopup)
'    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "设置合并标本")
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Manage_ThingDel, "清空所有合并标本")
    
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Me.rptList.Records.RemoveAt (Row.Index)
    Me.rptList.Populate
    Call RefreshUnion
End Sub

Public Function ZlSave() As Long
    Dim intLoop As Integer
    Dim lngSourceID As Long
    Dim strUnionID As String
    Dim varID() As String
    
    On Error GoTo errH
    
    '是否大于两个标本
    If Me.rptList.Records.Count < 2 Then
        MsgBox "标本少于两个不能合并请选择标本后再点合并!", vbQuestion, gstrSysName
        Exit Function
    End If
    
    '是否有有主标本
    If CheckName = 0 Then
        MsgBox "请选择一个有主标本后,再合并!", vbInformation, gstrSysName
        Exit Function
    End If
    
    '得到合并ID
    For intLoop = 1 To Me.rptList.Records.Count
        If rptList.Records(intLoop - 1).Item(mCol.合并状态).Value = "合并标本" Then
            lngSourceID = Me.rptList.Records(intLoop - 1).Item(mCol.标本ID).Value
        Else
            strUnionID = strUnionID & ";" & Me.rptList.Records(intLoop - 1).Item(mCol.标本ID).Value
        End If
    Next
    strUnionID = Mid(strUnionID, 2)
    
    
                        
    If lngSourceID > 0 And strUnionID <> "" Then
        gcnOracle.BeginTrans
        
        varID = Split(strUnionID, ";")
        
        For intLoop = 0 To UBound(varID)
            gstrSql = "Zl_检验标本记录_Union(" & lngSourceID & "," & varID(intLoop) & ")"
            zlDatabase.ExecuteProcedure gstrSql, gstrSysName
        Next
        
        gcnOracle.CommitTrans
    End If
    
    Me.rptList.Records.DeleteAll
    Me.rptList.Populate
    Call RefreshUnion
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function
Private Function CheckName() As Integer
    '查找列表中是否有有病人姓名的记录存在
    Dim intLoop As Integer
    
    For intLoop = 1 To Me.rptList.Records.Count
        If Me.rptList.Records(intLoop - 1).Item(mCol.姓名).Value <> "" Then
            CheckName = intLoop
            Exit For
        End If
    Next


End Function

Private Sub RefreshUnion()
    '功能       刷新合并和被合并结果
    Dim rsTmp As New adodb.Recordset
    Dim lngKey As Long      '合并标本ID
    Dim lngUnionKey As Long '被合并标本ID
    Dim intLoop As Integer
    Dim Record As ReportRecord

    
    Me.RptComPare1.Records.DeleteAll
    Me.RptComPare2.Records.DeleteAll
    Me.RptComPare1.Populate
    Me.RptComPare2.Populate
    If Me.rptList.Rows.Count = 0 Then Exit Sub
    
    
    With Me.rptList
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mCol.合并状态).Value = "合并标本" Then
                lngKey = Val(.Rows(intLoop).Record(mCol.标本ID).Value)
            End If
            If .Rows(intLoop).Record(mCol.合并状态).Value = "被合并标本" Then
                lngUnionKey = Val(.Rows(intLoop).Record(mCol.标本ID).Value)
                If .Rows(intLoop).Selected = True Then
                    Exit For
                End If
            End If
            
        Next
    End With
    
    
    '合并
    If lngKey > 0 Then
        gstrSql = "Select C.中文名 || Decode(C.英文名, Null, '', '(' || C.英文名 || ')') As 检验项目, B.检验结果, C.单位," & vbNewLine & _
                "       Trim(Replace(Replace(' ' || Zlgetreference(C.ID, A.标本类型, Decode(A.性别, '男', 1, '女', 2, 0), A.出生日期, A.仪器id, A.年龄,a.申请科室id)," & vbNewLine & _
                "                             ' .', '0.'), '～.', '～0.')) As 参考, " & vbNewLine & _
                " DECODE(B.结果标志,3,'↑',2,'↓',1,'',4,'异常',5,'↓↓',6,'↑↑','') AS 标志,B.结果标志 " & vbNewLine & _
                "From 检验标本记录 A, 检验普通结果 B, 诊治所见项目 C, 诊疗项目目录 D, 检验项目 E" & vbNewLine & _
                "Where A.ID = B.检验标本id And B.检验项目id = C.ID And B.诊疗项目id = D.ID(+) And B.检验项目id = E.诊治项目id And A.ID = [1]" & vbNewLine & _
                "Order By Decode(E.排列序号, Null, Nvl(D.编码, 9999999999), E.排列序号), B.排列序号 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
            
        Do Until rsTmp.EOF
            Set Record = Me.RptComPare1.Records.Add
            For intLoop = 0 To Me.RptComPare1.Columns.Count
                Record.AddItem ""
            Next
            Record(mRCol.检验项目).Value = Nvl(rsTmp("检验项目"))
            Record(mRCol.结果).Value = Nvl(rsTmp("检验结果"))
            Record(mRCol.标志).Value = Nvl(rsTmp("标志"))
            Record(mRCol.单位).Value = Nvl(rsTmp("单位"))
            Record(mRCol.参考).Value = Nvl(rsTmp("参考"))
            Call ApplyResultColor(Record, Val(Nvl(rsTmp("结果标志"))))
            rsTmp.MoveNext
        Loop
    End If
    
    If lngKey > 0 And lngUnionKey > 0 Then
        gstrSql = "Select C.中文名 || Decode(C.英文名, Null, '', '(' || C.英文名 || ')') As 检验项目, B.检验结果, C.单位," & vbNewLine & _
                    "       Trim(Replace(Replace(' ' || Zlgetreference(C.ID, F.标本类型, Decode(F.性别, '男', 1, '女', 2, 0), F.出生日期, F.仪器id, F.年龄,f.申请科室id)," & vbNewLine & _
                    "                             ' .', '0.'), '～.', '～0.')) As 参考, " & vbNewLine & _
                    " DECODE(B.结果标志,3,'↑',2,'↓',1,'',4,'异常',5,'↓↓',6,'↑↑','') AS 标志,B.结果标志 " & vbNewLine & _
                    "From 检验标本记录 A, 检验普通结果 B, 诊治所见项目 C, 诊疗项目目录 D, 检验项目 E, 检验标本记录 F" & vbNewLine & _
                    "Where A.ID = B.检验标本id And B.检验项目id = C.ID And B.诊疗项目id = D.ID(+) And B.检验项目id = E.诊治项目id And F.ID = [2] And" & vbNewLine & _
                    "      A.ID = [1]" & vbNewLine & _
                    "Order By Decode(E.排列序号, Null, Nvl(D.编码, 9999999999), E.排列序号), B.排列序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngUnionKey, lngKey)
            
        Do Until rsTmp.EOF
            Set Record = Me.RptComPare2.Records.Add
            For intLoop = 0 To Me.RptComPare2.Columns.Count
                Record.AddItem ""
            Next
            Record(mRCol.检验项目).Value = Nvl(rsTmp("检验项目"))
            Record(mRCol.结果).Value = Nvl(rsTmp("检验结果"))
            Record(mRCol.标志).Value = Nvl(rsTmp("标志"))
            Record(mRCol.单位).Value = Nvl(rsTmp("单位"))
            Record(mRCol.参考).Value = Nvl(rsTmp("参考"))
            Call ApplyResultColor(Record, Val(Nvl(rsTmp("结果标志"))))
            rsTmp.MoveNext
        Loop

    End If
    
    Me.RptComPare1.Populate
    Me.RptComPare2.Populate
End Sub


Private Sub rptList_SelectionChanged()
    Call RefreshUnion
End Sub
Private Sub ApplyResultColor(Record As ReportRecord, bytMode As Byte)
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim lngColor As Long, lngForeColor As Long
    
    Select Case bytMode
        Case 0, 1
            lngColor = vbWhite
            lngForeColor = COLOR.默认前景色
        Case 5, 6 '异常低、高
            lngColor = COLOR.报警背景色
            lngForeColor = vbWhite
        Case 2
            lngColor = COLOR.低标背景色
            lngForeColor = COLOR.超标前景色
        Case Else
            lngColor = COLOR.超标背景色
            lngForeColor = COLOR.超标前景色
    End Select
    
    Record.Item(mRCol.结果).BackColor = lngColor
    Record.Item(mRCol.结果).ForeColor = lngForeColor
End Sub
