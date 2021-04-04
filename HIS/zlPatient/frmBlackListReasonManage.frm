VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBlackListReasonManage 
   BorderStyle     =   0  'None
   Caption         =   "常用不良行为原因"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2955
      Left            =   420
      TabIndex        =   1
      Top             =   645
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBlackListReasonManage.frx":0000
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
   Begin XtremeSuiteControls.ShortcutCaption stcTitle 
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _Version        =   589884
      _ExtentX        =   10398
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "基础设置>不良行为常用原因"
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
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   0
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "frmBlackListReasonManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mlngModule As Long
Private mstrPrivs As String
Public Event zlActivate(ByVal frmSubForm As Form) '事件触发

Public Sub zlInitComm(frmMain As Form, cbsThis As Object, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化接口
    '入参:objPati-调用主窗口
    '     cbsThis-菜单对象
    '     strPrivs-权限串
    '     lngModule-模块号
    '编制:刘兴洪
    '日期:2018-11-08 11:28:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain: Set mcbsMain = cbsThis
    mstrPrivs = strPrivs: mlngModule = lngModule
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Sub zlCancelBands()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:控件解绑
    '编制:刘兴洪
    '日期:2018-11-15 15:48:53
    '主要是在重建前，删除控件后，可能存在绑定的控件还在工具栏这个容器中，造成删除时，会儿控件一并删除
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrToolBar As CommandBar
    On Error GoTo errHandle
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "工具栏")
    If cbrToolBar Is Nothing Then Exit Sub
    cbrToolBar.Controls.DeleteAll
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Function zlLoadData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-13 15:33:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    On Error GoTo errHandle
    zlLoadData = LoadDataToGrid
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGridColumnHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格列头
    '编制:刘兴洪
    '日期:2018-11-08 15:13:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsGrid
        .Clear: .Rows = 2: .Cols = 4
        i = 0
        .TextMatrix(0, i) = "编码": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "名称": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "简码": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "是否系统固定": .ColWidth(i) = 1000: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "常用原因列表"
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Function LoadDataToGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2018-11-08 16:17:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, lngRow As Long
    Dim strName As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 编码,名称,简码,是否固定 From 常用不良行为原因 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsGrid
        If .Row > 0 And .Row <= .Rows - 1 Then
            strName = .TextMatrix(.Row, .ColIndex("名称"))
        End If
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("编码")) = Nvl(rsTemp!编码)
            .TextMatrix(lngRow, .ColIndex("名称")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("简码")) = Nvl(rsTemp!简码)
            .TextMatrix(lngRow, .ColIndex("是否系统固定")) = IIf(Val(Nvl(rsTemp!是否固定)) = 1, "√", "")
            If strName = .TextMatrix(lngRow, .ColIndex("名称")) Then .Row = lngRow
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    LoadDataToGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHandle
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    
    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加常用原因(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改常用原因(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除常用原因(&D)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "工具栏")
    If cbrToolBar Is Nothing Then
        Set cbrToolBar = mcbsMain.Add("工具栏", xtpBarTop)
    End If
    
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup And cbrControl.Index > 1 Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加常用原因", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改常用原因", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除常用原因", cbrControl.Index + 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
'    Set objPopup = cbrToolBar.Controls.Add(xtpControlButtonPopup, conMenu_View_FindType, "按科室过滤↓")
'    objPopup.flags = xtpFlagRightAlign
'    '被绑定的控件必须动态加载，因为工具栏一但被删除，被绑定的控件的句柄就会变成0
'    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
'
'    If txtFind.UBound > 0 Then Unload txtFind(1)
'    Load txtFind(1)
'    objCustom.Handle = txtFind(1).hWnd
'    objCustom.flags = xtpFlagRightAlign
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function IsAllowEdit(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定行是否允许编辑
    '入参:lngRow-指定行
    '返回:允许编辑返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
        
    If lngRow <= 0 Or lngRow > vsGrid.Rows - 1 Then Exit Function
    With vsGrid
        IsAllowEdit = .TextMatrix(lngRow, .ColIndex("名称")) <> "" And .TextMatrix(lngRow, .ColIndex("是否系统固定")) = ""
    End With
    Exit Function
errHandle:
    Exit Function
End Function

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置功能菜单的Eanbled属性和visible属性
    '编制:刘兴洪
    '日期:2018-11-08 16:55:37
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim blnVisible As Boolean, blnEnable As Boolean
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "编辑常用原因")
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If vsGrid.Rows >= 2 Then
           Control.Enabled = vsGrid.TextMatrix(1, vsGrid.ColIndent("名称")) <> ""
        Else
           Control.Enabled = False
        End If
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowEdit(vsGrid.Row)
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowEdit(vsGrid.Row)
    End Select
End Sub
Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关功能操作
    '编制:刘兴洪
    '日期:2018-11-08 16:56:26
    '---------------------------------------------------------------------------------------------------------------------------------------------

      
    Err = 0: On Error GoTo errHandle
    
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem: Call ExecuteAddItem
    Case conMenu_Edit_Modify: Call ExecuteModifyItem
    Case conMenu_Edit_Delete: Call ExcuteDelete
    Case conMenu_View_Refresh: LoadDataToGrid
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ExecuteAddItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行增加常用原因操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListReasonEdit
    On Error GoTo errHandle
    If Not frmEdit.zlShowEdit(mfrmMain, 0) Then Exit Function
    Call LoadDataToGrid
    ExecuteAddItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteModifyItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行修改常用原因操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListReasonEdit
    Dim strCode As String
    On Error GoTo errHandle
    With vsGrid
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        If .TextMatrix(.Row, .ColIndex("是否系统固定")) <> "" Then
            MsgBox "不允许对系统固定的常用原因进行修改!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strCode = Trim(.TextMatrix(.Row, .ColIndex("编码")))
    End With
    If strCode = "" Then Exit Function
    
    If Not frmEdit.zlShowEdit(mfrmMain, 1, strCode) Then Exit Function
    Call LoadDataToGrid
    ExecuteModifyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行删除常用原因操作
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 17:10:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCode As String, strName As String, lngRow As Long
    Dim strSQL As String
    
    On Error GoTo errHandle
    With vsGrid
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Function
        If .TextMatrix(.Row, .ColIndex("是否系统固定")) <> "" Then
            MsgBox "不允许对系统固定的常用原因进行删除!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strCode = Trim(.TextMatrix(.Row, .ColIndex("编码")))
        strName = Trim(.TextMatrix(.Row, .ColIndex("名称")))
    End With
    If strCode = "" Then Exit Function
     
    
    If MsgBox("你确定要对常用原因为『" & strName & "』进行删除操作 吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strSQL = "Zl_常用不良行为原因_Delete('" & strCode & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    With vsGrid
        lngRow = .Row
        If lngRow > .Rows - 1 And .Rows <= 2 Then
            .Clear 1: .Rows = 2
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        ElseIf lngRow > .Rows - 1 Then
            .RemoveItem lngRow
            .Row = .Rows - 1
        ElseIf lngRow <= .Rows - 1 Then
            .RemoveItem lngRow
            .Row = lngRow - 1
        End If
    End With
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
     
Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then vsGrid.SetFocus
    RaiseEvent zlActivate(Me)
End Sub

Private Sub Form_Load()

    Err = 0: On Error GoTo errHandle
    RestoreWinState Me, App.ProductName
    
    Call InitGridColumnHead
    Call LoadDataToGrid
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    stcTitle.Move 0, 0, Me.ScaleWidth
    With vsGrid
        .Left = 10: .Top = stcTitle.Top + stcTitle.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "常用原因列表"
    Err = 0: On Error Resume Next
    Set mcbsMain = Nothing
    Set mfrmMain = Nothing
End Sub

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "常用原因列表"
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "常用原因列表"
End Sub

Private Sub vsGrid_DblClick()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:双击修改
    '编制:刘兴洪
    '日期:2018-11-08 17:35:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ExecuteModifyItem
End Sub

 

Private Sub zlDataPrint(bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2018-11-08 17:37:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    If UserInfo.姓名 = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo errHandle
    objOut.Title.Text = "常用不良行为原因清单"
    Set objOut.Body = vsGrid
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub stcTitle_GotFocus()
    On Error Resume Next
    If vsGrid.Visible Then vsGrid.SetFocus
End Sub


Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandle
    If Not (Button = vbRightButton) Or Not (Me.Visible And Me.Enabled) Then Exit Sub
    
    Me.SetFocus:   RaiseEvent zlActivate(Me)
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
