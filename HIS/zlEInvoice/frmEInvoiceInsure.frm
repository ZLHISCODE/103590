VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceInsure 
   BorderStyle     =   0  'None
   Caption         =   "支付类别对照"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vs支付类别 
      Height          =   1080
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2715
      _cx             =   4789
      _cy             =   1905
      Appearance      =   0
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
      FormatString    =   $"frmEInvoiceInsure.frx":0000
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
      ExplorerBar     =   2
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
      Begin VB.Image imgAdd 
         Height          =   240
         Left            =   -480
         Picture         =   "frmEInvoiceInsure.frx":00CC
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   360
      Top             =   2760
      Width           =   525
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "支付类别对照"
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
Attribute VB_Name = "frmEInvoiceInsure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long

Private Sub Form_Load()
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    vs支付类别.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Public Sub RefreshData()
    Dim strSQL As String, i As Integer, j As Integer
    Dim rs支付类别 As ADODB.Recordset
    Dim str保险名称 As String
    
    On Error GoTo errHandle
    strSQL = "Select c.名称 As 保险名称, a.Id As 保险大类id, a.名称, b.大类编码, b.大类名称" & vbNewLine & _
                "From 保险支付大类 A, 支付类别对照 B, 保险类别 C" & vbNewLine & _
                "Where a.Id = b.保险大类id(+) And a.险类 = c.序号" & vbNewLine & _
                "Order By c.名称, a.名称"

    Set rs支付类别 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    With vs支付类别
        .Clear 1
        .Rows = 2
        .OutlineBar = flexOutlineBarSymbolsLeaf
        .Subtotal flexSTClear
        .MultiTotals = True
        .SubtotalPosition = flexSTAbove
        .OutlineCol = .ColIndex("保险名称")
        .Rows = rs支付类别.RecordCount + 1
        j = 1
        For i = 1 To rs支付类别.RecordCount
            If rs支付类别!保险名称 <> str保险名称 Then
                str保险名称 = rs支付类别!保险名称
                .AddItem rs支付类别!保险名称, j
                .RowData(j) = 1
                .MergeCol(2) = True
                .RowOutlineLevel(j) = 1
                .IsSubtotal(j) = True
                j = j + 1
             End If
             .TextMatrix(j, .ColIndex("保险名称")) = rs支付类别!保险名称
            .TextMatrix(j, .ColIndex("保险大类id")) = Val(rs支付类别!保险大类id)
            .TextMatrix(j, .ColIndex("名称")) = Nvl(rs支付类别!名称)
            .TextMatrix(j, .ColIndex("大类编码")) = Nvl(rs支付类别!大类编码)
            .TextMatrix(j, .ColIndex("大类名称")) = Nvl(rs支付类别!大类名称)
            .RowOutlineLevel(j) = 2
             .IsSubtotal(j) = True
             rs支付类别.MoveNext
             j = j + 1
        Next
        .Outline 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)")
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
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)", cbrControl.index + 1)
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPopup()
    '显示弹出菜单
    Dim objPopup As CommandBarPopup
    Err = 0: On Error GoTo ErrHandler
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus
    
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

Public Sub AddNewInsure()
    '新增支付类别对照
    Dim frmEdit As New frmEInvoiceInsureSet
    Dim blnRefresh As Boolean
    Dim str支付名称 As String, str保险名称 As String
    Dim lng支付大类ID As Long
    
    On Error GoTo errHandle
    With vs支付类别
        str保险名称 = .TextMatrix(.Row, .ColIndex("保险名称"))
        str支付名称 = .TextMatrix(.Row, .ColIndex("名称"))
        lng支付大类ID = Val(.TextMatrix(.Row, .ColIndex("保险大类id")))
    End With
    If lng支付大类ID = 0 Then Exit Sub
    Call frmEdit.ShowMe(Me, 0, str保险名称, lng支付大类ID, str支付名称, , , blnRefresh)
    If blnRefresh Then Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteInsure()
    '删除支付类别对照
    On Error GoTo errHandle
    Dim lng大类ID As Long, str大类编码 As String
    Dim str大类名称 As String, strSQL As String
    With vs支付类别
        If .Row = 0 Then Exit Sub
        lng大类ID = Val(.TextMatrix(.Row, .ColIndex("保险大类id")))
        str大类名称 = .TextMatrix(.Row, .ColIndex("大类名称"))
        str大类编码 = .TextMatrix(.Row, .ColIndex("大类编码"))
        
        If str大类编码 = "" Then Exit Sub
        If MsgBox("你确认要删除大类名称为“" & str大类名称 & "”,大类编码为“" & str大类编码 & "”的支付类别对照吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Me.MousePointer = 11
        strSQL = "Zl_支付类别对照_Update(2," & lng大类ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Me.MousePointer = 0
    End With

    Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub ModifyInsure()
    '修改支付类别对照
    Dim frmEdit As New frmEInvoiceInsureSet
    Dim blnRefresh As Boolean
    Dim str保险名称 As String, str支付名称 As String
    Dim str大类编码 As String, str大类名称 As String
    Dim lng支付大类ID As Long
    
    On Error Resume Next
    With vs支付类别
        If .Row = 0 Then Exit Sub
            str保险名称 = .TextMatrix(.Row, .ColIndex("保险名称"))
            str支付名称 = .TextMatrix(.Row, .ColIndex("名称"))
            lng支付大类ID = Val(.TextMatrix(.Row, .ColIndex("保险大类id")))
            str大类编码 = .TextMatrix(.Row, .ColIndex("大类编码"))
            str大类名称 = .TextMatrix(.Row, .ColIndex("大类名称"))
        End With
    If str大类编码 = "" Then Exit Sub
    Call frmEdit.ShowMe(Me, 1, str保险名称, lng支付大类ID, str支付名称, str大类编码, str大类名称, blnRefresh)
    If blnRefresh Then Call RefreshData
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_NewItem '新增
        Call AddNewInsure
    Case conMenu_Edit_Modify  '修改
        Call ModifyInsure
    Case conMenu_Edit_Delete '删除
        Call DeleteInsure
    Case conMenu_View_Refresh '刷新数据
        Call RefreshData
    Case Else
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean, blnHaveData As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    With vs支付类别
        blnEnable = Val(.RowData(.Row)) = 0
        blnHaveData = .TextMatrix(.Row, .ColIndex("大类编码")) <> ""
    End With
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        Control.Enabled = False
    Case conMenu_Edit_NewItem
        Control.Enabled = blnEnable And Not blnHaveData
    Case conMenu_Edit_Modify
        Control.Enabled = blnEnable And blnHaveData
    Case conMenu_Edit_Delete
        Control.Enabled = blnEnable And blnHaveData
    Case Else
    End Select
End Sub

Private Sub vs支付类别_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs支付类别, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs支付类别_DblClick()
    Dim str大类编码 As String
    
    With vs支付类别
        If .Row <= 0 Then Exit Sub
         If .RowData(.Row) = 1 Then Exit Sub
        str大类编码 = .TextMatrix(.Row, .ColIndex("大类编码"))
    End With
    If str大类编码 = "" Then
        Call AddNewInsure
    Else
        Call ModifyInsure
    End If
End Sub

Private Sub vs支付类别_GotFocus()
    If vs支付类别.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs支付类别, &HFFEBD7
End Sub

Private Sub vs支付类别_LostFocus()
    If vs支付类别.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs支付类别
    OS.OpenIme False
End Sub

Private Sub vs支付类别_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub
