VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceChannel 
   BorderStyle     =   0  'None
   Caption         =   "收费渠道对照"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vs收费渠道 
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoiceChannel.frx":0000
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
      Editable        =   2
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
         Picture         =   "frmEInvoiceChannel.frx":0123
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
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
      Caption         =   "收费渠道对照"
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
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   360
      Top             =   2760
      Width           =   525
   End
End
Attribute VB_Name = "frmEInvoiceChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    vs收费渠道.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Private Sub Form_Load()
    Call RefreshData
End Sub

Public Sub RefreshData()
    Dim strSql As String, i As Integer
    Dim rs收费渠道 As ADODB.Recordset

    On Error GoTo errHandle
    strSql = "Select a.名称 As 结算方式, a.性质 As 结算性质, c.Id As 卡类别id, Nvl(c.名称, '无') As 卡类别名称," & vbNewLine & _
                "      b.渠道编码, 0 As 附加标志" & vbNewLine & _
                "From 结算方式 A, 收费渠道对照 B, 医疗卡类别 C" & vbNewLine & _
                "Where a.名称 = b.结算方式(+) And a.名称 = c.结算方式(+) And b.卡类别id Is Null" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.结算方式, c.性质 As 结算性质, a.卡类别id, Nvl(b.名称, '无') As 卡类别名称, a.渠道编码," & vbNewLine & _
                "       Decode(a.结算方式, b.结算方式, 0, 1) As 附加标志" & vbNewLine & _
                "From 收费渠道对照 A, 医疗卡类别 B, 结算方式 C" & vbNewLine & _
                "Where a.结算方式 = c.名称 And a.卡类别id = b.Id And c.性质 = 8" & vbNewLine & _
                "Order By 卡类别id, 附加标志"

    Set rs收费渠道 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    With vs收费渠道
        .Clear 1
        .Rows = 2
        For i = 1 To rs收费渠道.RecordCount
            .TextMatrix(i, .ColIndex("结算方式")) = rs收费渠道!结算方式
            .TextMatrix(i, .ColIndex("原结算方式")) = rs收费渠道!结算方式
            .TextMatrix(i, .ColIndex("结算性质")) = Val(rs收费渠道!结算性质)
            .TextMatrix(i, .ColIndex("卡类别ID")) = Val(NVL(rs收费渠道!卡类别id))
            If Val(rs收费渠道!结算性质) = 8 And Val(NVL(rs收费渠道!卡类别id)) > 0 And Val(NVL(rs收费渠道!附加标志)) = 0 Then
                .CellButtonPicture = imgAdd: .ComboList = "..."
            End If
            .TextMatrix(i, .ColIndex("卡类别名称")) = NVL(rs收费渠道!卡类别名称)
            .TextMatrix(i, .ColIndex("渠道编码")) = NVL(rs收费渠道!渠道编码)
            .TextMatrix(i, .ColIndex("附加标志")) = NVL(rs收费渠道!附加标志)
            rs收费渠道.MoveNext
            If i < rs收费渠道.RecordCount Then .Rows = .Rows + 1
        Next
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

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    blnEnable = vs收费渠道.TextMatrix(vs收费渠道.Row, vs收费渠道.ColIndex("渠道编码")) <> ""

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        Control.Enabled = False
    Case conMenu_Edit_NewItem
        Control.Enabled = Not blnEnable
    Case conMenu_Edit_Modify
        Control.Enabled = blnEnable
    Case conMenu_Edit_Delete
        Control.Enabled = blnEnable
    Case Else
    End Select
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

Private Sub vs收费渠道_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With vs收费渠道
        If .ColIndex("结算方式") <> Col Then Exit Sub
        If .TextMatrix(Row, .ColIndex("结算方式")) = "" Then Exit Sub
        If .TextMatrix(Row, .ColIndex("渠道编码")) = "" Then Exit Sub
        If Val(.TextMatrix(Row, .ColIndex("卡类别id"))) = 0 Then Exit Sub
        If Modify收费渠道对照(.TextMatrix(Row, .ColIndex("结算方式")), Val(.TextMatrix(Row, .ColIndex("卡类别id"))), .TextMatrix(Row, .ColIndex("渠道编码")), _
           .TextMatrix(Row, .ColIndex("原结算方式"))) = False Then Exit Sub
        Call RefreshData
    End With
End Sub

Private Sub vs收费渠道_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim str结算方式 As String
    Dim rs三方卡结算 As ADODB.Recordset
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs收费渠道, OldRow, NewRow, OldCol, NewCol
    
    With vs收费渠道
        If NewCol <> .ColIndex("结算方式") Then
            .ComboList = "..."
            Exit Sub
        End If
        If Val(.TextMatrix(NewRow, .ColIndex("结算性质"))) = 8 And Val(.TextMatrix(NewRow, .ColIndex("卡类别id"))) > 0 Then
            If Val(.TextMatrix(NewRow, .ColIndex("附加标志"))) = 0 Then
                .CellButtonPicture = imgAdd: .ComboList = "..."
                .ColComboList(.ColIndex("结算方式")) = ""
            Else
                str结算方式 = .TextMatrix(NewRow - 1, .ColIndex("结算方式"))
                Set rs三方卡结算 = Get三方结算方式(str结算方式)
                .ColComboList(.ColIndex("结算方式")) = .BuildComboList(rs三方卡结算, "结算方式", "结算方式")
            End If
        End If
    End With
End Sub

Private Sub vs收费渠道_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str结算方式 As String
    Dim rs三方卡结算  As New ADODB.Recordset
    With vs收费渠道
        Select Case Col
        Case .ColIndex("结算方式")
            If Not (Val(.TextMatrix(Row, .ColIndex("结算性质"))) = 8 And Val(.TextMatrix(Row, .ColIndex("卡类别id"))) > 0) Then Cancel = True: Exit Sub
            If Val(.TextMatrix(Row, .ColIndex("附加标志"))) = 1 Then
                str结算方式 = .TextMatrix(Row - 1, .ColIndex("结算方式"))
                If str结算方式 <> "" Then
                    Set rs三方卡结算 = Get三方结算方式(str结算方式)
                    .ColComboList(.ColIndex("结算方式")) = .BuildComboList(rs三方卡结算, "结算方式", "结算方式")
                    Exit Sub
                End If
            End If
            If CheckThirdCard(Val(.TextMatrix(Row, .ColIndex("卡类别id")))) = False Then
                Cancel = True: Exit Sub
            Else
                .CellButtonPicture = imgAdd: .ComboList = "..."
            End If
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub vs收费渠道_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vs收费渠道
        If .ColIndex("结算方式") <> Col Then Exit Sub
        If Not (Val(.TextMatrix(Row, .ColIndex("结算性质"))) = 8 And Val(.TextMatrix(Row, .ColIndex("卡类别id"))) > 0) Then Exit Sub
        If CheckThirdCard(Val(.TextMatrix(Row, .ColIndex("卡类别id")))) = False Then Exit Sub
        .AddItem Val(.TextMatrix(Row, .ColIndex("结算性质"))) & vbTab & .TextMatrix(Row, .ColIndex("卡类别id")) & vbTab & .TextMatrix(Row, .ColIndex("卡类别名称")) & vbTab & .TextMatrix(Row, .ColIndex("原结算方式")) & vbTab & "" & vbTab & "" & vbTab & "1", Row + 1
    End With
End Sub

Private Sub vs收费渠道_DblClick()
    Dim str渠道编码 As String
    With vs收费渠道
        If .Row <= 0 Then Exit Sub
        str渠道编码 = .TextMatrix(.Row, .ColIndex("渠道编码"))
    End With
    If str渠道编码 = "" Then
        Call AddNewChannel
    Else
        Call ModifyChannel
    End If
End Sub

Private Sub vs收费渠道_GotFocus()
    If vs收费渠道.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs收费渠道, &HFFEBD7
End Sub

Private Sub vs收费渠道_KeyDown(KeyCode As Integer, Shift As Integer)
    With vs收费渠道
        If .Row < 1 Then Exit Sub
        If .Col <> .ColIndex("结算方式") Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("渠道编码")) = "" And Val(.TextMatrix(.Row, .ColIndex("附加标志"))) = 1 Then
            If KeyCode = vbKeyDelete Then .RemoveItem .Row
        End If
    End With
End Sub

Private Sub vs收费渠道_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vs收费渠道
        Select Case Col
        Case .ColIndex("渠道编码")
            If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End Select
    End With
End Sub

Private Sub vs收费渠道_LostFocus()
    If vs收费渠道.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs收费渠道
    OS.OpenIme False
End Sub

Private Sub vs收费渠道_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub

Private Function Get三方结算方式(ByVal str结算方式 As String) As ADODB.Recordset
    Dim strSql As String
    Dim rsTmp  As ADODB.Recordset '当前有效的三方卡支付方式
    If str结算方式 = "" Then Exit Function
    On Error GoTo ErrHandler

    strSql = " Select Rownum As 序号, a.编码, a.名称 As 结算方式, b.名称 As 卡名称, Decode(Nvl(b.名称, '-'), '-', 1, 0) As 附加标志 " & _
                   " From 结算方式 A, 医疗卡类别 B " & _
                   " Where a.性质 = 8 And a.名称 = b.结算方式(+) and a.名称<>[1]"
    Set Get三方结算方式 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str结算方式)

    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_NewItem '新增
        Call AddNewChannel
    Case conMenu_Edit_Modify  '修改
        Call ModifyChannel
    Case conMenu_Edit_Delete '删除
        Call DeleteChannel
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

Private Function CheckThirdCard(ByVal lng卡类别id As Long) As Boolean
    '功能：检查是否重复录入三方卡
    Dim i As Integer
    
    If lng卡类别id = 0 Then Exit Function
    With vs收费渠道
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("卡类别id")) = lng卡类别id And Val(.TextMatrix(i, .ColIndex("附加标志"))) = 1 Then
                Exit Function
            End If
        Next
    End With
    CheckThirdCard = True
End Function

Public Sub AddNewChannel()
    '新增收据费目对照
    Dim frmEdit As New frmEInvoiceChannelSet
    Dim blnRefresh As Boolean
    Dim str结算方式 As String, str卡类别名称 As String
    Dim lng卡类别id As Long
    
    On Error GoTo errHandle
    With vs收费渠道
        str结算方式 = .TextMatrix(.Row, .ColIndex("结算方式"))
        str卡类别名称 = .TextMatrix(.Row, .ColIndex("卡类别名称"))
        lng卡类别id = Val(.TextMatrix(.Row, .ColIndex("卡类别id")))
    End With
    If str结算方式 = "" Then
        MsgBox "未选择结算方式，请先选择结算方式！", vbInformation, gstrSysName
        Exit Sub
    End If
    Call frmEdit.ShowMe(Me, 0, lng卡类别id, str卡类别名称, str结算方式, , blnRefresh)
    If blnRefresh Then Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteChannel()
    '删除收据费目对照
    On Error GoTo errHandle
    Dim lng卡类别id As Long, str渠道编码 As String
    Dim str结算方式 As String, strSql As String
    Dim str卡类别 As String
    With vs收费渠道
        If .Row = 0 Then Exit Sub
        lng卡类别id = Val(.TextMatrix(.Row, .ColIndex("卡类别ID")))
        str结算方式 = .TextMatrix(.Row, .ColIndex("结算方式"))
        str渠道编码 = .TextMatrix(.Row, .ColIndex("渠道编码"))
        str卡类别 = .TextMatrix(.Row, .ColIndex("卡类别名称"))
        
        If str渠道编码 = "" Then Exit Sub
        If MsgBox("你确认要删除结算方式为“" & str结算方式 & "”" & IIf(lng卡类别id = 0, "", "卡类别名称为“" & str卡类别 & "”") & "的收费渠道对照吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Me.MousePointer = 11
        strSql = "Zl_收费渠道对照_Update(2,'" & str结算方式 & "'," & IIf(lng卡类别id = 0, "NULL", lng卡类别id) & ",'" & str渠道编码 & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
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

Public Sub ModifyChannel()
    '修改收据费目对照
    Dim frmEdit As New frmEInvoiceChannelSet
    Dim blnRefresh As Boolean
    Dim str结算方式 As String, str卡类别名称 As String
    Dim lng卡类别id As Long, str渠道编码 As String
    
    On Error Resume Next
    With vs收费渠道
        If .Row = 0 Then Exit Sub
        str结算方式 = .TextMatrix(.Row, .ColIndex("结算方式"))
        str卡类别名称 = .TextMatrix(.Row, .ColIndex("卡类别名称"))
        lng卡类别id = Val(.TextMatrix(.Row, .ColIndex("卡类别id")))
        str渠道编码 = .TextMatrix(.Row, .ColIndex("渠道编码"))
    End With
    If str渠道编码 = "" Then Exit Sub
    Call frmEdit.ShowMe(Me, 1, lng卡类别id, str卡类别名称, str结算方式, str渠道编码, blnRefresh)
    If blnRefresh Then Call RefreshData
End Sub

Private Function Modify收费渠道对照(ByVal str结算方式 As String, ByVal lng卡类别id As Long, _
                           ByVal str渠道编码 As String, ByVal str原结算方式 As String) As Boolean
    Dim strSql As String
    
    If str结算方式 = str原结算方式 Then Exit Function
    If lng卡类别id = 0 Then Exit Function
    If str渠道编码 = "" Then Exit Function
    If str原结算方式 = "" Then Exit Function
    
    On Error GoTo errHandle
    '调整收费渠道对照的结算方式
    strSql = "Zl_收费渠道对照_Update("
    '操作类型_In In Number,
    strSql = strSql & 3 & ","
    '结算方式_In In 收费渠道对照.结算方式%Type,
    strSql = strSql & "'" & str结算方式 & "',"
    '卡类别id_In In 收费渠道对照.卡类别id%Type,
    strSql = strSql & lng卡类别id & ","
    '渠道编码_In In 收费渠道对照.渠道编码%Type
    strSql = strSql & "'" & str渠道编码 & "',"
    '原结算方式_In In 收费渠道对照.结算方式%Type := Null
    strSql = strSql & "'" & str原结算方式 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "收费渠道对照")
    
    Modify收费渠道对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
