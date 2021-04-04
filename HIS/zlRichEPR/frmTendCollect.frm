VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmTendCollect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分组汇总设置"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   Icon            =   "frmTendCollect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboSelect 
      BackColor       =   &H80000018&
      Height          =   300
      ItemData        =   "frmTendCollect.frx":000C
      Left            =   690
      List            =   "frmTendCollect.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1305
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1305
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   825
   End
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   5115
      _cx             =   9022
      _cy             =   6165
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTendCollect.frx":0022
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
      AutoSizeMouse   =   0   'False
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   30
      Picture         =   "frmTendCollect.frx":0084
      Top             =   390
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   $"frmTendCollect.frx":094E
      Height          =   540
      Left            =   570
      TabIndex        =   3
      Top             =   435
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmTendCollect.frx":09F8
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTendCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnInit As Boolean
Private mblnEdit As Boolean

Private Const conMenu_删除 = 1
Private Const conMenu_保存 = 2
Private Const conMenu_恢复 = 3
Private Const conMenu_帮助 = 4
Private Const conMenu_退出 = 5
Private mrsTemp As New ADODB.Recordset
Private mrsCollect As New ADODB.Recordset

Private Enum colMenu
    项目序号
    项目名称
    绑定分组
    选择
End Enum

Private Sub cboSelect_Click()
    Dim arrResult() As String
    Dim strValue As String
    Dim lngLoop As Long
    If Not mblnEdit Then mblnEdit = (VsfData.TextMatrix(VsfData.Row, 0) <> cboSelect.Text)
    With mrsCollect
    mrsCollect.Filter = "项目名称 = '" & cboSelect.Text & "'"
    Do While Not .EOF
        If NVL(!项目值域) <> "" Then
            arrResult = Split(NVL(!项目值域), ";")
            strValue = Trim(VsfData.TextMatrix(VsfData.Row, 绑定分组))
            
            For lngLoop = 0 To UBound(arrResult)
                If strValue = "" Then
                    strValue = arrResult(lngLoop)
                Else
                    If Not InStr(1, "|" & strValue & "|", "|" & arrResult(lngLoop) & "|") > 0 Then
                        strValue = strValue & "|" & arrResult(lngLoop)
                    End If
                End If
            Next
            
            VsfData.TextMatrix(VsfData.Row, 绑定分组) = strValue
            
        Else
            strValue = Trim(VsfData.TextMatrix(VsfData.Row, 绑定分组))
            If strValue = "" Then
                VsfData.TextMatrix(VsfData.Row, 绑定分组) = NVL(!项目名称)
            Else
                If Not InStr(1, "|" & strValue & "|", "|" & NVL(!项目名称) & "|") > 0 Then
                    VsfData.TextMatrix(VsfData.Row, 绑定分组) = strValue & "|" & NVL(!项目名称)
                End If
            End If
        End If
        .MoveNext
    Loop
    End With
    VsfData.Col = 绑定分组
    Call InitCons
End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboSelect.ListIndex < 0 Then Exit Sub
        Call cboSelect_Click
        VsfData.Col = 1
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_保存
        If Not CheckData Then Exit Sub
        If Not SaveData Then Exit Sub
    Case conMenu_恢复
        Call LoadData
    Case conMenu_帮助
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_保存
        Control.Enabled = mblnEdit
    Case conMenu_恢复
        Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call MainDefCommandBar
    Call LoadData
End Sub

Private Function CheckData() As Boolean
    
    CheckData = True
End Function

Private Function SaveData() As Boolean
    Dim lngOrder As Long
    Dim lngLoop As Long
    Dim strGCollect As String
    Dim strSQL() As String
    Dim blnTran As Boolean
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
         '新增
    '    项目序号_IN IN  护理记录项目.项目序号%TYPE,
    '    项目名称_IN IN  护理记录项目.项目名称%TYPE,
    '    项目类型_IN IN  护理记录项目.项目类型%TYPE,
    '    项目长度_IN IN  护理记录项目.项目长度%TYPE,
    '    项目小数_IN IN  护理记录项目.项目小数%TYPE,
    '    项目单位_IN IN  护理记录项目.项目单位%TYPE,
    '    项目表示_IN IN  护理记录项目.项目表示%TYPE,
    '    项目值域_IN IN  护理记录项目.项目值域%TYPE,
    '    护理等级_IN   IN  护理记录项目.护理等级%TYPE,
    '    分组名_IN   IN  护理记录项目.分组名%TYPE,
    '    项目ID_IN   IN  护理记录项目.项目ID%TYPE
    
    If VsfData.TextMatrix(lngLoop, 绑定分组) <> "" Then
        For lngLoop = VsfData.FixedRows To VsfData.Rows - 1
            lngOrder = Val(VsfData.TextMatrix(lngLoop, 项目序号))
            With mrsTemp
                mrsTemp.Filter = "项目序号=" & lngOrder
                If mrsTemp.RecordCount > 0 Then
                    strSQL(ReDimArray(strSQL)) = "ZL_护理记录项目_UPDATE(" & lngOrder & ",'" & _
                    NVL(!项目名称) & "'," & NVL(!项目类型) & "," & NVL(!项目长度) & "," & NVL(!项目小数) & ",'" & _
                    NVL(!项目单位) & "'," & NVL(!项目表示) & ",'" & NVL(!项目值域) & "'," & _
                    NVL(!护理等级) & ",'" & NVL(!分组名) & "','" & NVL(!项目ID) & "'," & NVL(!应用方式) & "," & _
                    NVL(!适用病人) & "," & NVL(!项目性质) & "," & NVL(!应用场合) & ",'" & NVL(!说明) & "','" & _
                    NVL(!缺省值) & "','" & VsfData.TextMatrix(lngLoop, 绑定分组) & "')"
                
                End If
            End With
        Next
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveData = True
    mblnEdit = False
    
    Exit Function
    
errHand:
    '出错处理
    
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadData()
    
    On Error GoTo errHand
    
    mblnEdit = False
    mblnInit = True
    Call InitCons
    With VsfData
        .Clear
        .Rows = 2
        .Cols = 4
        .TextMatrix(0, 项目序号) = "项目序号"
        .TextMatrix(0, 项目名称) = "项目名称"
        .TextMatrix(0, 绑定分组) = "分组内容"
        .TextMatrix(0, 选择) = "选择"
        .ColWidth(项目名称) = 1000
        .ColWidth(绑定分组) = 2500
        .ColWidth(选择) = 1500
        .ColHidden(项目序号) = True
        .ColAlignment(项目名称) = flexAlignLeftCenter
        .ColAlignment(绑定分组) = flexAlignLeftCenter
        .ColAlignment(选择) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 项目序号, 0, 选择) = flexAlignCenterCenter
    End With
    
    '添加数据
    gstrSQL = "" & _
            " select A.项目序号,A.项目名称,A.项目类型,A.项目长度,A.项目小数,A.项目单位,A.项目表示,A.项目值域,A.护理等级," & _
            " A.分组名,A.项目id,A.应用方式,A.适用病人,A.项目性质,A.应用场合,A.说明,A.缺省值,A.分组汇总" & _
            " from 护理记录项目 A" & _
            " Where A.项目表示=4 " & _
            " Order By A.项目序号"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有汇总项目")
    With mrsTemp
        Do While Not .EOF
            If VsfData.TextMatrix(.AbsolutePosition, 项目名称) = "" Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition, 项目序号) = CStr(!项目序号)
            VsfData.TextMatrix(.AbsolutePosition, 项目名称) = CStr(!项目名称)
            VsfData.TextMatrix(.AbsolutePosition, 绑定分组) = NVL(!分组汇总)
            .MoveNext
        Loop
    End With
    VsfData.Rows = VsfData.Rows - 1
    
    gstrSQL = "" & _
        " select A.项目序号,A.项目名称,A.项目值域 " & _
        " from 护理记录项目 A" & _
        " Where A.项目类型= 1 and A.项目表示 in (0,2) " & _
        " Order By A.项目序号"
    '为下拉框添加活动项目
    Set mrsCollect = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有汇总项目")
    With mrsCollect
        Me.cboSelect.Clear
        Do While Not .EOF
            cboSelect.AddItem !项目名称
            cboSelect.ItemData(cboSelect.NewIndex) = !项目序号
            .MoveNext
        Loop
    End With
    mblnInit = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup, objFile As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim lngHandel As Long

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    'cbsMain
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
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
    cbsMain.Icons = imgPublic.Icons
    
    '工具栏定义
    '-----------------------------------------------------
    cbsMain.DeleteAll
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)      '固有
    objBar.EnableDocking xtpFlagStretched
    objBar.Closeable = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_保存, "保存"): objControl.STYLE = xtpButtonIconAndCaption: objControl.ToolTipText = "保存数据": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_恢复, "恢复"): objControl.STYLE = xtpButtonIconAndCaption: objControl.ToolTipText = "取消保存"
        Set objControl = .Add(xtpControlButton, conMenu_帮助, "帮助"): objControl.STYLE = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_退出, "退出"): objControl.STYLE = xtpButtonIconAndCaption
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_保存             '保存
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("修改的数据还未保存，你确定要退出吗？" & vbCrLf & "点“是”则放弃修改并退出，点“否”继续修改！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngWidth, lngHeight)
    imgNote.Move lngLeft + 20, lngTop - 20
    lblNote.Move lngLeft + imgNote.Width, lngTop
    With VsfData
        .Left = lngLeft
        .Top = lngTop + lblNote.Height + 30
        .Height = lngHeight - lngTop
        .Width = lngWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsTemp Is Nothing Then Set mrsTemp = Nothing
    If Not mrsCollect Is Nothing Then Set mrsCollect = Nothing
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    txtInput.Text = VsfData.TextMatrix(VsfData.Row, 绑定分组)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput.Text) = "" And Trim(VsfData.TextMatrix(VsfData.Row, 绑定分组)) = "" Then Exit Sub
    If Not mblnEdit Then mblnEdit = (VsfData.TextMatrix(VsfData.Row, 绑定分组) <> Trim(txtInput.Text))
    VsfData.TextMatrix(VsfData.Row, 绑定分组) = Trim(txtInput.Text)
    If Trim(txtInput.Text) = "" Then VsfData.TextMatrix(VsfData.Row, 绑定分组) = " "
    VsfData.Col = 绑定分组
    If VsfData.Row + 1 <= VsfData.Rows - 1 Then VsfData.Row = VsfData.Row + 1
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InitCons
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_EnterCell
    
'    Call VsfData_KeyDown(vbKeySpace, 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim objCon As Object
    Dim lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long
    
    Call InitCons
    If mblnInit Then Exit Sub
    If VsfData.Col < 绑定分组 Then Exit Sub
    
    If Not VsfData.RowIsVisible(VsfData.Row) Then VsfData.TopRow = VsfData.Row
    lngLeft = VsfData.Left + VsfData.CellLeft + 10
    lngTop = VsfData.Top + VsfData.CellTop + 10
    lngHeight = VsfData.CellHeight - 10
    lngWidth = VsfData.CellWidth - 10
    
    Select Case VsfData.Col
    Case 选择
        Set objCon = Me.cboSelect
    Case 绑定分组
        Set objCon = Me.txtInput
    End Select
    
    With objCon
        .Left = lngLeft
        .Top = lngTop
        If VsfData.Col <> 选择 Then .Height = lngHeight
        .Width = lngWidth
        
        On Error Resume Next
        Err = 0
        .Text = VsfData.TextMatrix(VsfData.Row, VsfData.Col)
        
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub InitCons()
    cboSelect.Visible = False
    txtInput.Visible = False
End Sub


