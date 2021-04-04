VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmAnimalPartMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体温部位设置"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   Icon            =   "frmAnimalPartMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   690
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox cboSelect 
      BackColor       =   &H80000018&
      Height          =   300
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1110
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   510
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAnimalPartMan.frx":6852
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
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   420
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmAnimalPartMan.frx":68B4
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmAnimalPartMan.frx":AE9A
   End
End
Attribute VB_Name = "frmAnimalPartMan"
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

Private Enum colMenu
    项目名称
    部位
    缺省项
    固定项
End Enum

'固定项不允许编辑与删除,必须指定缺省项
'非固定项可不必设定缺省项

Private Sub cboSelect_Click()
    If Not mblnEdit Then mblnEdit = (VsfData.TextMatrix(VsfData.ROW, 0) <> cboSelect.Text)
    VsfData.TextMatrix(VsfData.ROW, 0) = cboSelect.Text
    VsfData.RowData(VsfData.ROW) = cboSelect.ItemData(cboSelect.ListIndex)
End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboSelect.ListIndex < 0 Then Exit Sub
        Call cboSelect_Click
        VsfData.COL = 1
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_删除
        VsfData.RemoveItem VsfData.ROW
        mblnEdit = True
    Case conMenu_保存
        If Not CheckData Then Exit Sub
        If Not SaveData Then Exit Sub
    Case conMenu_恢复
        Call LoadData
    Case conMenu_帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_删除
        Control.Enabled = (VsfData.RowData(VsfData.ROW))
        If Control.Enabled Then Control.Enabled = (VsfData.TextMatrix(VsfData.ROW, 固定项) = "")
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
    Dim strPara As String
    Dim lngOrder As Long
    Dim lngRow As Long, lngCount As Long
    On Error GoTo errHand
    
    '格式：项目序号;部位'缺省'固定,部位'缺省'固定|项目序号;部位'缺省'固定,
    strPara = ","
    lngCount = VsfData.Rows - 1
    For lngRow = 1 To lngCount
        If lngOrder <> VsfData.RowData(lngRow) Then
            If VsfData.TextMatrix(lngRow, 固定项) <> "" Or VsfData.TextMatrix(lngRow, 部位) <> "" Then
                lngOrder = VsfData.RowData(lngRow)
                strPara = Mid(strPara, 1, Len(strPara) - 1) & "|" & lngOrder & ";" & VsfData.TextMatrix(lngRow, 部位) & _
                "'" & IIf(VsfData.TextMatrix(lngRow, 缺省项) = "", 0, 1) & "'" & IIf(VsfData.TextMatrix(lngRow, 固定项) = "", 0, 1) & ","
            End If
        Else
            If VsfData.TextMatrix(lngRow, 固定项) <> "" Or VsfData.TextMatrix(lngRow, 部位) <> "" Then
                strPara = strPara & VsfData.TextMatrix(lngRow, 部位) & _
                "'" & IIf(VsfData.TextMatrix(lngRow, 缺省项) = "", 0, 1) & "'" & IIf(VsfData.TextMatrix(lngRow, 固定项) = "", 0, 1) & ","
            End If
        End If
    Next
    If strPara <> "," Then
        strPara = Mid(strPara, 2, Len(strPara) - 2)
    Else
        strPara = ""
    End If
    
    gstrSQL = "ZL_体温部位_UPDATE('" & strPara & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存体温部位")
    
    mblnEdit = False
    SaveData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadData()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mblnEdit = False
    mblnInit = True
    Call InitCons
    With VsfData
        .Clear
        .Rows = 2
        .Cols = 4
        .TextMatrix(0, 项目名称) = "活动项目"
        .TextMatrix(0, 部位) = "部位"
        .TextMatrix(0, 缺省项) = "缺省"
        .TextMatrix(0, 固定项) = "固定"
        .ColWidth(项目名称) = 1000
        .ColWidth(部位) = 2500
        .ColWidth(缺省项) = 500
        .ColWidth(固定项) = 500
        .ColAlignment(项目名称) = flexAlignLeftCenter
        .ColAlignment(部位) = flexAlignLeftCenter
        .ColAlignment(缺省项) = flexAlignCenterCenter
        .ColAlignment(固定项) = flexAlignCenterCenter
    End With
    
    '添加数据
    gstrSQL = " Select A.项目序号,A.部位,B.项目名称,A.缺省项,A.固定项 " & _
              " From 体温部位 A,护理记录项目 B " & _
              " Where A.项目序号=B.项目序号 " & _
              " Order by 项目序号,部位"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取体温部位")
    With rsTemp
        Do While Not .EOF
            If VsfData.TextMatrix(.AbsolutePosition, 项目名称) = "" Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition, 项目名称) = CStr(!项目名称)
            VsfData.TextMatrix(.AbsolutePosition, 部位) = NVL(!部位)
            VsfData.TextMatrix(.AbsolutePosition, 缺省项) = IIf(NVL(!缺省项, 0) = 1, "√", "")
            VsfData.TextMatrix(.AbsolutePosition, 固定项) = IIf(NVL(!固定项, 0) = 1, "√", "")
            VsfData.RowData(.AbsolutePosition) = CLng(!项目序号)
            .MoveNext
        Loop
    End With
    
    '为下拉框添加活动项目
    gstrSQL = " Select 项目序号,项目名称 From 护理记录项目 Where 项目性质=2 Order by 项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取活动项目")
    With rsTemp
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

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
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
        Set objControl = .Add(xtpControlButton, conMenu_删除, "删除"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "删除行"
        Set objControl = .Add(xtpControlButton, conMenu_保存, "保存"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "保存数据": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_恢复, "恢复"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "取消保存"
        Set objControl = .Add(xtpControlButton, conMenu_帮助, "帮助"): objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_退出, "退出"): objControl.Style = xtpButtonIconAndCaption
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
    With VsfData
        .Left = lngLeft
        .Top = lngTop
        .Height = lngHeight - lngTop
        .Width = lngWidth
    End With
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput.Text) = "" Then Exit Sub
    
    If Not mblnEdit Then mblnEdit = (VsfData.TextMatrix(VsfData.ROW, 部位) <> Trim(txtInput.Text))
    VsfData.TextMatrix(VsfData.ROW, 部位) = Trim(txtInput.Text)
    If VsfData.ROW = VsfData.Rows - 1 Then
        VsfData.Rows = VsfData.Rows + 1
        VsfData.TextMatrix(VsfData.ROW + 1, 项目名称) = VsfData.TextMatrix(VsfData.ROW, 项目名称)
        VsfData.RowData(VsfData.ROW + 1) = VsfData.RowData(VsfData.ROW)
    End If
    VsfData.ROW = VsfData.ROW + 1
    VsfData.COL = 0
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InitCons
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(vbKeySpace, 0)
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngOrder As Long
    Dim intRow As Integer, intCount As Integer
    If VsfData.COL = 缺省项 And KeyCode = vbKeySpace Then
        If VsfData.TextMatrix(VsfData.ROW, 缺省项) = "√" Then Exit Sub
        
        '取消该项目的所有部位的缺省项标志
        lngOrder = VsfData.RowData(VsfData.ROW)
        intCount = VsfData.Rows - 1
        For intRow = 1 To intCount
            If lngOrder = VsfData.RowData(intRow) Then
                If intRow = VsfData.ROW Then
                    VsfData.TextMatrix(intRow, 缺省项) = "√"
                Else
                    VsfData.TextMatrix(intRow, 缺省项) = ""
                End If
            End If
        Next
        mblnEdit = True
        
    End If
End Sub

Private Sub VsfData_EnterCell()
    Dim objCon As Object
    Dim lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long
    
    Call InitCons
    If mblnInit Then Exit Sub
    If VsfData.TextMatrix(VsfData.ROW, 固定项) <> "" Then Exit Sub
    If VsfData.COL > 部位 Then Exit Sub
    
    If Not VsfData.RowIsVisible(VsfData.ROW) Then VsfData.TopRow = VsfData.ROW
    lngLeft = VsfData.Left + VsfData.CellLeft + 10
    lngTop = VsfData.Top + VsfData.CellTop + 10
    lngHeight = VsfData.CellHeight - 10
    lngWidth = VsfData.CellWidth - 10
    
    Select Case VsfData.COL
    Case 0
        Set objCon = Me.cboSelect
    Case 1
        Set objCon = Me.txtInput
    End Select
    
    With objCon
        .Left = lngLeft
        .Top = lngTop
        If VsfData.COL <> 0 Then .Height = lngHeight
        .Width = lngWidth
        
        On Error Resume Next
        Err = 0
        .Text = VsfData.TextMatrix(VsfData.ROW, VsfData.COL)
        If Err <> 0 Then .ListIndex = 0
        
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub InitCons()
    cboSelect.Visible = False
    txtInput.Visible = False
End Sub
