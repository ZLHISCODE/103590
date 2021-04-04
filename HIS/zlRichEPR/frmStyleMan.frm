VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmStyleMan 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "病历常用样式"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   5085
      Left            =   150
      TabIndex        =   0
      Top             =   495
      Width           =   3390
      _cx             =   5980
      _cy             =   8969
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
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483634
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Line LinTop 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   3285
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label lblList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请选择要应用的样式"
      Height          =   180
      Left            =   105
      TabIndex        =   1
      Top             =   150
      Width           =   1620
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   240
      Top             =   5820
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmStyleMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conCol编号 = 0
Const conCol名称 = 1
Const conCol段落 = 2
Const conCol字体 = 3
Const conCol系统 = 4

'公共事件
Public Event DblClick(ByVal lngStyleCode As Long)   '双击或按回车选择制定样式
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCode As Long
    
    Select Case Control.ID
    Case conMenu_Tool_Apply
        Call vfgList_DblClick
    Case conMenu_Edit_NewItem
        lngCode = frmStyleSet.ShowMe(Me, True)
        If lngCode = 0 Then Exit Sub
        Call zlRefList(lngCode)
    Case conMenu_Edit_Modify
        Err = 0: On Error Resume Next
        lngCode = Me.vfgList.TextMatrix(Me.vfgList.ROW, conCol编号)
        If Err <> 0 Or lngCode = 0 Then MsgBox "没有选定样式！", vbExclamation, gstrSysName: Exit Sub
        Err = 0: On Error GoTo 0
        lngCode = frmStyleSet.ShowMe(Me, False, lngCode)
        If lngCode = 0 Then Exit Sub
        Call zlRefList(lngCode)
    Case conMenu_Edit_Delete
        Err = 0: On Error Resume Next
        lngCode = Me.vfgList.TextMatrix(Me.vfgList.ROW, conCol编号)
        If Err <> 0 Or lngCode = 0 Then MsgBox "没有选定样式！", vbExclamation, gstrSysName: Exit Sub
        If MsgBox("真的删除常用样式“" & Me.vfgList.TextMatrix(Me.vfgList.ROW, conCol名称) & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        Err = 0: On Error GoTo ErrHand
        gstrSQL = "Zl_病历常用样式_Delete(" & lngCode & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "frmStyleMan")
        Call zlRefList
    End Select
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Left = -120
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    Me.lblList.Move 90, 150
    With Me.LinTop
        .X1 = 0: .Y1 = Me.lblList.Top + Me.lblList.Height + 45
        .X2 = Me.ScaleWidth: .Y2 = .Y1
    End With
    With Me.vfgList
        .Left = 90: .Width = lngScaleRight - .Left * 2
        .Top = Me.lblList.Top + Me.lblList.Height + 150: .Height = lngScaleBottom - .Top - 90
        .ColWidth(conCol名称) = .Width - Screen.TwipsPerPixelX * 1 ' - 250
    End With
End Sub

Private Sub cbsThis_SpecialColorChanged()

End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    With Me.vfgList
        Select Case Control.ID
        Case conMenu_Edit_NewItem
            Control.Enabled = (InStr(1, gstrPrivsEpr, "病历样式设置") > 0)
        Case conMenu_Edit_Modify
            Control.Enabled = (InStr(1, gstrPrivsEpr, "病历样式设置") > 0)
            If Control.Enabled Then Control.Enabled = (.Rows > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.ROW, conCol编号)) > 0)
        Case conMenu_Edit_Delete
            Control.Enabled = (InStr(1, gstrPrivsEpr, "病历样式设置") > 0)
            If Control.Enabled Then Control.Enabled = (.Rows > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.ROW, conCol编号)) > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.ROW, conCol系统)) = 0)
        Case conMenu_Tool_Apply
                Control.Enabled = InStr(1, gstrPrivsEpr, ";字体格式设置;") > 0
        End Select
    End With
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
    '-----------------------------------------------------
    '内部菜单工具栏定义
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlcommfun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "应用(&A)"): cbrControl.STYLE = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.STYLE = xtpButtonCaption
    End With
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '-----------------------------------------------------
    '样式列表填写
    Call zlRefList
End Sub

Private Sub vfgList_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Err = 0: On Error Resume Next
    With Me.vfgList
        .CellBorderRange OldRowSel, conCol名称, OldRowSel, conCol名称, RGB(255, 255, 255), 0, 0, 0, 0, 0, 0
        .CellBorderRange NewRowSel, conCol名称, NewRowSel, conCol名称, RGB(0, 64, 128), 2, 2, 2, 2, 0, 0
        .ForeColorSel = .Cell(flexcpForeColor, NewRowSel, conCol名称, NewRowSel, conCol名称)
    End With
End Sub

Private Sub vfgList_DblClick()
    Dim lngCode As Long
    If InStr(1, gstrPrivsEpr, ";字体格式设置;") = 0 Then Exit Sub
    Err = 0: On Error Resume Next
    lngCode = Me.vfgList.TextMatrix(Me.vfgList.ROW, conCol编号)
    If Err <> 0 Or lngCode = 0 Then Exit Sub
    RaiseEvent DblClick(lngCode)
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vfgList_DblClick
End Sub

Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cbrPopupBar As CommandBar
Dim cbrControl As CommandBarControl
    If Button <> vbRightButton Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("弹出", xtpBarPopup)
    With cbrPopupBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "应用(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub zlRefList(Optional lngCode As Long)
    '-----------------------------------------------------
    '功能：刷新装入样式列表
    '参数： lngCode，要选中的样式号
    '-----------------------------------------------------
Dim objFont As New StdFont
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
Dim aryFormat() As String
    
    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select 编号, 名称, 段落样式, 字体样式, 系统 From 病历常用样式 Order By 编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmStyleMan")
    
    With Me.vfgList
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(conCol编号) = 0: .ColWidth(conCol段落) = 0: .ColWidth(conCol字体) = 0: .ColWidth(conCol系统) = 0
        .ColWidth(conCol名称) = .Width - Screen.TwipsPerPixelX
        .ColAlignment(conCol名称) = flexAlignLeftCenter
        
        For lngCount = .FixedRows To .Rows - 1
        
            '设置样式字体
            Set objFont = Nothing
            aryFormat = Split(.TextMatrix(lngCount, conCol字体), ";")
            If UBound(aryFormat) = 5 Then
                If Trim(aryFormat(0)) <> "" Then objFont.Name = aryFormat(0)
                If Val(aryFormat(1)) > 0 Then objFont.Size = Val(aryFormat(1))
                objFont.Bold = IIf(Mid(aryFormat(2), 1, 1) = 1, True, False)
                objFont.Italic = IIf(Mid(aryFormat(2), 2, 1) = 1, True, False)
'                objFont.Hidden = IIf(Mid(aryFormat(2), 3, 1) = 1, True, False)
'                objFont.Protected = IIf(Mid(aryFormat(2), 4, 1) = 1, True, False)
'                objFont.Link = IIf(Mid(aryFormat(2), 5, 1) = 1, True, False)
                objFont.Strikethrough = IIf(Mid(aryFormat(2), 6, 1) = 1, True, False)
'                objFont.Superscript = IIf(Mid(aryFormat(2), 7, 1) = 1, True, False)
'                objFont.Subscript = IIf(Mid(aryFormat(2), 8, 1) = 1, True, False)
                objFont.Underline = Val(aryFormat(3))
                If Val(aryFormat(4)) >= 0 Then .Cell(flexcpBackColor, lngCount, conCol名称, lngCount, conCol名称) = Val(aryFormat(4))
                If Val(aryFormat(5)) >= 0 Then .Cell(flexcpForeColor, lngCount, conCol名称, lngCount, conCol名称) = Val(aryFormat(5))
            End If
            Set .Cell(flexcpFont, lngCount, conCol名称, lngCount, conCol名称) = objFont
            
            '计算设置样式高度
            .ROWHEIGHT(lngCount) = zlStyleHeight(objFont.Size, .TextMatrix(lngCount, conCol段落))
            Select Case Val(Left(.TextMatrix(lngCount, conCol段落) & " ", 1))
            Case 2: .Cell(flexcpAlignment, lngCount, conCol名称, lngCount, conCol名称) = flexAlignRightCenter
            Case 1: .Cell(flexcpAlignment, lngCount, conCol名称, lngCount, conCol名称) = flexAlignCenterCenter
            Case Else: .Cell(flexcpAlignment, lngCount, conCol名称, lngCount, conCol名称) = flexAlignLeftCenter
            End Select
            
            '选中指定样式
            If .TextMatrix(lngCount, conCol编号) = lngCode Then .ROW = lngCount
        Next
        If .Rows > 0 Then Call vfgList_AfterSelChange(.ROW, 0, .ROW, 0)
        .TopRow = .ROW
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function zlStyleHeight(lngStdHeight As Long, strParagraph As String) As Long
    '按段落样式计算返回行高度
Dim lngRowHeight As Long
Dim aryFormat() As String
    aryFormat = Split(strParagraph, ";")
    
    '行高度计算
    lngRowHeight = lngStdHeight * 1.3
    Select Case Val(Mid(aryFormat(0), 3, 1))
    Case 0: lngRowHeight = lngStdHeight         '单倍行距
    Case 1: lngRowHeight = lngStdHeight * 1.5   '1.5倍行距
    Case 2: lngRowHeight = lngStdHeight * 2     '两倍行距
    Case 3                                      '最小行距为1行，否则显示精确值。
        If Val(aryFormat(7)) <= 0 Then
            lngRowHeight = lngStdHeight
        ElseIf Val(aryFormat(7)) < lngStdHeight Then
            lngRowHeight = lngStdHeight
        Else
            lngRowHeight = Val(aryFormat(7))
        End If
    Case 4                                      '精确行距。
        If Val(aryFormat(7)) <= 0 Then
            lngRowHeight = lngStdHeight
        Else
            lngRowHeight = Val(aryFormat(7))
        End If
    Case 5      '多倍行距
        If Val(aryFormat(7)) <= 0 Then
            lngRowHeight = lngStdHeight
        Else
            lngRowHeight = lngStdHeight * Val(aryFormat(7))
        End If
    End Select
    '段前段后高度
    If Val(aryFormat(9)) > 0 Then lngRowHeight = lngRowHeight + Val(aryFormat(9))
    If Val(aryFormat(10)) > 0 Then lngRowHeight = lngRowHeight + Val(aryFormat(10))
    zlStyleHeight = lngRowHeight * 20 + Screen.TwipsPerPixelY * 2
End Function
