VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmProcEditMark 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2370
      Index           =   0
      Left            =   405
      ScaleHeight     =   2370
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   585
      Width           =   2055
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1755
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   405
         Width           =   1935
         _cx             =   3413
         _cy             =   3096
         Appearance      =   1
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmProcEditMark.frx":0000
      Left            =   510
      Top             =   45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcEditMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mobjMain As Object
Private mlngKey As Long
Public Event AfterAdd(ByVal strSign As String)
Public Event AfterDelete(ByVal strSign As String)
Public Event AfterChanged()

Private Sub InitCommandBar()
    '******************************************************************************************************************
    '功能：初始菜单工具栏
    '参数：无
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    objBar.SetIconSize 16, 16
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新建")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
End Sub

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, objPane)
    objPane.Title = "基本信息"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

End Sub

Public Function InitData(ByVal objMain As Object, ByVal lngKey As Long)
    Set mobjMain = objMain
    mlngKey = lngKey
    Call ExecuteCommand("初始控件")
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("刷新数据")
End Function

Public Function SaveData(ByVal lngKey As Long) As Boolean
    If mlngKey = 0 Then
        mlngKey = lngKey
    End If
    If ExecuteCommand("保存数据") Then
        SaveData = True
    End If
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    Dim objItem As Object
    Dim intRow As Integer
    Dim rsSQL As ADODB.Recordset
    
    On Error GoTo errHand
    
    Call gclsBase.SQLRecord(rsSQL)
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Call InitCommandBar
        Call InitDockPannel
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("标识", 800, flexAlignLeftCenter, flexDTString, , "标识", True)
            Call .AppendColumn("说明", 2000, flexAlignLeftCenter, flexDTString, , "说明", True)
            
            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(.ColIndex("说明"), True, vbVsfEditText)
            
            .AppendRows = True
        End With
    Case "刷新数据"
        strSQL = "Select 过程ID as ID,标识,说明 From zlProcedureNote Where 过程ID=[1]"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mlngKey)
        If rs.BOF = False Then
            Call mclsVsf.LoadGrid(rs)
        End If
    Case "删除标识"
        RaiseEvent AfterDelete(vsf(0).TextMatrix(vsf(0).Row, vsf(0).ColIndex("标识")))
        Call mclsVsf.DeleteRow(vsf(0).Row)
    Case "保存数据"
        With vsf(0)
            strSQL = "Zl_zlProcedureNote_Delete(" & mlngKey & ")"
            Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            For intRow = 1 To .Rows - 1
                strSQL = "Zl_zlProcedureNote_Update(" & mlngKey & ",'" & .TextMatrix(intRow, .ColIndex("标识")) & "','" & .TextMatrix(intRow, .ColIndex("说明")) & "')"
                Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
            Next
        End With
        Call SQLRecordExecute(rsSQL)
    End Select
    ExecuteCommand = True
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    Case conMenu_Edit_NewItem
        Call mclsVsf.AppendRow
        RaiseEvent AfterAdd(vsf(0).TextMatrix(vsf(0).Row, vsf(0).ColIndex("标识")))
    Case conMenu_Edit_Delete
        Call ExecuteCommand("删除标识")
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case 1
        Item.handle = picPane(0).hwnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
End Sub

Private Sub mclsVsf_AfterNewRow(ByVal Row As Long, Col As Long)
    '自动产生标识
    vsf(0).TextMatrix(Row, vsf(0).ColIndex("标识")) = CStr(Val(vsf(0).TextMatrix(Row - 1, vsf(0).ColIndex("标识"))) + 1)
    RaiseEvent AfterChanged
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (vsf(0).TextMatrix(Row, vsf(0).ColIndex("标识")) = "")
    If Cancel = True Then
        vsf(0).TextMatrix(Row, vsf(0).ColIndex("标识")) = "1001"
    End If
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vsf(0).Move 15, 15, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight - 30
    End Select
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    RaiseEvent AfterChanged
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub


