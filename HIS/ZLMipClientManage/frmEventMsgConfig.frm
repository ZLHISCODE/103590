VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmEventMsgConfig 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   1
      Left            =   405
      ScaleHeight     =   2685
      ScaleWidth      =   5220
      TabIndex        =   0
      Top             =   1320
      Width           =   5220
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   225
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
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
         RowHeightMin    =   270
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
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
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
         WordWrap        =   -1  'True
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
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEventMsgConfig.frx":0000
      Left            =   690
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEventMsgConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mlngModualCode As Long
Private mstrPrivs As String
Private mstrSQL As String
Private WithEvents mclsVsf As zlVSFlexGrid.clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mblnStartUp As Boolean
Private mlngTmp As Long
Private mblnShowAll As Boolean
Private mblnShowStop As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean
Private mfrmParent As Object
Private mlngMoudalCode As Long
Private mstrDataKey As String
Private mrsPara As ADODB.Recordset
Private mintMaxOutlineLevel As Integer
Private mintAryOutlineLevel(0 To 11) As Integer
Private mobjToolBar As Object

Private Enum Command
    初始控件
    读注册表
    刷新数据

End Enum

'######################################################################################################################

Public Function InitForm(ByVal frmParent As Object, ByVal lngMoudalCode As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngMoudalCode = lngMoudalCode
    
    InitForm = True
    
End Function

Public Function ReadData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrDataKey = strDataKey
    
    Call ExecuteCommand(Command.刷新数据)
    
    ReadData = True
    
End Function


Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant
    
    On Error GoTo errHand
            
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件
        '
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新数据
        
        mintMaxOutlineLevel = 0
        
        Set mrsPara = zlCommFun.CreateCondition
        Call zlCommFun.SetCondition(mrsPara, "事件消息id", mstrDataKey)
    
        Call mclsVsf.LoadGrid(gclsBusiness.EventMsgConfigRead("事件消息", mrsPara))
        Call mclsVsf.VsfObject.AutoSize(mclsVsf.ColIndex("数据赋值"), mclsVsf.ColIndex("数据赋值"))
        mintMaxOutlineLevel = mclsVsf.ShowOutline(mclsVsf.ColIndex("id"), mclsVsf.ColIndex("parent_id"))
        mobjToolBar.Visible = (mintMaxOutlineLevel > 0)
        
        For intLoop = mintMaxOutlineLevel To 1 Step -1
            If intLoop < 12 Then mintAryOutlineLevel(intLoop) = 1
            Call mclsVsf.Outline(intLoop)
        Next
                
    End Select
    
    
    GoTo EndHand

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    '初始网格控件
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(0), True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 720, flexAlignCenterCenter, flexDTString, , "[序号]", False, False, False)
        Call .AppendColumn("节点类型", 2100, flexAlignLeftCenter, flexDTString, , "节点类型", True)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("parent_id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        
        Call .AppendColumn("节点标题", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("数据类型", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("重复频率", 810, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("数据重复", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("数据赋值", 6000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("节点说明", 900, flexAlignLeftCenter, flexDTString, , "", True)
                
        .VsfObject.OutlineCol = .ColIndex("节点类型")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
    End With

    
    InitGrid = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 300, DockTopOf, Nothing)
    objPane.Title = "SQL"
    objPane.Options = PaneNoCaption
       
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    
    Set mobjToolBar = cbsMain.Add("工具栏", xtpBarTop)
    mobjToolBar.ContextMenuPresent = False
    mobjToolBar.ShowTextBelowIcons = False
    mobjToolBar.EnableDocking xtpFlagStretched
        
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlLabel, 0, "层次:", True, , xtpButtonCaption)
    For intLoop = 1 To 10
        Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, 1, " " & intLoop & " ", , , xtpButtonCaption, "选中展开当前层，不选中则收拢当前层")
        objControl.Parameter = intLoop
    Next
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, 1, "更多", , , xtpButtonCaption, "选中展开当前层，不选中则收拢当前层")
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim intIndex As Integer
    Dim intMaxIndex As Integer
    
    Select Case Control.ID
    Case 1
        intIndex = Val(Control.Parameter)
        If mintAryOutlineLevel(intIndex) = 1 Then
            '展开,前面的自动展开
            
            With vsf(0)
                If intIndex > 10 Then
                    intMaxIndex = mintMaxOutlineLevel
                Else
                    intMaxIndex = intIndex
                End If
                
                For lngRow = 1 To .Rows - 1
                    If .IsSubtotal(lngRow) = True And .RowOutlineLevel(lngRow) <= intMaxIndex Then
                        .IsCollapsed(lngRow) = flexOutlineExpanded
                    End If
                Next
            End With
            
            For intLoop = intIndex To 1 Step -1
                mintAryOutlineLevel(intLoop) = -1
            Next
            
        Else
            If intIndex > 10 Then
                For intLoop = 11 To mintMaxOutlineLevel
                    mclsVsf.Outline intLoop
                Next
            Else
                mclsVsf.Outline intIndex
            End If
            
            mintAryOutlineLevel(intIndex) = 1
        End If
        
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case 1
        
        Control.Checked = (mintAryOutlineLevel(Val(Control.Parameter)) = -1)
        Control.Visible = (Val(Control.Parameter) > 0 And Val(Control.Parameter) <= mintMaxOutlineLevel)
        
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Call InitGrid
    Call InitDockPannel
    Call InitCommandBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf = Nothing
    
    Set mobjFindKey = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub


Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub


Private Sub vsf_DblClick(Index As Integer)
    Dim lngRow As Long
    
    Call mclsVsf.DbClick
    
    With vsf(Index)
        
        Select Case .Col
        Case .ColIndex("数据重复"), .ColIndex("数据赋值")
        
        Case Else
            
            lngRow = .Row
            
            If .IsSubtotal(lngRow) = True Then
                .IsCollapsed(lngRow) = IIf(.IsCollapsed(lngRow) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
            End If
            
        End Select
    End With
    
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
        
    With vsf(Index)
        If KeyAscii = vbKeySpace Then
            Select Case .Col
            Case .ColIndex("数据重复"), .ColIndex("数据赋值")
            
            Case Else
                Call vsf_DblClick(Index)
            End Select
        End If
    End With
End Sub
