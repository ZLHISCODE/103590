VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmChildStationDiagnose 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2985
      Index           =   0
      Left            =   840
      ScaleHeight     =   2985
      ScaleWidth      =   4290
      TabIndex        =   2
      Top             =   870
      Width           =   4290
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   285
         Width           =   3990
         _cx             =   7038
         _cy             =   3784
         Appearance      =   1
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2985
      Index           =   1
      Left            =   4320
      ScaleHeight     =   2985
      ScaleWidth      =   4290
      TabIndex        =   0
      Top             =   2670
      Width           =   4290
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   285
         Width           =   3990
         _cx             =   7038
         _cy             =   3784
         Appearance      =   1
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1050
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmChildStationDiagnose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'（１）窗体级变量定义

Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mfrmMain As Object
Private mblnAllowModify As Boolean
Private mlngKey As Long
Private mstrPrivs As String
Private mobjStateInfo As CommandBarControl

Private WithEvents mclsVsfBefore As clsVsf
Attribute mclsVsfBefore.VB_VarHelpID = -1
Private WithEvents mclsVsfAfter As clsVsf
Attribute mclsVsfAfter.VB_VarHelpID = -1

Public Event AfterDataChanged()
Public Event AfterMakeCharge()

'######################################################################################################################
'（２）自定义过程或函数

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
        
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
    
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    
    If ExecuteCommand("读取数据") = False Then Exit Function

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    
    ValidData = True
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Call ExecuteCommand("清空数据")
    
    ClearData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    Dim lngRow As Long
    Dim lngOrderKey As Long
    Dim rs As New ADODB.Recordset
    Dim lng病人id As Long
    Dim lng主页id As Long
    Dim lng挂号id As Long
    Dim intNo As Integer
    
    On Error GoTo errHand
       
    strSQL = "Select 医嘱id,病人id,主页id From 病人手术记录 a Where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, mlngKey)
    If rs.BOF = False Then
        lngOrderKey = zlCommFun.NVL(rs("医嘱id").Value, 0)
        lng病人id = zlCommFun.NVL(rs("病人id").Value, 0)
        lng主页id = zlCommFun.NVL(rs("主页id").Value, 0)
    End If
    If lngOrderKey = 0 Then Exit Function
    
    strSQL = "SELECT b.ID FROM 病人医嘱记录 a,病人挂号记录 b WHERE a.挂号单=b.NO And a.病人来源=1 and a.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngOrderKey)
    If rs.BOF = False Then lng挂号id = zlCommFun.NVL(rs("ID").Value, 0)
    
    '术前诊断
    '------------------------------------------------------------------------------------------------------------------
    intNo = 0
    strSQL = "ZL_病人诊断记录_DELETE2(" & lngOrderKey & ",8)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 And (Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) > 0 Or Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) > 0) Then
                
                intNo = intNo + 1
                strSQL = "ZL_病人诊断记录_INSERT(" & lng病人id & "," & IIf(lng主页id > 0, lng主页id, ZVal(lng挂号id)) & ",1,NULL,8," & Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) & "," & Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) & ",NULL,'" & .TextMatrix(lngRow, .ColIndex("术前诊断")) & "',NULL,NULL,NULL,SYSDATE," & lngOrderKey & "," & intNo & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    
    '术后诊断
    '------------------------------------------------------------------------------------------------------------------
    intNo = 0
    strSQL = "ZL_病人诊断记录_DELETE2(" & lngOrderKey & ",9)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(1)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 And (Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) > 0 Or Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) > 0) Then
                
                intNo = intNo + 1
                strSQL = "ZL_病人诊断记录_INSERT(" & lng病人id & "," & IIf(lng主页id > 0, lng主页id, ZVal(lng挂号id)) & ",1,NULL,9," & Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) & "," & Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) & ",NULL,'" & .TextMatrix(lngRow, .ColIndex("术后诊断")) & "',NULL,NULL,NULL,SYSDATE," & lngOrderKey & "," & intNo & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With

    SaveData = True
    
    Exit Function
    
    '
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function


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
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    
    cbsMain.Options.LargeIcons = False
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_MakeCharge, "按已行手术生成费用", True, , xtpButtonIconAndCaption)

    Set mobjStateInfo = NewToolBar(objBar, xtpControlLabel, 0, "")
    mobjStateInfo.Flags = xtpFlagRightAlign
        
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim objArray As Variant
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        '手术人员
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfBefore = New clsVsf
        With mclsVsfBefore
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("疾病编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("疾病id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("术前诊断", 1500, flexAlignLeftCenter, flexDTString, "", "诊断描述", True)
            .AppendRows = True
        End With
        
        Set mclsVsfAfter = New clsVsf
        With mclsVsfAfter
            Call .Initialize(Me.Controls, vsf(1), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("疾病编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("疾病id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("术后诊断", 1500, flexAlignLeftCenter, flexDTString, "", "诊断描述", True)
            .AppendRows = True
        End With
        
        Call InitCommandBar
        
        Dim objPane As Pane
    
        Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing)
        objPane.Title = "术前诊断"
        objPane.Options = PaneNoCaption
    
        Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane)
        objPane.Title = "术后诊断"
        objPane.Options = PaneNoCaption
    
        dkpMain.SetCommandBars cbsMain
        dkpMain.Options.ThemedFloatingFrames = True
        dkpMain.Options.UseSplitterTracker = False
        dkpMain.Options.AlphaDockingContext = True
        dkpMain.Options.HideClient = True
    
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
    
        mclsVsfBefore.ClearGrid
        mclsVsfAfter.ClearGrid
        mobjStateInfo.Caption = " "
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 Then blnAllowModify = False
        
        If blnAllowModify Then

            With mclsVsfBefore
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("疾病编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("诊断编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("术前诊断"), True, vbVsfEditText)

                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End With

            With mclsVsfAfter
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("疾病编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("诊断编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("术后诊断"), True, vbVsfEditText)

                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End With

        Else

            mclsVsfBefore.AllowEdit = False
            mclsVsfAfter.AllowEdit = False

        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
    
        mobjStateInfo.Caption = " "
        
        gstrSQL = "Select No,记录性质 From 病人手术单据 Where 记录id=[1] And 单据类型=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey, 4)
        If rs.BOF = False Then
            If IsNull(rs("No").Value) = False Then
                Select Case rs("记录性质").Value
                Case 1
                    mobjStateInfo.Caption = "已生成收费单，单据号：" & rs("No").Value
                Case 2
                    mobjStateInfo.Caption = "已生成记帐单，单据号：" & rs("No").Value
                End Select
            End If
        End If
        cbsMain.RecalcLayout
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"

        '拟行诊断
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfBefore.ClearGrid
        gstrSQL = GetPublicSQL(SQL.病人诊断记录)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey, 8)
        If rs.BOF = False Then Call mclsVsfBefore.LoadGrid(rs)

        '已行诊断
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfAfter.ClearGrid
        gstrSQL = GetPublicSQL(SQL.病人诊断记录)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey, 9)
        If rs.BOF = False Then Call mclsVsfAfter.LoadGrid(rs)
        
        Call ExecuteCommand("刷新状态")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "生成手术收费单"
    
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("如果已生成，则会自动删除或作废，确定要生成已行手术费用为收费单吗（新单）？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        gstrSQL = "Select 医嘱id From 病人手术记录 Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            strTmp = MakeChargeBill(Val(rs("医嘱id").Value), 1, "手术", False, mstrPrivs)
            If strTmp <> "" Then
                ShowSimpleMsg "手术收费单已经生成，单据号：" & strTmp
                Call ExecuteCommand("刷新状态")
                RaiseEvent AfterMakeCharge
            End If
        End If
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "生成手术记帐单"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("如果已生成，则会自动删除或作废，确定要生成已行手术费用为记帐单吗（新单）？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        gstrSQL = "Select 医嘱id From 病人手术记录 Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            strTmp = MakeChargeBill(Val(rs("医嘱id").Value), 2, "手术", False, mstrPrivs)
            If strTmp <> "" Then
                ShowSimpleMsg "手术记帐单已经生成，单据号：" & strTmp
                Call ExecuteCommand("刷新状态")
                RaiseEvent AfterMakeCharge
            End If
        End If
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "生成手术零费单"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("如果已生成，则会自动删除或作废，确定要生成已行手术费用为零费单吗（新单）？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        gstrSQL = "Select 医嘱id From 病人手术记录 Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
        
            strTmp = MakeChargeBill(Val(rs("医嘱id").Value), 2, "手术", True, mstrPrivs)
            If strTmp <> "" Then
                ShowSimpleMsg "手术零费单已经生成，单据号：" & strTmp
                Call ExecuteCommand("刷新状态")
                RaiseEvent AfterMakeCharge
            End If
        
        End If
        
        Exit Function
        
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 1

        Call ExecuteCommand("生成手术收费单")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 2

        Call ExecuteCommand("生成手术记帐单")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 3
                
        Call ExecuteCommand("生成手术零费单")
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_MakeCharge

        With CommandBar.Controls

            .DeleteAll
                        
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 1, "收费单据(&1)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 2, "记帐单据(&2)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 3, "零耗费用(&3)")
            With cbsMain.KeyBindings
                .Add FCONTROL, vbKeyN, conMenu_Edit_MakeCharge * 2 + 1
                .Add FCONTROL, vbKeyB, conMenu_Edit_MakeCharge * 2 + 2
            End With
            
        End With
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    blnAllowModify = mblnAllowModify
    If mlngKey = 0 Then blnAllowModify = False

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge
        
        Control.Enabled = blnAllowModify And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 1
        
'        Control.Visible = (mstr病人来源 = "门诊" And IsPrivs(mstrPrivs, "生成附费"))
        Control.Enabled = blnAllowModify And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3

        Control.Enabled = blnAllowModify And Control.Visible
        
    End Select

errHand:

End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

'######################################################################################################################
'（３）窗体及其控件的事件处理

Private Sub Form_Unload(Cancel As Integer)
    Set mobjStateInfo = Nothing
    Set mclsVsfBefore = Nothing
    Set mclsVsfAfter = Nothing
End Sub

Private Sub mclsVsfAfter_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(1).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfAfter_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(1)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub mclsVsfBefore_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(0).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfBefore_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(0)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsfBefore.AppendRows = True
    Case 1
        vsf(1).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsfAfter.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfBefore.AfterEdit(Row, Col)
    Case 1
        Call mclsVsfAfter.AfterEdit(Row, Col)
    End Select

    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case Index
    Case 0
        Call mclsVsfBefore.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Case 1
        Call mclsVsfAfter.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Select Case Index
    Case 0
        mclsVsfBefore.AppendRows = True
    Case 1
        mclsVsfAfter.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0
        mclsVsfBefore.AppendRows = True
    Case 1
        mclsVsfAfter.AppendRows = True
    End Select
End Sub


Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
        mclsVsfBefore.AppendRows = True
    Case 1
        mclsVsfAfter.AppendRows = True
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strTmp As String

    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0, 1                               '术前术后诊断
            If Col = .ColIndex("疾病编码") Or Col = .ColIndex("诊断编码") Then
                Select Case Col
                '------------------------------------------------------------------------------------------------------
                Case .ColIndex("疾病编码")

                    gstrSQL = GetPublicSQL(SQL.疾病编码选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "D")

                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码选择", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then

                        If Index = 2 Then
                            If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                Exit Sub
                            End If
                        End If
                        .EditText = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("编码").Value)
                        If Index = 0 Then
                            .TextMatrix(Row, .ColIndex("术前诊断")) = zlCommFun.NVL(rs("名称").Value)
                        Else
                            .TextMatrix(Row, .ColIndex("术后诊断")) = zlCommFun.NVL(rs("名称").Value)
                        End If
                        .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("ID").Value, 0)

                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                        gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)

                        DataChanged = True
                    End If
                '----------------------------------------------------------------------------------------------------------
                Case .ColIndex("诊断编码")

                    gstrSQL = GetPublicSQL(SQL.疾病诊断选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)

                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\疾病诊断选择", "请从下表中选择一个疾病诊断项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If

                        .EditText = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("编码").Value)
                        If Index = 0 Then
                            .TextMatrix(Row, .ColIndex("术前诊断")) = zlCommFun.NVL(rs("名称").Value)
                        Else
                            .TextMatrix(Row, .ColIndex("术后诊断")) = zlCommFun.NVL(rs("名称").Value)
                        End If
                        .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("ID").Value, 0)

                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                        gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))

                        DataChanged = True
                    End If
                End Select

                '----------------------------------------------------------------------------------------------------------
                If bytRet = 1 Then
                    If rsData.BOF = False Then
                        .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("疾病编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("诊断编码").Value)
                        .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("疾病id").Value, 0)
                        .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("诊断id").Value, 0)
                    End If
                End If

            End If
        End Select
    End With
End Sub
Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        Call mclsVsfBefore.KeyDown(KeyCode, Shift)
    Case 1
        Call mclsVsfAfter.KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim bytRet As Byte
    Dim strClass As String

    With vsf(Index)
        If KeyCode = vbKeyReturn Then

            If InStr(.EditText, "'") > 0 Then
                KeyCode = 0
                .EditText = ""
                Exit Sub
            End If
            strText = UCase(.EditText)
            bytMode = GetApplyMode(strText)
            strText = strText & "%"
            strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, strText, "%" & strText)

            Select Case Index
            Case 0, 1

                If Col = .ColIndex("疾病编码") Or Col = .ColIndex("诊断编码") Then

                    Select Case Col
                    '--------------------------------------------------------------------------------------------------
                    Case .ColIndex("疾病编码")

                        gstrSQL = GetPublicSQL(SQL.疾病编码过滤, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "D")

                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码过滤", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        If bytRet = 1 Then

                            If Index = 2 Then
                                If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            Else
                                If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            End If

                            .EditText = zlCommFun.NVL(rs("编码").Value)
                            .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("编码").Value)
                            If Index = 0 Then
                                .TextMatrix(Row, .ColIndex("术前诊断")) = zlCommFun.NVL(rs("名称").Value)
                            Else
                                .TextMatrix(Row, .ColIndex("术后诊断")) = zlCommFun.NVL(rs("名称").Value)
                            End If
                            .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("ID").Value, 0)

                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                            gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)

                            DataChanged = True
                        End If
                    '--------------------------------------------------------------------------------------------------
                    Case .ColIndex("诊断编码")
                        gstrSQL = GetPublicSQL(SQL.疾病诊断过滤, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\疾病诊断过滤", "请从下表中选择一个疾病诊断项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        If bytRet = 1 Then

                            If Index = 2 Then
                                If mclsVsfBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            Else
                                If mclsVsfAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            End If

                            .EditText = zlCommFun.NVL(rs("编码").Value)
                            .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("编码").Value)
                            If Index = 0 Then
                                .TextMatrix(Row, .ColIndex("术前诊断")) = zlCommFun.NVL(rs("名称").Value)
                            Else
                                .TextMatrix(Row, .ColIndex("术后诊断")) = zlCommFun.NVL(rs("名称").Value)
                            End If
                            .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("ID").Value, 0)

                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                            gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))

                            DataChanged = True
                        End If
                    End Select

                    If bytRet = 1 Then

                        '--------------------------------------------------------------------------------------------------
                        If rsData.BOF = False Then

                            .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("疾病编码").Value)
                            .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("诊断编码").Value)

                            .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("疾病id").Value, 0)
                            .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("诊断id").Value, 0)
                        End If

                    Else
                        KeyCode = 0
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If

                End If

            End Select
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfBefore.KeyPress(KeyAscii)
    Case 1
        Call mclsVsfAfter.KeyPress(KeyAscii)
    End Select
    
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfBefore.KeyPressEdit(KeyAscii)
    Case 1
        Call mclsVsfAfter.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
        Case 0
            Call mclsVsfBefore.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 1
            Call mclsVsfAfter.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfBefore.EditSelAll
    Case 1
        Call mclsVsfAfter.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfBefore.BeforeEdit(Row, Col, Cancel)
    Case 1
        Call mclsVsfAfter.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfBefore.ValidateEdit(Col, Cancel)
    Case 1
        Call mclsVsfAfter.ValidateEdit(Col, Cancel)
    End Select
End Sub


