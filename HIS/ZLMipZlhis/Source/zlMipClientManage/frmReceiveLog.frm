VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmReceiveLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   3
      Left            =   630
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   12
      Top             =   4170
      Width           =   2700
      Begin RichTextLib.RichTextBox rtb 
         Height          =   1440
         Left            =   255
         TabIndex        =   13
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2540
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmReceiveLog.frx":0000
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1770
      Index           =   1
      Left            =   6510
      ScaleHeight     =   1770
      ScaleWidth      =   2385
      TabIndex        =   10
      Top             =   3795
      Width           =   2385
      Begin XtremeSuiteControls.TabControl tbc 
         Height          =   1365
         Left            =   195
         TabIndex        =   11
         Top             =   135
         Width           =   1770
         _Version        =   589884
         _ExtentX        =   3122
         _ExtentY        =   2408
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   2
      Left            =   3450
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   8
      Top             =   975
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
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
         GridColor       =   12632256
         GridColorFixed  =   12632256
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
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   0
      Left            =   3495
      ScaleHeight     =   1980
      ScaleWidth      =   2700
      TabIndex        =   7
      Top             =   3765
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsfXML 
         Height          =   1395
         Left            =   255
         TabIndex        =   14
         Top             =   285
         Width           =   2190
         _cx             =   3863
         _cy             =   2461
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
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483634
         GridColor       =   12632256
         GridColorFixed  =   12632256
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483634
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   8
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
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   3390
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   75
      Width           =   1305
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   1
         Left            =   -30
         TabIndex        =   6
         Top             =   -30
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   130416643
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   1380
      ScaleHeight     =   240
      ScaleWidth      =   1245
      TabIndex        =   3
      Top             =   0
      Width           =   1275
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   -30
         TabIndex        =   4
         Top             =   -30
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   130416643
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   285
      Width           =   870
      Begin VB.ComboBox cboPeiord 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   -30
         Width           =   900
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7425
      TabIndex        =   0
      Top             =   585
      Width           =   1575
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   210
      Top             =   750
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmReceiveLog.frx":009D
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmReceiveLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义

Private Enum Command
    初始控件
    读注册表
    删除日志
    刷新内容
    刷新日志
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mclsVsfXML As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnReading As Boolean

Public Event AfterClose()
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'接口方法
Public Function ShowForm()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call Form_Activate
End Function

'######################################################################################################################
Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim strTmp As String
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim blnMuliSelect As Boolean
    Dim strTemp As String
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strLine As String
        
    On Error GoTo errHand
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.初始控件
        
        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel
        
        With tbc
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .BoldSelected = True
                .Color = xtpTabColorDefault
                .ColorSet.ButtonSelected = Color.白色
                .ShowIcons = True
            End With
                        
            .InsertItem 0, "XML ", picPane(0).hWnd, 0
            .InsertItem 1, "文本 ", picPane(3).hWnd, 0

            .Item(0).Selected = True
        End With
        
        mblnReading = True
        With cboPeiord
            .Clear
            .AddItem "今  天"
            .AddItem "昨  天"
            .AddItem "前三天"
            .AddItem "本  周"
            .AddItem "前一周"
            .AddItem "前半月"
            .AddItem "本  月"
            .AddItem "前一月"
            .AddItem "前二月"
            .AddItem "本  季"
            .AddItem "前三月"
            .AddItem "本半年"
            .AddItem "前半年"
            .AddItem "自定义"
        End With
        If cboPeiord.ListCount > 0 And cboPeiord.ListIndex = -1 Then cboPeiord.ListIndex = 0
        mblnReading = False
        
        dtp(0).Value = Format(GetBasePeriod(cboPeiord.Text, 1), dtp(0).CustomFormat)
        dtp(1).Value = Format(GetBasePeriod(cboPeiord.Text, 2), dtp(1).CustomFormat)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除日志
                
        Set rsPara = zlCommFun.CreateParameter
        
        With vsf(0)
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("选择")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            If blnMuliSelect = True Then
                If MsgBox("您确认要删除已经勾选的发送日志吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                                                        
                    strTemp = ""
                    lngCount = 0
                    
                    For lngLoop = 1 To .Rows - 1
                        If Val(Abs(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And .TextMatrix(intRow, .ColIndex("id")) <> "" Then
                            strLine = ""
                            strLine = Trim(.TextMatrix(lngLoop, .ColIndex("id")))
                            If LenB(strTemp & ";" & strLine) > 3500 Then
                                If strTemp <> "" Then
                                    lngCount = lngCount + 1
                                    strTemp = Mid(strTemp, 2)
                                    Call zlCommFun.SetParameter(rsPara, "SendLog_" & lngCount, strTemp)
                                    strTemp = ""
                                End If
                            End If
                            strTemp = strTemp & ";" & strLine
                        End If
                    Next
                    If strTemp <> "" Then
                        lngCount = lngCount + 1
                        strTemp = Mid(strTemp, 2)
                        Call zlCommFun.SetParameter(rsPara, "SendLog_" & lngCount, strTemp)
                    End If
                    Call zlCommFun.SetParameter(rsPara, "SendLog_SegmentCount", lngCount)
                    If gclsBusiness.ReceiveLogEdit("Delete", rsPara) Then
                        Call ExecuteCommand(Command.刷新日志)
                    End If
                End If
            ElseIf .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                Call zlCommFun.SetParameter(rsPara, "SendLog_1", .TextMatrix(.Row, .ColIndex("id")))
                Call zlCommFun.SetParameter(rsPara, "SendLog_SegmentCount", 1)
                If gclsBusiness.ReceiveLogEdit("Delete", rsPara) Then
                    Call ExecuteCommand(Command.刷新日志)
                End If
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新日志
        
        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
            
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "Start_Date", Format(dtp(0).Value, dtp(0).CustomFormat))
            Call zlCommFun.SetCondition(rsCondition, "End_Date", Format(dtp(1).Value, dtp(1).CustomFormat))
            
            If Trim(txtLocation.Text) = "" Then
                ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ReceiveLogRead("FilterData", rsCondition))
            Else
                Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
                Call zlCommFun.SetCondition(rsCondition, "FilterText", Trim(txtLocation.Text))
                ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.ReceiveLogRead("FilterData", rsCondition))
            End If

            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新内容
        
        With vsf(0)
            Dim objXML As New clsMessageXML
            Dim intMaxOutlineLevel As Integer
            
            Call mclsVsfXML.ClearGrid
            rtb.Text = ""
            
            If Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" Then
                strTemp = gclsBusiness.GetReceiveLogContent(Trim(.TextMatrix(.Row, .ColIndex("id"))))
                If strTemp <> "" Then
                    rtb.Text = strTemp
                    Call objXML.LoadString(strTemp)
                    Set rsTmp = objXML.GetNodeData(".//message")
                    
                    With mclsVsfXML
                        Call .LoadGrid(rsTmp)
                        Call .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"), .VsfObject.BackColor)
                        Call ExpandAllOutline(vsfXML)
                        .VsfObject.AutoSize 1, .VsfObject.Cols - 1
                        .VsfObject.Cell(flexcpForeColor, 1, .ColIndex("节点内容"), .VsfObject.Rows - 1, .ColIndex("节点内容")) = Color.兰色
                    End With
                End If

            End If
            
        End With
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
    On Error GoTo errHand
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[序号]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("接收时间", 1830, flexAlignLeftCenter, flexDTString, , "receive_date", True)
        Call .AppendColumn("消息用户", 900, flexAlignLeftCenter, flexDTString, , "receive_mipuser", True)
        Call .AppendColumn("产品用户", 1500, flexAlignLeftCenter, flexDTString, , "receive_user", True)
        Call .AppendColumn("接收站点", 2400, flexAlignLeftCenter, flexDTString, , "receive_station", True)
        Call .AppendColumn("接收模块", 3000, flexAlignLeftCenter, flexDTString, , "receive_module", True)
        Call .AppendColumn("接收消息", 1500, flexAlignLeftCenter, flexDTString, , "item_identity", True)
        
        Call .AppendColumn("说明", 1500, flexAlignLeftCenter, flexDTString, , "receive_note", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
        
        .AppendRows = True
    End With
            
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsfXML = New zlVSFlexGrid.clsVsf
    With mclsVsfXML
        Call .Initialize(Me.Controls, vsfXML, True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 720, flexAlignCenterCenter, flexDTString, , "[序号]", False, False, False)
        Call .AppendColumn("id", 1200, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("parent_id", 1200, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("节点名称", 2100, flexAlignLeftCenter, flexDTString, , "nodename", True)
        Call .AppendColumn("节点内容", 2100, flexAlignLeftCenter, flexDTString, , "nodetext", True)
                                        
        .VsfObject.OutlineCol = .ColIndex("节点名称")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        .AppendRows = False
    End With
    
    InitGrid = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
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
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = zlCommFun.GetPubIcons
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
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Parameter, "参数", True)
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "全选", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "全清")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "清除", True)
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "时间：", , , xtpButtonCaption)
    objControl.IconId = 99999999
    objControl.BeginGroup = True
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(2).hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "从", , , xtpButtonCaption)
    objControl.IconId = 99999999
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(3).hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "到", , , xtpButtonCaption)
    objControl.IconId = 99999999
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(4).hWnd
    
        
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "接收站点")
    If mstrFindKey = "" Then mstrFindKey = "接收站点"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, False, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.消息用户"): objControl.Parameter = "消息用户"
    objControl.IconId = 99999999
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.产品用户"): objControl.Parameter = "产品用户"
    objControl.IconId = 99999999
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.接收站点"): objControl.Parameter = "接收站点"
    objControl.IconId = 99999999
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&4.接收模块"): objControl.Parameter = "接收模块"
    objControl.IconId = 99999999
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&5.接收消息"): objControl.Parameter = "接收消息"
    objControl.IconId = 99999999
    
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Close, "关闭")
    objControl.flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh           '刷新
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete
    End With
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing)
    objPane.Title = "记录"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane)
    objPane.Title = "内容"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub


Private Sub cboPeiord_Click()
    If mblnReading Then Exit Sub
    
    If cboPeiord.Text <> "自定义" Then
        dtp(0).Value = Format(GetBasePeriod(cboPeiord.Text, 1), dtp(0).CustomFormat)
        dtp(1).Value = Format(GetBasePeriod(cboPeiord.Text, 2), dtp(1).CustomFormat)
        
        Call ExecuteCommand(Command.刷新日志)
        Call ExecuteCommand(Command.刷新内容)
    Else
        mblnDataChanged = True
    End If
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Parameter
        Call frmReceiveLogPara.ShowConfigDialog(Me)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 1
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
    
        Call ExecuteCommand(Command.删除日志)
                    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新
                
        Call ExecuteCommand(Command.刷新日志)
        Call ExecuteCommand(Command.刷新内容)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
        RaiseEvent AfterClose
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem
            Control.Checked = (mstrFindKey = Control.Parameter)
        End Select
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(2).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub dtp_Change(Index As Integer)
    '更改时间段名称为“自定义“
    mblnReading = True
    
    Select Case Index
    Case 0, 1
        Call zlControl.CboLocate(cboPeiord, "自定义")
    End Select
    
    mblnReading = False
    
    mblnDataChanged = True
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    Call ExecuteCommand(Command.刷新日志)
    Call ExecuteCommand(Command.刷新内容)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1005
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

    Call zlComLib.RestoreWinState(Me, App.ProductName)
'    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
'    Call zlDataBase.ShowReportMenu(Me, ParamInfo.系统号, ParamInfo.模块号, UserInfo.模块权限)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 2, 300, 15, 500, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mclsVsfXML = Nothing
    Set mobjFindKey = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vsfXML.Move 0, 15, picPane(Index).Width, picPane(Index).Height
    Case 1
        tbc.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 3
        rtb.Move 0, 15, picPane(Index).Width, picPane(Index).Height
    Case 2
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

'        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Refresh, True)
            If obj Is Nothing Then Exit Sub
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

'        End If
'        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf(Index).AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If NewRow <> OldRow Then
        Call ExecuteCommand(Command.刷新内容)
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey, .ColIndex("id"))
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        mclsVsf(Index).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
    End With
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Call mclsVsf(Index).BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim objMenu As CommandBarControl
    
    Set objMenu = cbsMain.FindControl(, conMenu_Edit_Modify, False)
    If Not (objMenu Is Nothing) Then
        If objMenu.Enabled = True Then
            Call cbsMain_Execute(objMenu)
        End If
    End If
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf(Index).KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mclsVsf(Index).MoveColumn = (vsf(Index).MouseRow = 0)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call zlCommFun.SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    
    Select Case bytPlace
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '

        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&S)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "全部勾选(&A)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "全部不选(&U)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "清除日志(&D)")
        cbrPopupItem.BeginGroup = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "重新刷新(&R)")
        cbrPopupItem.BeginGroup = True
        
        
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(0).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).ValidateEdit(Col, Cancel)
End Sub

Private Sub vsfXML_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
'    mclsVsfXML.AppendRows = True
End Sub

Private Sub vsfXML_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfXML.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfXML_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsfXML.AppendRows = True
End Sub

Private Sub vsfXML_AfterSort(ByVal Col As Long, Order As Integer)
'    mclsVsfXML.AppendRows = True
End Sub

Private Sub vsfXML_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'    mclsVsfXML.AppendRows = True
End Sub
