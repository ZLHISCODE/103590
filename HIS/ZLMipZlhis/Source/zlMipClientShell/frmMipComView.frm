VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMipComView 
   Caption         =   "通用消息查阅"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13410
   Icon            =   "frmMipComView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   13410
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4245
      Index           =   1
      Left            =   240
      ScaleHeight     =   4245
      ScaleWidth      =   2370
      TabIndex        =   10
      Top             =   1695
      Width           =   2370
      Begin XtremeSuiteControls.TaskPanel tpl 
         Height          =   4770
         Left            =   345
         TabIndex        =   11
         Top             =   495
         Width           =   3210
         _Version        =   589884
         _ExtentX        =   5662
         _ExtentY        =   8414
         _StockProps     =   64
         Behaviour       =   1
         ItemLayout      =   2
         HotTrackStyle   =   3
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3885
      Index           =   0
      Left            =   5505
      ScaleHeight     =   3885
      ScaleWidth      =   5070
      TabIndex        =   3
      Top             =   3900
      Width           =   5070
      Begin RichTextLib.RichTextBox txtText 
         Height          =   1515
         Left            =   15
         TabIndex        =   9
         Top             =   945
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   2672
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMipComView.frx":0A02
      End
      Begin VB.PictureBox picBack 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   915
         Index           =   1
         Left            =   15
         ScaleHeight     =   915
         ScaleWidth      =   5040
         TabIndex        =   4
         Top             =   15
         Width           =   5040
         Begin VB.Label lblLinkTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药品付款单据"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   915
            MouseIcon       =   "frmMipComView.frx":0A9F
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label lblLinkType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "链    接："
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   45
            TabIndex        =   12
            Top             =   630
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主题内容"
            Height          =   180
            Index           =   1
            Left            =   915
            TabIndex        =   8
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主    题："
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "接收时间："
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   6
            Top             =   90
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2014-01-10 16:58:00"
            Height          =   180
            Index           =   2
            Left            =   915
            TabIndex        =   5
            Top             =   90
            Width           =   1710
         End
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7425
      TabIndex        =   2
      Top             =   585
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2010
      Index           =   2
      Left            =   5535
      ScaleHeight     =   2010
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   1680
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1785
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   2670
         _cx             =   4710
         _cy             =   3149
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   3285
      Top             =   2445
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMipComView.frx":0DA9
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
      Bindings        =   "frmMipComView.frx":20CF
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMipComView"
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
    删除消息
    阅读消息
    标记阅读
    刷新内容
    刷新消息
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mstrDataFile As String
Private mclsMipReceiptData As clsMipReceiptData
Private mstrCurrentGroup As String
Private mobjParentForm As Object

Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
Public Event AfterReadMessage()

'######################################################################################################################
'接口方法
Public Function ShowForm(ByVal objParentForm As Object, ByVal strDataFile As String, Optional ByVal blnOnlyNew As Boolean = False)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrDataFile = strDataFile
    
    mstrCurrentGroup = "G1"
    If blnOnlyNew Then
        mstrCurrentGroup = "G2"
        If tpl.Groups.Count > 0 Then
            tpl.Groups(1).Items(1).Selected = False
            tpl.Groups(1).Items(2).Selected = True
        End If
    End If
    
    Set mobjParentForm = objParentForm
    Me.Show , mobjParentForm
    
    Call ExecuteCommand(Command.刷新消息)
    Call ExecuteCommand(Command.刷新内容)
    
End Function

'######################################################################################################################
Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, gfrmMipResource.ils16)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[序号]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTBoolean, "", "[选择]", False)
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("阅读标记", 0, flexAlignLeftCenter, flexDTString, , "receive_read", True, , , True)
        Call .AppendColumn("文本内容", 0, flexAlignLeftCenter, flexDTString, , "receive_text", True, , , True)
        Call .AppendColumn("链接类型", 0, flexAlignLeftCenter, flexDTString, , "receive_lnk_type", True, , , True)
        Call .AppendColumn("链接标题", 0, flexAlignLeftCenter, flexDTString, , "receive_lnk_title", True, , , True)
        Call .AppendColumn("链接参数", 0, flexAlignLeftCenter, flexDTString, , "receive_lnk_para", True, , , True)
        
        Call .AppendColumn("时间", 1800, flexAlignLeftCenter, flexDTString, , "receive_date", True)
        Call .AppendColumn("主题", 1800, flexAlignLeftCenter, flexDTString, , "receive_topic", True)
                        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("序号")
        .ConstCol = .ColIndex("序号")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
        
    End With
        
    InitGrid = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
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
        Call InitTaskPanel
        
        Set mclsMipReceiptData = New clsMipReceiptData
        Call mclsMipReceiptData.Initialize(mstrDataFile)
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.删除消息
        With vsf(0)
                        
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("选择")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            
            If mclsMipReceiptData.OpenDataFile = True Then
                If blnMuliSelect = True Then
                    If MsgBox("您确认要删除已经勾选的接收消息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        For lngLoop = 1 To .Rows - 1
                            If Abs(Val(.TextMatrix(lngLoop, .ColIndex("选择")))) = 1 And .TextMatrix(lngLoop, .ColIndex("id")) <> "" Then
                                Call mclsMipReceiptData.DeleteReceiveMessage(.TextMatrix(lngLoop, .ColIndex("id")))
                            End If
                        Next
                        Call ExecuteCommand(Command.刷新消息)
                    End If
                ElseIf .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                    If MsgBox("您确认要删除当前选中行的接收消息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If mclsMipReceiptData.DeleteReceiveMessage(.TextMatrix(.Row, .ColIndex("id"))) Then
                            Call ExecuteCommand(Command.刷新消息)
                        End If
                    End If
                End If
                mclsMipReceiptData.CloseDataFile
            End If
            
            
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.阅读消息
        With vsf(0)
            If mclsMipReceiptData.OpenDataFile = True Then
                blnMuliSelect = False
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 Then
                        blnMuliSelect = True
                        Exit For
                    End If
                Next
                
                If blnMuliSelect = True Then
                    If MsgBox("您确认要将这些消息都标记为已阅读吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        For lngLoop = 1 To .Rows - 1
                            If Abs(Val(.TextMatrix(intRow, .ColIndex("选择")))) = 1 And .TextMatrix(intRow, .ColIndex("id")) <> "" And .TextMatrix(intRow, .ColIndex("阅读标记")) <> "1" Then
                                If mclsMipReceiptData.UpdateReceiveMessageReaded(.TextMatrix(intRow, .ColIndex("id"))) Then
                                    .TextMatrix(intRow, .ColIndex("阅读标记")) = "1"
                                    .Cell(flexcpFontBold, intRow, 1, intRow, .Cols - 1) = False
                                End If
                            End If
                        Next
                        RaiseEvent AfterReadMessage
                    End If
                ElseIf .TextMatrix(.Row, .ColIndex("id")) <> "" And .TextMatrix(.Row, .ColIndex("阅读标记")) <> "1" Then
                    
                    If mclsMipReceiptData.UpdateReceiveMessageReaded(.TextMatrix(.Row, .ColIndex("id"))) Then
                        .TextMatrix(.Row, .ColIndex("阅读标记")) = "1"
                        .Cell(flexcpFontBold, .Row, 1, .Row, .Cols - 1) = False
                        RaiseEvent AfterReadMessage
                    End If
                    
                End If
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新消息
        
        With vsf(0)
            
            If mclsMipReceiptData.OpenDataFile() = True Then
            
                mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
                mclsVsf(0).ClearGrid
                Set rsCondition = zlCommFun.CreateCondition
                
                
                rsTmp = mclsMipReceiptData.ReadReceiveMessage("Count", rsCondition)
                If rsTmp.DataSet.BOF = False Then
                    If Val(rsTmp.DataSet("未读数").Value) > 0 Then
                        tpl.Groups(1).Items(2).Caption = "未读消息[" & Val(rsTmp.DataSet("未读数").Value) & "]"
                        tpl.Groups(1).Items(2).Bold = True
                    Else
                        tpl.Groups(1).Items(2).Caption = "未读消息"
                        tpl.Groups(1).Items(2).Bold = False
                    End If
                    tpl.Reposition
                End If
                
'                Call zlCommFun.SetCondition(rsCondition, "Start_Date", Format(dtp(0).Value, dtp(0).CustomFormat))
'                Call zlCommFun.SetCondition(rsCondition, "End_Date", Format(dtp(1).Value, dtp(1).CustomFormat))

                
                Call zlCommFun.SetCondition(rsCondition, "receive_read", IIf(mstrCurrentGroup = "G2", 1, 0))
                
                If Trim(txtLocation.Text) = "" Then
                    
                    rsTmp = mclsMipReceiptData.ReadReceiveMessage("FilterData", rsCondition)
                    If rsTmp.DataSet.BOF = False Then ExecuteCommand = mclsVsf(0).LoadDataSource(rsTmp.DataSet.DataSource)

                Else
                    Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
                    Call zlCommFun.SetCondition(rsCondition, "FilterText", Trim(txtLocation.Text))
                    
                    rsTmp = mclsMipReceiptData.ReadReceiveMessage("FilterData", rsCondition)
                    If rsTmp.DataSet.BOF = False Then ExecuteCommand = mclsVsf(0).LoadDataSource(rsTmp.DataSet.DataSource)
                    
                End If
    
                Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
                
                For intRow = 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, .ColIndex("阅读标记"))) = 0 Then
                        .Cell(flexcpFontBold, intRow, 1, intRow, .Cols - 1) = True
                    End If
                Next
                
                mclsMipReceiptData.CloseDataFile
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.刷新内容
        With vsf(0)
            txtText.Text = ""
            lblLinkType.Tag = ""
            lblLinkTitle.Caption = ""
            lblLinkTitle.Tag = ""
            txtText.Text = .TextMatrix(.Row, .ColIndex("文本内容"))
            lbl(2).Caption = .TextMatrix(.Row, .ColIndex("时间"))
            lbl(1).Caption = .TextMatrix(.Row, .ColIndex("主题"))
            lblLinkType.Tag = .TextMatrix(.Row, .ColIndex("链接类型"))
            lblLinkTitle.Caption = .TextMatrix(.Row, .ColIndex("链接标题"))
            lblLinkTitle.Tag = .TextMatrix(.Row, .ColIndex("链接参数"))
            lblLinkTitle.Visible = (lblLinkTitle.Caption <> "")
            
            '如果为未阅读，处理为已阅读
            Call ExecuteCommand(Command.阅读消息, 0)

    
        End With
    End Select
        
    GoTo EndHand

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mclsMipReceiptData.CloseDataFile
    '------------------------------------------------------------------------------------------------------------------
EndHand:
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
'    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = frmMipResource.imgPublic.Icons
    cbsMain.Options.LargeIcons = True
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap


    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True, , , "退出运行日志查阅功能")
    
    '------------------------------------------------------------------------------------------------------------------
    '分类
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "全选(&A)", , , , "将当前列表中的所有数据置为勾选状态")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "全清(&C)", , , , "将当前列表中的所有数据置为非勾选状态")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "清除(&D)", True, , , "清除当前行或者勾选中的通用消息")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "标记为已读(&M)", True, , , "将当前行或者勾选中的通用消息设置为已读状态")
    
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = zlCommFun.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", , , , "显示/隐藏工具栏按钮")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", , , , "显示/隐藏工具栏按钮上的文字内容")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", , , , "设置工具栏按钮图标为大图标或小图标")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)", , , , "显示/隐藏状态栏")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True, , , "按当前设置的条件重新刷新通用消息数据")
    
    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)", , , , "显示关于通用消息查阅的操作说明")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True, , , "显示有关通用消息的相关说明")
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        
            
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "全选", True, , , , , "将当前列表中的所有数据置为勾选状态")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "全清", , , , , , "将当前列表中的所有数据置为非勾选状态")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "清除", True, , , , , "清除当前行或者勾选中的通用消息")
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "标记为已读", True, , , , , "将当前行或者勾选中的通用消息设置为已读状态")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "主题", False, , xtpButtonIconAndCaption)
    objControl.IconId = conMenu_View_Find
        
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新", , , , , , "按当前设置的条件重新刷新通用消息数据")
            
    cbsMain.StatusBar.Visible = True
    cbsMain.StatusBar.IdleText = "准备"
    Call cbsMain.StatusBar.AddPane(0)
    Call cbsMain.StatusBar.SetPaneText(0, cbsMain.StatusBar.IdleText)
    Call cbsMain.StatusBar.SetPaneStyle(0, SBPS_STRETCH)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_CAPS)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_NUM)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_SCRL)

    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add FCONTROL, vbKeyDelete, conMenu_Edit_Delete     '清除
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll          '全选
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll       '全清
    End With
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
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

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "导航"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, objPane)
    objPane.Title = "记录"
    objPane.Options = PaneNoCaption
        
    Set objPane = dkpMain.CreatePane(3, 800, 100, DockBottomOf, objPane)
    objPane.Title = "内容"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Sub InitTaskPanel()
    
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    
    With tpl
        .SetIconSize 24, 24
        Call .Icons.AddIcons(ImageManager1.Icons)
        .VisualTheme = xtpTaskPanelThemeNativeWinXP
        .Behaviour = xtpTaskPanelBehaviourToolbox
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        
        .SetMargins 5, 5, 5, 5, 5
        .SetItemInnerMargins 0, 5, 0, 5
        .SelectItemOnFocus = True
                        
        Set objGroup = .Groups.Add(0, "类型")
        objGroup.Expandable = False
        objGroup.CaptionVisible = False
        
        Set objItem = objGroup.Items.Add(1, "接收消息", xtpTaskItemTypeLink, 3)
        objItem.Tag = "G1"
        objItem.Tooltip = "当前工作站所接收到的消息"
        If mstrCurrentGroup = objItem.Tag Then objItem.Selected = True
                
        Set objItem = objGroup.Items.Add(2, "未读消息", xtpTaskItemTypeLink, 2)
        objItem.Tag = "G2"
        objItem.Tooltip = "当前工作站所接收到的未读消息"
        If mstrCurrentGroup = objItem.Tag Then objItem.Selected = True
        
        .Reposition
    
    End With
    
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As Object
    Dim blnMuliSelect As Boolean
    
    Select Case Control.id
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
    
        Call ExecuteCommand(Command.删除消息)
                        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        
        blnMuliSelect = False
        With vsf(0)
            For lngLoop = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(lngLoop, .ColIndex("选择")))) = 1 Then
                    blnMuliSelect = True
                    Call ExecuteCommand(Command.阅读消息, 1)
                    Exit For
                End If
            Next
        End With
        If blnMuliSelect = False Then Call ExecuteCommand(Command.阅读消息, 0)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新
                
        Call ExecuteCommand(Command.刷新消息)
        Call ExecuteCommand(Command.刷新内容)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size      '大图标
    
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar
        cbsMain.StatusBar.Visible = Not cbsMain.StatusBar.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Close
    
        Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsf(0)
        Select Case Control.id
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll
            
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button            '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = cbsMain(2).Visible
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Text              '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Size              '大图标
            Control.Checked = cbsMain.Options.LargeIcons
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_StatusBar                 '状态栏
            Control.Checked = cbsMain.StatusBar.Visible
        
        End Select
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Item.Handle = picPane(2).hWnd
    Case 3
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1005
    
    Call ExecuteCommand(Command.初始控件)
    Call ExecuteCommand(Command.读注册表)

'    If Not (gcnOracle Is Nothing) Then Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 100, Me.ScaleHeight)
    Call zlCommFun.SetPaneRange(dkpMain, 3, 15, 100, Me.ScaleWidth, 200)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
            
End Sub

Private Sub lblLinkTitle_Click()
    
    RaiseEvent OpenLink(Val(lblLinkType.Tag), lblLinkTitle.Tag)
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        picBack(1).Move 15, 15, picPane(Index).Width - 30
        txtText.Move 15, picBack(1).Top + picBack(1).Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (picBack(1).Top + picBack(1).Height + 15) - 15
    Case 1
        tpl.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 2
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub tpl_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    mstrCurrentGroup = Item.Tag
    Call ExecuteCommand(Command.刷新消息)
    Call ExecuteCommand(Command.刷新内容)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        txtLocation.Tag = ""
        
        Dim obj As CommandBarControl
        
        Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
        If Not (obj Is Nothing) Then
            If obj.Enabled = True Then Call cbsMain_Execute(obj)
        End If
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
    If OldRow <> NewRow Then
        '如果为未阅读，处理为已阅读
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
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "全部勾选(&A)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "全部不选(&U)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "清除消息(&D)")
        cbrPopupItem.BeginGroup = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "重新刷新(&R)")
        cbrPopupItem.BeginGroup = True
                
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
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



