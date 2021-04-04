VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmChildQuestion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbo抽查次数 
      Height          =   300
      Left            =   5145
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   435
      Width           =   2625
   End
   Begin VB.Timer tmr 
      Interval        =   60000
      Left            =   2490
      Top             =   5430
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   3990
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   450
      Width           =   1140
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   3
      Left            =   810
      ScaleHeight     =   1935
      ScaleWidth      =   2340
      TabIndex        =   6
      Top             =   2880
      Width           =   2340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   150
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Height          =   1935
      Index           =   2
      Left            =   255
      ScaleHeight     =   1935
      ScaleWidth      =   2340
      TabIndex        =   4
      Top             =   1635
      Width           =   2340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   150
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Height          =   3585
      Index           =   1
      Left            =   4005
      ScaleHeight     =   3585
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   1155
      Width           =   3135
      Begin XtremeSuiteControls.TabControl tbcQuestion 
         Height          =   1830
         Left            =   255
         TabIndex        =   3
         Top             =   450
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   0
      Left            =   150
      ScaleHeight     =   1935
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   510
      Width           =   2340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   150
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Left            =   -30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmChildQuestion.frx":0000
      Left            =   810
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmMain            As Object
Private mlngKey             As Long
Private mstr文件id          As String
Private mlng医嘱id          As Long
Private mlng科室ID          As Long
Private mlngReferKey        As Long
Private mblnReading         As Boolean
Private mstrSQL             As String
Private mblnDataChanged     As Boolean
Private mblnAllowModify     As Boolean
Private mlngMoudal          As Long
Private mclsVsf(2)          As New clsVsf
Private mlng提交Id          As Long
Private mlng病人ID          As Long
Private mlng主页ID          As Long
Private mstrObject          As String
Private mlngTmp             As Long
Private mintIndex           As Integer
Private mintPreTime         As String
Private mstrStart           As String
Private mstrEnd             As String
Private mblnCurrentPatient  As Boolean
Private mlngIntenal         As Long
Private mlngLoop            As Long
Private mstrDepts           As String
Private mrsCondition        As ADODB.Recordset
Private mlngCurNum          As Long '当前次数
Private mstr日期选择        As String
Private mstr抽查开始时间    As String
Private mstr抽查结束时间    As String
Private mstr反馈人          As String
Private mblnRef             As Boolean
Private mblnAuditEnter  As Boolean              '是否允许录入审查意见
Private mstrPrivs       As String
Private mblnDataExecute As Boolean

Private WithEvents mfrmChildQuestionEdit As frmChildQuestionEdit
Attribute mfrmChildQuestionEdit.VB_VarHelpID = -1

Public Event AfterSaveQuestion(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
Public Event AfterDeleteQuestion(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
Public Event AfterDataChanged()
Public Event LocationDocument(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt反馈对象 As Byte, ByVal lng文件id As Long, ByVal lng医嘱id As Long, ByVal lng科室ID As Long)
Public Event AfterQuestionType(ByVal blnQuestionType As Boolean)

'######################################################################################################################
Public Property Get 提交Id() As Long
    提交Id = mlng提交Id
End Property

Public Property Get Depts() As String
    Depts = mstrDepts
End Property

Public Property Let Depts(ByVal vDepts As String)
    mstrDepts = vDepts
End Property

Public Property Get CurrentPatient() As Boolean
    CurrentPatient = mblnCurrentPatient
End Property

Public Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildQuestionEdit.DataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    If Not (mfrmChildQuestionEdit Is Nothing) Then
        DataChanged = mfrmChildQuestionEdit.DataChanged
    End If
End Property

Public Property Let AllowModify(ByVal blnData As Boolean)
    mblnAllowModify = blnData
    
    mfrmChildQuestionEdit.AllowModify = blnData
    
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Function InitData(ByVal frmMain As Object, ByVal lngMoudal As Long, ByVal blnAllowModify As Boolean, ByVal blnAuditEnter As Boolean, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAuditEnter = blnAuditEnter
    Set mfrmMain = frmMain
    mblnAllowModify = blnAllowModify
    mlngMoudal = lngMoudal
    mstrPrivs = strPrivs
    
    mstr日期选择 = "前一月"
    mstr抽查开始时间 = GetDateTime("前一月", 1)
    mstr抽查结束时间 = GetDateTime("前一月", 2)
    
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
    Call ExecuteCommand("读注册表")
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("刷新次数")
    mintPreTime = cbo.Text
    If cbo.Text <> "[指定...]" Then
        mstrStart = GetDateTime(mintPreTime, 1)
        mstrEnd = GetDateTime(mintPreTime, 2)
    End If
    Call ExecuteCommand("读取完成反馈")
    
    DataChanged = False
    
End Function

Public Function SetParamter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strObject As String, ByVal strParam As String, Optional ByVal lng提交Id As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim varParam As Variant
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng提交Id = lng提交Id
    mstrObject = strObject
    
    Select Case mstrObject
    Case "首页记录", "住院医嘱"
                
        mstr文件id = 0
        mlng医嘱id = 0
        mlng科室ID = 0
            
    Case "住院病历", "护理病历", "知情文件", "疾病证明"
                
        If strParam <> "" Then
            varParam = Split(strParam, ";")
            mstr文件id = varParam(0)
            mlng医嘱id = 0
            mlng科室ID = 0
        End If
    
    Case "医嘱报告"
        'strParam：报告id;医嘱id
        If strParam <> "" Then
            varParam = Split(strParam, ";")
            If UBound(varParam) >= 1 Then
                mstr文件id = varParam(0)
                mlng医嘱id = Val(varParam(1))
                mlng科室ID = 0
            End If
        End If
    
    Case "护理记录"
        
        'strParam：科室id;保留;开始~截止;文件id
        
        If strParam <> "" Then
            varParam = Split(strParam, ";")
            If UBound(varParam) >= 1 Then
                mstr文件id = Val(varParam(3))
                mlng医嘱id = 0
                mlng科室ID = Val(varParam(0))
            End If
            
        End If
    Case Else
        mstr文件id = 0
        mlng医嘱id = 0
        mlng科室ID = 0
    End Select
    
    SetParamter = True
    
End Function

Public Function RefreshData(strDepts As String, rsCondition As ADODB.Recordset, ByVal blnAuditEnter As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAuditEnter = blnAuditEnter
    mstrDepts = strDepts
    Set mrsCondition = rsCondition
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("控件状态")
    
    If ExecuteCommand("刷新数据") = False Then Exit Function
    
    DataChanged = False
    
    RefreshData = True
    
End Function

'######################################################################################################################
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
    Dim objCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

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
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Manage_ThingAdd, "增加次数", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "选择次数", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加反馈", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_CopyNewItem, "再次反馈", , conMenu_Edit_NewItem, xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除反馈", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "结束反馈", , , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SendBack, "回退结束", , , xtpButtonIcon)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "保存更改", True, , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消更改", , , xtpButtonIcon)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "当前病人", True, , xtpButtonIcon)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "刷新反馈", , , xtpButtonIcon)
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_View_Find, "时间", , 1, xtpButtonCaption)
    objControl.Flags = xtpFlagRightAlign
    Set objCustom = NewToolBar(objBar, xtpControlCustom, conMenu_View_Find, "", , , xtpButtonCaption)
    objCustom.Handle = cbo.hWnd
    objCustom.Flags = xtpFlagRightAlign
    
    
    Set objBar = cbsMain.Add("抽查", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, conMenu_View_FindType, "抽查次数", , 1, xtpButtonCaption)
    objControl.Flags = xtpFlagAlignLeft
    
    Set objCustom = NewToolBar(objBar, xtpControlCustom, conMenu_View_FindType, "", , , xtpButtonCaption)
    objCustom.Handle = cbo抽查次数.hWnd
    objCustom.Flags = xtpFlagAlignLeft
    
    
    
    
End Function

Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs              As New ADODB.Recordset
    Dim rsSQL           As New ADODB.Recordset
    Dim blnAllowModify  As Boolean
    Dim intRow          As Integer
    Dim strTmp          As String
    Dim strDept         As String
    Dim i               As Integer
    Dim mlngTmpCurNum   As Long
    On Error GoTo errHand
    
    mblnReading = True
    Call SQLRecord(rsSQL)
    
    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf(0) = New clsVsf
        With mclsVsf(0)
            Call .Initialize(Me.Controls, vsf(0), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[图标]", False)
            Call .AppendColumn("反馈意见", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈对象", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("姓名", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("提交id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("相关id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("文件id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("医嘱id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("科室id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("反馈对象id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("科室", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf(1) = New clsVsf
        With mclsVsf(1)
            Call .Initialize(Me.Controls, vsf(1), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("提交id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("相关id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("文件id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("医嘱id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("科室id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[图标]", False)
            Call .AppendColumn("姓名", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈意见", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈对象", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈对象id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("科室", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsf(2) = New clsVsf
        With mclsVsf(2)
            Call .Initialize(Me.Controls, vsf(2), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("提交id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("相关id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("文件id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("医嘱id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("科室id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[图标]", False)
            Call .AppendColumn("姓名", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈意见", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈对象", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("反馈对象id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("科室", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        Call InitCommandBar
            
        '划分停靠区域
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "问题": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, Nothing): objPane.Title = "详细": objPane.Options = PaneNoCaption

        dkpMain.SetCommandBars cbsMain
        Call DockPannelInit(dkpMain)
        
        Call TabControlInit(tbcQuestion)
        With tbcQuestion
            .PaintManager.BoldSelected = True
                           
            .InsertItem 0, "未改", picPane(0).hWnd, 6
            .InsertItem 1, "未审", picPane(2).hWnd, 7
            .InsertItem 2, "结束", picPane(3).hWnd, 8
            .Item(0).Selected = True
        End With
        
        With cbo
            .AddItem "今  天"
            .AddItem "昨  天"
            .AddItem "本  周"
            .AddItem "本  月"
            .AddItem "本  季"
            .AddItem "本半年"
            .AddItem "本  年"
            .AddItem "前三天"
            .AddItem "前一周"
            .AddItem "前半月"
            .AddItem "前一月"
            .AddItem "前二月"
            .AddItem "前三月"
            .AddItem "前半年"
            .AddItem "前一年"
            .AddItem "前二年"
            .AddItem "[指定...]"
            .ListIndex = 0
        End With
        
        mintPreTime = "今  天"
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        mlngIntenal = Val(GetPara("未复查刷新频率", mfrmMain.模块号, "5"))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        If tbcQuestion.Enabled <> Not DataChanged Then
            
            tbcQuestion.Enabled = Not DataChanged
            vsf(0).Enabled = Not DataChanged
            vsf(1).Enabled = Not DataChanged
            vsf(2).Enabled = Not DataChanged
            
            vsf(0).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
            vsf(1).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
            vsf(2).ForeColor = IIf(DataChanged, COLOR.深灰色, COLOR.黑色)
            
            RaiseEvent AfterDataChanged

        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
    
        With vsf(0)
            If tbcQuestion.ItemCount > 0 Then
                If Val(.RowData(.Row)) > 0 Then
                    tbcQuestion.Item(0).Caption = "未改(" & .Rows - 1 & ")"
                Else
                    tbcQuestion.Item(0).Caption = "未改"
                End If
            End If
        End With
        
        With vsf(1)
            If tbcQuestion.ItemCount > 0 Then
                If Val(.RowData(.Row)) > 1 Then
                    tbcQuestion.Item(1).Caption = "未审(" & .Rows - 1 & ")"
                Else
                    tbcQuestion.Item(1).Caption = "未审"
                End If
            End If
        End With

        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        Call ExecuteCommand("读取未改反馈")
        Call ExecuteCommand("读取未审反馈")
        Call ExecuteCommand("读取完成反馈")
        Call ExecuteCommand("读取反馈内容")
        Call ExecuteCommand("刷新状态")
        
        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新指定反馈"
                
        Set rs = gclsPackage.GetQuestion(mrsCondition, "", 1, mlngTmp)
        If rs.BOF = True Then Exit Function

        intRow = mclsVsf(0).FindRow(mlngTmp, -1)
        With vsf(0)
            If intRow > 0 Then
                '已加载
                .Row = intRow
            Else
                '未加载
                If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            Call mclsVsf(0).LoadGridRow(.Row, rs)
        End With

        Call ExecuteCommand("读取反馈内容")
        Call ExecuteCommand("刷新状态")
                
    '------------------------------------------------------------------------------------------------------------------
    Case "增加反馈记录"
        
        '检查是否已经启动了方案
        If gclsPackage.GetExamineStartUse = False Then
            Call MsgBox("必须在电子病案审查项目中启动一个审查方案,才能添加病案反馈!", vbQuestion + vbDefaultButton2, ParamInfo.系统名称)
            GoTo endHand
        End If
        
        With vsf(0)
            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
            .Row = .Rows - 1
            .ShowCell .Row, IIf(.Col = -1, 1, .Col)
        End With

        Call ExecuteCommand("读取反馈内容")
        
        mlngCurNum = GetMaxNumNEW(mlng病人ID, mlng主页ID, 0)
        Call mfrmChildQuestionEdit.SetCurNum(mlngCurNum)
        
        
        Call mfrmChildQuestionEdit.NewData(mstrObject, mstr文件id, mlng医嘱id, mlng科室ID, mlngReferKey, mlngCurNum)
        mclsVsf(mintIndex).AppendRows = True
        
        GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
    Case "复制反馈记录"
        
        With vsf(mintIndex)
            '检查并获取相关的反馈记录
            Set rs = gclsPackage.GetRelevanceID(Val(.RowData(.Row)))
            If Not rs.EOF Then
                If NVL(rs!相关ID) = "" Then
                    mlngReferKey = -1
                Else
                    mlngReferKey = NVL(rs!相关ID, 0)
                End If
            Else
                mlngReferKey = Val(.RowData(.Row))
            End If
            Call vsf_DblClick(mintIndex)
            If mlngReferKey > 0 Or mlngReferKey = -1 Then tbcQuestion.Item(0).Selected = True
        End With

        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "删除反馈记录"
        
        With vsf(0)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("您是否真的要删除当前反馈问题吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                gstrSQL = "zl_病案反馈记录_Delete(" & Val(.RowData(.Row)) & ")"
                Call SQLRecordAdd(rsSQL, gstrSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
                Call ExecuteCommand("刷新次数")
                Call SetCob(mlngCurNum)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "完成反馈记录"
        
        With vsf(mintIndex)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("您是否真的要结束当前反馈问题吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                gstrSQL = "zl_病案反馈记录_Finish(" & Val(.RowData(.Row)) & ",'" & gstrUserName & "')"
                Call SQLRecordAdd(rsSQL, gstrSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "回退完成反馈"
        
        With vsf(mintIndex)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("您是否真的要回退当前已结束的反馈问题吗？", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.系统名称) = vbYes Then
                gstrSQL = "zl_病案反馈记录_RollBackFinish(" & Val(.RowData(.Row)) & ")"
                Call SQLRecordAdd(rsSQL, gstrSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "移除反馈记录"
        
        With vsf(mintIndex)
            If .Rows > 2 Then
                .RemoveItem .Row
                mclsVsf(mintIndex).AppendRows = True
            Else
                Call mclsVsf(mintIndex).ClearGrid
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取未改反馈"
        
        mlngTmpCurNum = mlngCurNum
        mclsVsf(0).ClearGrid
        mlngCurNum = mlngTmpCurNum
        
        If mrsCondition Is Nothing Then GoTo endHand
        
        If mblnCurrentPatient Then
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 2, , , , mlng病人ID, mlng主页ID, mlngCurNum, mstr抽查开始时间, mstr抽查结束时间, mstr反馈人)
        Else
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 2, , , , , , mlngCurNum, mstr抽查开始时间, mstr抽查结束时间, mstr反馈人)
        End If
        
        If rs.BOF = False Then
            Call mclsVsf(0).LoadGrid(rs)
        End If
                
    '------------------------------------------------------------------------------------------------------------------
    Case "读取未审反馈"
        
        mclsVsf(1).ClearGrid
        
        If mrsCondition Is Nothing Then GoTo endHand
        
        If mblnCurrentPatient Then
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 3, , , , mlng病人ID, mlng主页ID, mlngCurNum, mstr抽查开始时间, mstr抽查结束时间, mstr反馈人)
        Else
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 3, , , , , , mlngCurNum, mstr抽查开始时间, mstr抽查结束时间, mstr反馈人)
        End If
        If rs.BOF = False Then
            Call mclsVsf(1).LoadGrid(rs)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取完成反馈"
        
        mclsVsf(2).ClearGrid
        '
        If mrsCondition Is Nothing Then GoTo endHand
        
        If mstrStart = "" Then
            mstrStart = GetDateTime("今  天", 1)
            mstrEnd = GetDateTime("今  天", 2)
        End If
        
        If mblnCurrentPatient Then
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 4, , mstrStart, mstrEnd, mlng病人ID, mlng主页ID, mlngCurNum, mstr抽查开始时间, mstr抽查结束时间, mstr反馈人)
        Else
            Set rs = gclsPackage.GetQuestion(mrsCondition, mstrDepts, 4, , mstrStart, mstrEnd, , , mlngCurNum, mstr抽查开始时间, mstr抽查结束时间, mstr反馈人)
        End If
        
        If rs.BOF = False Then
            Call mclsVsf(2).LoadGrid(rs)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取反馈内容"
        
        With vsf(mintIndex)
            Call mfrmChildQuestionEdit.RefreshData(Val(.RowData(.Row)), mblnAuditEnter)
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新次数"
        
        If mstr日期选择 = "自定义" Or mstr日期选择 = "所有" Then
        
        Else
            mstr抽查开始时间 = GetDateTime(mstr日期选择, 1)
            mstr抽查结束时间 = GetDateTime(mstr日期选择, 2)
        End If
        
        Call Init抽查次数(mstr抽查开始时间, mstr抽查结束时间)
        Call SetCob(mlngCurNum)
'        call ExecuteCommand("刷新数据")
          
    '------------------------------------------------------------------------------------------------------------------
    Case "恢复数据"
            
        If mfrmChildQuestionEdit.DataChanged Then
            With vsf(0)
                If Val(.RowData(.Row)) = 0 And .Rows > 2 Then
                    .Rows = .Rows - 1
                    .Row = .Rows - 1
                End If
            End With
            Call ExecuteCommand("读取反馈内容")
            mfrmChildQuestionEdit.DataChanged = False
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "校验数据"
    
        '1.
        '--------------------------------------------------
        If mfrmChildQuestionEdit.DataChanged Then
            If mfrmChildQuestionEdit.ValidData = False Then GoTo endHand
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
                    
        If mfrmChildQuestionEdit.DataChanged Then
        
            With vsf(0)
                mlngTmp = Val(.RowData(.Row))
                If mlngTmp > 0 Then
                    
                    '修改保存,病人id,主页id,提交id用当前记录的值
                    If mfrmChildQuestionEdit.SaveData(rsSQL, mlngTmp, Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), Val(.TextMatrix(.Row, .ColIndex("提交id"))), mlngCurNum) = False Then GoTo endHand
                
                Else
                
                    '新增保存时,病人id,主页id,提交id用当前外界传入的值
                    If mfrmChildQuestionEdit.SaveData(rsSQL, mlngTmp, mlng病人ID, mlng主页ID, mlng提交Id, mlngCurNum) = False Then GoTo endHand
                    
                End If
            End With
        
            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            Call ExecuteCommand("刷新次数")
            Call SetCob(mlngCurNum)
        End If

        GoTo endHand
    '------------------------------------------------------------------------------------------------------------------
    Case "读注册表"
        
        On Error Resume Next
        
        strTmp = GetPara("完成问题范围", mfrmMain.模块号, "今  天")
        If Left(strTmp, 7) = "[指定...]" Then
            cbo.Text = "[指定...]"
            mstrStart = Split(strTmp, ";")(1)
            mstrEnd = Split(strTmp, ";")(2)
        Else
            cbo.Text = strTmp
        End If
        
        mblnCurrentPatient = (Val(zlDatabase.GetPara("当前病人", glngSys, mlngMoudal, "0")) = 1)
        
        On Error GoTo errHand

        If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
            '使用个性化设置
            mclsVsf(0).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_0_20081113", ""))
            mclsVsf(1).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_1_20081113", ""))
            mclsVsf(2).LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_2_20081113", ""))
        End If
        
        mlngCurNum = GetRegister(私有模块, Me.Name, "当前次数", 1)
        mstr日期选择 = GetRegister(私有模块, Me.Name, "日期选择", "前一月")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "写注册表"
        
        If cbo.Text = "[指定...]" Then
            Call SetPara("完成问题范围", cbo.Text & ";" & mstrStart & ";" & mstrEnd, mfrmMain.模块号)
        Else
            Call SetPara("完成问题范围", cbo.Text, mfrmMain.模块号)
        End If
        Call SetPara("当前病人", IIf(mblnCurrentPatient, 1, 0), mfrmMain.模块号)
        
        Call SetRegister(私有模块, Me.Name, "表格参数_0_20081113", mclsVsf(0).SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_1_20081113", mclsVsf(1).SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "表格参数_2_20081113", mclsVsf(2).SaveStateToString)
        Call SetRegister(私有模块, Me.Name, "当前次数", mlngCurNum)
        Call SetRegister(私有模块, Me.Name, "日期选择", mstr日期选择)
    End Select

    ExecuteCommand = True
    
    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function


Private Sub cbo_Click()
    Dim strTmp As String
    Dim dStart As Date
    Dim dEnd As Date
    
    If mblnReading Then Exit Sub
    
    If cbo.Text = "[指定...]" Then
        
        If mstrStart = "" Then
            mstrStart = GetDateTime("今  天", 1)
            mstrEnd = GetDateTime("今  天", 2)
        End If
        
        dStart = CDate(mstrStart)
        dEnd = CDate(mstrEnd)
        If Not frmQuestionTime.ShowMe(Me, dStart, dEnd) Then
            '取消时恢复原来的选择
            Call zlControl.CboLocate(cbo, mintPreTime)
            vsf(mintIndex).SetFocus
            Exit Sub
        Else
            mstrStart = Format(dStart, "yyyy-MM-dd HH:mm:ss")
            mstrEnd = Format(dEnd, "yyyy-MM-dd HH:mm:ss")
            vsf(mintIndex).SetFocus
            mintPreTime = cbo.Text
        End If
        
    Else
        mintPreTime = cbo.Text
        mstrStart = GetDateTime(mintPreTime, 1)
        mstrEnd = GetDateTime(mintPreTime, 2)
    End If

    Call ExecuteCommand("读取完成反馈")
    
End Sub

Private Sub cbo抽查次数_Click()
    If mblnRef Then Exit Sub
    mlngCurNum = CLng(cbo抽查次数.ItemData(cbo抽查次数.ListIndex))
    
    If mlngCurNum > 0 Then
        '重新分析时间
         If cbo抽查次数.Text <> "" Then
            mstr抽查开始时间 = GetAnalyseTime(cbo抽查次数.Text, 1)
            mstr抽查结束时间 = GetAnalyseTime(cbo抽查次数.Text, 2)
         End If
    Else
        If mstr日期选择 = "自定义" Or mstr日期选择 = "所有" Then
            mstr抽查开始时间 = Format("2000-01-01 00:00:00", "yyyy-MM-dd HH:mm:SS")
            mstr抽查结束时间 = Format("3000-01-01 23:59:59", "yyyy-MM-dd HH:mm:SS")

        Else
            mstr抽查开始时间 = GetDateTime(mstr日期选择, 1)
            mstr抽查结束时间 = GetDateTime(mstr日期选择, 2)
        End If
    End If
    
    
    Call ExecuteCommand("读取未改反馈")
    Call ExecuteCommand("读取反馈内容")
    Call ExecuteCommand("刷新状态")
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_ThingAdd        '增加次数
    '获取当日 最大次数+1
    '刷新列表
        mlngCurNum = GetMaxNum(mstr抽查开始时间, mstr抽查结束时间)
        Call ExecuteCommand("刷新数据")
        Call mfrmChildQuestionEdit.SetCurNum(mlngCurNum)
        
        mlngReferKey = 0
        Call ExecuteCommand("增加反馈记录")

        DataChanged = False
    Case conMenu_File_Preview              '过滤条件
    '--------------------------------------------------------------------------------------------------------------
    '显示条件过滤
    '根据条件刷新列表
        Dim blnFilter As Boolean
        blnFilter = frmChildQuestionFilter.ShowPara(Me, mstr抽查开始时间, mstr抽查结束时间, mstr日期选择, mlngCurNum, mstr反馈人)
        If blnFilter Then
            Call Init抽查次数(mstr抽查开始时间, mstr抽查结束时间)
            Call SetCob(mlngCurNum)
            If ExecuteCommand("刷新数据") = False Then Exit Sub
            DataChanged = False
        End If
    Case conMenu_Edit_NewItem               '增加反馈记录
    '--------------------------------------------------------------------------------------------------------------
        mlngReferKey = 0
        Call ExecuteCommand("增加反馈记录")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_CopyNewItem               '复制反馈记录
        
        mlngReferKey = 0
        Call ExecuteCommand("复制反馈记录")
        If mlngReferKey > 0 Or mlngReferKey = -1 Then Call ExecuteCommand("增加反馈记录")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                '删除反馈记录

        If ExecuteCommand("删除反馈记录") Then
            Call ExecuteCommand("移除反馈记录")
            Call ExecuteCommand("刷新状态")
'            RaiseEvent AfterDeleteQuestion(mlng病人ID, mlng主页ID)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Send            '完成反馈记录
        
        If ExecuteCommand("完成反馈记录") Then
            Call ExecuteCommand("移除反馈记录")
            Call ExecuteCommand("读取完成反馈")
            Call ExecuteCommand("刷新状态")
            RaiseEvent AfterDeleteQuestion(mlng病人ID, mlng主页ID)
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SendBack
                
        If ExecuteCommand("回退完成反馈") Then
            Call ExecuteCommand("移除反馈记录")
            Call ExecuteCommand("读取未改反馈")
            Call ExecuteCommand("读取未审反馈")
            Call ExecuteCommand("刷新状态")
            
            RaiseEvent AfterDeleteQuestion(mlng病人ID, mlng主页ID)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save           '保存数据
    
        If ExecuteCommand("校验数据") And DataChanged Then
            If ExecuteCommand("保存数据") Then
                
                DataChanged = False
                
                Call ExecuteCommand("刷新指定反馈")
                Call ExecuteCommand("刷新状态")
                
'                RaiseEvent AfterSaveQuestion(mlng病人ID, mlng主页ID)
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle         '恢复数据
    
        Call ExecuteCommand("恢复数据")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Find                '过滤数据
        
        If ExecuteCommand("过滤数据") Then
            Call ExecuteCommand("刷新数据")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter                  '当前病人
        
        mblnCurrentPatient = Not mblnCurrentPatient
        mlngCurNum = CLng(cbo抽查次数.ItemData(cbo抽查次数.ListIndex))
        Call ExecuteCommand("刷新数据")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               '刷新数据
        Call ExecuteCommand("刷新数据")
    End Select
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand
    
    With vsf(mintIndex)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Find
            Control.Visible = (mintIndex = 2)
            Control.Enabled = (Control.Visible And DataChanged = False)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_ThingAdd
            Control.Visible = False ' (mintIndex = 0 And AllowModify)
'            Control.Enabled = (Control.Visible And DataChanged = False And AllowModify And mlng病人ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview
            Control.Visible = (mintIndex = 0 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And AllowModify And mlng病人ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem
            Control.Visible = (mintIndex = 0 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And AllowModify And mlng病人ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_CopyNewItem
            Control.Visible = (mintIndex = 1 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify And mlng病人ID > 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            Control.Visible = (mintIndex = 0 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Send
            Control.Visible = (mintIndex <> 2 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SendBack
            Control.Visible = (mintIndex = 2 And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
            Control.Visible = ((mintIndex = 0 Or mintIndex = 1) And AllowModify)
            Control.Enabled = (Control.Visible And DataChanged = True And AllowModify)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Filter
            Control.Visible = (IsPrivs(mstrPrivs, "院级反馈") And IsPrivs(mstrPrivs, "科级反馈")) Or (IsPrivs(mstrPrivs, "院级反馈") And IsPrivs(mstrPrivs, "科级反馈") = False)
        
            Control.Checked = Control.Visible And mblnCurrentPatient
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_FindType '抽查次数
            Control.Enabled = (mintIndex = 0 And AllowModify)
'            If (mintIndex = 0 And AllowModify) Then
''                Me.cbsMain.ActiveMenuBar.Visible = False
'            Else
'
'            End If
'            Control.Checked = mblnCurrentPatient
        End Select
    End With
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Set mfrmChildQuestionEdit = New frmChildQuestionEdit
        Call mfrmChildQuestionEdit.InitData(mfrmMain, AllowModify, mstrPrivs)
        Item.Handle = mfrmChildQuestionEdit.hWnd
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 2, 100, 325, Me.ScaleWidth, 325)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExecuteCommand("写注册表")
    Unload mfrmChildQuestionEdit
End Sub

Private Sub mfrmChildQuestionEdit_AfterDataChanged()
    Call ExecuteCommand("控件状态")
End Sub

Private Sub mfrmChildQuestionEdit_AfterQuestionType(ByVal blnQuestionType As Boolean)
    'blnQuestionType=True 院级反馈 =Flase 科级反馈
    Dim lngCurNum As Long
    If blnQuestionType Then
        lngCurNum = GetMaxNumNEW(mlng病人ID, mlng主页ID, 0)
        Call mfrmChildQuestionEdit.SetCurNum(lngCurNum)
    Else
        lngCurNum = GetMaxNumNEW(mlng病人ID, mlng主页ID, 1)
        Call mfrmChildQuestionEdit.SetCurNum(lngCurNum)
    End If
    RaiseEvent AfterQuestionType(blnQuestionType)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(0).AppendRows = True
    Case 1
        tbcQuestion.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 2
        vsf(1).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(1).AppendRows = True
    Case 3
        vsf(2).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(2).AppendRows = True
    End Select
End Sub

Private Sub tbcQuestion_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mintIndex = Item.Index
    Call ExecuteCommand("读取反馈内容")
End Sub

Private Sub tmr_Timer()
        
    mlngLoop = mlngLoop + 1
    If mlngLoop = mlngIntenal And mlngIntenal > 0 Then
    
        '自动刷新未复查
        
        Call ExecuteCommand("读取未审反馈")
        
        If tbcQuestion.Item(1).Selected Then Call ExecuteCommand("读取反馈内容")
        mlngIntenal = Val(GetPara("未复查刷新频率", mfrmMain.模块号, "5"))
        
    End If
    
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Call ExecuteCommand("读取反馈内容")
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey)
    vsf(Index).ShowCell vsf(Index).Row, vsf(Index).Col
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf(Index).SaveKey = Val(vsf(Index).RowData(vsf(Index).Row))
End Sub

Private Sub vsf_DblClick(Index As Integer)
    With vsf(Index)
        RaiseEvent LocationDocument(Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), Val(.TextMatrix(.Row, .ColIndex("反馈对象id"))), Val(.TextMatrix(.Row, .ColIndex("文件id"))), Val(.TextMatrix(.Row, .ColIndex("医嘱id"))), Val(.TextMatrix(.Row, .ColIndex("科室id"))))
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick(Index)
    End If
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar     As CommandBar
    Dim cbrPopupItem    As CommandBarControl
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)

        If Not mclsVsf(Index).MoveColumn Then
            
            '弹出菜单处理
            Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "增加反馈")
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "再次反馈")
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "删除反馈")
            
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "完成反馈"): cbrPopupItem.BeginGroup = True
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SendBack, "回退完成")
            
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存更改"): cbrPopupItem.BeginGroup = True
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消更改")
            
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Filter, "当前病人"): cbrPopupItem.BeginGroup = True
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "刷新反馈")

            cbrPopupBar.ShowPopup
        End If

    End Select
End Sub

'获取当天最大次数
Private Function GetMaxNum(ByVal str抽查开始时间 As String, ByVal str抽查结束时间 As String) As Long
    On Error GoTo errH
    Dim rsData          As ADODB.Recordset
    Dim strSQL          As String
    Dim strStart        As String
    Dim strEnd          As String
    strSQL = "select max(反馈次数) as 次数 from 病案反馈记录 where 反馈时间 BetWeen [1] And [2]"
    strStart = str抽查开始时间 ' GetDateTime("今  天", 1)
    strEnd = str抽查结束时间 ' GetDateTime("今  天", 2)
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "获取当天最大次数", CDate(strStart), CDate(strEnd))

    If rsData.EOF = False Then
        GetMaxNum = IIf(IsNull(rsData!次数), 1, rsData!次数 + 1)
    Else
        GetMaxNum = 2
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Function

'自动获取该病人的当天的评分次数
Private Function GetMaxNumNEW(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng方式 As Long) As Long
'以一个自然日做为一个评分次，对一个自然日内所有的评分项累加合计评分(病人+日期+评分级别区分评分次)
    On Error GoTo errH
    Dim rsData          As ADODB.Recordset
    Dim strSQL          As String
    Dim strNow          As String

    strNow = zlDatabase.Currentdate
    
    strSQL = "Select Sum(A.今日次数) as 今日次数,Sum(A.以往最大次数) as 以往最大次数 From (" & vbNewLine & _
        "select max(反馈次数) as 今日次数,0 AS 以往最大次数 from 病案反馈记录 where 病人ID=[1] And 主页ID=[2] And nvl(评分级别,0)=[3]" & vbNewLine & _
        "And 反馈时间 BetWeen To_Date([4], 'yyyy-mm-dd') And To_Date([4], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
        "Union All" & vbNewLine & _
        "select 0 AS 今日次数,max(反馈次数) AS 以往最大次数 from 病案反馈记录 where 病人ID=[1] And 主页ID=[2] And nvl(评分级别,0)=[3]" & vbNewLine & _
        "And 反馈时间< To_Date([4], 'yyyy-mm-dd')) A"

    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "获取当天最大次数", lng病人ID, lng主页ID, lng方式, Format(strNow, "yyyy-mm-dd"))

    If rsData.EOF = False Then
        If rsData!今日次数 > 0 Then
            GetMaxNumNEW = rsData!今日次数
        Else
            If rsData!以往最大次数 > 0 Then
                GetMaxNumNEW = rsData!以往最大次数 + 1
            Else
                GetMaxNumNEW = 1
            End If
        End If
    Else
        GetMaxNumNEW = 1
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Function


'获取抽查次数汇总信息
Private Sub Init抽查次数(ByVal str抽查开始时间 As String, ByVal str抽查结束时间 As String)
    On Error GoTo errH
        Dim rs As ADODB.Recordset
        Dim lngCount As Long '记录次数
        mblnRef = True
        cbo抽查次数.Clear
        lngCount = 0
        gstrSQL = "select distinct(反馈次数),Sum(A.分值) as 总扣分数,Min(A.反馈时间) as 最早反馈时间 from 病案反馈记录 A where A.反馈时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400 group by A.反馈次数 order by A.反馈次数 Asc"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(str抽查开始时间, "yyyy-mm-dd"), Format(str抽查结束时间, "yyyy-mm-dd"))
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            cbo抽查次数.AddItem "所有"
            
                Do Until rs.EOF
                If NVL(rs!反馈次数, 0) = 0 Then
                    cbo抽查次数.AddItem "第" & NVL(rs!反馈次数, 0) & "次-" & Format(NVL(rs!最早反馈时间, Now()), "YYYY-MM-DD") & "(" & NVL(rs!总扣分数, 0) & ")"
                    cbo抽查次数.ItemData(cbo抽查次数.NewIndex) = NVL(rs!反馈次数, 0)
                End If
                rs.MoveNext
            Loop
            
            rs.MoveFirst
            Do Until rs.EOF
                    If lngCount >= 10 Then Exit Do
'                        Call AddComboData(cbo抽查次数, rs, "最早反馈时间", "次数", , False)
                        If NVL(rs!反馈次数, 0) <> 0 Then
                            cbo抽查次数.AddItem "第" & NVL(rs!反馈次数, 0) & "次-" & Format(NVL(rs!最早反馈时间, Now()), "YYYY-MM-DD") & "(" & NVL(rs!总扣分数, 0) & ")"
                            cbo抽查次数.ItemData(cbo抽查次数.NewIndex) = NVL(rs!反馈次数, 0)
                        End If
                    lngCount = lngCount + 1
                    rs.MoveNext
            Loop
            cbo抽查次数.ListIndex = 0
        Else
            cbo抽查次数.AddItem "所有"
            cbo抽查次数.ListIndex = 0
            cbo抽查次数.ItemData(cbo抽查次数.NewIndex) = 0
        End If
        mblnRef = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnRef = False
    Err.Clear
End Sub

Private Sub SetCob(ByVal lngCurNum As Long)
    Dim i As Integer
    mblnRef = True
    For i = 0 To cbo抽查次数.ListCount - 1
        If cbo抽查次数.ItemData(i) = lngCurNum Then
            cbo抽查次数.ListIndex = i
            mblnRef = False
            Exit Sub
        End If
    Next
    mblnRef = False
End Sub

Private Function GetAnalyseTime(ByVal strTime As String, ByVal lngMode As Long) As String
    Dim strTemp As String
    Dim i As Integer
    '获取时间值
    i = InStrRev(strTime, "次")
    If i > 0 Then
        strTemp = Right(strTime, Len(strTime) - i - 1)
        i = InStrRev(strTemp, "(")
        If i > 0 Then
            strTemp = Left(strTemp, i - 1)
            
            Select Case lngMode
            Case 1
                GetAnalyseTime = Format(strTemp, "yyyy-MM-dd 00:00:00")
            Case 2
                GetAnalyseTime = Format(strTemp, "yyyy-MM-dd 23:59:59")
            End Select
        End If
    End If
    
End Function
