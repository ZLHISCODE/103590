VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInDoctorAdvice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   ControlBox      =   0   'False
   Icon            =   "frmInDoctorAdvice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7155
      TabIndex        =   6
      Top             =   210
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frmInDoctorAdvice.frx":000C
         ToolTipText     =   "选择需要显示的列(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Frame fraAdviceUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   240
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   5040
      Width           =   7275
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4875
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   6240
      _cx             =   11007
      _cy             =   8599
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
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
      FixedCols       =   2
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInDoctorAdvice.frx":055A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   345
         Top             =   645
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":05F5
               Key             =   "紧急"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":080F
               Key             =   "补录"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":0D29
               Key             =   "未申请"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":1243
               Key             =   "已申请"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   975
         Top             =   645
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":175D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":1A57
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":1D51
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":204B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":2345
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSign 
         Left            =   1635
         Top             =   645
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":263F
               Key             =   "签名"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   45
      ScaleHeight     =   600
      ScaleWidth      =   630
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   630
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAppend 
      Height          =   1425
      Left            =   225
      TabIndex        =   2
      Top             =   5505
      Width           =   6270
      _cx             =   11060
      _cy             =   2514
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsColumn 
      Height          =   3495
      Left            =   6885
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInDoctorAdvice.frx":2991
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   End
   Begin MSComctlLib.TabStrip tabAppend 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   529
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱发送明细(&S)"
            Key             =   "医嘱发送明细"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱签名记录(&G)"
            Key             =   "医嘱签名记录"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInDoctorAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mfrmParent As Object
Public mstrPrivs As String
Private WithEvents mfrmEdit As Form
Attribute mfrmEdit.VB_VarHelpID = -1

'上次刷新数据时的病人信息
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng病区ID As Long
Private mlng科室ID As Long
Private mbln出院 As Boolean
Private mlng前提ID As Long
Private mblnShowAll As Boolean

Private mblnMoved As Boolean
Private mvInDate As Date

Private Enum Menu_Advice
    mnu新开医嘱 = 0
    mnu补录医嘱 = 1
    mnu修改医嘱 = 2
    mnu删除医嘱 = 3
    mnu医嘱审核 = 5 '-
    mnu医嘱停止 = 7 '-
    mnu医嘱作废 = 8
    mnu医嘱暂停 = 9 '-
    mnu医嘱启用 = 11
    mnu临嘱发送 = 13 '-
    mnu医嘱回退 = 14
    mnu复制到文本 = 16 '-
End Enum

'报表菜单项索引
Private Enum Menu_Report
    mnu长期医嘱单 = 0
    mnu临时医嘱单 = 1
    mnu医嘱记录本 = 2
End Enum

'固定列
Private Const COL_F标志 = 0
Private Const COL_F申请 = 1
'隐藏列
Private Const COL_ID = 2
Private Const COL_相关ID = COL_ID + 1
Private Const COL_组ID = COL_ID + 2
Private Const COL_组号 = COL_ID + 3
Private Const COL_婴儿ID = COL_ID + 4
Private Const COL_医嘱状态 = COL_ID + 5
Private Const COL_诊疗类别 = COL_ID + 6
Private Const COL_操作类型 = COL_ID + 7
Private Const COL_毒理分类 = COL_ID + 8
Private Const COL_标志 = COL_ID + 9
'可见列
Private Const COL_警示 = COL_ID + 10 'Pass
Private Const COL_期效 = COL_ID + 11
Private Const COL_开始时间 = COL_ID + 12
Private Const COL_医嘱内容 = COL_ID + 13
Private Const COL_皮试 = COL_ID + 14
Private Const COL_总量 = COL_ID + 15
Private Const COL_单量 = COL_ID + 16
Private Const COL_频率 = COL_ID + 17
Private Const COL_用法 = COL_ID + 18
Private Const COL_医生嘱托 = COL_ID + 19
Private Const COL_执行时间 = COL_ID + 20
Private Const COL_终止时间 = COL_ID + 21
Private Const COL_执行科室 = COL_ID + 22
Private Const COL_执行性质 = COL_ID + 23
Private Const COL_上次执行 = COL_ID + 24
Private Const COL_状态 = COL_ID + 25
Private Const COL_开嘱医生 = COL_ID + 26
Private Const COL_开嘱时间 = COL_ID + 27
Private Const COL_校对护士 = COL_ID + 28
Private Const COL_校对时间 = COL_ID + 29
Private Const COL_停嘱医生 = COL_ID + 30
Private Const COL_停嘱时间 = COL_ID + 31
Private Const COL_停嘱护士 = COL_ID + 32
Private Const COL_确认停嘱时间 = COL_ID + 33
'隐藏列
Private Const COL_单据ID = COL_ID + 34 '对应病历文件目录.ID
Private Const COL_申请项 = COL_ID + 35 '诊疗单据是否有申请项
Private Const COL_报告项 = COL_ID + 36 '诊疗单据是否有报告项
Private Const COL_申请ID = COL_ID + 37 '对应病人病历记录.ID
Private Const COL_前提ID = COL_ID + 38
Private Const COL_签名否 = COL_ID + 39

Private Enum COL发送清单
    cs发送号 = 0
    cs发送时间 = 1
    cs发送医嘱 = 2
    cs单据号 = 3
    cs收费项目 = 4
    cs数次 = 5
    cs计费状态 = 6
    cs执行状态 = 7
    cs执行科室 = 8
    cs首次时间 = 9
    cs末次时间 = 10
    cs发送人 = 11
    cs记录性质 = 12
End Enum

Public Function zlRefresh(lng病人ID As Long, lng主页ID As Long, lng病区ID As Long, _
    lng科室ID As Long, Optional bln出院 As Boolean, Optional ByVal lng前提ID As Long = 0, Optional ByVal ifShowAll As Boolean = True) As Boolean
'功能：刷新或清除医嘱清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID
    mlng病区ID = lng病区ID: mlng科室ID = lng科室ID
    mbln出院 = bln出院: mlng前提ID = lng前提ID
    mblnShowAll = ifShowAll
    
    '判断病人是否已转出及入院时间
    '因为该函数内外都在调用,参数不好变,直接读取
    mblnMoved = False
    If lng病人ID <> 0 Then
        strSQL = "Select 入院日期,数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病人ID, lng主页ID)
        On Error GoTo 0
        mblnMoved = Nvl(rsTmp!数据转出, 0) <> 0
        mvInDate = rsTmp!入院日期
    End If
    
    If mlng病人ID = 0 Then
        '清除医嘱清单
        Call ClearAdviceData
        Call ClearAppendData
    Else
        '显示医嘱清单
        Call LoadAdvice
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlButtonClick(objButton As Button) As Boolean
'功能：执行医嘱按钮功能
    Select Case objButton.Key
        Case "新开"
            Call FuncAdviceAdd
        Case "修改"
            Call FuncAdviceModi
        Case "删除"
            Call FuncAdviceDel
        Case "审核"
            Call FuncAdviceAudit
        Case "停止"
            Call FuncAdviceStop
        Case "作废"
            Call FuncAdviceRevoke
        Case "发送"
            Call FuncAdviceSend
        Case "回退"
            '在主窗体中处理了
        Case "签名"
            Call FuncAdviceSign
    End Select
End Function

Public Function zlMenuClick(objMenu As Menu) As Boolean
'功能：执行医嘱菜单功能
    Dim strText As String
    
    If objMenu.Caption Like "*(&*)*" Then
        strText = Split(objMenu.Caption, "(")(0)
    Else
        strText = objMenu.Caption
    End If
        
    If objMenu.Name = "mnuAdviceFuncRoll" Then
        '回退子菜单的处理
        Call FuncAdviceRoll(objMenu.Index)
    ElseIf objMenu.Name = "mnuViewAdviceAppend" Then
        '显示/隐藏附加表格
        objMenu.Checked = Not objMenu.Checked
        fraAdviceUD.Visible = objMenu.Checked
        vsAppend.Visible = objMenu.Checked
        Call Form_Resize
        
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Else
        Select Case strText
            Case "新开医嘱"
                Call FuncAdviceAdd
            Case "补录医嘱"
                Call FuncAdviceSupply
            Case "修改医嘱"
                Call FuncAdviceModi
            Case "删除医嘱"
                Call FuncAdviceDel
            Case "医嘱审核"
                Call FuncAdviceAudit
            Case "医嘱停止"
                Call FuncAdviceStop
            Case "医嘱作废"
                Call FuncAdviceRevoke
            Case "医嘱暂停"
                Call FuncAdvicePause
            Case "医嘱启用"
                Call FuncAdviceResume
            Case "临嘱发送"
                Call FuncAdviceSend
            Case "长期医嘱单"
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1204_1", mfrmParent, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID)
            Case "临时医嘱单"
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1204_2", mfrmParent, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID)
            Case "医嘱记录本"
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1204_3", mfrmParent, "病人科室=" & mlng科室ID)
            Case "复制到文本"
                Call FuncCopyToText
            Case "电子签名"
                Call FuncAdviceSign
            Case "验证签名"
                Call FuncAdviceSignVerify
            Case "取消签名"
                Call FuncAdviceSignErase
        End Select
    End If
    
    zlMenuClick = True
End Function

Private Sub SetFuncEnabled()
'功能：根据当前病人或数据情况，设置功能可用性
    Dim blnAdvice As Boolean, blnEnabled As Boolean
    
    '避免其它地方调用
    On Error Resume Next
    
    With mfrmParent
        '1.无病人的情况:mlng病人ID <> 0
        '2.病人已出院的情况:Not mbln出院
        '3.无数据的情况
        blnAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        
        .mnuAdviceFunc(mnu新开医嘱).Enabled = mlng病人ID <> 0 And Not mbln出院
        .mnuAdviceFunc(mnu补录医嘱).Enabled = mlng病人ID <> 0 And Not mbln出院
        
        '未校对才可以修改
        blnEnabled = mlng病人ID <> 0 And Not mbln出院 And blnAdvice
        If blnEnabled Then
            If InStr(",1,2,", vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 0 Then blnEnabled = False
        End If
        If blnEnabled Then '未签名
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then blnEnabled = False
        End If
        .mnuAdviceFunc(mnu修改医嘱).Enabled = blnEnabled
        
        '未校对才可以删除
        blnEnabled = mlng病人ID <> 0 And blnAdvice
        If blnEnabled Then
            If InStr(",1,2,", vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 0 Then blnEnabled = False
        End If
        If blnEnabled Then '未签名
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then blnEnabled = False
        End If
        .mnuAdviceFunc(mnu删除医嘱).Enabled = blnEnabled
        
        '医嘱审核:与新开一致
        .mnuAdviceFunc(mnu医嘱审核).Enabled = mlng病人ID <> 0 And Not mbln出院
        
        .mnuAdviceFunc(mnu医嘱停止).Enabled = mlng病人ID <> 0
        .mnuAdviceFunc(mnu医嘱作废).Enabled = mlng病人ID <> 0
        .mnuAdviceFunc(mnu医嘱暂停).Enabled = mlng病人ID <> 0 And Not mbln出院
        .mnuAdviceFunc(mnu医嘱启用).Enabled = mlng病人ID <> 0 And Not mbln出院
        
        '出院病人不允许回退操作:预出院病人可以回退出院医嘱发送
        blnEnabled = False
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "Z" _
            And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型))) > 0 Then
            If Val(mfrmParent.lvwPati.SelectedItem.ListSubItems(4).Tag) = 3 Then
                If mfrmParent.mnuAdviceFuncRoll(0).Tag <> "" Then
                    If Val(Split(mfrmParent.mnuAdviceFuncRoll(0).Tag, "|")(0)) = 0 _
                        And Val(Split(mfrmParent.mnuAdviceFuncRoll(0).Tag, "|")(1)) <> 0 Then
                        blnEnabled = True
                    End If
                End If
            End If
        End If
        .mnuAdviceFunc(mnu医嘱回退).Enabled = mlng病人ID <> 0 And blnAdvice And (Not mbln出院 Or blnEnabled)
        .mnuAdviceFunc(mnu临嘱发送).Enabled = mlng病人ID <> 0 And Not mbln出院
        
        .mnuReportItem(mnu长期医嘱单).Enabled = mlng病人ID <> 0
        .mnuReportItem(mnu临时医嘱单).Enabled = mlng病人ID <> 0
        .mnuReportItem(mnu医嘱记录本).Enabled = mlng病人ID <> 0
        .mnuAdviceFunc(mnu复制到文本).Enabled = mlng病人ID <> 0 And blnAdvice
                
        '----------------------------------------------------------------------
        '电子签名部份
        blnEnabled = mlng病人ID <> 0 And blnAdvice And tabAppend.SelectedItem.Index = 2
        If blnEnabled Then
            If vsAppend.RowData(vsAppend.Row) = 0 Then blnEnabled = False
        End If
        .mnuSignVerify.Enabled = blnEnabled
        .mnuSignErase.Enabled = blnEnabled
        .mnuSignNew.Enabled = mlng病人ID <> 0
        .tbrSys.Buttons("签名").Enabled = .mnuSignNew.Enabled
        '----------------------------------------------------------------------
        .tbrSys.Buttons("新开").Enabled = .mnuAdviceFunc(mnu新开医嘱).Enabled
        .tbrSys.Buttons("修改").Enabled = .mnuAdviceFunc(mnu修改医嘱).Enabled
        .tbrSys.Buttons("删除").Enabled = .mnuAdviceFunc(mnu删除医嘱).Enabled
        .tbrSys.Buttons("审核").Enabled = .mnuAdviceFunc(mnu医嘱审核).Enabled
        .tbrSys.Buttons("停止").Enabled = .mnuAdviceFunc(mnu医嘱停止).Enabled
        .tbrSys.Buttons("作废").Enabled = .mnuAdviceFunc(mnu医嘱作废).Enabled
        .tbrSys.Buttons("发送").Enabled = .mnuAdviceFunc(mnu临嘱发送).Enabled
        .tbrSys.Buttons("回退").Enabled = .mnuAdviceFunc(mnu医嘱回退).Enabled
    End With
End Sub

Private Sub FuncAdviceSend()
'功能：临嘱发送
    Dim blnRefresh As Boolean
    
    If frmInAdviceSend.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mlng主页ID, mlng前提ID, blnRefresh) Then
        If blnRefresh Then
            Call mfrmParent.mnuViewRefresh_Click
        Else
            Call LoadAdvice
        End If
    End If
End Sub

Private Sub FuncAdviceRoll(Index As Integer)
'功能：医嘱回退
'参数：Index=回退内容在菜单上的索引
    Dim strSQL As String, strOper As String
    Dim lngFlag As Long, blnBat As Boolean
    Dim int类型 As Integer, lng医嘱ID As Long, lng发送号 As Long
    Dim vOperDate As Date, vOperName As String
    Dim lng签名ID As Long, strSign As String
    
    If Val(mfrmParent.mnuAdviceFunc(mnu医嘱回退).Tag) = 0 Then Exit Sub
    If mfrmParent.mnuAdviceFuncRoll(Index).Tag = "" Then Exit Sub
    
    '(组ID)取一组医嘱中相关ID为空的医嘱ID(给药途径,中药用法,主要手术,检查项目,及独立医嘱)
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)) <> 0 Then
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID))
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If lng医嘱ID = 0 Then Exit Sub
    
    int类型 = Val(Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(0))
    lng发送号 = Val(Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(1))
    vOperDate = CDate(Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(2))
    vOperName = Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(3)
        
    '医生只能回退自已的操作,对电子签名同时也判断了是否回退本人的签名
    If vOperName <> UserInfo.姓名 Then
        MsgBox "你不能回退其他人对医嘱的操作：" & vbCrLf & vbCrLf & mfrmParent.mnuAdviceFuncRoll(Index).Caption & vbTab, vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '根据医生是否作电子签名作检查和提示
    '-------------------------------------------------------
    blnBat = AdviceCanBatchRoll(lng医嘱ID, int类型, lng发送号, vOperDate, lng签名ID) '其它一起操作的医嘱是否有签名
    strOper = Decode(int类型, 0, "发送", 4, "作废", 5, "重整", 6, "暂停", 7, "启用", 8, "停止", 9, "确认停止", 10, "填写皮试结果")
    If MsgBox("确实要回退以下操作吗？" & vbCrLf & vbCrLf & mfrmParent.mnuAdviceFuncRoll(Index).Caption & vbTab & _
        IIF(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 And lng签名ID <> 0, _
            vbCrLf & vbCrLf & "提示：该医嘱" & strOper & "时已签名，将同时回退与它一起" & strOper & "并签名的其它医嘱。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    '批量回退提示
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 And lng签名ID <> 0 Then
        '当前及其它医嘱一起操作并签名,固定一起回退(blnBat=True)
    Else
        If blnBat Then
            If MsgBox("还有其它医嘱和当前医嘱一起被同时" & strOper & "，要同时回退这些医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnBat = False
        End If
    End If
    
    '对医嘱费用的结帐情况进行检查
    If int类型 = 0 And lng发送号 <> 0 Then
        If Not CheckAdviceBalanceRoll(lng发送号, lng医嘱ID, blnBat) Then Exit Sub
    End If
    
    If int类型 = 8 Then '临嘱不会直接回退自动停止
        If blnBat Then
            If MsgBox("要保留这些医嘱的执行终止时间吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lngFlag = 1
            End If
        Else
            If RowIs配方行(vsAdvice.Row) Then
                lngFlag = 1 '中药配方始终保留执行终止时间
            Else
                If MsgBox("要保留医嘱的执行终止时间吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    lngFlag = 1
                End If
            End If
        End If
    End If
    
    '如涉及回退已签名的操作，先取消签名
    '-------------------------------------------------------
    If blnBat Then
        If lng签名ID = 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then
            lng签名ID = GetAdviceSign(lng医嘱ID, int类型, vOperName, vOperDate)
        End If
    Else
        lng签名ID = 0
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then
            lng签名ID = GetAdviceSign(lng医嘱ID, int类型, vOperName, vOperDate)
        End If
    End If
    If lng签名ID <> 0 Then
        strSign = "zl_医嘱签名记录_Delete(" & lng签名ID & ")"
    End If
    
    '检查能否回退签名
    If strSign <> "" Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "系统没有设置电子签名认证中心，回退操作不能继续。", vbInformation, gstrSysName
            Else
                MsgBox "电子签名部件未能正确安装，回退操作不能继续。", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '如果是回退发送且已计费,暂未处理销帐上传问题(1.可能是部分销帐,2.也可不管,预结时自动上传)
    If blnBat Then
        strSQL = "zl_病人医嘱记录_批量回退(" & lng医嘱ID & "," & int类型 & "," & _
            "To_Date('" & Format(vOperDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & lng发送号 & "," & lngFlag & ")"
    Else
        strSQL = "zl_病人医嘱记录_回退(" & lng医嘱ID & "," & lngFlag & ")"
    End If
    
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans
    If strSign <> "" Then
        Call zlDatabase.ExecuteProcedure(strSign, Me.Name)
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "Z" _
        And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型))) > 0 _
        And int类型 = 0 And lng发送号 <> 0 Then
        '回退出院医嘱刷新主界面
        Call mfrmParent.mnuViewRefresh_Click
    Else
        Call LoadAdvice
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function AdviceCanBatchRoll(ByVal lng医嘱ID As Long, ByVal int类型 As Integer, ByVal lng发送号 As Long, ByVal dat时间 As Date, lng签名ID As Long) As Boolean
'功能：检查指定医嘱当前操作是否与其它医嘱一起批量执行的,以判断是否可以批量回退
'参数：lng医嘱ID=相关ID为空的医嘱的ID(一组医嘱的ID)
'      int类型=医嘱操作类型
'      dat时间=医嘱操作的时间
'返回：是否有可以一起回退的其它医嘱
'      lng签名ID=这些要回退的医嘱是否已签名(作废,停止),如有则返回签名ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    lng签名ID = 0
    If int类型 = 0 Then
        strSQL = "Select 医嘱ID From 病人医嘱发送 A Where 发送号=[2]" & _
            " And Not Exists(Select ID From 病人医嘱记录 B Where B.ID=A.医嘱ID And (ID=[1] Or 相关ID=[1]))"
    Else
        strSQL = "Select 操作类型,操作时间,操作人员 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[3] And 操作时间=[4]"
        strSQL = "Select 医嘱ID,Nvl(签名ID,0) as 签名ID From 病人医嘱状态 A Where (操作类型,操作时间,操作人员)=(" & strSQL & ")" & _
            " And Not Exists(Select ID From 病人医嘱记录 B Where B.ID=A.医嘱ID And (ID=[1] Or 相关ID=[1] Or (A.操作类型=8 And 医嘱期效=1)))"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID, lng发送号, int类型, dat时间)
    If Not rsTmp.EOF Then
        If int类型 = 0 Then
'            '不能通过批量回退已出院或预出院病人的医嘱发送
'            strSQL = "Select C.病人ID,C.主页ID From 病人医嘱发送 A,病人医嘱记录 B,病案主页 C" & _
'                " Where A.医嘱ID=B.ID And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
'                " And (C.出院日期 is Not NULL Or C.状态=3) And A.发送号=[1] And Rownum=1"
'            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng发送号)
'            If Not rsTmp.EOF Then Exit Function
        ElseIf int类型 <> 0 Then
            rsTmp.Filter = "签名ID<>0"
            If Not rsTmp.EOF Then lng签名ID = rsTmp!签名ID
        End If
        AdviceCanBatchRoll = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncCopyToText()
    Dim strCopy As String, intRow As Integer
    
    With vsAdvice
        strCopy = ""
        For intRow = .FixedRows To .Rows - 1
            If InStr(",5,6,", .TextMatrix(intRow, COL_诊疗类别)) > 0 Then
                strCopy = strCopy & .TextMatrix(intRow, COL_医嘱内容) _
                        & " " & .TextMatrix(intRow, COL_单量) _
                        & " " & .TextMatrix(intRow, COL_频率) _
                        & " " & .TextMatrix(intRow, COL_用法) _
                        & vbCrLf
            Else
                strCopy = strCopy & .TextMatrix(intRow, COL_医嘱内容) & vbCrLf
            End If
        Next
    End With
    If strCopy <> "" Then
        VB.Clipboard.Clear
        VB.Clipboard.SetText strCopy
    End If
End Sub

Private Function RowIs配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否中药配方行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='7' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs配方行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否检验组合行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='C' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs检验行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlItemRef()
'功能：调用诊疗参考
    Dim lng诊疗项目ID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_诊疗类别) = "E" And (RowIs配方行(.Row) Or RowIs检验行(.Row)) Then
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    Call ShowClinicHelp(0, mfrmParent, lng诊疗项目ID)
End Sub

Public Sub zlPrintSetup()
    Call zlPrintSet
End Sub

Public Sub zlExcel()
    Call OutputList(3)
End Sub

Public Sub zlPreview()
    Call OutputList(2)
End Sub

Public Sub zlPrint()
    Call OutputList(1)
End Sub

Private Sub Form_Activate()
    If vsColumn.Visible Then
        vsColumn.SetFocus '列选择器
    Else
        picFocus.SetFocus '这样设置后本窗体内的焦点顺序才有效
        vsAdvice.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
    vsColumn.Visible = False '列选择器
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objMenu As Object
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    
    If KeyCode = vbKeyEscape Then '列选择器
        If vsColumn.Visible Then
            vsColumn.Visible = False
            vsAdvice.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then
        Call imgColSel_MouseUp(1, 0, 0, 0)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu新开医嘱)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyM Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu修改医嘱)
    ElseIf KeyCode = vbKeyDelete Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu删除医嘱)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu医嘱停止)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyZ Then
        Set objMenu = mfrmParent.mnuAdviceFuncRoll(0)
    ElseIf KeyCode = vbKeyF5 Then
        Call LoadAdvice
    ElseIf KeyCode = vbKeyF6 Then
        Call zlItemRef
    ElseIf KeyCode = vbKeyF9 Then
        If mfrmParent.mnuViewAdviceFilter.Visible And mfrmParent.mnuViewAdviceFilter.Enabled Then
            Call mfrmParent.mnuViewAdviceFiler_Click
        End If
    ElseIf KeyCode = vbKeyF8 Then
        Call mfrmParent.mnuViewAdviceCyc_Click
    End If
    
    If Not objMenu Is Nothing Then
        If objMenu.Enabled And objMenu.Visible Then
            Call zlMenuClick(objMenu)
        End If
    End If
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsAppend.Height - y < 60 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        tabAppend.Top = tabAppend.Top + y
        vsAppend.Top = vsAppend.Top + y
        vsAppend.Height = vsAppend.Height - y
        Me.Refresh
    End If
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsColumn
            If .Visible Then
                .Visible = False
                vsAdvice.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsAdvice.ColHidden(.RowData(i)) Or vsAdvice.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub mfrmEdit_Unload(Cancel As Integer)
    If Not Cancel Then
        If frmInAdviceEdit.mblnOK Then Call LoadAdvice
        Set mfrmEdit = Nothing
        
        '因PACS调用加上
        On Error Resume Next
        
        If mfrmParent.tabFunc.SelectedItem.Key = "医嘱" Then
            Call BringWindowToTop(Me.Hwnd)
        End If
    End If
End Sub

Private Function CheckWindow() As Boolean
'功能：检查医嘱编辑窗口是否已经打开
    If Not mfrmEdit Is Nothing Then
        '当前窗口打开了
        MsgBox "医嘱编辑窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        '定位到当前的窗口
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        If mfrmEdit.Visible Then mfrmEdit.SetFocus
        Exit Function
    Else
        '其它窗口打开了
        If Not CheckAdviceWindow("住院医嘱编辑") Then Exit Function
    End If
    CheckWindow = True
End Function

Private Sub FuncAdviceDel()
'删除：删除当前医嘱
'说明：在主界面删除,对检查组合,手术组合,中药配方,是整个删除,一并给药只删除当前药品
    Dim strSQL As String, lng医嘱ID As Long
    Dim blnGroup As Boolean, i As Long
    Dim lngRow As Long
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以删除。", vbInformation, gstrSysName
            Exit Sub
        End If

        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查是否可以删除
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "你不能删除该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If InStr(",1,2,", .TextMatrix(.Row, COL_医嘱状态)) = 0 Then
            MsgBox "当前选择的医嘱已经过校对，不能删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能删除
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能删除。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If

        '无执业资格的医生只能删除修改未审核的医嘱。
        If Not HaveAuditPriv And HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_开嘱医生))) Then
            MsgBox "你没有资格删除当前选择的医嘱，或者当前选择的医嘱已经过审核，不能删除。", vbInformation, gstrSysName
            Exit Sub
        End If

        If InStr(",5,6,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If blnGroup Then
                If MsgBox("医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """与其它药品一并给药,确实要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not blnGroup Then
            If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_病人医嘱记录_Delete(" & lng医嘱ID & ",1)"
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With vsAdvice
        '界面上直接删除
        .Redraw = False
        
        '删除一并给药第一行时的显示处理
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_相关ID)) = Val(.TextMatrix(.Row + 1, COL_相关ID)) Then
                If .TextMatrix(.Row, COL_开始时间) <> "" And .TextMatrix(.Row + 1, COL_开始时间) = "" Then
                    .TextMatrix(.Row + 1, COL_期效) = .TextMatrix(.Row, COL_期效)
                    .TextMatrix(.Row + 1, COL_开始时间) = .TextMatrix(.Row, COL_开始时间)
                    .TextMatrix(.Row + 1, COL_频率) = .TextMatrix(.Row, COL_频率)
                    .TextMatrix(.Row + 1, COL_用法) = .TextMatrix(.Row, COL_用法)
                End If
            End If
        End If
        
        lngRow = .Row
        .RemoveItem .Row
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        
        Call .ShowCell(.Row, .Col)
        .Redraw = True
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col) '颜色及附表更新
    End With
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSign()
'功能：对医嘱进行电子签名
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lng签名ID As Long, lng证书ID As Long
    Dim intRule As Integer
    
    If mlng病人ID = 0 Then Exit Sub
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign Is Nothing Then Exit Sub
    
    '获取签名医嘱源文
    intRule = ReadAdviceSignSource(1, mlng病人ID, mlng主页ID, strIDs, 0, mblnMoved, strSource, mlng前提ID)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "该病人目前没有可以签名的医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
    If strSign <> "" Then
        lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
        strSQL = "zl_医嘱签名记录_Insert(" & lng签名ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        Call LoadAdvice '刷新界面
        MsgBox "已完成电子签名。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'功能：校验医嘱的电子签名(可对已转移的数据)
    Dim strSource As String
    
    If mlng病人ID = 0 Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 2 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '获取签名医嘱源文
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '验证签名
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub

Private Sub FuncAdviceSignErase()
'功能：取消医嘱的电子签名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If mlng病人ID = 0 Then Exit Sub
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 2 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '作废和停止医嘱的签名不能取消
        If InStr(",4,8,", .Cell(flexcpData, .Row, 0)) > 0 Then
            MsgBox "不能直接取消作废或停止医嘱的签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        '新开签名必须是在新开或校对疑问状态
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If InStr(",1,2,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态))) = 0 Then
                MsgBox "由于医嘱已经经过校对，该签名不能取消。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '不能取消医技下达的签名
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) <> 0 Then
            MsgBox "你不能取消医技科室下达医嘱的签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        '只能取消自已签的名
        If .TextMatrix(.Row, 2) <> UserInfo.姓名 Then
            MsgBox "该签名人不是你本人，不能取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要取消这次签名吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_医嘱签名记录_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice '刷新界面
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsConsultation() As Boolean
'功能：判断当前执行功能的病人是否会诊病人
    On Error Resume Next
    IsConsultation = mfrmParent.tabPati.SelectedItem.Key = "会诊病人"
    Err.Clear: On Error GoTo 0
End Function

Private Sub FuncAdviceAudit()
'功能：审核医嘱
    If Not CheckWindow Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    
    If Not HaveAuditPriv Then
        MsgBox "你不具有审核医嘱的资格！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mlng主页ID, mlng前提ID, , , , , , IsConsultation, True)
End Sub

Private Sub FuncAdviceAdd()
'功能：增加新的医嘱
    If Not CheckWindow Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
        
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mlng主页ID, mlng前提ID, , , , , , IsConsultation)
End Sub

Private Sub FuncAdviceModi()
'功能：修改当前医嘱
    Dim lng医嘱ID As Long
    
    If Not CheckWindow Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then Exit Sub
        
        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '医技下达的医嘱
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "你不能修改该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已校对或已废止
        If InStr(",4,8,9,", .TextMatrix(.Row, COL_医嘱状态)) > 0 Then
            MsgBox "当前选择的医嘱已经作废或停止，不能修改。", vbInformation, gstrSysName
            Exit Sub
        ElseIf InStr(",1,2,", .TextMatrix(.Row, COL_医嘱状态)) = 0 Then
            MsgBox "当前选择的医嘱已经过校对，不能修改。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能修改
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能修改。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '无执业资格的医生只能删除修改未审核的医嘱。
        If Not HaveAuditPriv And HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_开嘱医生))) Then
            MsgBox "你没有资格修改当前选择的医嘱，或者当前选择的医嘱已经过审核，不能修改。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set mfrmEdit = frmInAdviceEdit
        Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mlng主页ID, mlng前提ID, , , Val(.TextMatrix(.Row, COL_婴儿ID)), lng医嘱ID, , IsConsultation)
    End With
End Sub

Private Sub FuncAdviceRevoke()
'功能：医嘱作废
    If mlng病人ID = 0 Then Exit Sub
    
    If mlng前提ID = 0 Then '用于医生站
        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        frmAdviceOperate.mstrPrivs = mstrPrivs
        frmAdviceOperate.mint类型 = 0
        frmAdviceOperate.mlng病区ID = mlng病区ID
        frmAdviceOperate.mlng病人ID = mlng病人ID
        frmAdviceOperate.mlng主页ID = mlng主页ID
        frmAdviceOperate.mlng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        frmAdviceOperate.Show 1, Me
        If frmAdviceOperate.mblnOK Then Call LoadAdvice
    Else
        If FuncAdviceRevoke0(2) Then Call LoadAdvice
    End If
End Sub

Private Function FuncAdviceRevoke0(ByVal int病人来源 As Integer) As Boolean
'删除：当前医嘱作废(一组医嘱作废)
    Dim strSQL As String, lng医嘱ID As Long
    
    Dim str医嘱ID As String, intRule As Integer
    Dim lng签名ID As Long, lng证书ID As Long
    Dim strSource As String, strSign As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以作废。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '检查是否可以作废
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "该医嘱不允许操作。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If int病人来源 = 1 Then
            If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 8 Then
                MsgBox "当前选择的门诊医嘱尚未发送或已经作废。", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If InStr(",1,2,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
                MsgBox "当前选择的住院医嘱尚未校对，请直接删除。", vbInformation, gstrSysName
                Exit Function
            End If
            If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
                MsgBox "当前选择的住院医嘱已经作废或停止。", vbInformation, gstrSysName
                Exit Function
            End If
            If .TextMatrix(.Row, COL_上次执行) <> "" Then
                MsgBox "当前选择的住院医嘱已经发送，不能再作废。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '电子签名检查和提示
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
                Else
                    MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
        End If
        
        If RowIn一并给药(.Row, 0, 0) Then
            If MsgBox("该组一并给药的医嘱将会一起作废，确实要作废吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("确实要作废医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        strSQL = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ")"
        
        '作废时的电子签名
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '获取签名医嘱源文
            str医嘱ID = lng医嘱ID '组ID,返回为明细ID
            intRule = ReadAdviceSignSource(4, mlng病人ID, mlng主页ID, str医嘱ID, 0, mblnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
            If strSign <> "" Then
                lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
                strSign = "zl_医嘱签名记录_Insert(" & lng签名ID & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱ID & "')"
            Else
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    FuncAdviceRevoke0 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceStop()
'功能：停止医嘱
    If mlng病人ID = 0 Then Exit Sub
    
    If mlng前提ID = 0 Then '用于医生站
        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Sub
        End If
                
        frmAdviceOperate.mstrPrivs = mstrPrivs
        frmAdviceOperate.mint类型 = 1
        frmAdviceOperate.mlng病区ID = mlng病区ID
        frmAdviceOperate.mlng病人ID = mlng病人ID
        frmAdviceOperate.mlng主页ID = mlng主页ID
        frmAdviceOperate.mlng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        frmAdviceOperate.Show 1, Me
        If frmAdviceOperate.mblnOK Then Call LoadAdvice
    Else
        If FuncAdviceStop0() Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdvicePause()
'功能：暂停医嘱
    If mlng病人ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmAdviceOperate.mstrPrivs = mstrPrivs
    frmAdviceOperate.mint类型 = 5
    frmAdviceOperate.mlng病区ID = mlng病区ID
    frmAdviceOperate.mlng病人ID = mlng病人ID
    frmAdviceOperate.mlng主页ID = mlng主页ID
    frmAdviceOperate.mlng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    frmAdviceOperate.mbln护士站 = True
    frmAdviceOperate.Show 1, Me
    If frmAdviceOperate.mblnOK Then Call LoadAdvice
End Sub

Private Sub FuncAdviceResume()
'功能：启用医嘱
    If mlng病人ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmAdviceOperate.mstrPrivs = mstrPrivs
    frmAdviceOperate.mint类型 = 6
    frmAdviceOperate.mlng病区ID = mlng病区ID
    frmAdviceOperate.mlng病人ID = mlng病人ID
    frmAdviceOperate.mlng主页ID = mlng主页ID
    frmAdviceOperate.mlng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    frmAdviceOperate.mbln护士站 = True
    frmAdviceOperate.Show 1, Me
    If frmAdviceOperate.mblnOK Then Call LoadAdvice
End Sub

Private Function FuncAdviceStop0() As Boolean
'删除：当前医嘱停止(仅用于住院长嘱)
    Dim strSQL As String, lng医嘱ID As Long
    Dim strStopTime As String
    
    Dim str医嘱ID As String, intRule As Integer
    Dim lng签名ID As Long, lng证书ID As Long
    Dim strSource As String, strSign As String
    Dim colStopTime As New Collection
    
    With vsAdvice
        '检查是否可以作废
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以停止。", vbInformation, gstrSysName
            Exit Function
        End If
                        
        If mblnMoved Then
            MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
                        
        If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
            MsgBox "该医嘱不允许操作。", vbInformation, gstrSysName
            Exit Function
        End If
                        
        '检查
        If .TextMatrix(.Row, COL_期效) <> "长嘱" Then
            MsgBox "当前选择的医嘱不是住院长期医嘱。", vbInformation, gstrSysName
            Exit Function
        End If
        If .TextMatrix(.Row, COL_总量) <> "" Then
            MsgBox "中药配方在发送后会自动停止。", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",1,2,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
            MsgBox "当前选择的住院医嘱尚未校对，请直接删除。", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
            MsgBox "当前选择的住院医嘱已经作废或停止。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '电子签名检查和提示
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "停止已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能停止。", vbInformation, gstrSysName
                Else
                    MsgBox "停止已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能停止。", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
        End If
        
        If RowIn一并给药(.Row, 0, 0) Then
            If MsgBox("该组一并给药的医嘱将会一起停止，确实要停止吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("确实要停止医嘱""" & .TextMatrix(.Row, COL_医嘱内容) & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        '停嘱时缺省的医嘱终止时间
        If .TextMatrix(.Row, COL_终止时间) = "" Then
            If gbln长期医嘱次日生效 Then
                strStopTime = Format(zlDatabase.Currentdate + 1, "yyyy-MM-dd 00:00")
            Else
                strStopTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            End If
        Else
            strStopTime = .TextMatrix(.Row, COL_终止时间)
        End If
        strSQL = "ZL_病人医嘱记录_停止(" & lng医嘱ID & ",To_Date('" & strStopTime & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.姓名 & "')"
        
        '停止时的电子签名
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '获取签名医嘱源文
            str医嘱ID = lng医嘱ID '组ID,返回为明细ID
            colStopTime.Add Format(strStopTime, "yyyy-MM-dd HH:mm:00"), "_" & lng医嘱ID
            intRule = ReadAdviceSignSource(8, mlng病人ID, mlng主页ID, str医嘱ID, 0, mblnMoved, strSource, , colStopTime)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要停止的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID)
            If strSign <> "" Then
                lng签名ID = zlDatabase.GetNextId("医嘱签名记录")
                strSign = "zl_医嘱签名记录_Insert(" & lng签名ID & ",8," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱ID & "')"
            Else
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    FuncAdviceStop0 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSupply()
'功能：补录医嘱
    If Not CheckWindow Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng病人ID, mlng主页ID, mlng前提ID, , True, , , , IsConsultation)
End Sub

Private Sub tabAppend_Click()
    If Val(vsAppend.Tag) = tabAppend.SelectedItem.Index Then Exit Sub
    
    If Visible Then
        Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
        
    vsAppend.Tag = tabAppend.SelectedItem.Index
    If tabAppend.SelectedItem.Index = 1 Then
        Call InitSendTable
    ElseIf tabAppend.SelectedItem.Index = 2 Then
        Call InitSignTable
    End If
    
    If Visible Then
        Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then vsAdvice.SetFocus
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next '为了外部系统调用增加，By：赵彤宇

    If NewRow = OldRow Then Exit Sub
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_开始时间)
    End If
    If vsAdvice.Redraw <> flexRDNone Then
        If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
            '显示医嘱附加表格的内容
            If mfrmParent.mnuViewAdviceAppend.Checked Then
                If tabAppend.SelectedItem.Index = 1 Then
                    Call ShowSendList(NewRow)
                ElseIf tabAppend.SelectedItem.Index = 2 Then
                    Call ShowSignList(NewRow)
                End If
            End If
            '显示医嘱可回退内容
            Call ShowRollList(NewRow)
        ElseIf mfrmParent.mnuViewAdviceAppend.Checked Then
            Call ClearAppendData
            vsAppend.Row = vsAppend.FixedRows
        End If
        
        Call SetFuncEnabled
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    If Col = COL_医嘱内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_皮试 Then
            Cancel = True
        ElseIf Col = COL_警示 Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, SysColor2RGB(.BackColorFixed)

            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '擦除一并给药相关行列的边线及内容
            lngLeft = COL_期效: lngRight = COL_开始时间
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_频率: lngRight = COL_用法
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                '为了支持预览输出
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '为了外部系统调用增加，By：赵彤宇
    On Error Resume Next
    If Button = 2 And mfrmParent.mnuAdvice.Visible Then PopupMenu mfrmParent.mnuAdvice, 2
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mlng病人ID = 0 Then Exit Sub
    
    '表头
    objOut.Title.Text = "病人医嘱清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    With mfrmParent.lvwPati.SelectedItem
        Set objRow = New zlTabAppRow
        objRow.Add "病人：" & .Text & " 性别：" & .SubItems(4) & " 年龄：" & .SubItems(5)
        objRow.Add "住院号：" & .SubItems(1) & " 床号：" & .SubItems(2)
        objOut.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "入院日期：" & .SubItems(8)
        objRow.Add "出院日期：" & .SubItems(9)
        objOut.UnderAppRows.Add objRow
    End With
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsAdvice
    
    '输出
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
        
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call InitColumnSelect '列选择器
    Call tabAppend_Click
    Call RestoreWinState(Me, App.ProductName)
    
    On Error Resume Next
    fraAdviceUD.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    tabAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    vsAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    Err.Clear: On Error GoTo 0
    '电子签名记录
    If gobjESign Is Nothing Then tabAppend.Visible = False
    
    Set mfrmEdit = Nothing
    Call InitSysPar '初始化系统参数
End Sub

Private Sub Form_Resize()
    Dim PriceH As Long
    
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    
    PriceH = IIF(vsAppend.Visible, vsAppend.Height + fraAdviceUD.Height + IIF(tabAppend.Visible, tabAppend.Height, 0), 0)
    
    vsAdvice.Left = 0
    vsAdvice.Top = 0
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - PriceH
    
    '列选择器
    With vsAdvice
        fraColSel.Left = .Left + (.ColWidth(0) + .ColWidth(1) - fraColSel.Width) / 2 + 30
        fraColSel.Top = .Top + (.RowHeight(0) - fraColSel.Height) / 2 + 30
    End With
    
    fraAdviceUD.Left = 0
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = Me.ScaleWidth
    
    tabAppend.Left = 0
    tabAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tabAppend.Width = Me.ScaleWidth
    
    vsAppend.Left = 0
    If tabAppend.Visible Then
        vsAppend.Top = tabAppend.Top + tabAppend.Height
    Else
        vsAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    End If
    vsAppend.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub GetAdviceWhere(int婴儿 As Integer, str期效 As String, str状态 As String, bln临嘱 As Boolean)
'功能：读取医嘱过滤设置条件,病人医嘱记录的表别名为"A"
'参数：bln临嘱=返回是否只显示临嘱
    Dim strWhere As String, strReg As String
    Dim strTmp As String, i As Long
    
    int婴儿 = -1: str期效 = "": str状态 = "": bln临嘱 = False
    
    '婴儿条件
    strReg = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\住院医嘱过滤", "病人婴儿", "1")
    If Val(strReg) = 1 Then
        int婴儿 = 0
    ElseIf Val(strReg) > 1 Then
        int婴儿 = Val(strReg) - 1
    End If
    
    '医嘱期效
    strReg = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\住院医嘱过滤", "医嘱期效", "11")
    If strReg <> "11" Then
        strTmp = ""
        For i = 1 To Len(strReg)
            If Val(Mid(strReg, i, 1)) = 1 Then
                strTmp = strTmp & "," & i - 1
            End If
        Next
        If strTmp <> "" Then str期效 = Mid(strTmp, 2)
        If strReg = "01" Then bln临嘱 = True
    End If
            
    '医嘱状态
    strReg = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\住院医嘱过滤", "医嘱状态", "111")
    If strReg <> "111" Then
        strTmp = ""
        For i = 1 To Len(strReg)
            If Val(Mid(strReg, i, 1)) = 1 Then
                If i = 1 Then strTmp = strTmp & ",1,2"  '未校对
                If i = 2 Then strTmp = strTmp & ",3,5,6,7"  '已校对
                If i = 3 Then strTmp = strTmp & ",4,8,9"  '已废止
            End If
        Next
        If strTmp <> "" Then str状态 = Mid(strTmp, 2)
    End If
End Sub

Private Sub ClearAdviceData()
'功能：清除医嘱清单数据
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitColumnSelect()
'功能：根据医嘱清单原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vsAdvice
        For i = .FixedCols To .Cols - 1
            If Not (.ColHidden(i) Or .ColWidth(i) = 0) Then
                If .TextMatrix(0, i) <> "" Then '审查结果,皮试
                    vsColumn.Rows = vsColumn.Rows + 1
                    lngRow = vsColumn.Rows - 1
                    vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                    vsColumn.RowData(lngRow) = i
                    
                    '固定显示列
                    If InStr(",开始时间,医嘱内容,开嘱医生,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsColumn.TextMatrix(lngRow, 0) = 1
                        vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                    End If
                End If
            End If
        Next
    End With
    vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 130
    vsColumn.Row = 1
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ID;相关ID;组ID;组号;婴儿ID;医嘱状态;诊疗类别;操作类型;毒理分类;标志;" & _
        ",240,4;期效,500,4;开始时间,1080,1;医嘱内容,3000,1;,375,4;" & _
        "总量,850,1;单量,850,1;频率,1000,1;用法,1000,1;医生嘱托,1000,1;执行时间,1000,1;" & _
        "终止时间,1560,1;执行科室,850,1;执行性质,850,1;上次执行,1560,1;状态,500,4;" & _
        "开嘱医生,850,1;开嘱时间,1080,1;校对护士,850,1;校对时间,1080,1;停嘱医生,850,1;" & _
        "停嘱时间,1080,1;停嘱护士,850,1;确认停嘱时间,1180,1;单据ID;申请项;报告项;申请ID;前提ID;签名否"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i) '记录原始列宽用于列选择器
        Next
        .ColHidden(COL_警示) = Not (gblnPass And InStr(mstrPrivs, "合理用药监测") > 0) 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(0) = 9 * Screen.TwipsPerPixelX
        .ColWidth(1) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Sub ClearAppendData()
'功能：清除附加表格数据
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
End Sub

Private Sub InitSendTable()
'功能：初始化发送清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "发送号;发送时间,1080,1;发送医嘱,1800,1;单据号,850,1;收费项目,1800,1;发送数次,850,1;计费状态,850,1;执行状态,850,1;执行科室,850,1;首次时间,1080,1;末次时间,1080,1;发送人,800,1;记录性质"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
End Sub

Private Sub InitSignTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "签名类型,1150,1;签名时间,1900,1;签名人,800,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = False
        .MergeCol(1) = False
    End With
End Sub

Private Function LoadAdvice() As Boolean
'功能：根据当前界面设置读取并显示医嘱清单
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim lngTop As Long, i As Long, j As Long
    Dim strFormat As String, strTmp As String
    Dim bln给药途径 As Boolean, bln中药用法 As Boolean
    Dim bln采集方法 As Boolean, bln申请项 As Boolean, bln已申请 As Boolean
    Dim str状态 As String, lng医嘱ID As Long
    Dim blnFirst As Boolean, strBill As String
    Dim str医嘱期效 As String, str医嘱状态 As String
    Dim int婴儿 As Integer, bln临嘱 As Boolean
    Dim bln重整 As Boolean, dat重整 As Date
    Dim blnDo As Boolean, strCurr As String, strTime As String
    
    If mlng病人ID = 0 Then Exit Function
    
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) '记录当前行
    
    '医嘱过滤条件
    Call GetAdviceWhere(int婴儿, str医嘱期效, str医嘱状态, bln临嘱)
    strWhere = ""
    If int婴儿 <> -1 Then
        strWhere = strWhere & " And Nvl(A.婴儿,0)=[4]"
    End If
    If str医嘱期效 <> "" Then
        strWhere = strWhere & " And Instr([5],','||Nvl(A.医嘱期效,0)||',')>0"
    End If
    If str医嘱状态 <> "" Then
        strWhere = strWhere & " And Instr([6],','||Nvl(A.医嘱状态,0)||',')>0"
    End If
    
    bln重整 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\住院医嘱过滤", "重整医嘱", "0")) <> 0
    If bln重整 Then
        '求最近次重整时间
        strSQL = "Select Max(B.操作时间) as 时间 From 病人医嘱记录 A,病人医嘱状态 B" & _
            " Where A.ID=B.医嘱ID And B.操作类型=5 And A.病人ID=[1] And A.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID)
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!时间) Then
                dat重整 = Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        If dat重整 = CDate(0) Then
            bln重整 = False
        Else
            strSQL = _
                " Select Distinct" & _
                "   A.ID,A.相关ID,A.病人ID,A.主页ID,A.序号,A.婴儿,A.医嘱状态,A.医嘱期效,A.诊疗类别," & _
                "   A.紧急标志,A.审查结果,A.开始执行时间,A.医嘱内容,A.皮试结果,A.总给予量,A.单次用量," & _
                "   A.执行频次,A.医生嘱托,A.执行时间方案,A.执行终止时间,A.执行性质,A.上次执行时间," & _
                "   A.开嘱医生,A.开嘱时间,A.校对护士,A.校对时间,A.停嘱医生,A.停嘱时间,A.确认停嘱时间," & _
                "   A.诊疗项目ID,A.执行科室ID,A.收费细目ID,A.申请ID,A.前提ID,A.病人来源" & _
                " From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And (B.操作时间>=[3] Or A.医嘱状态 IN(1,2))" & _
                " And A.病人ID=[1] And A.主页ID=[2]" & strWhere
            strSQL = "(" & strSQL & ")"
        End If
    End If
    
    '诊疗单据：对应诊疗单据,及申请项,报告项
    strBill = "Select A.ID as 医嘱ID,B.病历文件ID as 单据ID," & _
        " Max(Decode(C.填写时机,1,1,0)) as 申请项," & _
        " Max(Decode(C.填写时机,2,1,0)) as 报告项" & _
        " From 病人医嘱记录 A,诊疗单据应用 B,病历文件组成 C" & _
        " Where A.诊疗项目ID=B.诊疗项目ID And B.应用场合=2 And B.病历文件ID=C.病历文件ID(+)" & _
        " And A.病人ID=[1] And A.主页ID=[2]" & strWhere & _
        " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
        " Group by A.ID,B.病历文件ID"
        
    '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法
    str状态 = "Decode(A.医嘱状态,1,'新开',2,'疑问',3,'校对',4,'作废',5,'重整',6,'暂停',7,'启用',8,'停止',9,'确认停止')"
    strSQL = _
        "Select /*+ RULE */ A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " Nvl(A.婴儿,0) as 婴儿ID,A.医嘱状态,Nvl(A.诊疗类别,'*') as 诊疗类别,B.操作类型,C.毒理分类,A.紧急标志 as 标志," & _
            " A.审查结果,Decode(Nvl(A.医嘱期效,0),0,'长嘱','临嘱') as 期效," & _
            " To_Char(A.开始执行时间,'MM-DD HH24:MI') as 开始时间,A.医嘱内容,A.皮试结果 as 皮试," & _
            " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位),'5',Round(A.总给予量/D.住院包装,5)||D.住院单位,'6',Round(A.总给予量/D.住院包装,5)||D.住院单位,A.总给予量||B.计算单位)) as 总量," & _
            " Decode(A.单次用量,NULL,NULL,A.单次用量||B.计算单位) as 单量," & _
            " A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('246',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法," & _
            " A.医生嘱托,A.执行时间方案 as 执行时间,To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间," & _
            " Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室," & _
            " Decode(Instr('567E',Nvl(A.诊疗类别,'*')),0,NULL,A.执行性质) as 执行性质," & _
            " To_Char(A.上次执行时间,'YYYY-MM-DD HH24:MI') as 上次执行," & str状态 & " as 状态," & _
            " A.开嘱医生,To_Char(A.开嘱时间,'MM-DD HH24:MI') as 开嘱时间,A.校对护士,To_Char(A.校对时间,'MM-DD HH24:MI') as 校对时间," & _
            " A.停嘱医生,To_Char(A.停嘱时间,'MM-DD HH24:MI') as 停嘱时间,F.操作人员 as 停嘱护士," & _
            " To_Char(A.确认停嘱时间,'MM-DD HH24:MI') as 确认停嘱时间," & _
            " Y.单据ID,Y.申请项,Y.报告项,A.申请ID,A.前提ID,Decode(S.签名ID,NULL,0,1) as 签名否" & _
        " From " & IIF(bln重整, strSQL, "病人医嘱记录") & " A,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B," & _
            " 病人医嘱状态 F,病人医嘱状态 S,病人医嘱记录 X,(" & strBill & ") Y" & _
        " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+)" & _
            " And A.收费细目ID=D.药品ID(+) And A.相关ID=X.ID(+) And A.ID=Y.医嘱ID(+)" & _
            " And Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL)" & _
            " And A.ID=F.医嘱ID(+) And F.操作类型(+)=9 And A.ID=S.医嘱ID And S.操作类型=1" & _
            " And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And A.病人来源<>3" & strWhere & _
            IIF(mlng前提ID = 0 Or mblnShowAll, "", " And A.前提ID=[7]") & _
        " Order by Nvl(A.婴儿,0),组号,组ID,A.序号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID, dat重整, int婴儿, "," & str医嘱期效 & ",", "," & str医嘱状态 & ",", mlng前提ID)
    
    If Not rsTmp.EOF Then
        strCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        With vsAdvice
            .Redraw = False
            
            '绑定时按设计时的FormatString恢复一些缺省值(固定行列数，固定行列文字及行列对齐,尺寸,可见)
            'FormatString在运行时赋值无效
            '如果AutoResize=True,则所有列宽或行高被自动调整(根据AutoSizeMode)
            '如果WordWrap=True,则行高会被自动调整
            .WordWrap = False
            strFormat = GetColFormat(vsAdvice)
            Call ClearAdviceData
            .ScrollBars = flexScrollBarNone
            Set .DataSource = rsTmp
            .ScrollBars = flexScrollBarBoth
            If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
                gcnOracle.Errors.Clear '怪,绑定时固定有此错误
            End If
            Call SetColFormat(vsAdvice, strFormat)
            .TextMatrix(0, COL_皮试) = ""
            .TextMatrix(0, COL_警示) = "" 'Pass
            
            '自动调整行高
            .WordWrap = True
            .AutoSize COL_医嘱内容
            
            '处理每行医嘱
            i = .FixedRows
            Do While i <= .Rows - 1
                '成药及中药的一些处理
                bln给药途径 = False: bln中药用法 = False
                bln采集方法 = False: bln申请项 = False: bln已申请 = False '仅用于检验组合
                If .TextMatrix(i, COL_诊疗类别) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln给药途径 = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '显示成药的给药途径
                                    .TextMatrix(j, COL_用法) = .TextMatrix(i, COL_用法)
                                    '显示成药的执行性质
                                    If Val(.TextMatrix(j, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                        .TextMatrix(j, COL_执行性质) = "自备药"
                                    ElseIf Val(.TextMatrix(j, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                        .TextMatrix(j, COL_执行性质) = "离院带药"
                                    Else
                                        .TextMatrix(j, COL_执行性质) = ""
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln中药用法 = .TextMatrix(i - 1, COL_诊疗类别) = "7" '中药用法行
                            bln采集方法 = .TextMatrix(i - 1, COL_诊疗类别) = "C" '采集方法行
                            
                            '显示中药配方或检验组合的执行科室
                            .TextMatrix(i, COL_执行科室) = .TextMatrix(i - 1, COL_执行科室)
                            
                            If bln中药用法 Then
                                '显示中药配方执行性质
                                If Val(.TextMatrix(i - 1, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                    .TextMatrix(i, COL_执行性质) = "自备药"
                                ElseIf Val(.TextMatrix(i - 1, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                    .TextMatrix(i, COL_执行性质) = "离院带药"
                                Else
                                    .TextMatrix(i, COL_执行性质) = ""
                                End If
                            Else
                                .TextMatrix(i, COL_执行性质) = ""
                            End If
                            
                            '删除单味中药行,以及检验组合中的检验项目;同时判断检验申请
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    If .TextMatrix(j, COL_诊疗类别) = "C" Then
                                        If Val(.TextMatrix(j, COL_申请项)) = 1 Then
                                            bln申请项 = True
                                            If Val(.TextMatrix(j, COL_申请ID)) <> 0 Then
                                                bln已申请 = True
                                            End If
                                        End If
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        .TextMatrix(i, COL_执行性质) = ""
                    End If
                End If
                                                                
                '处理可见行的的一些标识:排开不可见但暂时未删除的行
                If Not bln给药途径 And .TextMatrix(i, COL_诊疗类别) <> "7" Then
                
                    '行高：为了支持zl9PrintMode:Resize之后,取RowHeight可能小于RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    '处理小数点问题,暂未想到办法
                    If Left(.TextMatrix(i, COL_总量), 1) = "." Then
                        .TextMatrix(i, COL_总量) = "0" & .TextMatrix(i, COL_总量)
                    End If
                    If Left(.TextMatrix(i, COL_单量), 1) = "." Then
                        .TextMatrix(i, COL_单量) = "0" & .TextMatrix(i, COL_单量)
                    End If
                    
                    '可申请医嘱标识(不管药品及相关医嘱,且只管主要医嘱)
                    If Not bln中药用法 And InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                        If bln采集方法 Then '利用前面取的结果
                            If bln申请项 Then
                                If Not bln已申请 Then
                                    Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("未申请").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("已申请").Picture
                                End If
                            End If
                        ElseIf Val(.TextMatrix(i, COL_申请项)) = 1 Then
                            If Val(.TextMatrix(i, COL_申请ID)) = 0 Then
                                Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("未申请").Picture
                            Else
                                Set .Cell(flexcpPicture, i, COL_F申请) = imgFlag.ListImages("已申请").Picture
                            End If
                        End If
                    End If
                    
                    '医嘱颜色
                    blnDo = False
                    If Val(.TextMatrix(i, COL_医嘱状态)) = 2 Then
                        '校对疑问
                        If lngTop = 0 Then lngTop = i '有删除行也不会影响取值
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80& '深红
                        blnDo = True
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 4 Then
                        '已作废
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                        .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                        blnDo = True
                    ElseIf InStr(",8,9,", Val(.TextMatrix(i, COL_医嘱状态))) > 0 Then
                        '已停止,已确认停止:长嘱都以终止时间进行判断
                        If strCurr >= .TextMatrix(i, COL_终止时间) Or .TextMatrix(i, COL_期效) = "临嘱" Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 6 Then
                        '已暂停
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 6), "yyyy-MM-dd HH:mm")
                        If strCurr >= strTime Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '深绿
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 7 Then
                        '已启用
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 7), "yyyy-MM-dd HH:mm")
                        If strCurr < strTime Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '深绿
                            blnDo = True
                        End If
                    End If
                    If Not blnDo Then
                        If lngTop = 0 Then lngTop = i
                        If Val(.TextMatrix(i, COL_医嘱状态)) <> 1 Then
                            '已通过校对(也包含后续的多个状态)
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '深蓝
                        End If
                    End If
                    
                    '校对后术后医嘱红色显示
                    If .TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 4 _
                        And InStr(",1,2,4,", Val(.TextMatrix(i, COL_医嘱状态))) = 0 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '红色
                    End If
                    
                    '毒麻精药品标识:中药配方及组成味中药不处理
                    If .TextMatrix(i, COL_毒理分类) <> "" Then
                        If InStr(",麻醉药,毒性药,精神药,", .TextMatrix(i, COL_毒理分类)) > 0 Then
                            .Cell(flexcpFontBold, i, COL_医嘱内容) = True
                        End If
                    End If
                    
                    '皮试结果标识
                    If .TextMatrix(i, COL_皮试) = "(+)" Then
                        .Cell(flexcpForeColor, i, COL_皮试) = vbRed
                    ElseIf .TextMatrix(i, COL_皮试) = "(-)" Then
                        .Cell(flexcpForeColor, i, COL_皮试) = vbBlue
                    End If
                    
                    '紧急标志:一并给药只显示在第一行
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_标志)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("紧急").Picture
                        ElseIf Val(.TextMatrix(i, COL_标志)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = imgFlag.ListImages("补录").Picture
                        End If
                    End If
                    
                    'Pass:根据审查结果显示警示灯
                    If .TextMatrix(i, COL_警示) <> "" Then
                        Set .Cell(flexcpPicture, i, COL_警示) = imgPass.ListImages(Val(.TextMatrix(i, COL_警示)) + 1).Picture
                        .TextMatrix(i, COL_警示) = ""
                    End If
                    
                    '电子签名标识
                    If Val(.TextMatrix(i, COL_签名否)) = 1 Then
                        Set .Cell(flexcpPicture, i, COL_医嘱内容) = imgSign.ListImages(1).Picture
                    End If
                End If
                
                If bln给药途径 Then
                    .RemoveItem i
                Else
                    i = i + 1
                End If
            Loop
            
            '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '电子签名图标对齐
            .Cell(flexcpPictureAlignment, .FixedRows, COL_医嘱内容, .Rows - 1, COL_医嘱内容) = 0
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
    
    '只有临嘱时才用红色表线
    vsAdvice.GridColor = IIF(bln临嘱, &H8080FF, vsAdvice.GridColorFixed)
        
    '缺省定位
    vsAdvice.Redraw = flexRDNone
    If lng医嘱ID <> 0 Then
        lng医嘱ID = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
        If lng医嘱ID <> -1 Then vsAdvice.Row = lng医嘱ID
    End If
    If lng医嘱ID = -1 Or lng医嘱ID = 0 Then
        If lngTop <> 0 Then
            vsAdvice.Row = lngTop
            vsAdvice.TopRow = lngTop
        Else
            vsAdvice.Row = vsAdvice.FixedRows
        End If
    End If
    vsAdvice.Col = vsAdvice.FixedCols
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Redraw = flexRDDirect
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Refresh
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的发送记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeRestrictRows
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部份执行')"
        strExe2 = "Decode(Nvl(B.执行状态,0),0,'未执行',1,'执行完成',2,'拒绝执行',3,'正在执行')"
        strState = "Decode(A.记录性质,1,Decode(A.记录状态,0,'收费划价',1,'已收费',3,'已退费'),2,Decode(A.记录状态,0,'记帐划价',1,'已记帐',3,'已销帐'),'已计费')"
        
        '药嘱对应的药品计价按住院包装显示,非药嘱对应的药品计价按零售单位显示
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Not RowIn一并给药(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '成药部份:填写了发送记录,但可能无对应费用(如自备药,但医嘱有规格)
            strSub = "Select A.*,B.住院包装,B.住院单位" & _
                " From 病人费用记录 A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL And A.收费类别 IN('5','6','7')" & _
                " And A.收费细目ID=B.药品ID And A.医嘱序号=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "病人费用记录", "H病人费用记录")
            ElseIf MovedByDate(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "病人费用记录", "H病人费用记录")
            End If
                
            strSQL = _
                " Select C.相关ID,C.标本部位,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                " Nvl(A.住院单位,D.住院单位) as 单位," & _
                " Nvl(A.数次/Nvl(A.住院包装,1),B.发送数次/Nvl(D.剂量系数,1)/Nvl(D.住院包装,1)) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID,Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态," & _
                " B.首次时间,B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费',1," & strState & ") as 计费状态," & _
                " B.发送人,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别" & _
                " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,药品规格 D" & _
                " Where B.医嘱ID=C.ID And C.收费细目ID=D.药品ID(+)" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And A.医嘱序号(+)=B.医嘱ID" & _
                " And C.ID=[1]"

            '在一并给药的首行才显示给药途径的发送
            If lngRow = lngBegin Then
                '给药途径部份:填写了发送记录(叮嘱无),但不一定有费用
                strSub = "Select A.*,B.住院包装,B.住院单位" & _
                    " From 病人费用记录 A,药品规格 B" & _
                    " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                    " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[2]"
                If mblnMoved Then
                    strSub = Replace(strSub, "病人费用记录", "H病人费用记录")
                ElseIf MovedByDate(mvInDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, "病人费用记录", "H病人费用记录")
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select C.相关ID,C.标本部位,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                    " Decode(Nvl(Instr('567',A.收费类别),0),0,D.计算单位,Nvl(A.住院单位,E.住院单位)) as 单位," & _
                    " Decode(Nvl(Instr('567',A.收费类别),0),0,B.发送数次," & _
                    "   Nvl(A.数次/Nvl(A.住院包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.住院包装,1))) as 发送数次," & _
                    " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID,Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态," & _
                    " B.首次时间,B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费',1," & strState & ") as 计费状态," & _
                    " B.发送人,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别" & _
                    " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D,药品规格 E" & _
                    " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+)" & _
                    " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID" & _
                    " And C.ID=[2]"
            End If
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        Else
            '其它医嘱(包括配方及检查，手术一组医嘱):填写了发送记录(叮嘱无),但不一定有费用
            '中药自备药也是无对应费用(但医嘱有规格)
            strSub = _
                " Select A.*,B.住院包装,B.住院单位" & _
                " From 病人费用记录 A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[1]"
            strSub = strSub & " Union ALL " & _
                " Select A.*,B.住院包装,B.住院单位" & _
                " From 病人费用记录 A,药品规格 B,病人医嘱记录 C" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=C.ID" & _
                " And C.相关ID=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "病人费用记录", "H病人费用记录")
            ElseIf MovedByDate(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "病人费用记录", "H病人费用记录")
            End If
            
            strSQL = _
                " Select * From 病人医嘱记录 Where ID=[1]" & _
                " Union ALL " & _
                " Select * From 病人医嘱记录 Where 相关ID=[1]"
            strSQL = _
                " Select C.相关ID,C.标本部位,B.发送时间,B.NO,B.记录性质,A.收费细目ID," & _
                " Decode(Nvl(Instr('567',A.收费类别),0),0,D.计算单位,Nvl(A.住院单位,E.住院单位)) as 单位," & _
                " Decode(Nvl(Instr('567',A.收费类别),0),0,B.发送数次," & _
                "   Nvl(Nvl(A.付数,1)*A.数次/Nvl(A.住院包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.住院包装,1))) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID,Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态," & _
                " B.首次时间,B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费',1," & strState & ") as 计费状态," & _
                " B.发送人,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别" & _
                " From (" & strSub & ") A,病人医嘱发送 B,(" & strSQL & ") C,诊疗项目目录 D,药品规格 E" & _
                " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+)" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        End If
        
        strSQL = "Select /*+ RULE */ A.发送序号,A.费用序号," & _
            " A.相关ID,A.诊疗类别,F.名称 as 类别名称,D.名称 as 诊疗项目,A.标本部位,A.发送时间,A.NO,A.记录性质," & _
            " Nvl(G.名称,B.名称)||Decode(B.产地,NULL,NULL,'('||B.产地||')')||Decode(B.规格,NULL,NULL,' '||B.规格) as 收费项目," & _
            " A.单位,A.发送数次 as 数量,C.名称 as 执行科室,A.执行状态,A.首次时间,A.末次时间,A.计费状态,A.发送人,A.发送号" & _
            " From (" & strSQL & ") A,收费项目目录 B,部门表 C,诊疗项目目录 D,诊疗项目类别 F,收费项目别名 G" & _
            " Where A.收费细目ID=B.ID(+) And A.执行部门ID=C.ID(+)" & _
            " And A.诊疗项目ID=D.ID And A.诊疗类别=F.编码" & _
            " And A.收费细目ID=G.收费细目ID(+) And G.码类(+)=1 And G.性质(+)=" & IIF(gbln商品名, 3, 1) & _
            " Order by A.发送号 Desc,A.诊疗类别,A.发送序号,A.费用序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
        
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, cs发送号) = Nvl(rsTmp!发送号, 0)
                .TextMatrix(i, cs发送时间) = Format(Nvl(rsTmp!发送时间), "MM-dd HH:mm")
                
                '发送医嘱
                If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, cs发送医嘱) = "药品医嘱-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, cs发送医嘱) = "给药途径-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, cs发送医嘱) = "采集方法-" & rsTmp!诊疗项目
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, cs发送医嘱) = "中药煎法-" & rsTmp!诊疗项目
                    Else
                        .TextMatrix(i, cs发送医嘱) = "中药用法-" & rsTmp!诊疗项目
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, cs发送医嘱) = "检验项目-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        .TextMatrix(i, cs发送医嘱) = "检查部位-" & Nvl(rsTmp!标本部位)
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, cs发送医嘱) = "附加手术-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, cs发送医嘱) = "麻醉项目-" & rsTmp!诊疗项目
                    End If
                Else
                    .TextMatrix(i, cs发送医嘱) = rsTmp!类别名称 & "医嘱-" & rsTmp!诊疗项目
                End If
               
                .TextMatrix(i, cs单据号) = Nvl(rsTmp!NO)
                .TextMatrix(i, cs收费项目) = Nvl(rsTmp!收费项目)
                .TextMatrix(i, cs数次) = FormatEx(Nvl(rsTmp!数量), 5) & Nvl(rsTmp!单位)
                .TextMatrix(i, cs计费状态) = Nvl(rsTmp!计费状态)
                .TextMatrix(i, cs执行状态) = Nvl(rsTmp!执行状态)
                .TextMatrix(i, cs执行科室) = Nvl(rsTmp!执行科室)
                .TextMatrix(i, cs首次时间) = Format(Nvl(rsTmp!首次时间), "MM-dd HH:mm")
                .TextMatrix(i, cs末次时间) = Format(Nvl(rsTmp!末次时间), "MM-dd HH:mm")
                .TextMatrix(i, cs发送人) = Nvl(rsTmp!发送人)
                .TextMatrix(i, cs记录性质) = Nvl(rsTmp!记录性质)
                rsTmp.MoveNext
            Next
        End If
        
        .Row = 1: .Col = cs发送医嘱
        .Redraw = True
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的签名记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.签名ID,A.操作类型,B.签名时间,B.签名人," & _
            " Decode(A.操作类型,1,'新开医嘱',4,'作废医嘱',8,'停止医嘱','其它操作') as 签名类型" & _
            " From 病人医嘱状态 A,医嘱签名记录 B Where A.医嘱ID=[1] And A.签名ID=B.ID Order by B.签名时间"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!签名ID)
                .TextMatrix(i, 0) = rsTmp!签名类型
                .Cell(flexcpData, i, 0) = Val(rsTmp!操作类型)
                .TextMatrix(i, 1) = Format(rsTmp!签名时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!签名人
                Set .Cell(flexcpPicture, i, 0) = imgSign.ListImages(1).Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowRollList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱可以回退的内容在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objMenu As Menu
    
    For i = mfrmParent.mnuAdviceFuncRoll.UBound To 0 Step -1
        mfrmParent.mnuAdviceFuncRoll(i).Tag = ""
        If i = 0 Then
            mfrmParent.mnuAdviceFuncRoll(i).Caption = "<无内容>"
        Else
            On Error Resume Next
            Unload mfrmParent.mnuAdviceFuncRoll(i)
            If Err.Number <> 0 Then
                Err.Clear
                mfrmParent.mnuAdviceFuncRoll(i).Visible = False
                mfrmParent.mnuAdviceFuncRoll(i).Tag = ""
            End If
            On Error GoTo 0
        End If
    Next
    mfrmParent.mnuAdviceFunc(mnu医嘱回退).Tag = ""
    If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
        ShowRollList = True: Exit Function
    End If
    
    On Error GoTo errH
    
    '可回退医嘱操作和发送,医嘱本身的操作优先(如发送后自动停止)
    '临嘱不可回退自动停止,回退发送时自动回退停止
    strSQL = " And (A.ID=[1] Or A.相关ID=[1])"
    strSQL = _
        " Select Distinct 0 as 发送号,B.操作人员 as 人员,B.操作时间 as 时间,B.操作类型," & _
        " Decode(B.操作类型,4,'作废医嘱',5,'重整医嘱',6,'暂停医嘱',7,'启用医嘱',8,'停止医嘱',9,'确认停止',10,'皮试结果') as 内容" & _
        " From 病人医嘱记录 A,病人医嘱状态 B" & _
        " Where A.ID=B.医嘱ID" & strSQL & _
        " And (Nvl(A.医嘱期效,0)=0 And B.操作类型 Not IN(1,2,3)" & _
            " Or Nvl(A.医嘱期效,0)=1 And B.操作类型 Not IN(1,2,3,8))" & _
        " Union ALL" & _
        " Select Distinct B.发送号,B.发送人 as 人员,B.发送时间 as 时间,0 as 操作类型,'发送医嘱' as 内容" & _
        " From 病人医嘱记录 A,病人医嘱发送 B" & _
        " Where A.ID=B.医嘱ID" & strSQL & _
        " Order by 时间 Desc,发送号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_组ID)))
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If i > 1 Then
                If mfrmParent.mnuAdviceFuncRoll.UBound >= i - 1 Then
                    mfrmParent.mnuAdviceFuncRoll(i - 1).Visible = True
                Else
                    Load mfrmParent.mnuAdviceFuncRoll(i - 1)
                End If
            End If
            Set objMenu = mfrmParent.mnuAdviceFuncRoll(mfrmParent.mnuAdviceFuncRoll.UBound)
            objMenu.Caption = "操作人:" & rsTmp!人员 & ",时间:" & Format(rsTmp!时间, "MM-dd HH:mm") & ",操作:" & rsTmp!内容
            '记录回退类型,发送号,发送时间,人员
            objMenu.Tag = rsTmp!操作类型 & "|" & rsTmp!发送号 & "|" & Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!人员
            If i = 1 Then
                '医生只能回退作废、停止、自已的临嘱发送操作
                If InStr(",0,4,8,", Nvl(rsTmp!操作类型, 0)) > 0 Then
                    objMenu.Enabled = True
                Else
                    objMenu.Enabled = False
                End If
            Else
                objMenu.Enabled = False
            End If
            rsTmp.MoveNext
        Next
        mfrmParent.mnuAdviceFunc(mnu医嘱回退).Tag = rsTmp.RecordCount
    End If
    
    ShowRollList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            vsAdvice.ColWidth(lngCol) = vsAdvice.ColData(lngCol)
            vsAdvice.ColHidden(lngCol) = False
        Else
            vsAdvice.ColWidth(lngCol) = 0
            vsAdvice.ColHidden(lngCol) = True
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub
