VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBloodExec 
   BorderStyle     =   0  'None
   Caption         =   "执行登记"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VSFlex8Ctl.VSFlexGrid vsExec 
      Height          =   1485
      Left            =   -30
      TabIndex        =   0
      Top             =   825
      Width           =   7125
      _cx             =   12568
      _cy             =   2619
      Appearance      =   2
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
      BackColorSel    =   16444122
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
      MergeCells      =   2
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
   Begin ComctlLib.ImageList imgList 
      Left            =   1635
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":0000
            Key             =   "未执行"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":059A
            Key             =   "已执行"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":0B34
            Key             =   "拒绝执行"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmBloodExec.frx":10CE
            Key             =   "正在执行"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsExec 
      Left            =   0
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsVsf As clsVsf
Private mlngSys As Long
Private mlngModul As Long
Private mlng医嘱ID As Long
Private mlng医护科室ID As Long
Private mlng病区ID As Long, mint场合 As Integer
Private mlngFontSize As Long
'医嘱发送相关
Private mlng发送号 As Long, mint执行状态 As Integer
Private mint记录性质 As Integer, mint门诊记帐 As Integer, mstrNO As String, mlng执行部门ID As Long
Private mint计费状态 As Integer, mlng病人ID As Long, mlng主页id As Long, mlng组ID As Long, mstr发送时间 As String

Private mstrPrivs As String
Private mblnMoved As Boolean
Private mblnLoad As Boolean
Private mfrmParent As Object

Private mblnShowExec As Boolean
Private mblnExecFresh As Boolean  '是否是执行过程刷新(主要是避免重复刷新，改变量在调用窗口赋值)
Private Enum CMD_EXEC
    ID_显示执行 = 1
    ID_完成执行 = 2
    ID_取消完成 = 3
    ID_执行记录 = 4
    ID_执行调整 = 5
    ID_执行删除 = 6
    ID_执行前核对 = 7 '输血前核对
    ID_取消执行前核对 = 8
    ID_执行中核对 = 9 '执行中核对
    ID_取消执行中核对 = 10
End Enum

Private Enum Enum_ExecState
    E_清除状态 = 0
    E_记录执行 = 1
    E_删除执行 = 2
    E_执行完成 = 3
    E_取消完成 = 4
    E_执行核对 = 5
    E_取消核对 = 6
End Enum
Private mintAdviceExecState As Enum_ExecState

Public Event ShowExec(ByVal blnShow As Boolean, ByVal lngHeight As Long)

Public Property Get AdviceExecState() As Integer
'执行状态(供临床界面刷新使用)
    AdviceExecState = mintAdviceExecState
End Property

Public Property Let AdviceExecState(intAdviceExecState As Integer)
    mintAdviceExecState = intAdviceExecState
End Property

Public Property Let ExecFresh(blnFresh As Boolean)
'是否是执行过程刷新
    mblnExecFresh = blnFresh
End Property

Public Property Get IsShowExec() As Boolean
    IsShowExec = mblnShowExec
End Property

Public Property Let IsShowExec(blnValue As Boolean)
    Call SetShowExec(blnValue)
    mblnShowExec = blnValue
    RaiseEvent ShowExec(mblnShowExec, Me.Height)
End Property

Public Function zlRefresh(ByVal frmParent As Object, ByVal lngSys As Long, ByVal lngModul As Enum_Inside_Program, ByVal lng医嘱ID As Long, ByVal lng医护科室ID As Long, ByVal strPrivs As String, _
   Optional ByVal int场合 As Integer = 2, Optional ByVal lng病区ID As Long, Optional ByVal blnMoved As Boolean = False, Optional ByVal lngFontSize As Long = 9) As Boolean
'功能：刷新对应医嘱的血液信息
'frmParent 调用对象主窗体，主要是供界面刷新使用(该窗体要求具有一个timer控件，名称为timBRefresh)
' int场合 =1 门诊调用,2-住院调用, 当场合=2时，传入lng病区ID
    Dim lngCount As Long
    
    If mblnExecFresh = False Then  '执行过程中避免重复刷新，应为本窗体内部已经调用了刷新
        Set mfrmParent = frmParent
        mlngSys = lngSys
        mlngModul = lngModul
        mlng医嘱ID = lng医嘱ID
        mlng医护科室ID = lng医护科室ID
        mint场合 = int场合
        mlng病区ID = lng病区ID
        mstrPrivs = strPrivs
        mblnMoved = blnMoved
        mlngFontSize = lngFontSize
        mlng发送号 = 0
        
        If mint场合 = 2 Then
             '删除现在的工具栏及顶级菜单项
            For lngCount = cbsExec.ActiveMenuBar.Controls.Count To 1 Step -1
                cbsExec.ActiveMenuBar.Controls(lngCount).Delete
            Next
            For lngCount = cbsExec.Count To 2 Step -1
                cbsExec(lngCount).Delete
            Next
            Call InitExecBar
        End If
        Call RefreshData
        Call SetFontSize(mlngFontSize)
    End If
    mblnExecFresh = False
    zlRefresh = True
End Function

Private Sub InitTable()
'表格初始化
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsExec, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)  '收发ID
        Call .AppendColumn("状态", 810, flexAlignLeftCenter, flexDTString, , "血液状态") '接收执行状态
        Call .AppendColumn("血液名称", 1800, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("规格", 810, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("ABO", 810, flexAlignLeftCenter, flexDTString, , "ABO", True)
        Call .AppendColumn("Rh(D)", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
        Call .AppendColumn("血袋编号", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("效期", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "血液效期", True)
        Call .AppendColumn("数量", 500, flexAlignRightCenter, flexDTDecimal, , , , , , False)
        
        '执行记录
        Call .AppendColumn("核查者", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("复查者", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("核查时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
'        Call .AppendColumn("执行科室", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("开始执行人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("开始时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
'        Call .AppendColumn("前15分钟滴速", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("输注前输血反应", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("输注前反应时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
'        Call .AppendColumn("输注中执行人", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("后15分钟滴速", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("输注后输血反应", 1200, flexAlignLeftCenter, flexDTString)
'        Call .AppendColumn("输注后反应时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("结束执行人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("结束时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        
        If Mid(gstr医嘱核对, 1, 1) = "1" Then
            Call .AppendColumn("核对人", 1200, flexAlignLeftCenter, flexDTString)
            Call .AppendColumn("核对时间", 1500, flexAlignLeftCenter, flexDTString)
        Else
            Call .AppendColumn("核对人", 0, flexAlignLeftCenter, flexDTString)
            Call .AppendColumn("核对时间", 0, flexAlignLeftCenter, flexDTString)
        End If
        
        '血液配发信息
        Call .AppendColumn("配血方法", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("配血结论", 1500, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("发血科室", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("发血人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("取血人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("发血时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("接收人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("核收人", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("接收时间", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        
        
        '隐藏列
        Call .AppendColumn("血液ID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("血液效期颜色", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("接收状态", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("执行状态", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("待执行科室ID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("执行科室ID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        vsExec.FrozenCols = vsExec.ColIndex("规格")
        .AppendRows = False
    End With
End Sub

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
   
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    Set cbsExec.Icons = gobjCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_显示执行, "显示执行内容")
        Set objControl = .Add(xtpControlButton, ID_完成执行, "执行完成")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_Complete
        Set objControl = .Add(xtpControlButton, ID_取消完成, "取消完成")
            objControl.IconId = conMenu_Edit_Untread
        
        Set objControl = .Add(xtpControlButton, ID_执行前核对, "核查")
            objControl.IconId = conMenu_Manage_ThingAudit
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_取消执行前核对, "取消核查")
            objControl.IconId = conMenu_Manage_ThingDelAudit
            
        Set objControl = .Add(xtpControlButton, ID_执行记录, "记录执行情况")
            objControl.IconId = conMenu_Manage_ThingAdd
        Set objControl = .Add(xtpControlButton, ID_执行删除, "删除执行情况")
            objControl.IconId = conMenu_Manage_ThingDel

        Set objControl = .Add(xtpControlButton, ID_执行中核对, "核对")
            objControl.IconId = conMenu_Manage_ThingAudit
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_取消执行中核对, "取消核对")
            objControl.IconId = conMenu_Manage_ThingDelAudit
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsExec.KeyBindings
        '.Add FCONTROL, vbKeyH, 0
    End With
End Sub

Private Sub RefreshData(Optional ByVal blnRefreshBlood As Boolean = True)
    '功能:刷新对应医嘱对应的血液信息
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    Dim lngSelectRowID As Long
    Dim lng组ID As Long
    Dim arrData, arrTmp() As String, i As Integer
    Dim strID As String
    On Error GoTo ErrHand
    If mblnLoad = False Then Call InitTable
    strSQL = _
        " Select A.发送号,B.相关ID,A.发送时间,A.执行状态,A.记录性质,A.NO,A.执行部门ID,A.计费状态,A.门诊记帐,B.病人ID,B.主页ID ,C.执行时间,C.核对人,C.核对时间" & vbNewLine & _
        " From 病人医嘱执行 C,病人医嘱发送 A,病人医嘱记录 B" & vbNewLine & _
        " Where A.医嘱ID=C.医嘱ID(+) and a.发送号=c.发送号(+) and a.医嘱ID=b.ID And b.id=[1]"
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "已发血液信息提取", mlng医嘱ID)
    If rsData.EOF Then
        MsgBox "该医嘱还未发送，不能进行执行登记！", vbInformation, gstrSysName
        Exit Sub
    End If
    lng组ID = Val("" & rsData!相关ID)
    mlng发送号 = rsData!发送号
    mstr发送时间 = Format(rsData!发送时间 & "", "YYYY-MM-DD HH:mm:ss")
    mint执行状态 = Val("" & rsData!执行状态)
    mint记录性质 = Val("" & rsData!记录性质)
    mstrNO = "" & rsData!NO
    mlng执行部门ID = Val("" & rsData!执行部门ID)
    mint计费状态 = Val("" & rsData!计费状态)
    mint门诊记帐 = Val("" & rsData!门诊记帐)
    mlng病人ID = Val("" & rsData!病人id)
    mlng主页id = Val("" & rsData!主页id)
    mlng组ID = lng组ID
    
    arrData = Array()
    Do While Not rsData.EOF
        ReDim Preserve arrData(UBound(arrData) + 1)
        arrData(UBound(arrData)) = "" & rsData!执行时间 & "'" & rsData!核对人 & "'" & rsData!核对时间
    rsData.MoveNext
    Loop
    
    '刷新前确定之前选择的血液
    If vsExec.Row >= vsExec.FixedRows And vsExec.Row < vsExec.Rows Then
        lngSelectRowID = Val(vsExec.RowData(vsExec.Row))
    End If
    
    If blnRefreshBlood = False Then Exit Sub
    strSQL = _
        " Select b.申请id, a.Id, a.血液id, a.Abo, a.Rh, To_Char(a.效期, 'YYYY-MM-DD hh24:mi') 血液效期, a.颜色 血液颜色, a.外观 血袋外观, a.配血人," & vbNewLine & _
        " Decode(Zl_血液失效_Check(k.效期报警,k.效期单位,a.效期),0," & COLOR.原始单据 & ",1," & COLOR.深灰色 & ",2," & COLOR.红色 & ") 血液效期颜色," & vbNewLine & _
        "       To_Char(a.配血日期, 'YYYY-MM-DD hh24:mi') 配血时间, Nvl(a.发血状态, 0) 发血状态, c.名称 发血科室, a.血袋编号," & vbNewLine & _
        "       Decode(Nvl(h.执行状态, 0)," & vbNewLine & _
        "               0," & vbNewLine & _
        "               " & IIf(gbln接收后才能执行 = True, "Decode(Nvl(h.接收状态, 0), 0, '待接收', 2, '拒绝接收', '已接收'),", "'等待执行',") & vbNewLine & _
        "               1," & vbNewLine & _
        "               '正在执行'," & vbNewLine & _
        "               2," & vbNewLine & _
        "               '完成执行'," & vbNewLine & _
        "               3," & vbNewLine & _
        "               '停止执行') 血液状态, a.实际数量 As 数量, e.名称 As 血液名称, e.规格," & vbNewLine & _
        "       (Select f_List2str(Cast(Collect(g.名称) As t_Strlist))" & vbNewLine & _
        "         From 诊疗项目目录 g, 血液配血方法 f" & vbNewLine & _
        "         Where f.配血方法id = g.Id(+) And f.收发id = a.Id) 配血方法," & vbNewLine & _
        "       (Select Max(f.配血结论) From 诊疗项目目录 g, 血液配血方法 f Where f.配血方法id = g.Id(+) And f.收发id = a.Id) 配血结论, h.发送人 发血人, h.取血人," & vbNewLine & _
        "       To_Char(h.发送时间, 'YYYY-MM-DD hh24:mi') 发血时间, h.接收状态, h.接收人, h.接收时间, h.核收人, h.核收时间, h.执行状态, h.执行科室id 待执行科室id," & vbNewLine & _
        "       g.名称 待执行部门,h.执行核对人 核查者, h.执行复查人 复查者, h.执行核对时间 核查时间" & vbNewLine & _
        " From 部门表 c, 收费项目目录 e, 血液品种 k, 血液规格 l, 血液收发记录 a, 部门表 g, 血液发送记录 h, 血液配血记录 b" & vbNewLine & _
        " Where c.Id = a.库房id And e.Id = a.血液id And k.品种id = l.品种id And l.规格id = a.血液id And a.Id = h.收发id  And" & vbNewLine & _
        "      h.执行科室id = g.Id(+) And h.配发id = b.Id And b.申请id = [1]" & vbNewLine & _
        " Order By a.配血日期, a.序号"

    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "已发血液信息提取", lng组ID)
    Call mclsVsf.LoadGrid(rsData, "", True)
    strID = ""
    For lngRow = vsExec.FixedRows To vsExec.Rows - 1
        Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = Nothing
        If Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID"))) > 0 Then
            strID = strID & "," & Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID")))
            Select Case Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("执行状态")))
                '0-未执行;1-正在执行;2-完成执行;3-停止执行
                Case 0
                   Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("未执行").Picture
                Case 1
                    Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("正在执行").Picture
                Case 2
                    Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("已执行").Picture
                Case 3
                    Set vsExec.Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgList.ListImages("拒绝执行").Picture
            End Select
            vsExec.Cell(flexcpForeColor, lngRow, mclsVsf.ColIndex("血液效期"), lngRow, mclsVsf.ColIndex("血液效期")) = Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("血液效期颜色")))
            
            '定位到上次选择的行次
            If Val(vsExec.RowData(lngRow)) = lngSelectRowID And lngSelectRowID > 0 Then
                vsExec.Row = lngRow
            End If
        End If
    Next lngRow
    If Left(strID, 1) = "," Then strID = Mid(strID, 2)
    If strID <> "" Then
        '提取血液执行信息
        strSQL = "Select /*+ CARDINALITY(B 10)*/ 收发id, 记录性质, 序号, 执行时间, 执行人, 执行科室id, 滴速, 输血反应, 反应时间, 输血部位是否渗漏 是否渗漏, 是否使用药物, 体温, 脉搏, 呼吸, 收缩压, 舒张压, 摘要, 登记人, 登记时间," & vbNewLine & _
            "       签名人, 签名时间" & vbNewLine & _
            " From 血液执行记录 a,Table(f_str2list([1])) B" & vbNewLine & _
            " Where a.收发id = b.Column_Value" & vbNewLine & _
            " Order By 收发id, 记录性质, 序号"
        Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "已发血液信息提取", strID)
        For lngRow = vsExec.FixedRows To vsExec.Rows - 1
            If Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID"))) > 0 Then
                rsData.Filter = "收发ID=" & Val(vsExec.TextMatrix(lngRow, mclsVsf.ColIndex("ID")))
                rsData.Sort = "记录性质,序号"
                Do While Not rsData.EOF
                    Select Case Val("" & rsData!记录性质)
                        Case 1
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("开始执行人")) = rsData!执行人 & ""
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("开始时间")) = Format("" & rsData!执行时间, "YYYY-MM-DD HH:mm")
                        Case 3
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("结束执行人")) = rsData!执行人 & ""
                            vsExec.TextMatrix(lngRow, vsExec.ColIndex("结束时间")) = Format("" & rsData!执行时间, "YYYY-MM-DD HH:mm")
                    End Select
                rsData.MoveNext
                Loop
                '加载执行中核对信息
                For i = 0 To UBound(arrData)
                    arrTmp = Split(CStr(arrData(i)), "'")
                    If Format(vsExec.TextMatrix(lngRow, vsExec.ColIndex("开始时间")), "YYYY-MM-DD HH:mm:ss") = Format(arrTmp(0), "YYYY-MM-DD HH:mm:ss") Then
                        vsExec.TextMatrix(lngRow, vsExec.ColIndex("核对人")) = arrTmp(1)
                        vsExec.TextMatrix(lngRow, vsExec.ColIndex("核对时间")) = arrTmp(2)
                        Exit For
                    End If
                Next i
            End If
        Next lngRow
    End If
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckDataMoved() As Boolean
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        CheckDataMoved = True
    End If
End Function

Private Function CheckItemOk() As Boolean
'功能：检查所有的项目是否都已经执行完成
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngTmp As Long
    
    strSQL = "Select a.收发id" & vbNewLine & _
    " From 血液发送记录 a, 血液配血记录 b" & vbNewLine & _
    " Where a.配发id = b.Id And (Nvl(a.执行状态, 0) = 0 or Nvl(a.执行状态, 0) = 1) And b.申请id = [1] And Rownum < 2"
    On Err GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "病人审核检查", mlng组ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该医嘱还存在未执行的血液记录，不能完成执行。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Mid(gstr医嘱核对, 1, 1)) = 1 Then
        strSQL = "Select 核对人 From 病人医嘱执行 Where 医嘱id = [1] And 发送号 = [2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
        If rsTmp.RecordCount = 1 Then
            If rsTmp!核对人 & "" = "" Then
                MsgBox "该医嘱还存在未核对的执行登记，必须核对了才能完成。", vbInformation, gstrSysName
                Exit Function
            End If
        ElseIf rsTmp.RecordCount > 1 Then
            lngTmp = rsTmp.RecordCount
            rsTmp.Filter = "核对人<>''"
            If lngTmp <> rsTmp.RecordCount Then
                MsgBox "该医嘱还存在未核对的执行登记，必须全部核对了才能完成。", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "当前医嘱还未记录执行情况，必须记录执行情况后核对了才能完成。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckItemOk = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckPatiIsAduit() As Boolean
'功能：检查病人是否开始审核
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim int审核标志 As Integer
    
    If mlng主页id = 0 Then CheckPatiIsAduit = True: Exit Function
    strSQL = "Select a.审核标志 From 病案主页 a" & _
                " Where a.病人ID=[1] And a.主页ID=[2]"
    On Err GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "病人审核检查", mlng病人ID, mlng主页id)
    If rsTmp.RecordCount > 0 Then
        If Val("" & rsTmp!审核标志) >= 1 And gbyt病人审核方式 = 1 Then
            MsgBox "该病人的费用正在审核或已经审核，不允许操作医嘱和费用。", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPatiIsAduit = True
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetShowExec(ByVal blnShow As Boolean)
    vsExec.Visible = blnShow
    If blnShow = True Then
        vsExec.Tag = "可见"
        Me.Height = vsExec.Top + vsExec.Height
    Else
        vsExec.Tag = ""
        Me.Height = vsExec.Top
    End If
End Sub

Public Sub SetFontSize(ByVal lngFontSize As Long)
'功能:设置医嘱清单的字体大小
    Dim bytSize As Byte
    bytSize = IIf(lngFontSize = 9, 0, 1)
    Call SetPublicFontSize(Me, bytSize)
End Sub

Private Function FuncExec(ByVal intExecId As CMD_EXEC) As Boolean
    Dim strSQL As String, rsTmp As New Recordset
    Dim byt来源 As Integer, blnIsAbnormal As Boolean
    Dim curMoney As Currency, str类别 As String, str类别名 As String
    Dim lngID As Long, str执行时间 As String, str核对人 As String
    Dim blnTrans As Boolean, blnOk As Boolean
    Dim i As Integer
    Dim arrSQL As Variant
    Dim blnFinish As Boolean
    '核对结果
    Dim strCheckOper As String, strCheckTime As String, strCheckResult As String
    On Error GoTo ErrHand
    Select Case intExecId
        Case ID_执行前核对, conMenu_Manage_ThingAudit * 100# + 1
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) <> "" Then
                MsgBox "该袋血液已经核对，不允许再次核对。", vbInformation, gstrSysName
                Exit Function
            End If
            If gbln接收后才能执行 = True Then
                If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("接收状态"))) <> 1 And Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("接收状态"))) <> 3 Then
                    MsgBox "该袋血液还未接收，不允许核对。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            str执行时间 = Format(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("接收时间")), "YYYY-MM-DD HH:mm")
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            blnOk = frmUserCheck.ShowMe(Me, mlngModul, mlng医护科室ID, mlng医护科室ID, str执行时间, "", True, 执行核对)
            If blnOk = True Then
                strCheckOper = frmUserCheck.SendAndTakeOper
                strCheckTime = frmUserCheck.SendTime
                strCheckResult = frmUserCheck.CheckResult
                strSQL = "Zl_血液执行记录_Check(" & lngID & ",'" & Split(strCheckOper, "'")(0) & "','" & Split(strCheckOper, "'")(1) & "',To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strCheckResult & "')"
                Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Else
                Exit Function
            End If
        Case ID_取消执行前核对, conMenu_Manage_ThingDelAudit * 100# + 1
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) = "" Then
                MsgBox "该袋血液还未核对，不能取消核对。", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("执行状态"))) <> 0 Then
                MsgBox "该袋血液已经开始执行，不能取消核对。", vbInformation, gstrSysName
                Exit Function
            End If
            strCheckOper = ""
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) <> UserInfo.姓名 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("复查者")) <> UserInfo.姓名 Then
                strCheckOper = gobjDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", mlngSys, mlngModul, "执行情况登记", , True)
                If strCheckOper = "" Then Exit Function
                If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) <> strCheckOper And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("复查者")) <> strCheckOper Then
                    MsgBox "只能取消自己核对或复查的血液，当前血液核对人是""" & vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) & """" & "复查人是""" & vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("复查者")) & """", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If MsgBox("你确定要取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            End If
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            strSQL = "Zl_血液执行记录_Uncheck(" & lngID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Case ID_执行中核对, conMenu_Manage_ThingAudit '执行中核对
            If Not Mid(gstr医嘱核对, 1, 1) = "1" Then
                MsgBox "不能核对输血医嘱，请在基础参数中勾选需核对输血医嘱参数。", vbInformation, gstrSysName
                Exit Function
            End If
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) <> "" Then
                MsgBox "该袋血液已经核对，不能再次核对。", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("执行状态"))) = 0 Or Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("执行状态"))) = 1 Then
                MsgBox "该袋血液还未完成执行情况登记，不能核对。", vbInformation, gstrSysName
                Exit Function
            End If
            str核对人 = gobjDatabase.UserIdentifyByUser(Me, "在核对执行情况前，请您先输入用户名和密码进行身份验证。", 100, IIf(mint场合 = 1, p医技工作站, p住院医嘱发送), "执行情况登记", , True)
            If str核对人 = "" Then Exit Function
            
            If str核对人 = vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("开始执行人")) Then
                MsgBox "执行人不能和审核人相同，不能核对。", vbInformation, gstrSysName
                Exit Function
            End If
            
            str执行时间 = Format(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("开始时间")), "yyyy-MM-dd HH:mm:ss")
            strSQL = "Zl_病人医嘱核对_Insert(" & mlng医嘱ID & "," & mlng发送号 & ",'" & str核对人 & "',To_Date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS'))"
            Call gobjDatabase.ExecuteProcedure(strSQL, "医嘱核对")
            Call SetExecState(E_执行核对)
        Case ID_取消执行中核对, conMenu_Manage_ThingDelAudit '取消执行中核对
            If Not Mid(gstr医嘱核对, 1, 1) = "1" Then
                MsgBox "不能取消输血医嘱核对，请在基础参数中勾选需核对输血医嘱参数。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) = "" Then
                MsgBox "该袋血液还未核对，不能取消核对。", vbInformation, gstrSysName
                Exit Function
            End If
            If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) <> UserInfo.姓名 Then
                str核对人 = gobjDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", 100, IIf(mint场合 = 1, p医技工作站, p住院医嘱发送), "执行情况登记", , True)
                If str核对人 = "" Then Exit Function
                If str核对人 <> vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) Then
                    MsgBox "只能取消自己核对的血液执行，当前血液核对人是""" & vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) & """", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If MsgBox("你确定要取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            End If
            strSQL = "Zl_病人医嘱核对_Delete(" & mlng医嘱ID & "," & mlng发送号 & ",To_Date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS'))"
            Call gobjDatabase.ExecuteProcedure(strSQL, "取消医嘱核对")
            Call SetExecState(E_取消核对)
        Case ID_执行记录, conMenu_Manage_ThingAdd
            If mint执行状态 = 1 Then
                MsgBox "该医嘱当前已经执行完成，不能再继续操作。", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(Mid(gstr医嘱核对, 1, 1)) > 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) <> "" Then
                MsgBox "该袋血液已经核对，请取消核对后再试。", vbInformation, gstrSysName
                Exit Function
            End If
            If CheckDataMoved Then Exit Function
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            blnFinish = False
            If frmBloodExecEdit.ShowEdit(Me, mlngModul, mlng医嘱ID, mlng发送号, mlng医护科室ID, lngID, mlng执行部门ID, mstrPrivs, , blnFinish) = False Then
                Exit Function
            End If
            If blnFinish = True Then
                Call SetExecState(E_执行完成)
            Else
                Call SetExecState(E_记录执行)
            End If
        Case ID_执行删除, conMenu_Manage_ThingDel
            If mint执行状态 = 1 Then
                MsgBox "该医嘱当前已经执行完成，不能再继续操作。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Val(Mid(gstr医嘱核对, 1, 1)) > 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) <> "" Then
                MsgBox "该袋血液已经核对，请取消核对后再试。", vbInformation, gstrSysName
                Exit Function
            End If
        
            If CheckDataMoved Then Exit Function
            If MsgBox("确实要删除该袋血液的执行情况吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            lngID = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID")))
            str执行时间 = vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("开始时间"))
            '是否存在该血袋消息，存在则设为已读
            If mint场合 = 1 Then
                strSQL = "select a.id,a.类型编码,a.就诊id,a.业务标识 from 业务消息清单 a,病人医嘱记录 b,病人挂号记录 c where a.病人id = [1] and a.就诊id = c.id and c.no = b.挂号单 and b.id = [2]" & vbNewLine & _
                        "and a.是否已阅 = 0 and a.类型编码 in ('ZLHIS_BLOOD_006','ZLHIS_BLOOD_007')  "
            ElseIf mint场合 = 2 Then
                strSQL = "select a.id,a.类型编码,a.就诊id,a.业务标识 from 业务消息清单 a,病人医嘱记录 b where a.病人id = [1] and a.就诊id = b.主页id and b.id = [2]" & vbNewLine & _
                        "and a.是否已阅 = 0 and a.类型编码 in ('ZLHIS_BLOOD_006','ZLHIS_BLOOD_007')  "
            End If
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "血袋相关消息", mlng病人ID, mlng医嘱ID)
            arrSQL = Array()
            For i = 0 To 1
                '将ZLHIS_BLOOD_006的消息设为已读
                If i = 0 Then rsTmp.Filter = "业务标识 = '" & mlng医嘱ID & ":" & Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("id"))) & "'"
                '将ZLHIS_BLOOD_007的消息设为已读
                If i = 1 Then rsTmp.Filter = "业务标识 = '" & mlng组ID & ":" & mlng医嘱ID & ":" & Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("id"))) & "'"
                If Not rsTmp.EOF Then
                    rsTmp.MoveFirst
                    Do While Not rsTmp.EOF
                        strSQL = "Zl_业务消息清单_Read(" & mlng病人ID & "," & rsTmp!就诊id & ",'" & rsTmp!类型编码 & "',"
                        strSQL = strSQL & IIf(mint场合 = 1, 4, 3) & ",'" & UserInfo.姓名 & "'," & mlng病区ID & ",NULL,"
                        strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                        rsTmp.MoveNext
                    Loop
                End If
            Next
            gcnOracle.BeginTrans
            blnTrans = True
            strSQL = "ZL_病人医嘱执行_Delete(" & mlng医嘱ID & "," & mlng发送号 & ",To_Date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS'),0,0," & mlng病区ID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            strSQL = "Zl_血液执行记录_Delete(" & lngID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            For i = 0 To UBound(arrSQL)
                Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans
            blnTrans = False
            Call SetExecState(E_删除执行)
        Case ID_完成执行, conMenu_Manage_Complete
            If mint执行状态 = 1 Then
                MsgBox "该医嘱当前已经执行完成，不能重复完成。", vbInformation, gstrSysName
                Exit Function
            End If
            If CheckDataMoved Then Exit Function
            '检查所有的项目是否都已经完成执行，启用了核对是否已经核对
            If Not CheckItemOk Then Exit Function
            '检查病人是否正在审核
            If Not CheckPatiIsAduit Then Exit Function
            
            '是否允许完成未收费病人的项目:不管记帐划价,因为要执行后审核,临嘱才可能发送到门诊收费
            If mint记录性质 = 1 And mint计费状态 > 0 Then
                If Not ItemHaveCash(2, True, mlng医嘱ID, mlng组ID, mlng发送号, "E", mstrNO, 1, 0, 0, mblnMoved, CDate(mstr发送时间), "", "", blnIsAbnormal) Then
                    If blnIsAbnormal Then
                        MsgBox "该病人还存在异常费用，请检查。", vbInformation, gstrSysName
                    Else
                        MsgBox "该病人还存在未收费的费用，请检查。", vbInformation, gstrSysName
                    End If
                    Exit Function
                End If
            End If
            If mint记录性质 = 2 Then
                curMoney = GetAdviceMoney(IIf(mlng组ID = 0, mlng医嘱ID, mlng组ID), mlng医嘱ID, mlng发送号, str类别, str类别名, True, IIf(mint门诊记帐 = 0, 2, 1))
                If curMoney > 0 Then
                    '住院出院病人费用控制
                    If Not PatiCanBilling(mlng病人ID, mlng主页id, GetInsidePrivs(100, p住院医嘱发送), p住院医嘱发送) Then Exit Function
                    '记帐报警
                    If InitObjPublicExpense(mlngSys) Then
                        If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, p住院医嘱发送, "", mstrNO, GetInsidePrivs(mlngSys, p住院医嘱发送), mlng病区ID) = False Then Exit Function
                    End If
                    
                    '门诊一卡通消费身份验证,只检查门诊记帐费用
                    If mint门诊记帐 = 1 Then
                        If InitObjPublicExpense(mlngSys) Then
                            If Not gobjPublicExpense.zlPatiIdentify(mlngModul, Me, mlng病人ID, curMoney) Then Exit Function
                        End If
                    End If
                End If
            End If
            
            If MsgBox("确认要将该医嘱执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            strSQL = "ZL_病人医嘱执行_Finish(" & mlng医嘱ID & "," & mlng发送号 & ",Null,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Call SetExecState(E_执行完成)
        Case ID_取消完成, conMenu_Manage_Undone
            If mint执行状态 <> 1 Then
                MsgBox "该医嘱当前不处于已执行状态，不能取消执行。", vbInformation, gstrSysName
                Exit Function
            End If
            If CheckDataMoved Then Exit Function
            '检查病人是否正在审核
            If Not CheckPatiIsAduit Then Exit Function
            
            If mint记录性质 <> 1 Then
                If mint门诊记帐 = 0 Then
                    byt来源 = 2
                Else
                    byt来源 = 1
                End If
                '费用结帐判断
                If Not ItemCanCancel(mlng医嘱ID, mlng发送号, mlng组ID, "E", True, mblnMoved, byt来源) Then Exit Function
            End If
            
            If MsgBox("确实要将该医嘱取消执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            strSQL = "ZL_病人医嘱执行_Cancel(" & mlng医嘱ID & "," & mlng发送号 & "," & "Null,0," & mlng病区ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            Call SetExecState(E_取消完成)
    End Select
    FuncExec = True
    Exit Function
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case ID_显示执行
            mblnShowExec = Not mblnShowExec
            Call SetShowExec(mblnShowExec)
            RaiseEvent ShowExec(mblnShowExec, Me.Height)
        Case ID_完成执行, ID_取消完成, conMenu_Manage_Complete, conMenu_Manage_Undone
            If FuncExec(Control.id) = True Then Call RefreshData(False)
        Case ID_执行记录, ID_执行删除, ID_执行前核对, ID_取消执行前核对, ID_执行中核对, ID_取消执行中核对, conMenu_Manage_ThingAdd, conMenu_Manage_ThingDel, _
                conMenu_Manage_ThingAudit, conMenu_Manage_ThingDelAudit, conMenu_Manage_ThingAudit * 100# + 1, conMenu_Manage_ThingDelAudit * 100# + 1
            If FuncExec(Control.id) = True Then Call RefreshData
    End Select
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Dim int血液执行状态 As Integer, int接收状态 As Integer, lng待执行科室 As Long, lng执行科室 As Long
    Dim bln允许执行 As Boolean
    
    blnEnable = False
    If vsExec.Row >= vsExec.FixedRows And mblnLoad = True Then
        blnEnable = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID"))) > 0
        If blnEnable = True Then
            int血液执行状态 = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("执行状态")))
            int接收状态 = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("接收状态")))
            lng待执行科室 = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("待执行科室ID")))
            lng执行科室 = Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("执行科室ID")))
        End If
    End If
    
    bln允许执行 = True
'    If int血液执行状态 > 0 Then '如果已经执行,则执行科室必须是当前执行科室
'        bln允许执行 = (lng执行科室 = mlng医护科室ID)
'    Else
'        '新增执行时，必须是接收时的执行科室
'        bln允许执行 = (lng待执行科室 = mlng医护科室ID) Or InStr(mstrPrivs, "执行他科项目")
'    End If
    
    Select Case Control.id
        Case ID_显示执行
            Control.Checked = mblnShowExec And mint场合 = 2 And Control.Visible
        Case ID_完成执行, conMenu_Manage_Complete
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "确认执行完成") = 0)
            Control.Enabled = blnEnable And (mint执行状态 = 0 Or mint执行状态 = 3) And Control.Visible
        Case ID_取消完成, conMenu_Manage_Undone
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "取消执行完成") = 0)
            Control.Enabled = blnEnable And mint执行状态 = 1 And Control.Visible
        Case ID_执行记录, conMenu_Manage_ThingAdd
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "执行情况登记") = 0)
            Control.Enabled = mblnShowExec And blnEnable And (mint执行状态 = 0 Or mint执行状态 = 3) And Control.Visible And bln允许执行
        Case ID_执行删除, conMenu_Manage_ThingDel
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "执行情况登记") = 0)
            Control.Enabled = mblnShowExec And blnEnable And (mint执行状态 = 0 Or mint执行状态 = 3) And (int血液执行状态 = 1 Or int血液执行状态 = 2) And Control.Visible And bln允许执行
        Case ID_执行前核对, conMenu_Manage_ThingAudit * 100# + 1 '执行前核查
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "执行情况登记") = 0)
            Control.Enabled = mblnShowExec And blnEnable And Control.Visible And int血液执行状态 = 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) = "" '已经接收的为核对血液才能核对
            If Control.Enabled And gbln接收后才能执行 = True Then
                Control.Enabled = (int接收状态 = 1 Or int接收状态 = 3)
            End If
        Case ID_取消执行前核对, conMenu_Manage_ThingDelAudit * 100# + 1 '取消执行前核查
            Control.Visible = Not (InStr(GetInsidePrivs(mlngSys, mlngModul), "执行情况登记") = 0)
            Control.Enabled = mblnShowExec And blnEnable And int血液执行状态 = 0 And vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核查者")) <> "" And Control.Visible
        Case ID_执行中核对, conMenu_Manage_ThingAudit, ID_取消执行中核对, conMenu_Manage_ThingDelAudit  '执行中核对'取消执行中核对
             If InStr(GetInsidePrivs(mlngSys, mlngModul), "执行情况登记") = 0 Or Val(Mid(gstr医嘱核对, 1, 1)) = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = blnEnable And (mint执行状态 = 0 Or mint执行状态 = 3) And IIf(mint场合 = 2, mblnShowExec, True)
                If (mint执行状态 = 0 Or mint执行状态 = 3) And IIf(mint场合 = 2, mblnShowExec, True) Then
                    If vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("核对人")) = "" Then
                        If Control.id = ID_执行中核对 Or Control.id = conMenu_Manage_ThingAudit Then Control.Enabled = True
                        If Control.id = ID_取消执行中核对 Or Control.id = conMenu_Manage_ThingDelAudit Then Control.Enabled = False
                    Else
                        If Control.id = ID_执行中核对 Or Control.id = conMenu_Manage_ThingAudit Then Control.Enabled = False
                        If Control.id = ID_取消执行中核对 Or Control.id = conMenu_Manage_ThingDelAudit Then Control.Enabled = True
                    End If
                End If
            End If
        Case conMenu_Manage_ThingModi '调整执行情况
            Control.Visible = False
            Control.Enabled = Control.Visible
    End Select
End Sub

Private Sub cbsExec_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    
    On Error Resume Next
    Call cbsExec.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If mint场合 = 1 Then lngTop = 0
    If mblnShowExec = True Then
        With vsExec
            .Left = lngLeft
            .Top = lngTop
            .Width = lngRight - lngLeft
            .Height = lngBottom - lngTop
        End With
    Else
        vsExec.Left = lngLeft
        vsExec.Top = lngTop
    End If
End Sub

Private Sub Form_Load()
    mblnShowExec = False
    mblnExecFresh = False
    mintAdviceExecState = 0
    mblnLoad = False
    Call InitTable
    mblnLoad = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf = Nothing
End Sub

Public Function zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsExec_Update(Control)
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsExec_Execute(Control)
End Sub

Private Sub vsExec_DblClick()
    Dim lngID As Long
    Dim objButton As CommandBarControl
    Dim blnReadOnly As Boolean
    
    If Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID"))) < 0 Or mblnShowExec = False Then Exit Sub
    If Not (Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("执行状态"))) > 0) Then Exit Sub
    If CheckDataMoved Then Exit Sub
    Call frmBloodExecEdit.ShowEdit(Me, mlngModul, mlng医嘱ID, mlng发送号, mlng医护科室ID, Val(vsExec.TextMatrix(vsExec.Row, vsExec.ColIndex("ID"))), mlng执行部门ID, mstrPrivs, True)
End Sub

Private Sub vsExec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
        
    If Button = 2 And mint场合 = 1 Then
        Set objPopup = cbsExec.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, ID_执行前核对, "核查"
            .Add xtpControlButton, ID_取消执行前核对, "取消核查"
            .Add xtpControlButton, ID_执行记录, "记录执行情况"
            .Add xtpControlButton, ID_执行删除, "删除执行情况"
            .Add xtpControlButton, ID_执行中核对, "核对"
            .Add xtpControlButton, ID_取消执行中核对, "取消核对"
        End With
        
        vsExec.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub SetExecState(ByVal intExecState As Integer)
    If Not mfrmParent Is Nothing Then
        On Error Resume Next
        mfrmParent.timBRefresh.Enabled = True
        If Err <> 0 Then Err.Clear
    End If
    mintAdviceExecState = intExecState
End Sub
