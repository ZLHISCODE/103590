VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareBrushManager 
   Caption         =   "结算卡刷卡管理"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmSquareBrushManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11745
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8025
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareBrushManager.frx":74F2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2715
      Left            =   105
      TabIndex        =   0
      Top             =   825
      Width           =   9885
      _cx             =   17436
      _cy             =   4789
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      Rows            =   10
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   350
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSquareBrushManager.frx":7D86
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   120
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
      ExplorerBar     =   7
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmSquareBrushManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintSucces As Integer
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private Const mconMenu_Edit_Affirm = 225
Private mrsBrushData As ADODB.Recordset
Private mrsFeeList As ADODB.Recordset
Private mdbl最大消费额 As Double
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mbytCall As Byte  '调用类型 0-  门诊费用调用 1-  住院结帐调用,3-其他
Private mstrTitle As String '用于窗体个性化保存的窗体名
Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的关联性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-31 10:45:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    '检查是否启用了相关的刷卡程序
    Set mobjBrushCard = New clsBrushSequareCard
    CheckDepend = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlShowBrushCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal rsFeeList As ADODB.Recordset, dbl最大消费额 As Double, rsRequare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷卡接口
    '入参:frmMain-调用的主窗体
    '     lngModule-调用的模块号
    '     strPrivs-调用的权限串
    '     dbl最大消费额-本次刷卡的最大刷卡额
    '     rsFeeList-费用详细信息()
    '出参:rsRequare-返回结算信息
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 10:33:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mdbl最大消费额 = dbl最大消费额
    Set mrsFeeList = rsFeeList  '费用明细
    If CheckDepend = False Then Exit Function
    Select Case mlngModule
    Case 1121 '  1121,'病人收费管理
        mbytCall = 0
    Case 1137  '病人结帐处理
        mbytCall = 1
    Case Else
        mbytCall = 3
    End Select
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    Set rsRequare = mrsBrushData
    zlShowBrushCard = mintSucces > 0
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件
    '编制:刘兴洪
    '日期:2009-12-23 10:57:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("卡号")) = "1|0"
        .ColData(.ColIndex("本次消费")) = "1|1"
        .Clear 1
        .Rows = 2
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(.ColIndex("结算卡类型")) = True
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 10:02:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup, rsTemp As ADODB.Recordset
    
      
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
       Set .Font = vsGrid.Font
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    cbsThis.ActiveMenuBar.Visible = False
        
  
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MoveCard, "移出刷卡记录"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_Affirm, "确定   "): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出  "): mcbrControl.BeginGroup = True
        mcbrControl.Flags = xtpFlagRightAlign
    End With
    
    
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    Set mcbrToolBar = cbsThis.Add("结算卡", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    mcbrToolBar.ShowTextBelowIcons = True
    With mcbrToolBar.Controls
        Set rsTemp = zlGet消费卡接口
        rsTemp.Sort = "自制卡,编号"
        Do While Not rsTemp.EOF
            Set mcbrControl = .Add(xtpControlButton, conMenu_Square_BrushCard + Val(rsTemp!编号), Nvl(rsTemp!名称)): mcbrControl.BeginGroup = True
            mcbrControl.IconId = 3816 ' conMenu_Square_BrushCard
            mcbrControl.Parameter = Val(rsTemp!编号)
 
            rsTemp.MoveNext
        Loop
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FALT, Asc("O"), mconMenu_Edit_Affirm
        .Add FALT, Asc("X"), conMenu_Edit_CardModify
 
         If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
         Do While Not rsTemp.EOF
            .Add FCONTROL, Asc(Trim(CStr(Chr(Val(rsTemp!编号) + 64)))), conMenu_Square_BrushCard + Val(rsTemp!编号)
            rsTemp.MoveNext
         Loop
     End With
         
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Err = 0: On Error Resume Next
    cbsThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    With vsGrid
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - stbThis.Height
    End With
End Sub

Private Sub Form_Load()
    mstrTitle = "结算卡刷卡管理"
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlDefCommandBars
    Call InitVsGrid
    Call zlInitBrushCardRec(mrsBrushData)
    Call vsGrid_GotFocus
End Sub
Private Function zlDeleteBrushCard(ByVal lng接口编号 As Long, Optional strCardNo As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除刷卡数据
    '入参:lng接口编号-接口编号
    '     strCardNo-卡号
    '出参:
    '返回:成功,返回ture,否则返回False
    '编制:刘兴洪
    '日期:2009-12-31 11:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    If mrsBrushData Is Nothing Then Exit Function
    If mrsBrushData.State <> 1 Then Exit Function
    If strCardNo = "" Then
        mrsBrushData.Filter = "接口编号=" & lng接口编号
    Else
        mrsBrushData.Filter = "接口编号=" & lng接口编号 & " and 卡号='" & strCardNo & "'"
    End If
    If mrsBrushData.EOF = False Then
        mrsBrushData.Delete (adAffectGroup)
    End If
    mrsBrushData.Filter = 0
    zlDeleteBrushCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlGet最大消费额(ByVal lng接口编号 As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口的最大消费额
    '入参:lng接口编号-接口编号
    '返回:
    '编制:刘兴洪
    '日期:2009-12-31 11:40:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl本次消费合计 As Double
    dbl本次消费合计 = 0
    Err = 0: On Error GoTo Errhand:
    If mrsBrushData Is Nothing Then GoTo ToCalc:
    If mrsBrushData.State <> 1 Then GoTo ToCalc:
    With mrsBrushData
        .Filter = "接口编号<>" & lng接口编号
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl本次消费合计 = dbl本次消费合计 + Val(Nvl(!结算金额))
            .MoveNext
        Loop
        .Filter = "接口编号=" & lng接口编号
        
    End With
ToCalc:
    zlGet最大消费额 = mdbl最大消费额 - dbl本次消费合计
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
'执行具体功能
Private Function zlExecuteBrushCard(ByVal lng接口编号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行具体功能
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-23 10:22:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, rsSquare As ADODB.Recordset, dbl本次消费合计 As Double, dbl最大消费额 As Double
    Err = 0: On Error GoTo Errhand:
    Set rsTemp = zlGet消费卡接口
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    rsTemp.Find "编号=" & lng接口编号, , , 1
    If rsTemp.EOF Then Exit Function
    If Val(Nvl(rsTemp!自制卡)) <> 1 Then
        '执行相关的三方接口
        If mobjBrushCard.zlInitInterFacel(lng接口编号) = False Then Exit Function
        
        '可通需要用到当前选择刷卡信息,因此，将其内容传入,以便查看(但不影响已经刷了的数据)
        Set rsSquare = zlDatabase.CopyNewRec(mrsBrushData)
        If mobjBrushCard.zlBrushCardSquare(mbytCall, Me, lng接口编号, mrsFeeList, zlGet最大消费额(lng接口编号), rsSquare) = False Then Exit Function
        dbl本次消费合计 = 0
        If Not rsSquare Is Nothing Then
            If rsSquare.State = 1 Then
                '需要检查是否超过了最大消费额
                rsSquare.Filter = "接口编号=" & lng接口编号
                If rsSquare.RecordCount <> 0 Then rsSquare.MoveFirst
                Do While Not rsSquare.EOF
                    dbl本次消费合计 = dbl本次消费合计 + Val(Nvl(rsSquare!结算金额))
                    If Val(Nvl(rsSquare!余额)) < Val(Nvl(rsSquare!结算金额)) Then
                        ShowMsgbox "注意:" & _
                                   "    " & rsTemp!名称 & " 的卡号为:" & Nvl(rsSquare!卡号) & "的余额(" & Format(Val(Nvl(rsSquare!余额)), gVbFmtString.FM_金额) & ")不足以支付刷卡金额(" & Format(Val(Nvl(rsSquare!结算金额)), gVbFmtString.FM_金额) & ")，请检查!"
                        
                        Exit Function
                    End If
                    rsSquare.MoveNext
                Loop
                If dbl本次消费合计 > dbl最大消费额 Then
                    ShowMsgbox "注意:" & vbCrLf & "    本次刷卡消费最大只能刷" & Format(dbl最大消费额, gVbFmtString.FM_金额) & "元,但你现在刷了" & Format(dbl本次消费合计, gVbFmtString.FM_金额) & "元,请检查!"
                    Exit Function
                End If
                '需要将rsSquare中的数据，更新到已经刷卡的数据中
                '删除数据;
                 Call zlDeleteBrushCard(lng接口编号, "")
                If rsSquare.RecordCount <> 0 Then rsSquare.MoveFirst
                Do While Not rsSquare.EOF
                    With mrsBrushData
                        .AddNew
                         !接口编号 = rsSquare!接口编号
                         !消费卡ID = rsSquare!消费卡ID
                         !卡号 = rsSquare!卡号
                         !结算方式 = rsTemp!结算方式
                         !卡名称 = rsSquare!卡名称
                         !余额 = rsSquare!余额
                         !结算金额 = rsSquare!结算金额
                         !交易时间 = rsSquare!交易时间
                         !备注 = rsSquare!备注
                         !结算标志 = 0
                         .Update
                    End With
                    rsSquare.MoveNext
                Loop
            Else
                Call zlDeleteBrushCard(lng接口编号, "")
            End If
        Else
            Call zlDeleteBrushCard(lng接口编号, "")
        End If
        GoTo BrushData:
    End If
    mrsBrushData.Filter = "接口编号=" & lng接口编号
     
    '自制卡,需要调用相关的刷卡界面
    If frmSquareBrushCard.zlShowBrushCard(Me, lng接口编号, mbytCall, mrsFeeList, zlGet最大消费额(lng接口编号), mrsBrushData) = False Then Exit Function
BrushData:
    Dim strCardNo As String
    lng接口编号 = 0
    With vsGrid
        If .Row > 0 Then
            lng接口编号 = Val(.Cell(flexcpData, .Row, .ColIndex("结算卡类型")))
            strCardNo = Trim(.Cell(flexcpData, .Rows - 1, .ColIndex("卡号")))
        End If
    End With
    Call FullDataToGrid(lng接口编号, strCardNo)
    zlExecuteBrushCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlMoveCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:移出当前刷卡的卡片信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-23 11:15:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCurRow  As Long
    Err = 0: On Error GoTo Errhand:
    
    With vsGrid
        If .Row < 0 Then Exit Function
        If .Rows < 2 Then Exit Function
        If Trim(.Cell(flexcpData, .Row, .ColIndex("卡号"))) <> "" Then
            '先找卡号
            mrsBrushData.Filter = "接口编号=" & Val(.Cell(flexcpData, .Row, .ColIndex("结算卡类型"))) & " and 卡号='" & Trim(.Cell(flexcpData, .Row, .ColIndex("卡号"))) & "'"
            If mrsBrushData.EOF = False Then
                mrsBrushData.Delete adAffectCurrent
                mrsBrushData.MoveNext
            End If
            mrsBrushData.Filter = 0
        End If
        lngCurRow = .Row
        Call FullDataToGrid
        If lngCurRow < .Rows - 1 Then
            lngCurRow = lngCurRow + 1
        Else
            lngCurRow = .Rows - 1
        End If
        If lngCurRow < 1 Then lngCurRow = 1
        If lngCurRow > 1 Then .Row = lngCurRow
    End With
    zlMoveCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub FullDataToGrid(Optional lngDefault接口编号 As Long = 0, Optional strDefaultCardNo As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新填数
    '入参:lngDefault接口编号-缺省指向的接口序号
    '     strDefaultCardNo-缺省指向的卡号
    '编制:刘兴洪
    '日期:2009-12-23 11:42:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng接口编号 As Long, lngRow As Long, dbl本次刷卡 As Double, dbl本次刷卡总计 As Double
    With vsGrid
        .Clear 1: .Rows = 2
        .TextMatrix(1, .ColIndex("结算卡类型")) = ""
        mrsBrushData.Filter = 0
        mrsBrushData.Sort = "接口编号,卡号"
        lngRow = 1: dbl本次刷卡 = 0: dbl本次刷卡总计 = 0
        If mrsBrushData.RecordCount <> 0 Then mrsBrushData.MoveFirst
        Do While Not mrsBrushData.EOF
            If lng接口编号 <> Val(Nvl(mrsBrushData!接口编号)) Then
                If lng接口编号 <> 0 Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("结算卡类型")) = .TextMatrix(.Rows - 2, .ColIndex("结算卡类型"))
                    .Cell(flexcpData, .Rows - 1, .ColIndex("结算卡类型")) = .Cell(flexcpData, .Rows - 2, .ColIndex("结算卡类型"))
                    .TextMatrix(.Rows - 1, .ColIndex("卡号")) = "小计"
                    .TextMatrix(.Rows - 1, .ColIndex("卡余额")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = ""
                    .TextMatrix(.Rows - 1, .ColIndex("本次消费")) = Format(dbl本次刷卡, gVbFmtString.FM_金额)
                    If lngDefault接口编号 = lngDefault接口编号 And "小计" = strDefaultCardNo Then
                        .Row = .Rows - 1
                    End If
                End If
                dbl本次刷卡 = 0
                lng接口编号 = Val(Nvl(mrsBrushData!接口编号))
            End If
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("结算卡类型"))) <> "" Then
                .Rows = .Rows + 1
            End If
            .TextMatrix(.Rows - 1, .ColIndex("结算卡类型")) = Nvl(mrsBrushData!卡名称)
            .Cell(flexcpData, .Rows - 1, .ColIndex("结算卡类型")) = Nvl(mrsBrushData!接口编号)
            .TextMatrix(.Rows - 1, .ColIndex("卡号")) = IIf(zlIsCardNoShowPW(Val(Nvl(mrsBrushData!接口编号))), "****", Nvl(mrsBrushData!卡号))
            .Cell(flexcpData, .Rows - 1, .ColIndex("卡号")) = Nvl(mrsBrushData!卡号)
            .TextMatrix(.Rows - 1, .ColIndex("卡余额")) = Format(Val(Nvl(mrsBrushData!余额)), gVbFmtString.FM_金额)
            .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = Nvl(mrsBrushData!结算方式)
            .TextMatrix(.Rows - 1, .ColIndex("本次消费")) = Format(Val(Nvl(mrsBrushData!结算金额)), gVbFmtString.FM_金额)
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = Nvl(mrsBrushData!备注)
            If lngDefault接口编号 = lngDefault接口编号 And Nvl(mrsBrushData!卡号) = strDefaultCardNo Then
                .Row = .Rows - 1
            End If
            dbl本次刷卡 = dbl本次刷卡 + Val(Nvl(mrsBrushData!结算金额))
            dbl本次刷卡总计 = dbl本次刷卡总计 + Val(Nvl(mrsBrushData!结算金额))
            mrsBrushData.MoveNext
        Loop
        If mrsBrushData.RecordCount <> 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("结算卡类型")) = .TextMatrix(.Rows - 2, .ColIndex("结算卡类型"))
            .Cell(flexcpData, .Rows - 1, .ColIndex("结算卡类型")) = .Cell(flexcpData, .Rows - 2, .ColIndex("结算卡类型"))
            .TextMatrix(.Rows - 1, .ColIndex("卡号")) = "小计"
            .TextMatrix(.Rows - 1, .ColIndex("卡余额")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("本次消费")) = Format(dbl本次刷卡, gVbFmtString.FM_金额)
            If lngDefault接口编号 = Val(.Cell(flexcpData, .Rows - 1, .ColIndex("结算卡类型"))) And "小计" = strDefaultCardNo Then
                .Row = .Rows - 1
            End If
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("结算卡类型")) = "合计"
            .Cell(flexcpData, .Rows - 1, .ColIndex("结算卡类型")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("卡号")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("卡余额")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("本次消费")) = Format(dbl本次刷卡总计, gVbFmtString.FM_金额)
            If lngDefault接口编号 = 0 And "合计" = strDefaultCardNo Then
                .Row = .Rows - 1
            End If
        End If
        If .Row < 0 And .Rows > 1 Then .Row = 1
    End With
End Sub
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.ID
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Edit_MoveCard '移除刷卡记录
            Call zlMoveCard
        Case mconMenu_Edit_Affirm
            mintSucces = mintSucces + 1
            Unload Me
        Case conMenu_File_Exit '
            mintSucces = 0
            Unload Me
        Case Else
            If Val(Control.Parameter) > 0 Then
                '执行具体功能:
                Call zlExecuteBrushCard(Val(Control.Parameter))
            End If
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
         
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
End Sub
 
Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub
