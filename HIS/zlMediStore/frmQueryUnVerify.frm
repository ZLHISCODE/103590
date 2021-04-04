VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmQueryUnVerify 
   Caption         =   "未审核单据查询"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11415
   Icon            =   "frmQueryUnVerify.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11415
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picData 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5895
      ScaleWidth      =   11055
      TabIndex        =   9
      Top             =   1320
      Width           =   11055
      Begin VB.Frame fraLineH1 
         Height          =   50
         Left            =   0
         TabIndex        =   12
         Top             =   4320
         Width           =   3405
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   9375
         _cx             =   16536
         _cy             =   4410
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmQueryUnVerify.frx":076A
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1245
         Left            =   0
         TabIndex        =   11
         Top             =   4560
         Width           =   11055
         _cx             =   19500
         _cy             =   2196
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmQueryUnVerify.frx":082A
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.TextBox Txt药品 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   6240
         MaxLength       =   50
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   120
         Width           =   3255
      End
      Begin VB.CommandButton cmd查询 
         Caption         =   "查询(&S)"
         Height          =   350
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton Cmd药品 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9480
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox chkDrug 
         BackColor       =   &H80000003&
         Caption         =   "药品"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   143
         Width           =   735
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   2040
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Text            =   "cboStock"
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "时间范围"
         Height          =   180
         Left            =   2520
         TabIndex        =   4
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblStock 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "库房"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "注意：红色表示指定药品将做出库！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   2880
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmQueryUnVerify.frx":08ED
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQueryUnVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnLoadDate As Boolean '数据是否加载完
Private mintChoose级数 As Byte          '0-售价单位;1-门诊单位;2-药库单位;3-住院单位
Private mintNumberDigit As Integer

Private Sub GetData()
    
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal cboStcokMain As ComboBox, ByVal intChoose级数 As Byte, ByVal intNumberDigit As Integer)
    Dim i As Integer
    
    cboStock.Clear
    For i = 1 To cboStcokMain.ListCount - 1 '排除第一个所有库房
        cboStock.AddItem cboStcokMain.List(i)
        cboStock.ItemData(cboStock.NewIndex) = cboStcokMain.ItemData(i)
    Next
    
    If cboStock.ListCount > 0 Then
        cboStock.ListIndex = IIf(cboStcokMain.ListIndex - 1 >= 0, cboStcokMain.ListIndex - 1, 0)
    End If
    
    mintChoose级数 = intChoose级数
    mintNumberDigit = intNumberDigit
    
    Me.Show vbModal, FrmMain
End Sub

Private Sub InitComandBars()
    '初始化菜单：加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

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
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPublic.Icons
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "预览")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "打印")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出")
        cbrControlMain.BeginGroup = True
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    cbsMain.Item(1).Visible = False
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.id
        Case mconMenu_File_Preview
            subPrint 2
        Case mconMenu_File_Print
            subPrint 1
        Case mconMenu_File_Exit
            Unload Me
    End Select
    
End Sub

Private Sub cbsFilePreView()
    '打印预览
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsFilePrint()
    '打印
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picCondition.Move lngLeft, lngTop, lngRight - lngLeft
    
    Me.lblMsg.Move 0, Me.ScaleHeight - lblMsg.Height - 50, lblMsg.Width, lblMsg.Height
    
    Me.picData.Move lngLeft, picCondition.Top + picCondition.Height + 50, lngRight - lngLeft, _
        Me.ScaleHeight - Me.picCondition.Top - Me.picCondition.Height - lblMsg.Height - 150
End Sub


Private Sub chkDrug_Click()
    Txt药品.Enabled = IIf(chkDrug.Value = 1, True, False)
    Cmd药品.Enabled = IIf(chkDrug.Value = 1, True, False)
    
    Txt药品.BackColor = IIf(Txt药品.Enabled, &H80000005, &H80000004)
End Sub

Private Sub cmd查询_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim intDay As Integer
    Dim lngStockid As Long
    Dim strSql As String
    
    On Error GoTo errHandle
    
    blnLoadDate = False
    
    lngStockid = Val(cboStock.ItemData(cboStock.ListIndex))
    
    Select Case cboTime.ListIndex
        Case 0 '显示7天内
            intDay = 7
        Case 1 '显示1个月内
            intDay = 30
        Case 2 '显示3个月内
            intDay = 90
        Case 3 '显示半年内
            intDay = 183
        Case 4 '显示1年内
            intDay = 365
    End Select
    
    If chkDrug.Value = 1 Then '勾选了药品
        If Val(Txt药品.Tag) <= 0 Then
            MsgBox "请选择要查询的药品！", vbInformation + vbOKOnly, gstrSysName
            Txt药品.SetFocus
            Exit Sub
        End If
        
        '根据传入参数设置显示单位
'        Select Case mintChoose级数
'            Case 1
'                strSql = ", Decode(Sign(a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)), -1, a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0), 0) As 数量 ,C.计算单位 as 单位"
'            Case 2
'                strSql = ", Decode(Sign(a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)), -1, a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0), 0)/D.门诊包装 As 数量  ,D.门诊单位 as 单位"
'            Case 3
'                strSql = ", Decode(Sign(a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)), -1, a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0), 0)/D.药库包装 As 数量  ,D.药库单位 as 单位"
'            Case 4
'                strSql = ", Decode(Sign(a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)), -1, a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0), 0)/D.住院包装 As 数量 ,D.住院单位 as 单位"
'        End Select
        Select Case mintChoose级数
            Case 1
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0) As 数量 ,C.计算单位 as 单位"
            Case 2
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)/D.门诊包装 As 数量  ,D.门诊单位 as 单位"
            Case 3
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)/D.药库包装 As 数量  ,D.药库单位 as 单位"
            Case 4
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)/D.住院包装 As 数量 ,D.住院单位 as 单位"
        End Select
        
        '查询指定药品的未审核单据级单据数量，和占用可用数量
        gstrSQL = "Select a.id, a.入出类别, Count(Distinct NO) As 单据数量, Sum(a.数量) As 实际数量,Max(a.单位) as 单位 " & vbNewLine & _
                "From (Select e.id ,e.名称 入出类别, a.No" & strSql & vbNewLine & _
                "       From 药品收发记录 A, 未审药品记录 B ,收费项目目录 C,药品规格 D, 药品入出类别 E" & vbNewLine & _
                "       Where a.Id = b.收发id And A.药品id = C.id And C.id = d.药品id And a.入出类别id = e.Id " & IIf(lngStockid = 0, "", " And b.库房id = [1] ") & " And b.药品id = [2] " & IIf(cboTime.ListIndex = 5, "", " And a.填制日期 > sysdate - [3]") & " And Not Exists" & vbNewLine & _
                "        (Select 1 From 药品收发记录 C Where b.收发id = c.Id And Nvl(c.发药方式, 0) = -1 And c.单据 In (8, 9, 10))) A" & vbNewLine & _
                "Group By a.入出类别,a.id"
    Else
        '查询未审核单据及单据数量
        gstrSQL = "Select e.id ,e.名称 入出类别, Count(Distinct a.No) As 单据数量" & vbNewLine & _
                "From 药品收发记录 A, 未审药品记录 B, 药品入出类别 E" & vbNewLine & _
                "Where a.Id = b.收发id And a.入出类别id = e.Id " & IIf(lngStockid = 0, "", " And b.库房id = [1] ") & "" & IIf(cboTime.ListIndex = 5, "", " And a.填制日期 > sysdate - [3] ") & vbNewLine & _
                "Group By e.id ,e.名称"
    End If
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", lngStockid, Val(Txt药品.Tag), intDay)
    
    With vsfList
        .rows = 1
        .rows = .rows + 1
        .Row = .rows - 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Row, .ColIndex("入出类别id")) = rsTemp!id
            .TextMatrix(.Row, .ColIndex("入出类别")) = rsTemp!入出类别
            .TextMatrix(.Row, .ColIndex("单据数量")) = rsTemp!单据数量
            If chkDrug.Value = 1 Then
                '颜色区分入出
                If rsTemp!实际数量 < 0 Then
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF '出库红色
                Else
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
                End If
                
                .TextMatrix(.Row, .ColIndex("实际数量")) = zlStr.FormatEx(Abs(rsTemp!实际数量), mintNumberDigit, False, True) & rsTemp!单位
                .Cell(flexcpFontBold, 1, .ColIndex("实际数量"), .rows - 1, .ColIndex("实际数量")) = True
            End If
            
            .rows = .rows + 1
            .Row = .Row + 1
            
            rsTemp.MoveNext
        Loop
        
        If Trim(.TextMatrix(.rows - 1, 0)) = "" Then .RemoveItem (.rows - 1) '删除最后的空行
    End With
    
    colHidden
    
    blnLoadDate = True
    
    vsfList_EnterCell '加载明细
    
    vsfList.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , True)
    
'    Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
    
    cmd查询.SetFocus
End Sub

Private Sub Form_Load()
    
    With cboTime
        .Clear
        
        .AddItem "0-显示7天内"
        .AddItem "1-显示1个月内"
        .AddItem "2-显示3个月内"
        .AddItem "3-显示半年内"
        .AddItem "4-显示1年内"
        .AddItem "5-显示所有"
        
        .ListIndex = 0
    End With
    
    Call InitComandBars
    
    colHidden
End Sub


Private Sub colHidden()
    '根据选择条件隐藏列
    With vsfList
        .colHidden(.ColIndex("实际数量")) = chkDrug.Value = 0 '未选择药品，隐藏“占用可用数量”列
        
        .ColWidth(.ColIndex("实际数量")) = IIf(.colHidden(.ColIndex("实际数量")), 0, 900)

    End With
    With vsfDetail
        .colHidden(.ColIndex("数量")) = chkDrug.Value = 0 '未选择药品，隐藏“数量”列
        
        .ColWidth(.ColIndex("数量")) = IIf(.colHidden(.ColIndex("数量")), 0, 1545)
    End With
End Sub

Private Sub fraLineH1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With fraLineH1
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    With vsfList
        .Height = fraLineH1.Top - .Top
    End With
    
    With vsfDetail
        .Top = fraLineH1.Top + fraLineH1.Height + 100
        .Height = ScaleHeight - .Top
    End With
    Me.Refresh
End Sub


Private Sub picData_Resize()
    On Error Resume Next
    
    With vsfList
        .Move 0, 0, picData.Width, 2500
    End With
    
    With fraLineH1
        .Move 0, vsfList.Top + vsfList.Height, picData.Width, fraLineH1.Height
    End With
    
    With vsfDetail
        .Move 0, fraLineH1.Top + fraLineH1.Height, picData.Width, picData.Height - fraLineH1.Top - fraLineH1.Height
    End With
End Sub


Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + picCondition.Left + Txt药品.Left
    sngTop = Me.Top + picCondition.Top + Txt药品.Top + Txt药品.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 3630
    End If
    
    strkey = Trim(Txt药品.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , True)
    
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
    
    cmd查询.SetFocus
    
End Sub

Private Sub vsfList_EnterCell()
    Dim rsTemp As New ADODB.Recordset
    Dim intDay As Integer
    Dim lngStockid As Long
    Dim int入出类别id As Integer
    Dim strSql As String
    
    On Error GoTo errHandle
    
    If Not blnLoadDate Then Exit Sub
    
    Select Case cboTime.ListIndex
        Case 0 '显示7天内
            intDay = 7
        Case 1 '显示1个月内
            intDay = 30
        Case 2 '显示3个月内
            intDay = 90
        Case 3 '显示半年内
            intDay = 183
        Case 4 '显示1年内
            intDay = 365
    End Select
    
    lngStockid = Val(cboStock.ItemData(cboStock.ListIndex))
    int入出类别id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("入出类别id")))
    
    If chkDrug.Value = 1 Then '勾选了药品
        '根据传入参数设置显示单位
        Select Case mintChoose级数
            Case 1
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0) As 数量 ,C.计算单位 as 单位"
            Case 2
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)/D.门诊包装 As 数量  ,D.门诊单位 as 单位"
            Case 3
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)/D.药库包装 As 数量  ,D.药库单位 as 单位"
            Case 4
                strSql = ", a.入出系数 * Nvl(a.付数, 1) * Nvl(a.实际数量, 0)/D.住院包装 As 数量 ,D.住院单位 as 单位"
        End Select
        
        gstrSQL = "Select a.No, a.填制人, a.填制日期, a.摘要, Sum(a.数量) 数量,Max(a.单位) as 单位" & vbNewLine & _
                "From (Select a.No" & strSql & vbNewLine & _
                "       , a.填制人, a.填制日期, a.摘要" & vbNewLine & _
                "       From 药品收发记录 A, 未审药品记录 B,收费项目目录 C,药品规格 D" & vbNewLine & _
                "       Where a.Id = b.收发id And A.药品id = C.id And C.id = d.药品id And b.库房id = [1] And a.入出类别id = [2] " & IIf(cboTime.ListIndex = 5, "", " And a.填制日期 > sysdate - [3]") & " And a.药品id = [4] And Not Exists" & vbNewLine & _
                "        (Select 1 From 药品收发记录 C Where b.收发id = c.Id And Nvl(c.发药方式, 0) = -1 And c.单据 In (8, 9, 10))) A" & vbNewLine & _
                "Group By a.No, a.填制人, a.填制日期, a.摘要, a.数量"


    Else
        gstrSQL = "Select Distinct a.No, a.填制人, a.填制日期, a.摘要" & vbNewLine & _
                "From 药品收发记录 A, 未审药品记录 B" & vbNewLine & _
                "Where a.Id = b.收发id And b.库房id = [1] And a.入出类别id = [2]" & IIf(cboTime.ListIndex = 5, "", " And a.填制日期 > sysdate - [3]") & vbNewLine & _
                "Order By NO"
    End If
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", lngStockid, int入出类别id, intDay, Val(Txt药品.Tag))
    
    With vsfDetail
        .rows = 1
        .rows = .rows + 1
        .Row = .rows - 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Row, .ColIndex("No")) = rsTemp!NO
            .TextMatrix(.Row, .ColIndex("填制人")) = rsTemp!填制人
            .TextMatrix(.Row, .ColIndex("填制日期")) = rsTemp!填制日期
            .TextMatrix(.Row, .ColIndex("摘要")) = "" & rsTemp!摘要
            If chkDrug.Value = 1 Then
                '颜色区分入出
                If rsTemp!数量 < 0 Then
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF '出库红色
                Else
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
                End If
                .TextMatrix(.Row, .ColIndex("数量")) = zlStr.FormatEx(Abs(rsTemp!数量), mintNumberDigit, False, True) & rsTemp!单位: .Cell(flexcpFontBold, 1, .ColIndex("数量"), .rows - 1, .ColIndex("数量")) = True
            End If
            
            .rows = .rows + 1
            .Row = .Row + 1
            
            rsTemp.MoveNext
        Loop
        
        If Trim(.TextMatrix(.rows - 1, 0)) = "" Then .RemoveItem (.rows - 1) '删除最后的空行
    End With
    
    colHidden
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow

    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "未审核单据"
        
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is vsfDetail Then
        Set objPrint.Body = vsfDetail
    Else
        Set objPrint.Body = vsfList
    End If
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

