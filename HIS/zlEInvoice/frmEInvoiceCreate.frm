VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEInvoiceCreate 
   BorderStyle     =   0  'None
   Caption         =   "电子票据开具"
   ClientHeight    =   10860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   8748
      Left            =   384
      ScaleHeight     =   8745
      ScaleWidth      =   12945
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   12948
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   468
         Left            =   168
         ScaleHeight     =   465
         ScaleWidth      =   12690
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   408
         Width           =   12684
         Begin VB.ComboBox cbo业务类型 
            Height          =   276
            Left            =   912
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   96
            Width           =   1812
         End
         Begin VB.ComboBox cbo收费员 
            Height          =   276
            Left            =   3720
            TabIndex        =   6
            Top             =   96
            Width           =   1812
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "刷新(&R)"
            Height          =   300
            Left            =   11640
            TabIndex        =   11
            Top             =   84
            Width           =   1000
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   276
            Left            =   6648
            TabIndex        =   8
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   139329539
            CurrentDate     =   43941
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   276
            Left            =   9168
            TabIndex        =   10
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   139329539
            CurrentDate     =   43941
         End
         Begin VB.Label lbl业务类型 
            AutoSize        =   -1  'True
            Caption         =   "业务类型"
            Height          =   180
            Left            =   144
            TabIndex        =   3
            Top             =   144
            Width           =   720
         End
         Begin VB.Label lbl收费员 
            AutoSize        =   -1  'True
            Caption         =   "收费员"
            Height          =   180
            Left            =   3120
            TabIndex        =   5
            Top             =   144
            Width           =   540
         End
         Begin VB.Label lbl收费时间 
            AutoSize        =   -1  'True
            Caption         =   "收费时间"
            Height          =   180
            Left            =   5880
            TabIndex        =   7
            Top             =   144
            Width           =   720
         End
         Begin VB.Label lbl业务时间_ 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Left            =   8928
            TabIndex        =   9
            Top             =   144
            Width           =   180
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfExse 
         Height          =   1356
         Left            =   432
         TabIndex        =   12
         Top             =   2184
         Width           =   6108
         _cx             =   10774
         _cy             =   2392
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1032
      Left            =   24
      Top             =   900
      Width           =   528
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10845
      _Version        =   589884
      _ExtentX        =   19129
      _ExtentY        =   529
      _StockProps     =   6
      Caption         =   "补开电子票据"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoiceCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form, mlngSys As Long, mlngModule As Long, mstrDBUser As String, mstrEInvPrivs As String
Private mcbsMain   As Object          'CommandBar控件
Private mobjEInvoice As clsEInvoiceModule
Private mobjPubEInvoice As Object 'zlPublicExpense.clsPubEInvoice
Private mblnPrinting As Boolean
Private mrs收费员 As ADODB.Recordset

Public Event ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
Public Event ShowInfo(ByVal strInfo As String)

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String, _
    ByVal strEInvPrivs As String, objEInvoice As Object, objPubEInvoice As Object)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
    mstrEInvPrivs = strEInvPrivs
    Set mobjEInvoice = objEInvoice
    Set mobjPubEInvoice = objPubEInvoice
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom

    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EInvoice, "开票(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清(&C)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
    End With

    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_EInvoice, "开票", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新", cbrControl.Index + 1): cbrControl.BeginGroup = True
    End With

    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_SelAll
        .Add FCONTROL, Asc("C"), conMenu_Edit_ClsAll
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
    Case conMenu_File_Preview '预览
        Call OutputList(2)
    Case conMenu_File_Print '打印
        Call OutputList(1)
    Case conMenu_File_Excel '输出到Excel…
        Call OutputList(3)
    Case conMenu_Edit_EInvoice '开票
        Call CreateEInvoice
    Case conMenu_Edit_SelAll '全选
        Call Grid_SelAllRecord(vsfExse, True)
    Case conMenu_Edit_ClsAll '全清
        Call Grid_SelAllRecord(vsfExse, False)
    Case conMenu_View_Refresh '刷新
        Call cmdRefresh_Click
    End Select
End Sub

Private Sub CreateEInvoice()
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lng结算ID As Long, strDate As String, bln补结算 As Boolean
    Dim i As Long, byt场合 As Byte, lng冲销ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    With vsfExse
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then blnChecked = True: Exit For
        Next
    
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        For i = .FixedRows To .Rows - 1
            lng结算ID = Val(.RowData(i)): lng冲销ID = 0
            If lng结算ID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                byt场合 = Val(NVL(.Cell(flexcpData, i, .ColIndex("业务类型")))) 'Array("1-收费", "2-预交", "3-结帐", "4-挂号", "5-就诊卡")
                If byt场合 = 1 Or byt场合 = 4 Then
                    bln补结算 = Val(NVL(.Cell(flexcpData, i, .ColIndex("单据号")))) = 1
                ElseIf byt场合 = 12 Then '余额退款
                    lng冲销ID = Val(NVL(.Cell(flexcpData, i, .ColIndex("单据号"))))
                End If
                
                byt场合 = byt场合 Mod 10
                If blnInit = False Then
                    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt场合) = False Then Exit Sub
                    blnInit = True
                End If
                
                If mobjPubEInvoice.zlGetEInvoiceIDFromBalanceID(byt场合, lng结算ID) <> 0 Then
                        .TextMatrix(i, .ColIndex("开票结果")) = "开票失败"
                        .TextMatrix(i, .ColIndex("开票说明")) = "本次结算已开具电子票据，请刷新后重试。"
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                Else
                    If GetSwapCollectFromBalanceID(byt场合, lng结算ID, cllSwapData, bln补结算, lng冲销ID, False, strErrMsg) = False Then
                        .TextMatrix(i, .ColIndex("开票结果")) = "开票失败"
                        .TextMatrix(i, .ColIndex("开票说明")) = strErrMsg
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                    Else
                        If mobjPubEInvoice.zlOnlyCreateEinvoice(Me, byt场合, cllSwapData, Nothing, False, strErrMsg) Then
                            .TextMatrix(i, .ColIndex("开票结果")) = "开票成功"
                            .TextMatrix(i, .ColIndex("开票说明")) = ""
                        Else
                            .TextMatrix(i, .ColIndex("开票结果")) = "开票失败"
                            .TextMatrix(i, .ColIndex("开票说明")) = strErrMsg
                            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                        End If
                    End If
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "请选择需要开具电子票据的费用记录。", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_EInvoice '开票
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "开具电子票据")
        Control.Enabled = Control.Visible And vsfExse.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_SelAll '全选
        Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_ClsAll '全清
        Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
    End Select
End Sub

Private Sub cbo收费员_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
     
    If cbo收费员.ListIndex <> -1 Then
        '弹出列表时,又在文本框输入了内容
        If UCase(cbo收费员.Text) <> UCase(cbo收费员.List(cbo收费员.ListIndex)) Then Call zlControl.CboSetIndex(cbo收费员.hWnd, -1)
    End If
    
    If cbo收费员.Text = "" Then
        cbo收费员.ListIndex = -1
    ElseIf cbo收费员.ListIndex = -1 Then
        If Select收费员(Me, mlngSys, mlngModule, cbo收费员, mrs收费员) = False Then
            KeyAscii = 0: zlControl.TxtSelAll cbo收费员: Exit Sub
        End If
    End If
    
    If cbo收费员.ListIndex = -1 Then cbo收费员.Text = ""
End Sub

Private Sub cbo收费员_LostFocus()
    If cbo收费员.Text <> "" And cbo收费员.ListIndex < 0 Then cbo收费员.Text = ""
End Sub

Private Sub cmdRefresh_Click()
    '加载未开具电子票据的费用数据
    Dim dtBegin As Date, dtEnd As Date
    
    '1.数据检查
    On Error GoTo ErrHandler
    dtBegin = dtp开始时间.Value: dtEnd = dtp结束时间.Value
    If dtp开始时间 > dtp结束时间 Then
        MsgBox "费用的开始时间不能大于结束时间！", vbInformation, gstrSysName
        zlControl.ControlSetFocus dtp结束时间:  Exit Sub
    End If
    
    If DateDiff("m", dtp开始时间, dtp结束时间) > 6 Then
        If MsgBox("对当前费用时间范围内的数据进行查询可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    '2.获取数据
    Dim rsExse As ADODB.Recordset
    If GetExseData(zlStr.NeedCode(cbo业务类型.Text), zlStr.NeedName(cbo收费员.Text), _
        dtp开始时间.Value, dtp结束时间.Value, rsExse) = False Then Exit Sub
        
    Dim lngOldRow As Long, lngOldCol As Long
    lngOldRow = vsfExse.Row: lngOldCol = vsfExse.Col
    
    '选择,业务类型,NO,姓名,性别,年龄,门诊号,住院号,收费员,收费时间,费用金额,开票结果,开票说明
    vsfExse.Clear 1
    vsfExse.Rows = vsfExse.FixedRows + 1
    
    With vsfExse
        .Redraw = flexRDNone
        Do While Not rsExse.EOF
            If .TextMatrix(.Rows - 1, .ColIndex("业务类型")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = Val(NVL(rsExse!结算ID))
            .TextMatrix(.Rows - 1, .ColIndex("业务类型")) = Decode(Val(NVL(rsExse!业务类型)) Mod 10, 1, "收费", 2, "预交", 3, "结帐", 4, "挂号", 5, "就诊卡")
            .Cell(flexcpData, .Rows - 1, .ColIndex("业务类型")) = Val(NVL(rsExse!业务类型))
            .TextMatrix(.Rows - 1, .ColIndex("单据号")) = NVL(rsExse!NO)
            Select Case Val(NVL(rsExse!业务类型)) Mod 10
            Case 1
                .Cell(flexcpData, .Rows - 1, .ColIndex("单据号")) = NVL(rsExse!补结算)
            Case 2
                .Cell(flexcpData, .Rows - 1, .ColIndex("单据号")) = NVL(rsExse!冲销ID)
            Case 4
                .Cell(flexcpData, .Rows - 1, .ColIndex("单据号")) = NVL(rsExse!补结算)
            End Select
            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = NVL(rsExse!姓名)
            .TextMatrix(.Rows - 1, .ColIndex("性别")) = NVL(rsExse!性别)
            .TextMatrix(.Rows - 1, .ColIndex("年龄")) = NVL(rsExse!年龄)
            .TextMatrix(.Rows - 1, .ColIndex("门诊号")) = NVL(rsExse!门诊号)
            .TextMatrix(.Rows - 1, .ColIndex("住院号")) = NVL(rsExse!住院号)
            .TextMatrix(.Rows - 1, .ColIndex("收费员")) = NVL(rsExse!操作员姓名)
            .TextMatrix(.Rows - 1, .ColIndex("收费时间")) = Format(NVL(rsExse!收款时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("费用金额")) = FormatEx(Val(NVL(rsExse!金额)), 2, , , 6)
        
            rsExse.MoveNext
        Loop
            
        .Cell(flexcpFontBold, .FixedRows, .ColIndex("开票结果"), .Rows - 1, .ColIndex("开票说明")) = True
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows - 1 > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngOldCol = 0 Or lngOldCol > .Cols - 1, .FixedCols, lngOldCol)
            .ShowCell .Row, .Col  '立刻显示到指定单元
        End If
        
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call InitExseGrid
    
    Dim varData As Variant, i As Integer
    varData = Array("1-收费", "2-预交", "3-结帐", "4-挂号", "5-就诊卡")
    cbo业务类型.Clear
    For i = 0 To UBound(varData)
        cbo业务类型.AddItem varData(i)
    Next
    cbo业务类型.ListIndex = 0
    
    Call Load收费员(cbo收费员, mrs收费员)

    dtp结束时间.Value = zlDatabase.Currentdate
    dtp开始时间.Value = Format(DateAdd("d", -1, dtp结束时间.Value), "yyyy-MM-dd 00:00:00")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
    Set mobjEInvoice = Nothing
    
    Set mrs收费员 = Nothing
End Sub

Private Function InitExseGrid() As Boolean
    '初始化VSFGrid表格控件
    Dim strHead As String, varData As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '列名1,对齐方式1,列宽1|列名2,对齐方式2,列宽2|...
    strHead = "选择,4,500|业务类型,1,900|单据号,1,1000|姓名,1,1000|性别,1,600|年龄,1,600|门诊号,1,1000|住院号,1,1000" & _
                    "|收费员,1,800|收费时间,4,2000|费用金额,7,1200" & _
                    "|开票结果,1,1000|开票说明,1,5000"
    With vsfExse
        .Redraw = flexRDNone '暂停表格显示刷新
        .Clear
        .Rows = 2
        .FixedRows = 1: .FixedCols = 0

        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = Split(varData(i), ",")(0)  '设置Key值,用于根据 ColIndex() 确定列
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = Split(varData(i), ",")(1)
        Next

        .AllowSelection = False '不允许多选
        .AllowBigSelection = False '不允许点击固定行/列选择整行/整列
        .SelectionMode = flexSelectionByRow '按行选择
        .AllowUserResizing = flexResizeColumns '允许用户调整列宽
        .BackColorSel = &HE0E0E0
        .ForeColorSel = vbBlack
        
        .Editable = flexEDKbdMouse

        .RowHeightMin = 300
        .ColDataType(.ColIndex("选择")) = flexDTBoolean

        .Redraw = flexRDBuffered '刷新表格显示
    End With
    InitExseGrid = True
    Exit Function
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub picMain_Resize()
    On Error Resume Next
    picFilter.Move 0, 0, picMain.ScaleWidth
    vsfExse.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth, picMain.ScaleHeight - (picFilter.Top + picFilter.Height)
End Sub

Private Sub vsfExse_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    On Error Resume Next
    vsfExse.ForeColorSel = vsfExse.CellForeColor
End Sub

Private Sub vsfExse_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsfExse.ColIndex("选择") Or vsfExse.TextMatrix(Row, 1) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vsfExse_GotFocus()
    Call SetActiveList(vsfExse)
End Sub

Private Sub vsfExse_LostFocus()
    Call SetActiveList(vsfExse, False)
End Sub

Private Sub SetActiveList(vsfGrid As VSFlexGrid, Optional ByVal blnGetFocus As Boolean = True)
    '设置控件选择行背景高亮色
    If blnGetFocus Then
        vsfExse.BackColorSel = &HE0E0E0

        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &H8000000D '&HC0C0C0
    Else
        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub vsfExse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(False)
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte
    Dim intCurrentRow As Integer, vsfGrid As VSFlexGrid
    
    On Error GoTo ErrHandler
    '表头
    Set objOut = New zlPrint1Grd
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    Set vsfGrid = vsfExse
    objOut.Title.Text = "未开具电子票据费用清单"
    
    '表项
    Set objRow = New zlTabAppRow
    objRow.Add "业务类型：" & cbo业务类型.Text
    objRow.Add "收费员：" & cbo收费员.Text
    objRow.Add "费用时间：" & Format(dtp开始时间, "yyyy-mm-dd") & " 至 " & Format(dtp结束时间, "yyyy-mm-dd")
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    vsfGrid.Redraw = False
    intCurrentRow = vsfGrid.Row
    mblnPrinting = True
    
    '表体
    Set objOut.Body = vsfGrid
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
End Sub
