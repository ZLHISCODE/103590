VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#4.1#0"; "zlIDKind.ocx"
Begin VB.Form frmEInvoicePrint 
   BorderStyle     =   0  'None
   Caption         =   "纸质票据管理"
   ClientHeight    =   10350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13785
   LinkTopic       =   "Form2"
   ScaleHeight     =   10350
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   8748
      Left            =   840
      ScaleHeight     =   8745
      ScaleWidth      =   12780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   12780
      Begin VB.Frame fraSplit 
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   1632
         MousePointer    =   7  'Size N S
         TabIndex        =   14
         Top             =   3864
         Width           =   1005
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   444
         Left            =   24
         ScaleHeight     =   450
         ScaleWidth      =   12690
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   168
         Width           =   12684
         Begin VB.TextBox txtPatient 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1248
            MaxLength       =   100
            TabIndex        =   5
            ToolTipText     =   "定位:F6,输入:-病人ID,*门诊号,+住院号,.挂号单号,例如:*2536表示按门诊号查找"
            Top             =   84
            Width           =   1470
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "刷新(&R)"
            Height          =   300
            Left            =   11568
            TabIndex        =   12
            Top             =   84
            Width           =   1000
         End
         Begin VB.ComboBox cbo票据类型 
            Height          =   276
            Left            =   3744
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   96
            Width           =   1812
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   276
            Left            =   6648
            TabIndex        =   9
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   73269251
            CurrentDate     =   43941
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   276
            Left            =   9168
            TabIndex        =   11
            Top             =   96
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   73269251
            CurrentDate     =   43941
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   300
            Left            =   600
            TabIndex        =   4
            Top             =   84
            Width           =   636
            _ExtentX        =   1111
            _ExtentY        =   529
            Appearance      =   2
            IDKindStr       =   $"frmEInvoicePrint.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "宋体"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            DefaultCardType =   "0"
            NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            MustSelectItems =   "姓名,就诊卡"
            BackColor       =   -2147483633
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            Caption         =   "姓名"
            Height          =   180
            Left            =   168
            TabIndex        =   3
            Top             =   144
            Width           =   360
         End
         Begin VB.Label lbl业务时间_ 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Left            =   8928
            TabIndex        =   10
            Top             =   144
            Width           =   180
         End
         Begin VB.Label lbl收费时间 
            AutoSize        =   -1  'True
            Caption         =   "收费时间"
            Height          =   180
            Left            =   5880
            TabIndex        =   8
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lbl票据类型 
            AutoSize        =   -1  'True
            Caption         =   "票据类型"
            Height          =   180
            Left            =   2976
            TabIndex        =   6
            Top             =   144
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfEInvoice 
         Height          =   1356
         Left            =   360
         TabIndex        =   13
         Top             =   1392
         Width           =   6108
         _cx             =   1983064598
         _cy             =   1983056216
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
      Begin VSFlex8Ctl.VSFlexGrid vsfExse 
         Height          =   1404
         Left            =   720
         TabIndex        =   15
         Top             =   4656
         Width           =   4404
         _cx             =   1983061592
         _cy             =   1983056300
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
      Height          =   1035
      Left            =   0
      Top             =   900
      Width           =   525
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
      Caption         =   "电子票据打印"
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
Attribute VB_Name = "frmEInvoicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form, mlngSys As Long, mlngModule As Long, mstrDBUser As String, mstrEInvPrivs As String
Private mcbsMain   As Object          'CommandBar控件
Private mobjEInvoice As clsEInvoiceModule
Private mobjPubEInvoice As Object ' zlPublicExpense.clsPubEInvoice
Private mblnPrinting As Boolean
Attribute mblnPrinting.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mobjSquareCard As Object
Private mcllResult As Collection

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
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintEInvoice, "打印电子发票(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendMsg, "信息推送(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintNotice, "打印告知单(&N)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_TurnPaper, "换开纸质票据(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ReTurnPaper, "重新换开票据(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CancelTurnPaper, "作废纸质票据(&C)")
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
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
    Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintEInvoice, "打印电子发票", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendMsg, "信息推送", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrintNotice, "打印告知单", cbrControl.index + 1): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清", cbrControl.index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With

    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_SelAll
        .Add FCONTROL, Asc("C"), conMenu_Edit_ClsAll
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim frmParaSet As frmEInvoiceParaSet
    
    Select Case Control.ID
    Case conMenu_File_Preview '预览
        Call OutputList(2)
    Case conMenu_File_Print '打印
        Call OutputList(1)
    Case conMenu_File_Excel '输出到Excel…
        Call OutputList(3)
    
    Case conMenu_File_Parameter '参数设置
        Set frmParaSet = New frmEInvoiceParaSet
        Call frmParaSet.ShowMe(Me, mlngSys, 1145)
        
    Case conMenu_Edit_PrintEInvoice '打印电子发票
        Call ExcutePrintEInvoice(zlStr.NeedCode(cbo票据类型.Text))
    Case conMenu_Edit_PrintNotice '打印告知单
        Call ExecutePrintNotice(zlStr.NeedCode(cbo票据类型.Text))
    Case conMenu_Edit_SendMsg '信息推送
        Call ExecuteSendMsg(zlStr.NeedCode(cbo票据类型.Text))
    Case conMenu_Edit_TurnPaper '换开纸质票据
        Call ExcuteTurnPaper(zlStr.NeedCode(cbo票据类型.Text), False)
    Case conMenu_Edit_ReTurnPaper '重新换开票据
        Call ExcuteTurnPaper(zlStr.NeedCode(cbo票据类型.Text), True)
    Case conMenu_Edit_CancelTurnPaper '作废纸质票据
        Call ExcuteCancelTurnPaper(zlStr.NeedCode(cbo票据类型.Text))
        
    Case conMenu_Edit_SelAll '全选
        Call Grid_SelAllRecord(vsfEInvoice, True)
    Case conMenu_Edit_ClsAll '全清
        Call Grid_SelAllRecord(vsfEInvoice, False)
        
    Case conMenu_View_Refresh '刷新
        Call cmdRefresh_Click
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        If Me.ActiveControl Is vsfExse Then
            Control.Enabled = vsfExse.TextMatrix(1, 1) <> ""
        Else
            Control.Enabled = vsfEInvoice.TextMatrix(1, 1) <> ""
        End If
    
    Case conMenu_Edit_PrintEInvoice '打印电子发票
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_SendMsg '信息推送
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_PrintNotice '打印告知单
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_TurnPaper '换开纸质票据
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "换开纸质票据")
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_ReTurnPaper '重新换开票据
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "重新换开票据")
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_CancelTurnPaper '作废纸质票据
        Control.Visible = zlStr.IsHavePrivs(mstrEInvPrivs, "作废纸质票据")
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    
    Case conMenu_Edit_SelAll '全选
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    Case conMenu_Edit_ClsAll '全清
        Control.Enabled = Control.Visible And vsfEInvoice.TextMatrix(1, 1) <> ""
    End Select
End Sub

Private Sub cbo票据类型_Click()
    Dim byt场合 As Byte
    
    If cbo票据类型.Tag = cbo票据类型.Text Then Exit Sub
    cbo票据类型.Tag = cbo票据类型.Text
    
    byt场合 = zlStr.NeedCode(cbo票据类型.Text)
    vsfExse.Visible = byt场合 <> 2
    fraSplit.Visible = byt场合 <> 2
    vsfExse.Tag = IIf(byt场合 = 2, "ExseGridHidden", "")
    Call picMain_Resize
    Call cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
    Set mcllResult = Nothing
    Call LoadEInvoiceData
End Sub

Private Function LoadEInvoiceData() As Boolean
    '显示电子票据数据
    Dim dtBegin As Date, dtEnd As Date
    
    '1.数据检查
    On Error GoTo ErrHandler
    dtBegin = dtp开始时间.Value: dtEnd = dtp结束时间.Value
    If dtp开始时间 > dtp结束时间 Then
        MsgBox "费用的开始时间不能大于结束时间！", vbInformation, gstrSysName
        zlControl.ControlSetFocus dtp结束时间:  Exit Function
    End If
    
    If DateDiff("m", dtp开始时间, dtp结束时间) > 6 Then
        If MsgBox("对当前费用时间范围内的数据进行查询可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    '2.获取数据
    Dim rsEInvoice As ADODB.Recordset, bytQueryType As Byte, varQueryValue As Variant
    'bytQueryType 查询类型，0-所有，1-按病人ID查询，2-按费用单据号查询，3-按电子票据号查
    Select Case IDKind.GetCurCard.名称
    Case "收费单据号"
        If Trim(txtPatient.Text) <> "" Then
            bytQueryType = 2
            varQueryValue = Trim(txtPatient.Text)
        End If
    Case "电子票据号"
        If Trim(txtPatient.Text) <> "" Then
            bytQueryType = 3
            varQueryValue = Trim(txtPatient.Text)
        End If
    Case Else
        If Val(txtPatient.Tag) <> 0 Then
            bytQueryType = 1
            varQueryValue = Val(txtPatient.Tag)
        End If
    End Select
    If GetEInvoiceData(zlStr.NeedCode(cbo票据类型.Text), dtp开始时间.Value, dtp结束时间.Value, rsEInvoice, 3, 1, bytQueryType, varQueryValue) = False Then Exit Function
    
    Dim lngOldRow As Long, lngOldCol As Long
    lngOldRow = vsfEInvoice.Row: lngOldCol = vsfEInvoice.Col
    vsfEInvoice.Clear 1
    vsfEInvoice.Rows = 1 '清除Data值
    vsfEInvoice.Rows = vsfEInvoice.FixedRows + 1
    
    vsfExse.Clear 1
    vsfExse.Rows = 1 '
    vsfExse.Rows = vsfExse.FixedRows + 1
    
    With vsfEInvoice
        .Redraw = flexRDNone
        '选择,NO,姓名,性别,年龄,门诊号,住院号,票据类型,票据代码,票据号码,检验码,票据金额,开票点,开票时间,换开纸质发票,纸质发票号
        Do While Not rsEInvoice.EOF
            If .TextMatrix(.Rows - 1, .ColIndex("单据号")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = Val(Nvl(rsEInvoice!结算ID))
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Val(Nvl(rsEInvoice!ID))
            .TextMatrix(.Rows - 1, .ColIndex("单据号")) = Nvl(rsEInvoice!No)
            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = Nvl(rsEInvoice!姓名)
            .TextMatrix(.Rows - 1, .ColIndex("性别")) = Nvl(rsEInvoice!性别)
            .TextMatrix(.Rows - 1, .ColIndex("年龄")) = Nvl(rsEInvoice!年龄)
            .TextMatrix(.Rows - 1, .ColIndex("门诊号")) = Nvl(rsEInvoice!门诊号)
            .TextMatrix(.Rows - 1, .ColIndex("住院号")) = Nvl(rsEInvoice!住院号)
            
            .TextMatrix(.Rows - 1, .ColIndex("票据类型")) = Decode(Val(Nvl(rsEInvoice!票种)), 1, "收费", 2, "预交", 3, "结帐", 4, "挂号", 5, "就诊卡")
            .Cell(flexcpData, .Rows - 1, .ColIndex("票据类型")) = Val(Nvl(rsEInvoice!票种))
            .TextMatrix(.Rows - 1, .ColIndex("票据代码")) = Nvl(rsEInvoice!票据代码)
            .TextMatrix(.Rows - 1, .ColIndex("票据号码")) = Nvl(rsEInvoice!票据号码)
            .TextMatrix(.Rows - 1, .ColIndex("检验码")) = Nvl(rsEInvoice!检验码)
            .TextMatrix(.Rows - 1, .ColIndex("票据金额")) = FormatEx(Val(Nvl(rsEInvoice!票据金额)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("开票点")) = Nvl(rsEInvoice!开票点)
            .TextMatrix(.Rows - 1, .ColIndex("开票时间")) = Format(Nvl(rsEInvoice!开票时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("换开纸质发票")) = IIf(Val(Nvl(rsEInvoice!是否换开)) = 1, "√", "")
            .TextMatrix(.Rows - 1, .ColIndex("纸质发票号")) = Nvl(rsEInvoice!纸质发票号)
            
            Select Case Val(Nvl(rsEInvoice!票种))
            Case 2
                .Cell(flexcpData, .Rows - 1, .ColIndex("单据号")) = Val(Nvl(rsEInvoice!退款ID))
            Case 1, 4
                .Cell(flexcpData, .Rows - 1, .ColIndex("单据号")) = Val(Nvl(rsEInvoice!补结算))
            End Select
            
            If Not mcllResult Is Nothing Then
                If CollectionExitsValue(mcllResult, "_" & Val(Nvl(rsEInvoice!ID))) Then
                    .TextMatrix(.Rows - 1, .ColIndex("打印结果")) = mcllResult("_" & Val(Nvl(rsEInvoice!ID)))(0)
                    .TextMatrix(.Rows - 1, .ColIndex("打印说明")) = mcllResult("_" & Val(Nvl(rsEInvoice!ID)))(1)
                    If mcllResult("_" & Val(Nvl(rsEInvoice!ID)))(0) Like "*失败*" Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    End If
                End If
            End If
        
            rsEInvoice.MoveNext
        Loop
        
        .Cell(flexcpFontBold, .FixedRows, .ColIndex("打印结果"), .Rows - 1, .ColIndex("打印说明")) = True
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows - 1 > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngOldCol = 0 Or lngOldCol > .Cols - 1, .FixedCols, lngOldCol)
            .ShowCell .Row, .Col  '立刻显示到指定单元
        End If
        
        .Redraw = flexRDBuffered
    End With
    LoadEInvoiceData = True
    Exit Function
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim strIDKindStr As String
    
    Call CreateSquareCardObject(Me, mlngModule)
    strIDKindStr = "姓|姓名或就诊卡|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;手|手机号|0;单|收费单据号|0;票|电子票据号|0"
    Call IDKind.zlInit(Me, mlngSys, mlngModule, gcnOracle, mstrDBUser, mobjSquareCard, strIDKindStr, txtPatient)
    
    Call InitEInvoiceGrid
    Call InitExseGrid
    
    Dim varData As Variant, i As Integer
    varData = Array("1-收费", "2-预交", "3-结帐", "4-挂号", "5-就诊卡")
    cbo票据类型.Clear
    For i = 0 To UBound(varData)
        cbo票据类型.AddItem varData(i)
    Next
    cbo票据类型.ListIndex = 0
    
    dtp结束时间.Value = zlDatabase.Currentdate
    dtp开始时间.Value = Format(DateAdd("d", -7, dtp结束时间.Value), "yyyy-MM-dd 00:00:00")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If vsfEInvoice.Height + Y < 1200 Or vsfExse.Height - Y < 1200 Then Exit Sub

    fraSplit.Top = fraSplit.Top + Y
    
    vsfEInvoice.Height = vsfEInvoice.Height + Y
    vsfExse.Top = vsfExse.Top + Y
    vsfExse.Height = vsfExse.Height - Y
    Me.Refresh
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picFilter.Move 0, 0, picMain.ScaleWidth
    If vsfExse.Tag = "ExseGridHidden" Then
        vsfEInvoice.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth, picMain.ScaleHeight - (picFilter.Top + picFilter.Height)
    Else
        vsfEInvoice.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth, picMain.ScaleHeight * 2 / 3
        fraSplit.Move 0, vsfEInvoice.Top + vsfEInvoice.Height, picMain.ScaleWidth
        vsfExse.Move 0, fraSplit.Top + fraSplit.Height, picMain.ScaleWidth, picMain.ScaleHeight - (fraSplit.Top + fraSplit.Height)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
    Set mobjEInvoice = Nothing
End Sub

Private Function InitEInvoiceGrid() As Boolean
    '初始化VSFGrid表格控件
    Dim strHead As String, varData As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '列名1,对齐方式1,列宽1|列名2,对齐方式2,列宽2|...
    strHead = "选择,4,500|ID,1,0|单据号,1,1000|姓名,1,1000|性别,1,1000|年龄,1,1000|门诊号,1,1000|住院号,1,1000" & _
                    "|票据类型,1,1000|票据代码,1,2000|票据号码,1,2000|检验码,1,2000|票据金额,7,1000" & _
                    "|开票点,1,1000|开票时间,4,2000|换开纸质发票,3,2000|纸质发票号,1,2000" & _
                    "|打印结果,1,1000|打印说明,1,5000"
    With vsfEInvoice
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
    InitEInvoiceGrid = True
    Exit Function
ErrHandler:
    vsfEInvoice.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitExseGrid() As Boolean
    '初始化VSFGrid表格控件
    Dim strHead As String, varData As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '列名1,对齐方式1,列宽1|列名2,对齐方式2,列宽2|...
    strHead = "单据号,1,1000|开单科室,1,1000|开单人,1,800|费别,1,500|类别,4,800|名称,1,3000|商品名,1,3000" & _
                    "|规格,1,1200|单位,4,1000|数量,7,800|单价,7,1000|应收金额,7,1000|实收金额,7,1000|执行科室,4,1500"

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

        .RowHeightMin = 300

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

Private Sub vsfEInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Or NewRow < vsfEInvoice.FixedRows Then Exit Sub
    
    Call ShowEInvoiceExse(zlStr.NeedCode(cbo票据类型.Text), Val(vsfEInvoice.RowData(NewRow)))
    
    On Error Resume Next
    vsfEInvoice.ForeColorSel = vsfEInvoice.CellForeColor
End Sub

Private Function ShowEInvoiceExse(ByVal byt场合 As Byte, ByVal lngEInvoice As Long) As Boolean
    '显示费用明细
    '入参：
    '   byt场合 1-收费，2-预交，3-结帐，4-挂号，5-就诊卡
    Dim rsExse As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If byt场合 = 2 Then ShowEInvoiceExse = True: Exit Function
    
    With vsfExse
        .Clear 1
        .Rows = .FixedRows + 1
        
        If GetEInvoiceExse(byt场合, lngEInvoice, rsExse) = False Then Exit Function
        
        'NO,开单科室,开单人,费别,类别,名称,商品名,规格,单位,数量,单价,应收金额,实收金额,执行科室
        .Redraw = flexRDNone
        Do While Not rsExse.EOF
            If .TextMatrix(.Rows - 1, .ColIndex("单据号")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = Val(Nvl(rsExse!序号))
            .TextMatrix(.Rows - 1, .ColIndex("单据号")) = Nvl(rsExse!No)
            .TextMatrix(.Rows - 1, .ColIndex("开单科室")) = Nvl(rsExse!开单科室)
            .TextMatrix(.Rows - 1, .ColIndex("开单人")) = Nvl(rsExse!开单人)
            .TextMatrix(.Rows - 1, .ColIndex("费别")) = Nvl(rsExse!费别)
            .TextMatrix(.Rows - 1, .ColIndex("类别")) = Nvl(rsExse!类别)
            .TextMatrix(.Rows - 1, .ColIndex("名称")) = Nvl(rsExse!名称)
            .TextMatrix(.Rows - 1, .ColIndex("商品名")) = Nvl(rsExse!商品名)
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = Nvl(rsExse!规格)
            .TextMatrix(.Rows - 1, .ColIndex("单位")) = Nvl(rsExse!单位)
            .TextMatrix(.Rows - 1, .ColIndex("数量")) = Nvl(rsExse!数量)
            .TextMatrix(.Rows - 1, .ColIndex("单价")) = FormatEx(Val(Nvl(rsExse!单价)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("应收金额")) = FormatEx(Val(Nvl(rsExse!应收金额)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("实收金额")) = FormatEx(Val(Nvl(rsExse!实收金额)), 2, , , 6)
            .TextMatrix(.Rows - 1, .ColIndex("执行科室")) = Nvl(rsExse!执行科室)
        
            rsExse.MoveNext
        Loop
        
        .Redraw = flexRDBuffered
    End With
    ShowEInvoiceExse = True
    Exit Function
ErrHandler:
    vsfExse.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfEInvoice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsfEInvoice.ColIndex("选择") Or vsfEInvoice.TextMatrix(Row, 1) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vsfEInvoice_GotFocus()
    Call SetActiveList(vsfEInvoice)
End Sub

Private Sub vsfEInvoice_LostFocus()
    Call SetActiveList(vsfEInvoice, False)
End Sub

Private Sub vsfEInvoice_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not (Me.ActiveControl Is vsfEInvoice And Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(False)
End Sub

Private Sub vsfExse_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    On Error Resume Next
    vsfExse.ForeColorSel = vsfExse.CellForeColor
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
        vsfEInvoice.BackColorSel = &HE0E0E0
        vsfExse.BackColorSel = &HE0E0E0

        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &H8000000D '&HC0C0C0
    Else
        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub ExcutePrintEInvoice(ByVal byt场合 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印电子票据(A4纸)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lngEInvoiceID As Long, strDate As String, bln补结算 As Boolean
    Dim i As Long, lng冲销ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt场合) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("打印结果"), 1, .ColIndex("打印说明")) = "打印结果" & vbTab & "打印说明"
        .Cell(flexcpText, .FixedRows, .ColIndex("打印结果"), .Rows - 1, .ColIndex("打印说明")) = ""
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                If mobjPubEInvoice.zlPrintEInvoice(Me, lngEInvoiceID, False, strErrMsg) Then
                    .TextMatrix(i, .ColIndex("打印结果")) = "打印成功"
                    .TextMatrix(i, .ColIndex("打印说明")) = ""
                Else
                    .TextMatrix(i, .ColIndex("打印结果")) = "打印失败"
                    .TextMatrix(i, .ColIndex("打印说明")) = strErrMsg
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "请选择需要打印电子票据的记录。", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExcuteTurnPaper(ByVal byt场合 As Byte, ByVal bln重新换开 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行纸质发票换开或重新换开
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngEInvoiceID As Long, strEInvoiceCode As String, strEInvoiceNO As String, strInvoiceNO_Out As String
    Dim strNO As String, lng原结帐ID As Long, blnHaveEInvoice As Boolean
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim cllSwapData As Collection, int操作状态 As Integer, strUseDate As String
    Dim bln补结算 As Boolean, lng冲销ID As Long, strErrMsg As String
    Dim lngCount As Long, blnChecked As Boolean, blnFirst As Boolean
    Dim cllPati As Collection, cllBalance As Collection
    
    On Error GoTo ErrHandler
    Set mcllResult = New Collection
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt场合) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("打印结果"), 1, .ColIndex("打印说明")) = "换开结果" & vbTab & "换开说明"
        .Cell(flexcpText, .FixedRows, .ColIndex("打印结果"), .Rows - 1, .ColIndex("打印说明")) = ""
        
        blnFirst = True
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                lng原结帐ID = Val(.RowData(i))
                strEInvoiceCode = .TextMatrix(i, .ColIndex("票据代码"))
                strEInvoiceNO = .TextMatrix(i, .ColIndex("票据号码"))
                
                lng冲销ID = 0: bln补结算 = False
                If byt场合 = 2 Then
                    lng冲销ID = Val(Nvl(.Cell(flexcpData, i, .ColIndex("单据号"))))
                ElseIf byt场合 = 1 Or byt场合 = 4 Then
                    bln补结算 = Val(Nvl(.Cell(flexcpData, i, .ColIndex("单据号")))) = 1
                End If
            
                If .TextMatrix(i, .ColIndex("换开纸质发票")) = "" And bln重新换开 Then
                    strErrMsg = "无纸质票据换开记录，请执行[换开纸质票据]。"
                    mcllResult.Add Array("换开失败", strErrMsg), "_" & lngEInvoiceID
                ElseIf .TextMatrix(i, .ColIndex("换开纸质发票")) <> "" And Not bln重新换开 Then
                    strErrMsg = "已换开纸质票据，请执行[重新换开票据]。"
                    mcllResult.Add Array("换开失败", strErrMsg), "_" & lngEInvoiceID
                        
                ElseIf GetSwapCollectFromBalanceID(byt场合, lng原结帐ID, cllSwapData, bln补结算, lng冲销ID, False, strErrMsg) = False Then
                    mcllResult.Add Array("换开失败", strErrMsg), "_" & lngEInvoiceID
                Else
                    Set cllPati = cllSwapData("_PatiInfo")
                    strInvoiceNO_Out = GetNextPaperInvoice(Me, cllPati, 0, byt场合, blnFirst)
                    If strInvoiceNO_Out = "" Then
                        strErrMsg = "获取下一个有效票据号失败"
                        mcllResult.Add Array("换开失败", strErrMsg), "_" & lngEInvoiceID
                        If blnFirst Then Call LoadEInvoiceData: Exit Sub
                        Exit For
                    End If
                    blnFirst = False
                    
                    Set cllBalance = cllSwapData("_BalanceInfo")
                    cllBalance.Remove "_发票号"
                    cllBalance.Add strInvoiceNO_Out, "_发票号"
                    
                    If mobjPubEInvoice.zlTurnPaperInvoice(Me, byt场合, cllSwapData, lngEInvoiceID, _
                        strEInvoiceCode, strEInvoiceNO, strInvoiceNO_Out, int操作状态, , False, strErrMsg) Then
                        mcllResult.Add Array("换开成功", IIf(bln重新换开, "回收票据号：" & .TextMatrix(i, .ColIndex("纸质发票号")), "")), "_" & lngEInvoiceID
                    Else
                        mcllResult.Add Array("换开失败", strErrMsg), "_" & lngEInvoiceID
                    End If
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "请选择需要" & IIf(bln重新换开, "重新", "") & "换开纸质发票的记录。", vbInformation, gstrSysName
    End If
    
    Call LoadEInvoiceData
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetNextPaperInvoice(ByVal frmMain As Object, ByVal cllPatiInfo As Collection, _
    ByRef lng领用ID As Long, ByVal byt场合 As Byte, ByVal blnFirst As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一下张发票号
    '入参:
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '返回:发票号
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String, strErrMsg_Out As String
     
    '根据票据领用读取
    On Error GoTo ErrHandler
    If mobjPubEInvoice.zlGetNextInvoiceNo(frmMain, byt场合, strInvoiceNO, cllPatiInfo, lng领用ID, False, strErrMsg_Out) = False Then Exit Function
    
    If strInvoiceNO = "" Then
        If frmInputBox.InputBox(frmMain, "发票号确认", "无法获取将要使用的发票号，" & _
                        vbCrLf & "请你输入换开将要使用的发票号码：", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    ElseIf blnFirst Then
        If frmInputBox.InputBox(frmMain, "发票号确认", "请确认换开将要使用的发票号：", 30, 1, False, False, strInvoiceNO) = False Then Exit Function
    End If
    GetNextPaperInvoice = strInvoiceNO
    Exit Function
ErrHandler:
    Err.Clear
End Function

Private Sub ExcuteCancelTurnPaper(ByVal byt场合 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行纸质发票作废操作
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceNO As String, blnHaveEInvoice As Boolean
    Dim strNO As String, lng原结帐ID As Long, i As Long
    Dim rsTemp As ADODB.Recordset, lngEInvoiceID As Long
    Dim bln补结算 As Boolean, strErrMsg As String
    Dim lngCount As Long, blnChecked As Boolean
     
    On Error GoTo ErrHandler
    Set mcllResult = New Collection
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt场合) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("打印结果"), 1, .ColIndex("打印说明")) = "作废结果" & vbTab & "作废说明"
        .Cell(flexcpText, .FixedRows, .ColIndex("打印结果"), .Rows - 1, .ColIndex("打印说明")) = ""
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                lng原结帐ID = Val(.RowData(i))
                If .TextMatrix(i, .ColIndex("换开纸质发票")) = "" Then
                    strErrMsg = "无纸质票据换开记录。"
                    mcllResult.Add Array("作废失败", strErrMsg), "_" & lngEInvoiceID
                Else
                    bln补结算 = False
                    If byt场合 = 1 Or byt场合 = 4 Then
                        bln补结算 = Val(Nvl(.Cell(flexcpData, i, .ColIndex("单据号")))) = 1
                    End If
                    
                    strInvoiceNO = .TextMatrix(i, .ColIndex("纸质发票号"))
                    If mobjPubEInvoice.zlCancelPaperInvoice(Me, byt场合, strInvoiceNO, lng原结帐ID, lngEInvoiceID, , , , bln补结算, , False, strErrMsg) Then
                        mcllResult.Add Array("作废成功", ""), "_" & lngEInvoiceID
                    Else
                        mcllResult.Add Array("作废失败", strErrMsg), "_" & lngEInvoiceID
                    End If
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "请选择需要作废纸质发票的记录。", vbInformation, gstrSysName
    End If
    
    Call LoadEInvoiceData
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExecutePrintNotice(ByVal byt场合 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印告知单
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lngEInvoiceID As Long, strDate As String, bln补结算 As Boolean
    Dim i As Long, lng冲销ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt场合) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("打印结果"), 1, .ColIndex("打印说明")) = "打印结果" & vbTab & "打印说明"
        .Cell(flexcpText, .FixedRows, .ColIndex("打印结果"), .Rows - 1, .ColIndex("打印说明")) = ""
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                If mobjPubEInvoice.zlPrintNotice(Me, byt场合, lngEInvoiceID, False, strErrMsg) Then
                    .TextMatrix(i, .ColIndex("打印结果")) = "打印成功"
                    .TextMatrix(i, .ColIndex("打印说明")) = ""
                Else
                    .TextMatrix(i, .ColIndex("打印结果")) = "打印失败"
                    .TextMatrix(i, .ColIndex("打印说明")) = strErrMsg
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "请选择需要打印告知单的记录。", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExecuteSendMsg(ByVal byt场合 As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:推送消息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllSwapData As Collection, strErrMsg As String
    Dim lngEInvoiceID As Long, strDate As String, bln补结算 As Boolean
    Dim i As Long, lng冲销ID As Long, blnInit As Boolean
    Dim lngCount As Long, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    If GetPubEInvoiceObject(Me, mlngSys, mlngModule, mobjPubEInvoice, byt场合) = False Then Exit Sub
    With vsfEInvoice
        blnChecked = False
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then blnChecked = True: Exit For
        Next
        
        .Cell(flexcpText, 0, .ColIndex("打印结果"), 1, .ColIndex("打印说明")) = "推送结果" & vbTab & "推送说明"
        .Cell(flexcpText, .FixedRows, .ColIndex("打印结果"), .Rows - 1, .ColIndex("打印说明")) = ""
        .Cell(flexcpForeColor, .FixedRows, 0, .Rows - 1, .Cols - 1) = vbBlack
        
        For i = .FixedRows To .Rows - 1
            lngEInvoiceID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lngEInvoiceID <> 0 And (.Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Or Not blnChecked And i = .Row) Then
                lngCount = lngCount + 1
                
                If mobjPubEInvoice.zlSendEinvoiceMsg(Me, lngEInvoiceID, False, strErrMsg) Then
                    .TextMatrix(i, .ColIndex("打印结果")) = "推送成功"
                    .TextMatrix(i, .ColIndex("打印说明")) = ""
                Else
                    .TextMatrix(i, .ColIndex("打印结果")) = "推送失败"
                    .TextMatrix(i, .ColIndex("打印说明")) = strErrMsg
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                End If
            
                If Not blnChecked Then Exit For
            End If
        Next
    End With
    
    If lngCount = 0 Then
        MsgBox "请选择需要推送消息的记录。", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    
    If Me.ActiveControl Is vsfExse Then
        Set vsfGrid = vsfExse
        objOut.Title.Text = "电子票据明细清单"
    Else
        Set vsfGrid = vsfEInvoice
        objOut.Title.Text = "电子票据费用明细清单"
    End If
    
    '表项
    If Me.ActiveControl Is vsfExse Then
        Set objRow = New zlTabAppRow
        objRow.Add "票据代码：" & vsfEInvoice.TextMatrix(vsfEInvoice.Row, vsfEInvoice.ColIndex("票据代码"))
        objRow.Add "票据号码：" & vsfEInvoice.TextMatrix(vsfEInvoice.Row, vsfEInvoice.ColIndex("票据号码"))
        objOut.UnderAppRows.Add objRow
    Else
        Set objRow = New zlTabAppRow
        objRow.Add "票据类型：" & cbo票据类型.Text
        objRow.Add "费用时间：" & Format(dtp开始时间, "yyyy-mm-dd") & " 至 " & Format(dtp结束时间, "yyyy-mm-dd")
        objOut.UnderAppRows.Add objRow
    End If
    
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

Private Sub txtPatient_Change()
    txtPatient.Tag = ""
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnCancel As Boolean
     
    On Error GoTo ErrHandler
    If txtPatient.Locked Then Exit Sub

    If IDKind.GetCurCard.名称 Like "姓名*" Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If

    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    Call FindPati(IDKind.GetCurCard, blnCard)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, rsPatient As ADODB.Recordset
    Dim strInput As String
    
    RaiseEvent ShowInfo("")
    strInput = Trim(txtPatient.Text)
    If Val(txtPatient.Tag) <> 0 Then strInput = "-" & txtPatient.Tag
    
    If objCard.名称 = "收费单据号" Or objCard.名称 = "电子票据号" Then
        '
    Else
        If Not GetPatient(objCard, strInput, rsPatient, blnCancel, blnCard) Then
            If blnCancel Then '取消输入
                txtPatient.Text = ""
                zlControl.ControlSetFocus txtPatient
                Exit Sub
            End If
            RaiseEvent ShowInfo("未找到该病人，请检查输入内容!")
            If blnCard Then
                txtPatient.Text = ""
                '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
                txtPatient.PasswordChar = ""
                txtPatient.IMEMode = 0
            Else
                zlControl.TxtSelAll txtPatient
            End If
            zlControl.ControlSetFocus txtPatient
            Exit Sub
        End If
        
        txtPatient.Text = Nvl(rsPatient!姓名)
        txtPatient.Tag = Val(Nvl(rsPatient!病人ID))
    End If
    
    Call cmdRefresh_Click
    
    zlControl.ControlSetFocus txtPatient
    zlControl.TxtSelAll txtPatient
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    ByRef rsPatient As ADODB.Recordset, ByRef blnCancel As Boolean, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:
    '   objCard-指定的卡类别
    '   strInput-输入的值
    '   blnCard-是否刷卡
    '出参:
    '   rsPatient-病人信息，字段：病人ID,姓名
    '   blnCancel-是否取消输入
    '返回:读取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng病人ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, lng卡类别ID As Long

    On Error GoTo ErrHandler
    blnCancel = False
    strWhere = ""
    If blnCard And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        If mobjSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then Exit Function
        If lng病人ID <= 0 Then Exit Function
        strInput = "-" & lng病人ID
        strWhere = strWhere & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  '住院号(对住(过)院的病人)
        strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
        strWhere = strWhere & " And A.门诊号=[1]"
        '75087,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNo(Mid(strInput, 2), 3)
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                strPati = _
                    " Select A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.住院号,B.名称 as 科室,A.当前床号 as 床号,A.出生日期,A.身份证号,A.家庭地址" & _
                    " From 病人信息 A,部门表 B" & _
                    " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And A.姓名 Like [1]" & _
                    " Order by A.姓名"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                If rsTmp Is Nothing Then Exit Function
                strInput = rsTmp!病人ID
                strWhere = strWhere & " And A.病人ID=[2]"
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.医保号=[2]"
            Case "身份证号", "二代身份证", "身份证"
                strInput = UCase(strInput)
                If mobjSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If mobjSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
                strInput = zlCommFun.GetFullNo(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If mobjSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Function
                Else
                    If mobjSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Function
                End If
                If lng病人ID <= 0 Then Exit Function
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
        End Select
    End If
    
    strSQL = "Select A.病人ID,A.姓名 From 病人信息 A Where A.停用时间 is NULL" & strWhere
    Set rsPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If rsPatient.EOF Then Exit Function
    
    GetPatient = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub IDKind_ItemClick(index As Integer, objCard As Card)
    txtPatient.IMEMode = 0
    txtPatient.Text = ""
    zlControl.ControlSetFocus txtPatient: zlControl.TxtSelAll txtPatient
End Sub

'10.35.130:Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
'10.35.140以后:
Private Sub IDKind_ReadCard(ByVal objCard As Card, objPatiInfor As clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Function CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    On Error Resume Next
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zlOneCardComLib.clsOneCardComLib")
        If Err <> 0 Then Exit Function
    End If
    
    'Public Function zlInitComponents(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
        ByVal cnOracle As ADODB.Connection, _
        Optional blnDeviceSet As Boolean = False, _
        Optional strExpand As String) As Boolean
    '功能:zlInitComponents (初始化接口部件)
    '入参: frmMain-调用的主窗体
    '        lngModule-HIS调用模块号
    '       lngSys-传入的系统号
    '       strDBUser-数据库用户名
    '       cnOracle -HIS/三方机构
    '       blnDeviceSet-设备设置调用初始化
    '       strExpand-扩展信息(可选传入:卡类别ID-不传时,表示全部初始化,传入时,只初始化指定的接口)
    '返回:函数返回True:调用成功,False:调用失败
    If mobjSquareCard.zlInitComponents(frmMain, lngModule, mlngSys, mstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Set mobjSquareCard = Nothing
         Exit Function
    End If
    CreateSquareCardObject = True
End Function
