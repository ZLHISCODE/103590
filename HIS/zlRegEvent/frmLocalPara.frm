VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLocalPara 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "本机参数设置"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "frmLocalPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab tbPage 
      Height          =   5385
      Left            =   45
      TabIndex        =   23
      Top             =   105
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9499
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frmLocalPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstDept"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDefaultSet"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkRigistHeadSort"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDeviceSetup"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPrintSet(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrintSet(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrintSet(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdPrintSet(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "共用票据"
      TabPicture(1)   =   "frmLocalPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCards"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraTitle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "预约挂号单打印设置"
         Height          =   345
         Index           =   2
         Left            =   2340
         TabIndex        =   14
         Top             =   3840
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "挂号凭条打印设置"
         Height          =   345
         Index           =   3
         Left            =   2340
         TabIndex        =   12
         Top             =   3405
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "病人条码打印设置"
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   3405
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "挂号票据打印设置"
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   3840
         Width           =   1875
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   4275
         Width           =   1425
      End
      Begin VB.CheckBox chkRigistHeadSort 
         Caption         =   "挂号安排表点击列头排序"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2325
      End
      Begin VB.Frame fraTitle 
         Caption         =   "共用挂号票据"
         Height          =   1845
         Left            =   -74835
         TabIndex        =   19
         Top             =   525
         Width           =   6675
         Begin MSComctlLib.ListView lvwBill 
            Height          =   1455
            Left            =   150
            TabIndex        =   20
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483630
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "领用人"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "领用日期"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "号码范围"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "剩余"
               Object.Width           =   1499
            EndProperty
         End
      End
      Begin VB.Frame fraDefaultSet 
         Caption         =   "缺省值"
         Height          =   2445
         Left            =   240
         TabIndex        =   25
         Top             =   825
         Width           =   4290
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1590
            Width           =   2865
         End
         Begin VB.ComboBox cboDefaultSex 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1995
            Width           =   1260
         End
         Begin VB.ComboBox cboDefaultPayMode 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   405
            Width           =   2865
         End
         Begin VB.ComboBox cboDefaultFeeType 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   780
            Width           =   2865
         End
         Begin VB.ComboBox cboDefaultBalance 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1185
            Width           =   2865
         End
         Begin VB.Label lblDefaultPayCard 
            AutoSize        =   -1  'True
            Caption         =   "发卡类型"
            Height          =   180
            Left            =   420
            TabIndex        =   7
            Top             =   1650
            Width           =   720
         End
         Begin VB.Label lblDefaultSex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            Height          =   180
            Left            =   780
            TabIndex        =   9
            Top             =   2055
            Width           =   360
         End
         Begin VB.Label lblDefaultPayMode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款方式"
            Height          =   180
            Left            =   420
            TabIndex        =   1
            Top             =   465
            Width           =   720
         End
         Begin VB.Label lblDefaultBalance 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算方式"
            Height          =   180
            Left            =   420
            TabIndex        =   5
            Top             =   1245
            Width           =   720
         End
         Begin VB.Label lblDefaultFeeType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费别"
            Height          =   180
            Left            =   780
            TabIndex        =   3
            Top             =   840
            Width           =   360
         End
      End
      Begin VB.ListBox lstDept 
         ForeColor       =   &H80000012&
         Height          =   4470
         Left            =   4785
         Style           =   1  'Checkbox
         TabIndex        =   17
         ToolTipText     =   "Ctrl+A全选,Ctrl+C全消,如果一个都未选则表示不限制科室"
         Top             =   690
         Width           =   2175
      End
      Begin VB.Frame fraCards 
         Caption         =   "本地共用医疗卡"
         Height          =   2655
         Left            =   -74805
         TabIndex        =   21
         Top             =   2550
         Width           =   6660
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2190
            Left            =   195
            TabIndex        =   22
            Top             =   300
            Width           =   6405
            _cx             =   11298
            _cy             =   3863
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmLocalPara.frx":0044
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
            ExplorerBar     =   2
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号科室"
         Height          =   180
         Left            =   4770
         TabIndex        =   16
         ToolTipText     =   "设定本机可挂哪些科室的号"
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   330
      Left            =   7260
      TabIndex        =   18
      Top             =   5100
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   7260
      TabIndex        =   26
      Top             =   885
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   330
      Left            =   7260
      TabIndex        =   24
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "frmLocalPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngModul As Long
Private Sub chkDeptBespeakOneNum_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDefaultBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDefaultFeeType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDefaultPayMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
 
Private Sub cboDefaultSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
 
Private Sub chkRigistHeadSort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1111)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim blnHavePrivs As Boolean
    
    On Error GoTo Hd
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '数据库存储的模块参数
    '-------------------------------------------------------------------------------------------
    zlDatabase.SetPara "缺省付款方式", cboDefaultPayMode.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省费别", cboDefaultFeeType.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省结算方式", cboDefaultBalance.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省性别", cboDefaultSex.Text, glngSys, mlngModul, blnHavePrivs
     '问题 43847
    zlDatabase.SetPara "允许列头排序", chkRigistHeadSort.Value, glngSys, mlngModul, blnHavePrivs
    
    strTmp = ""
    If lstDept.ListCount <> lstDept.SelCount Then
        For i = 0 To lstDept.ListCount - 1
            If lstDept.Selected(i) = True Then
                strTmp = strTmp & "," & lstDept.ItemData(i)
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End If
    zlDatabase.SetPara "挂号科室", strTmp, glngSys, mlngModul, blnHavePrivs
    
    '共用挂号票据批次
    strTmp = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTmp = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    zlDatabase.SetPara "共用挂号票据批次", strTmp, glngSys, mlngModul, blnHavePrivs
    
    Call SaveInvoice
    Call InitLocPar(mlngModul)
    gblnOk = True
    Unload Me
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Function LoadFactList(bytKind As Byte) As Boolean
'功能：读取可用公用挂号票据或就诊卡领用
'参数:bytKind=4-挂号票据,5-就诊卡
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim ObjItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = zlDatabase.GetPara("共用挂号票据批次", glngSys, mlngModul, 0, Array(lvwBill), InStr(mstrPrivs, "参数设置") > 0)
    Set rsTmp = GetShareInvoiceGroupID(bytKind)
    
    For i = 1 To rsTmp.RecordCount
        Set ObjItem = lvwBill.ListItems.Add(, "_" & rsTmp!id, rsTmp!领用人)
        ObjItem.SubItems(1) = Format(rsTmp!登记时间, "yyyy-MM-dd")
        ObjItem.SubItems(2) = rsTmp!开始号码 & "," & rsTmp!终止号码
        ObjItem.SubItems(3) = rsTmp!剩余数量
        If rsTmp!id = lngTmp Then
            ObjItem.Checked = True
            ObjItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        zlDatabase.SetPara IIf(bytKind = 4, "共用挂号票据批次", "共用就诊卡批次"), "0", glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdPrintSet_Click(index As Integer)
    On Error GoTo Hd
    Select Case index
    '病人条码打印
    Case 0:
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me)
    Case 1:
        '挂号收费打印
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 2:
        '预约挂号打印   '56274
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    Case 3:
        '68408,刘尔旋,2013-12-11,挂号凭条打印
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me)
    Case Else:
    End Select
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = False
            Next
        End If
    End If
End Sub

Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:冉俊明
    '日期:2014-07-02
    '问题号:74552
    '说明:挂号管理中设置默认结算方式时候可以选择结算方式性质为"7-一卡通结算"的结算方式
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String

    strSQL = _
        " Select B.编码,B.名称,Nvl(B.缺省标志,0) as 缺省,Nvl(B.性质,1) as 性质,Nvl(B.应付款,0) as 应付款" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式" & _
        "   And(B.性质<>7 Or B.性质=7 And Exists(Select 1 From 一卡通目录 C Where C.结算方式=B.名称 And C.启用=1))" & _
        "   and B.性质<>8 And Instr(',1,2,7,',','||B.性质||',')>0" & _
        " Order by 性质,lpad(编码,3,' ')"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "挂号")
    
    '获取三方卡的结算方式
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    
    varData = Split(strPayType, ";")
    With cboDefaultBalance
        .Clear
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True: Exit For
                End If
            Next
                         
            If Not blnFind Then
                .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = 1
                If Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称) = gstr结算方式 Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
                If Val(Nvl(rsTemp!缺省)) = 1 Then .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        
        '加载结算方式性质为“7-一卡通结算”的医疗卡类别
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                If varTemp(1) = gstr结算方式 Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp           As New ADODB.Recordset
    Dim strSQL          As String
    Dim i               As Integer
    Dim str科室ID       As String
    Dim strTmp          As String
    Dim blnParSet       As Boolean

    
    gblnOk = False
    
    blnParSet = InStr(mstrPrivs, "参数设置") > 0
    On Error GoTo errH
    'a.初始数据
    '----------------------------------------------------------------------------------------
    strSQL = "Select Distinct B.编码 ||'-'|| B.名称 as 名称,B.ID From 挂号安排 A,部门表 B Where A.科室ID=B.ID Order by 名称"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    zlcontrol.CboAddData lstDept, rsTmp, True
    
    strSQL = "Select '医疗付款方式' 分类,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 医疗付款方式" & _
            " Union All " & _
            " Select '性别' 分类,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别" & _
            " Union All " & _
            " Select '费别' 分类,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别" & _
            " Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(1,3)" & _
            " Order by 分类,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '缺省医疗付款方式
    rsTmp.Filter = "分类='医疗付款方式'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultPayMode.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then cboDefaultPayMode.ListIndex = cboDefaultPayMode.NewIndex
        rsTmp.MoveNext
    Next
     '缺省费别    '不是仅限初诊身份唯一性项目(包含了缺省费别),不管有效期间及科室
    rsTmp.Filter = "分类='费别'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultFeeType.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then cboDefaultFeeType.ListIndex = cboDefaultFeeType.NewIndex
        rsTmp.MoveNext
    Next
    
    '缺省性别
    rsTmp.Filter = "分类='性别'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultSex.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then cboDefaultSex.ListIndex = cboDefaultSex.NewIndex
        rsTmp.MoveNext
    Next
    cboDefaultSex.AddItem "无"
    '缺省结算方式
    Call Load支付方式

    strTmp = zlDatabase.GetPara("缺省付款方式", glngSys, mlngModul, , Array(cboDefaultPayMode), blnParSet)
    zlcontrol.CboLocate cboDefaultPayMode, strTmp
    strTmp = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , Array(cboDefaultFeeType), blnParSet)
    zlcontrol.CboLocate cboDefaultFeeType, strTmp
    strTmp = zlDatabase.GetPara("缺省性别", glngSys, mlngModul, , Array(cboDefaultSex), blnParSet)
    zlcontrol.CboLocate cboDefaultSex, strTmp
    If cboDefaultSex.ListIndex = -1 Or strTmp = "无" Then cboDefaultSex.ListIndex = cboDefaultSex.ListCount - 1
    strTmp = zlDatabase.GetPara("缺省结算方式", glngSys, mlngModul, , Array(cboDefaultBalance), blnParSet)
    zlcontrol.CboLocate cboDefaultBalance, strTmp
 
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    chkRigistHeadSort.Value = IIf(zlDatabase.GetPara("允许列头排序", glngSys, mlngModul, , Array(chkRigistHeadSort), blnParSet) = "1", 1, 0)
    '读取可用的挂号科室
    str科室ID = zlDatabase.GetPara("挂号科室", glngSys, mlngModul, , Array(lstDept), blnParSet)
    If str科室ID = "" Then
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = True
        Next
    Else
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = InStr(1, "," & str科室ID & ",", "," & lstDept.ItemData(i) & ",") > 0
        Next
    End If
    If lstDept.ListCount > 0 Then lstDept.TopIndex = 0: lstDept.ListIndex = 0
    
    '读取可用公用挂号票据领用
    Call LoadFactList(4)
    
    '读取公用的就诊卡领用
     Call InitShareInvoice
    If tbPage.TabVisible(0) Then tbPage.Tab = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
    Next
    Item.Selected = True
End Sub
Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , , True, intType))
    '90875:李南春,2016/11/8,医疗卡证件类型
    gstrSQL = "Select ID,编码,名称, nvl(是否固定,0) as 是否固定  from 医疗卡类别  Where nvl(是否启用,0)=1 And nvl(是否证件,0)=0 "
    On Error GoTo Hd
    Set rs医疗卡类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rs医疗卡类别.Filter = "名称='就诊卡' and 是否固定=1"
    If rs医疗卡类别.EOF = False Then
        str缺省医疗卡 = rs医疗卡类别!名称: lng缺省医疗卡 = Val(rs医疗卡类别!id)
    End If
    With rs医疗卡类别
        cboType.Clear
        rs医疗卡类别.Filter = 0
        If rs医疗卡类别.RecordCount <> 0 Then rs医疗卡类别.MoveFirst
        Do While Not .EOF
            cboType.AddItem Nvl(!名称)
            cboType.ItemData(cboType.NewIndex) = Nvl(!id)
            If Nvl(!名称) = "就诊卡" And cboType.ListIndex < 0 Then cboType.ListIndex = cboType.NewIndex
            If lngCardTypeID = Val(Nvl(!id)) Then
                cboType.ListIndex = cboType.NewIndex
            End If
            .MoveNext
        Loop
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
    strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, , , True)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,医疗卡类别ID1|领用IDn,医疗卡类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(5)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            If Val(Nvl(rsTemp!使用类别ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = str缺省医疗卡
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = lng缺省医疗卡
            Else
                rs医疗卡类别.Filter = "ID=" & Val(Nvl(rsTemp!使用类别ID))
                If Not rs医疗卡类别.EOF Then
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = Nvl(rs医疗卡类别!名称)
                Else
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = Nvl(rsTemp!使用类别)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = Val(Nvl(rsTemp!使用类别ID))
            End If
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("医疗卡类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long, lng卡类别ID As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '保存共享票据
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("医疗卡类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用医疗卡批次", strValue, glngSys, mlngModul, blnHavePrivs
    If cboType.ListIndex >= 0 Then
        lng卡类别ID = cboType.ItemData(cboType.ListIndex)
    End If
    Call zlDatabase.SetPara("缺省医疗卡类别", lng卡类别ID, glngSys, mlngModul, blnHavePrivs)
End Sub
 
