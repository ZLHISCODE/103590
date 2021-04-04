VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetPar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frmSetPar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7170
      TabIndex        =   2
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7200
      TabIndex        =   0
      Top             =   450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin TabDlg.SSTab sTab 
      Height          =   4830
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   8520
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "基本控制(&1)"
      TabPicture(0)   =   "frmSetPar.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFee"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSeekName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDeviceSetup"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNameDays"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkLedWelcome"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbo预交结算"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboFee"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdPrintSet(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdPrintSet(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "预交票据控制(&2)"
      TabPicture(1)   =   "frmSetPar.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "img16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraPrepay"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPrintSet(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "医疗卡票据控制(&3)"
      TabPicture(2)   =   "frmSetPar.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl缺省发卡"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cboType"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraTitle"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "病人腕带打印设置"
         Height          =   345
         Index           =   2
         Left            =   4845
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "病案首页打印设置"
         Height          =   345
         Index           =   1
         Left            =   4845
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "预交款票据打印设置"
         Height          =   345
         Index           =   0
         Left            =   -70215
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3090
         Width           =   1815
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用预交票据"
         Height          =   2475
         Left            =   -74895
         TabIndex        =   19
         Top             =   450
         Width           =   6510
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   2055
            Left            =   60
            TabIndex        =   20
            Top             =   300
            Width           =   6285
            _cx             =   11086
            _cy             =   3625
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
            FormatString    =   $"frmSetPar.frx":0060
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
      Begin VB.ComboBox cboFee 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1230
         Width           =   2580
      End
      Begin VB.ComboBox cbo预交结算 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1590
         Width           =   2580
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   465
         TabIndex        =   4
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   600
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用医疗卡"
         Height          =   3570
         Left            =   -74865
         TabIndex        =   15
         Top             =   450
         Width           =   6390
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   3150
            Left            =   60
            TabIndex        =   16
            Top             =   300
            Width           =   6150
            _cx             =   10848
            _cy             =   5556
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
            FormatString    =   $"frmSetPar.frx":013F
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
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73695
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   4260
         Width           =   2580
      End
      Begin VB.TextBox txtNameDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   180
         Left            =   2850
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "0表示查找时不限制时间"
         Top             =   885
         Width           =   285
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2835
         TabIndex        =   6
         Top             =   1080
         Width           =   285
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   4845
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1860
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -70560
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSetPar.frx":0221
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSeekName 
         AutoSize        =   -1  'True
         Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
         Height          =   180
         Left            =   465
         TabIndex        =   5
         Top             =   915
         Width           =   3960
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "缺省费别"
         Height          =   180
         Left            =   465
         TabIndex        =   8
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "缺省缴款方式"
         Height          =   225
         Left            =   105
         TabIndex        =   10
         Top             =   1665
         Width           =   1290
      End
      Begin VB.Label lbl缺省发卡 
         Caption         =   "缺省发卡类型"
         Height          =   225
         Left            =   -74865
         TabIndex        =   17
         Top             =   4320
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmSetPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngModul As Long



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1131)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    Call SaveInvoice
    
    zlDatabase.SetPara "姓名查找天数", Val(txtNameDays.Text), glngSys, mlngModul, IIf(txtNameDays.Enabled = True, True, False)

    'LED设备
    zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset, objItem As ListItem
    Dim blnBill As Boolean
    
    gblnOK = False
    On Error GoTo errH
    
    Call InitShareInvoice
    
    txtNameDays.Text = Val(zlDatabase.GetPara("姓名查找天数", glngSys, mlngModul, , Array(txtNameDays), InStr(mstrPrivs, "参数设置") > 0))
    txtNameDays.Enabled = Val(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModul)) = 1
    
    'LED设备
    chkLedWelcome.Value = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, "参数设置") > 0)

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
  

Private Sub sTab_Click(PreviousTab As Integer)
    If sTab.Tab = 0 Then

    ElseIf sTab.Tab = 1 Then
        If vsPrepay.Enabled And vsPrepay.Visible Then vsPrepay.SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
     End If
End Sub

Private Sub txtNameDays_GotFocus()
    Call zlControl.TxtSelAll(txtNameDays)
End Sub

Private Sub txtNameDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNameDays_Validate(Cancel As Boolean)
    If Val(txtNameDays.Text) <= 0 Then
        txtNameDays.Text = 0
    ElseIf Val(txtNameDays.Text) > 999 Then
        txtNameDays.Text = 999
    End If
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
    Dim lngTemp As Long, i As Long, strSql As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    Dim str缺省费别 As String
    
    On Error GoTo Errhand
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , Array(cboType), blnHavePrivs, intType))
    '90875:李南春,2016/11/8,医疗卡证件类型
    gstrSQL = "Select ID,编码,名称, nvl(是否固定,0) as 是否固定  from 医疗卡类别  Where nvl(是否启用,0)=1 And nvl(是否证件,0)=0 "
    
    Set rs医疗卡类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rs医疗卡类别.Filter = "名称='就诊卡' and 是否固定=1"
    If rs医疗卡类别.EOF = False Then
        str缺省医疗卡 = rs医疗卡类别!名称: lng缺省医疗卡 = Val(rs医疗卡类别!ID)
    End If
    With rs医疗卡类别
        cboType.Clear
        rs医疗卡类别.Filter = 0
        If rs医疗卡类别.RecordCount <> 0 Then rs医疗卡类别.MoveFirst
        Do While Not .EOF
            cboType.AddItem NVL(!名称)
            cboType.ItemData(cboType.NewIndex) = NVL(!ID)
            If NVL(!名称) = "就诊卡" Then cboType.ListIndex = cboType.NewIndex
            .MoveNext
        Loop
    End With
    '问题号:58776
    For i = 0 To cboType.ListCount - 1
        If Val(cboType.ItemData(i)) = lngCardTypeID Then
             cboType.ListIndex = i
        End If
    Next
    
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
    strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, , , True, intType)
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
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!使用类别ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = str缺省医疗卡
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = lng缺省医疗卡
            Else
                rs医疗卡类别.Filter = "ID=" & Val(NVL(rsTemp!使用类别ID))
                If Not rs医疗卡类别.EOF Then
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = NVL(rs医疗卡类别!名称)
                Else
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = NVL(rsTemp!使用类别)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = Val(NVL(rsTemp!使用类别ID))
            End If
            .TextMatrix(lngRow, .ColIndex("领用人")) = NVL(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(NVL(rsTemp!剩余数量)), "##0;-##0;;")
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
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModul, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,预交类别ID1|领用IDn,预交类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        rsTemp.Filter = " 使用类别<>1   "   '不包含预交门诊票据
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!使用类别, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "门诊和住院共用"
            ElseIf Val(NVL(rsTemp!使用类别, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
            Else
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = Val(NVL(rsTemp!使用类别))
            
            .TextMatrix(lngRow, .ColIndex("领用人")) = NVL(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(NVL(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("预交类型"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    '加载缺省缴款方式(预交款)
    Load缴款方式
    '加载费别
    strSql = "Select A.编码,A.名称,A.简码,Nvl(A.缺省标志,0) as 缺省 From 费别 A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
             " Where (A.服务对象 = B.Column_Value or A.服务对象 is null) And A.属性=1 And Nvl(A.仅限初诊,0)=0 And  " & _
             " (a.有效开始 Is Null And a.有效结束 Is Null Or Trunc(Sysdate) Between a.有效开始 And a.有效结束) Order by A.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "1,2,3")
    cboFee.Clear
    Do While Not rsTemp.EOF
        cboFee.AddItem rsTemp!名称
        If rsTemp!缺省 = 1 Then cboFee.ListIndex = cboFee.NewIndex
    rsTemp.MoveNext
    Loop
    If cboFee.ListCount > 0 And cboFee.ListIndex < 0 Then cboFee.ListIndex = 0
    str缺省费别 = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , blnHavePrivs)
    If str缺省费别 <> "" Then
        For i = 0 To cboFee.ListCount - 1
            If cboFee.List(i) = str缺省费别 Then
                cboFee.ListIndex = i
            End If
        Next
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    Dim lng卡类别ID As Long
    If cboType.ListIndex >= 0 Then
        lng卡类别ID = cboType.ItemData(cboType.ListIndex)
    End If
    zlDatabase.SetPara "缺省医疗卡类别", lng卡类别ID, glngSys, mlngModul, blnHavePrivs
        
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
    '保存预交票据
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Val(.Cell(flexcpData, i, .ColIndex("预交类型")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用预交票据批次", strValue, glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "缺省缴款方式", Trim(cbo预交结算.Text), glngSys, mlngModul, blnHavePrivs
    '69489
    zlDatabase.SetPara "缺省费别", Trim(cboFee.Text), glngSys, mlngModul, blnHavePrivs
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共用预交票据列表", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("医疗卡类别"))) = Trim(.Cell(flexcpData, i, .ColIndex("医疗卡类别"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("选择")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub
Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("预交类型"))) = Trim(.Cell(flexcpData, i, .ColIndex("预交类型"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("选择")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Public Sub Load缴款方式()
    Dim strTemp As String, str缺省预交款方式 As String
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim objSquareCard As Object
    Dim varData As Variant, varTemp As Variant
    Dim strPayType As String
    Dim j As Long, i As Long
    Dim blnFind As Boolean, blnHavePrivs As Boolean
    
    strTemp = "1,2,5,7,8" & IIf(InStr(mstrPrivs, ";保险病人登记;") > 0, ",3", "")

    
    strSql = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合 ='预交款'  And B.名称=A.结算方式  " & _
        "           And Nvl(B.性质,1) In(" & strTemp & ")" & _
        " Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cbo预交结算
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = NVL(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind Then
                .AddItem NVL(rsTemp!名称)
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                .ItemData(.NewIndex) = Val(NVL(rsTemp!性质))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                j = j + 1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
        str缺省预交款方式 = zlDatabase.GetPara("缺省缴款方式", glngSys, mlngModul, , blnHavePrivs)
        If str缺省预交款方式 <> "" Then
            For i = 0 To cbo预交结算.ListCount
                If cbo预交结算.List(i) = str缺省预交款方式 Then
                    cbo预交结算.ListIndex = i
                End If
            Next
        End If
    End With
End Sub
