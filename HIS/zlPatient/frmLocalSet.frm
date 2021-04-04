VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sTab 
      Height          =   4755
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   8387
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "输入控制(&1)"
      TabPicture(0)   =   "frmLocalSet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkAutoRefresh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "医疗卡票据控制(&2)"
      TabPicture(1)   =   "frmLocalSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboType"
      Tab(1).Control(1)=   "chkBrushCardVerfy"
      Tab(1).Control(2)=   "chkBruhCardBackCard"
      Tab(1).Control(3)=   "fraTitle"
      Tab(1).Control(4)=   "cmdDeviceSetup(0)"
      Tab(1).Control(5)=   "img16"
      Tab(1).Control(6)=   "lblDefaultPayCard"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "预交款设置(&3)"
      TabPicture(2)   =   "frmLocalSet.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPrepay"
      Tab(2).Control(1)=   "cmdDeviceSetup(1)"
      Tab(2).Control(2)=   "chkLedWelcome"
      Tab(2).Control(3)=   "cboDefaultBalance"
      Tab(2).Control(4)=   "lblEdit"
      Tab(2).ControlCount=   5
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用票据"
         Height          =   3315
         Left            =   -74925
         TabIndex        =   11
         Top             =   525
         Width           =   5865
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   2925
            Left            =   90
            TabIndex        =   12
            Top             =   270
            Width           =   5670
            _cx             =   10001
            _cy             =   5159
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
            FormatString    =   $"frmLocalSet.frx":0054
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
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Index           =   1
         Left            =   -70560
         TabIndex        =   16
         Top             =   4260
         Width           =   1500
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   -74835
         TabIndex        =   13
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   4020
         Value           =   1  'Checked
         Width           =   1710
      End
      Begin VB.CheckBox chkAutoRefresh 
         Caption         =   "切换病人类型选项卡时，自动刷新病人数据"
         Height          =   180
         Left            =   285
         TabIndex        =   3
         Top             =   555
         Width           =   3840
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73665
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4050
         Width           =   2580
      End
      Begin VB.CheckBox chkBrushCardVerfy 
         Caption         =   "退卡获取单据号后刷卡验证退卡"
         Height          =   180
         Left            =   -74835
         TabIndex        =   6
         Top             =   3540
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.CheckBox chkBruhCardBackCard 
         Caption         =   "发卡按“退”刷卡退卡"
         Height          =   240
         Left            =   -74835
         TabIndex        =   7
         Top             =   3795
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.ComboBox cboDefaultBalance 
         Height          =   300
         Left            =   -73725
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4290
         Width           =   1875
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用..."
         Height          =   2880
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   5745
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2445
            Left            =   60
            TabIndex        =   5
            Top             =   300
            Width           =   5595
            _cx             =   9869
            _cy             =   4313
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
            FormatString    =   $"frmLocalSet.frx":0131
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
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Index           =   0
         Left            =   -70605
         TabIndex        =   10
         Top             =   4020
         Width           =   1500
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -71145
         Top             =   855
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
               Picture         =   "frmLocalSet.frx":0212
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDefaultPayCard 
         Caption         =   "缺省发卡类型"
         Height          =   210
         Left            =   -74835
         TabIndex        =   8
         Top             =   4095
         Width           =   1290
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "缺省结算方式"
         Height          =   180
         Left            =   -74850
         TabIndex        =   14
         Top             =   4350
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6240
      TabIndex        =   2
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6240
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Width           =   1100
   End
End
Attribute VB_Name = "frmLocalSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlngModul As Long, mstrPrivs As String, mbln担保 As Boolean
Private mstrClass As String, mstrDeposit As String
Private mblnOK As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal strPrivs As String, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置
    '入参:mlngModul-1101-病人信息管理,1102-就诊卡管理,1103-预交款管理
    '出参:
    '返回:保存,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 14:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mstrPrivs = strPrivs: mlngModul = lngModule
    mbln担保 = InStr(mstrPrivs, ";担保信息;") > 0 And mlngModul = 1101
    Me.Show 1, frmMain
    zlSetPara = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDeviceSetup_Click(Index As Integer)
    Call zlCommFun.DeviceSetup(Me, 100, mlngModul)
End Sub

Private Sub cmdHelp_Click()
    Select Case mlngModul
        Case 1101 '病人信息
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet1"
        Case 1102 '就诊卡
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet2"
        Case 1103 '预交款
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet3"
    End Select
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
    IsValied = False
    
    On Error GoTo errHandle
    If mlngModul <> 1103 Then
        '检查每种使用种式只能一个选择
        With vsBill
            str类别 = "-"
            For i = 1 To vsBill.Rows - 1
                If str类别 <> Trim(.TextMatrix(i, .ColIndex("医疗卡类别"))) Then
                   str类别 = Trim(.TextMatrix(i, .ColIndex("医疗卡类别")))
                   lngSelCount = 0
                    For j = 1 To vsBill.Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("医疗卡类别"))) = Trim(.TextMatrix(j, .ColIndex("医疗卡类别"))) Then
                            If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                                lngSelCount = lngSelCount + 1
                            End If
                        End If
                    Next
                    If lngSelCount > 1 Then
                        MsgBox "注意:" & vbCrLf & "    医疗卡类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    If mlngModul = 1102 Then IsValied = True: Exit Function
  '检查每种使用预交只能一个选择
    With vsPrepay
        str类别 = "-"
        For i = 1 To .Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("预交类型"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("预交类型")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("预交类型"))) = Trim(.TextMatrix(j, .ColIndex("预交类型"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    预交类型为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If mlngModul <> 1103 Then
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
    End If
    If mlngModul = 1102 Then Exit Sub
    
    
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
    Dim intTYPE As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    Dim strBillFormat As String
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    On Error GoTo errHandle
    '恢复列宽度
    If mlngModul <> 1103 Then
            lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , , True, intTYPE))
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
                    If lngCardTypeID = Val(NVL(!ID)) Then
                        cboType.ListIndex = cboType.NewIndex
                    End If
                    .MoveNext
                Loop
            End With
            
            zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
            strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, , , True, intTYPE)
            '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
            vsBill.Tag = ""
            Select Case intTYPE
            Case 1, 3, 5, 15
                vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
                fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
                If intTYPE = 5 Then vsBill.Tag = ""
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
                    '105985:李南春,2017/4/10,以医疗卡名称区分票据
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
    End If
    If mlngModul = 1102 Then Exit Sub
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModul, , , True, intTYPE)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
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
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            '58071
            Select Case Val(NVL(rsTemp!使用类别, ""))
            Case 0 '不区分门诊和住院票据
                .TextMatrix(lngRow, .ColIndex("预交类型")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = 0
            Case 1  '门诊票据
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
                .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = 1
            Case Else   '住院票据
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
                .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = 2
            End Select
            
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
 
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, strTmp As String
    
    '本地共用就诊卡
    If IsValied = False Then Exit Sub
    Call SaveInvoice
    
    Select Case mlngModul
    Case 1101 '病人信息
        '76824，李南春，2014/8/19，医疗卡类别处理
        If cboType.ListIndex >= 0 Then
            zlDatabase.SetPara "缺省医疗卡类别", cboType.ItemData(cboType.ListIndex), glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        Else
            zlDatabase.SetPara "缺省医疗卡类别", 0, glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        End If
        '54701:刘鹏飞,2012-09-19
        zlDatabase.SetPara "自动刷新数据", chkAutoRefresh.Value, glngSys, mlngModul, IIf(chkAutoRefresh.Enabled = True, True, False)
    Case 1102   '就诊卡
        '问题28130、27929
        If chkBruhCardBackCard.Value And chkBrushCardVerfy.Value Then
            strTmp = "3"
        ElseIf chkBruhCardBackCard.Value Then
            strTmp = "1"
        ElseIf chkBrushCardVerfy.Value Then
            strTmp = "2"
        Else
            strTmp = "0"
        End If
        Call zlDatabase.SetPara("退卡刷卡", strTmp, glngSys, mlngModul, IIf(chkBruhCardBackCard.Enabled = True, True, False))
    Case 1103
        zlDatabase.SetPara "缺省预交结算方式", Trim(cboDefaultBalance.Text), glngSys, glngModul, IIf(cboDefaultBalance.Enabled = True, True, False)
    End Select
    'LED设备
    zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)

    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

 Private Sub Load缺省预交结算方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载代收款
    '编制:刘兴洪
    '日期:2011-07-19 15:13:59
    '问题:  34705
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, varData As Variant, varTemp As Variant, j As Long, strTmp As String
    
    str结算方式 = zlDatabase.GetPara("缺省预交结算方式", glngSys, glngModul, , Array(cboDefaultBalance), InStr(mstrPrivs, ";参数设置;") > 0)
     
     On Error GoTo errHandle
    '结算方式
    strSQL = _
    " Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
    " From 结算方式应用 A,结算方式 B" & _
    " Where A.应用场合='预交款' And B.名称=A.结算方式 And Nvl(B.性质,1) In(1,2,3,5,8)" & _
    " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboDefaultBalance
        Do While Not rsTmp.EOF
            .AddItem NVL(rsTmp!名称)
            If .ListIndex < 0 And Val(NVL(rsTmp!缺省)) = 1 Then .ListIndex = .NewIndex
            If str结算方式 = NVL(rsTmp!名称) Then .ListIndex = .NewIndex
            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
 End Sub

Private Sub Form_Load()
    Dim i As Long, lngCardTypeID As Long
    Dim strPrintMode As String '问题号:50656
    Dim strArr打印方式() As String '问题号:50656
    Dim strTmp As String
    gblnOK = False
    Me.sTab.TabVisible(2) = False   '34705
    sTab.TabVisible(0) = mlngModul = 1101
    sTab.TabVisible(2) = mlngModul = 1103    '34705
    sTab.TabVisible(1) = mlngModul <> 1103    '34705
    If mlngModul = 1103 Then Call Load缺省预交结算方式
    Call InitShareInvoice   '加载共用批票据信息
    
    'LED设备
    chkLedWelcome.Value = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, ";参数设置;") > 0)

    Select Case mlngModul
    Case 1101 ''病人信息
        lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, , Array(cboType), InStr(mstrPrivs, ";参数设置;") > 0))
        For i = 0 To cboType.ListCount - 1
            If cboType.ItemData(i) = lngCardTypeID Then cboType.ListIndex = i: Exit For
        Next
        
        '54701:刘鹏飞,2012-09-19
        chkAutoRefresh.Value = zlDatabase.GetPara("自动刷新数据", glngSys, mlngModul, 1, Array(chkAutoRefresh), InStr(mstrPrivs, ";参数设置;") > 0)
    
    Case 1102   '就诊卡
        '问题28130
        Select Case Val(zlDatabase.GetPara("退卡刷卡", glngSys, mlngModul, "0", Array(chkBruhCardBackCard, chkBrushCardVerfy), InStr(mstrPrivs, ";参数设置;") > 0))
        Case 0: chkBruhCardBackCard.Value = 0: chkBrushCardVerfy.Value = 0
        Case 1: chkBruhCardBackCard.Value = 1
        Case 2: chkBrushCardVerfy.Value = 1
        Case 3: chkBruhCardBackCard.Value = 1: chkBrushCardVerfy.Value = 1
        End Select
        chkBruhCardBackCard.Visible = True: chkBrushCardVerfy.Visible = True
    Case 1103  '预交款
    End Select
    chkLedWelcome.Visible = mlngModul = 1103
    Exit Sub
errH:
    If ErrCenter() = 1 Then
         Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln担保 = False
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
    zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
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
