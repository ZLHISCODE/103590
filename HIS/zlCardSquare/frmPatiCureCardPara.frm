VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPatiCureCardPara 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6110
   ScaleMode       =   0  'User
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPrepay 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   330
      ScaleHeight     =   2295
      ScaleWidth      =   7845
      TabIndex        =   11
      Top             =   1755
      Width           =   7845
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用预交票据"
         Height          =   1590
         Left            =   390
         TabIndex        =   13
         Top             =   105
         Width           =   7770
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1215
            Left            =   75
            TabIndex        =   14
            Top             =   255
            Width           =   7605
            _cx             =   13414
            _cy             =   2143
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
            FormatString    =   $"frmPatiCureCardPara.frx":0000
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
      Begin VB.CommandButton cmdPrepayPrintSet 
         Caption         =   "预交票据打印设置(&Y)"
         Height          =   420
         Left            =   5850
         TabIndex        =   15
         Top             =   1785
         Width           =   1950
      End
   End
   Begin VB.PictureBox pic参数控制 
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   690
      ScaleHeight     =   4540
      ScaleMode       =   0  'User
      ScaleWidth      =   7845
      TabIndex        =   4
      Top             =   2190
      Width           =   7845
      Begin VB.OptionButton opt剩余款缺省 
         Caption         =   "剩余款存为预交"
         Height          =   375
         Index           =   1
         Left            =   4500
         TabIndex        =   18
         Top             =   3780
         Width           =   1785
      End
      Begin VB.OptionButton opt剩余款缺省 
         Caption         =   "剩余款找补给病人"
         Height          =   375
         Index           =   0
         Left            =   2070
         TabIndex        =   17
         Top             =   3780
         Width           =   1965
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   405
         Left            =   4605
         TabIndex        =   10
         Top             =   4188
         Width           =   1305
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "卡收据打印设置(&P)"
         Height          =   405
         Left            =   5955
         TabIndex        =   12
         Top             =   4188
         Width           =   1815
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   30
         TabIndex        =   9
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   4195
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用医疗卡"
         Height          =   1965
         Left            =   45
         TabIndex        =   7
         Top             =   1785
         Width           =   7755
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1545
            Left            =   120
            TabIndex        =   8
            Top             =   270
            Width           =   7575
            _cx             =   13361
            _cy             =   2725
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
            FormatString    =   $"frmPatiCureCardPara.frx":00E0
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
      Begin VB.Frame fracharge 
         Caption         =   "本地共用发卡票据"
         Height          =   1590
         Left            =   45
         TabIndex        =   5
         Top             =   75
         Width           =   7770
         Begin VSFlex8Ctl.VSFlexGrid vsCharge 
            Height          =   1215
            Left            =   75
            TabIndex        =   6
            Top             =   270
            Width           =   7605
            _cx             =   13414
            _cy             =   2143
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
            FormatString    =   $"frmPatiCureCardPara.frx":01C4
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
      Begin VB.Label lbl剩余款缺省处理方式 
         Caption         =   "剩余款缺省处理方式："
         Height          =   384
         Left            =   40
         TabIndex        =   16
         Top             =   3892
         Width           =   1875
      End
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   5120
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   7995
      _Version        =   589884
      _ExtentX        =   14102
      _ExtentY        =   9031
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   6945
      TabIndex        =   1
      Top             =   5445
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   5445
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   5775
      TabIndex        =   2
      Top             =   5445
      Width           =   1100
   End
End
Attribute VB_Name = "frmPatiCureCardPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnOk As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '返回:参数设置成功,返回true,否则的返回False
    '编制:刘兴洪
    '日期:2011-07-14 17:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOk = False
    
    Me.Show 1, frmMain
    zlSetPara = mblnOk
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
    isValied = False
    On Error GoTo errHandle
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
    
    '检查每种使用发票只能一个选择
    With vsCharge
        str类别 = "-"
        For i = 1 To .Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("使用类别"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("使用类别")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("使用类别"))) = Trim(.TextMatrix(j, .ColIndex("使用类别"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    使用类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    
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
    isValied = True
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
    zlDatabase.SetPara "共用医疗卡批次", strValue, glngSys, mlngModule, blnHavePrivs
    '保存预交票据
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("预交类型")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用预交票据批次", strValue, glngSys, mlngModule, blnHavePrivs
    
    '104726:李南春,2017/4/17,保存门诊医疗卡票据
    strValue = ""
    With vsCharge
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = Val(.RowData(i)): Exit For
            End If
        Next
    End With
    zlDatabase.SetPara "共用门诊收据批次", strValue, glngSys, mlngModule, blnHavePrivs
    
End Sub
Private Sub InitShareInvoice()
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intTYPE As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ID,编码,名称, nvl(是否固定,0) as 是否固定  from 医疗卡类别  "
    Set rs医疗卡类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rs医疗卡类别.Filter = "名称='就诊卡' and 是否固定=1"
    If rs医疗卡类别.EOF = False Then
        str缺省医疗卡 = rs医疗卡类别!名称: lng缺省医疗卡 = Val(rs医疗卡类别!id)
    End If
    
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "共用医疗票据列表", False, False
    strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModule, , , True, intTYPE)
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
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            '99007:李南春,2016/7/29，共用医疗卡票据获取使用类别ID
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
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModule, , , True, intTYPE)
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
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            If Val(Nvl(rsTemp!使用类别, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "门诊和住院共用"
            ElseIf Val(Nvl(rsTemp!使用类别, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
            Else
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = Val(Nvl(rsTemp!使用类别))
            
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
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
    
    '共用门诊票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModule, vsCharge, Me.Name, "共用门诊收据列表", False, False
    
    lngTemp = Val(zlDatabase.GetPara("共用门诊收据批次", glngSys, mlngModule, , , True, intTYPE))
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsCharge.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsCharge.ForeColor = vbBlue: vsCharge.ForeColorFixed = vbBlue
        fracharge.ForeColor = vbBlue: vsCharge.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
    Case Else
        vsCharge.ForeColor = &H80000008: vsCharge.ForeColorFixed = &H80000008
        fracharge.ForeColor = &H80000008
    End Select
    With vsCharge
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With

    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(1)
    With vsCharge
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!使用类别, " ")
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            If .RowData(lngRow) = lngTemp Then
                .TextMatrix(lngRow, .ColIndex("选择")) = -1
            End If
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub cmdOK_Click()
    Dim blnHavePrivs As Boolean, intData As Integer, strControl As String
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If isValied = False Then Exit Sub
    
    zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.value, glngSys, mlngModule, blnHavePrivs
    Call SaveInvoice
    Call Save剩余款缺省
    mblnOk = True: Unload Me
End Sub
Private Sub InitPara()
    Dim blnHavePrivs As Boolean, i As Long
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    'LED设备
    chkLedWelcome.value = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModule, 1, Array(chkLedWelcome), blnHavePrivs)
    '剩余款缺省处理方式
    i = Val(zlDatabase.GetPara("剩余款缺省处理方式", glngSys, mlngModule, 0, Array(chkLedWelcome), blnHavePrivs))
    opt剩余款缺省(i).value = True
End Sub
Private Sub cmdPrepayPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdPrintSet_Click()
    '打印设置
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1107", Me)
End Sub

Private Sub Form_Load()
    Call InitTbPage
    Call InitShareInvoice
    Call InitPara
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用医疗票据列表", False, False
    zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
    zl_vsGrid_Para_Save mlngModule, vsCharge, Me.Name, "共用门诊收据列表", False, False
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用医疗票据列表", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用医疗票据列表", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModule, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub
 
Private Sub vsCharge_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModule, vsCharge, Me.Name, "共用门诊收据列表", False, False
End Sub

Private Sub vsCharge_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModule, vsCharge, Me.Name, "共用门诊收据列表", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Val(.Cell(flexcpData, Row, .ColIndex("医疗卡类别"))) = Val(.Cell(flexcpData, i, .ColIndex("医疗卡类别"))) _
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

Private Sub vsCharge_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsCharge
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsCharge_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsCharge
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

Public Sub InitTbPage()
    Dim strControl As String
    Dim intTYPE As Integer, objItem As TabControlItem, blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    Set objItem = tbPage.InsertItem(1, "常规", pic参数控制.hWnd, 0)
    Set objItem = tbPage.InsertItem(2, "预交票据", picPrepay.hWnd, 0)
    With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        '.PaintManager.StaticFrame = True
        ' .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Position = xtpTabPositionTop
    End With
End Sub

Private Sub Save剩余款缺省()
    Dim blnHavePrivs As Boolean

    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "剩余款缺省处理方式", IIf(opt剩余款缺省(0).value, 0, 1), glngSys, mlngModule, blnHavePrivs
End Sub

