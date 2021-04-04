VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
   Icon            =   "frmSquareCardParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7515
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraTitle 
      Caption         =   "本地共用消费卡"
      Height          =   1965
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   7365
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   1635
         Left            =   75
         TabIndex        =   7
         Top             =   270
         Width           =   7215
         _cx             =   12726
         _cy             =   2884
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
         FormatString    =   $"frmSquareCardParaSet.frx":030A
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
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "缴款单打印设置"
      Height          =   360
      Left            =   3750
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2185
      Width           =   1875
   End
   Begin VB.CheckBox chk连续充值 
      Caption         =   "充值后不退出充值界面(&N)"
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   2245
      Width           =   2400
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   -30
      TabIndex        =   3
      Top             =   2610
      Width           =   7875
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   5790
      TabIndex        =   2
      Top             =   2190
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5985
      TabIndex        =   0
      Top             =   3000
      Width           =   1100
   End
End
Attribute VB_Name = "frmSquareCardParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs   As String, mblnFirst As Boolean, mblnChange As Boolean

Public Sub ShowParaSet(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置入口
    '入参:frmMain-父窗口
    '     lngModule-模块号
    '     strPrivs-权限串

    '编制:刘兴洪
    '日期:2009-11-19 15:29:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnFirst = True
    Me.Show 1, frmMain
End Sub

Private Sub LoadParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数设置
    '编制:刘兴洪
    '日期:2009-12-10 17:03:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, varData As Variant
    Dim blnIsHavePriv As Boolean
    blnIsHavePriv = InStr(1, mstrPrivs, ";参数设置;") > 0
    chk连续充值.value = IIf(Val(zlDatabase.GetPara("连续充值", glngSys, mlngModule, , Array(chk连续充值), blnIsHavePriv)) = 1, 1, 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function SaveSet() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-12-10 16:59:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, i As Long
    Dim strValue As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    Err = 0: On Error GoTo Errhand:
    '保存共享消费卡
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("消费卡类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用消费卡批次", strValue, glngSys, mlngModule, blnHavePrivs
   
    Call zlDatabase.SetPara("连续充值", IIf(chk连续充值.value = 1, 1, 0), glngSys, mlngModule, blnHavePrivs)
    SaveSet = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
    
    On Error GoTo errHandle
    '检查每种使用种式只能一个选择
    With vsBill
        str类别 = "-"
        For i = 1 To vsBill.Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("消费卡类别"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("消费卡类别")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("消费卡类别"))) = Trim(.TextMatrix(j, .ColIndex("消费卡类别"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    消费卡类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
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

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_1503"
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub
 
Private Sub Form_Load()
    Call InitShareInvoice
    Call LoadParaSet
End Sub

Private Sub InitShareInvoice()
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSql As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    
    On Error GoTo errHandle
    zl_vsGrid_Para_Restore mlngModule, vsBill, Me.Name, "共用消费卡列表", False, False
    strShareInvoice = zlDatabase.GetPara("共用消费卡批次", glngSys, mlngModule, , , True, intType)
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
    
    '格式:领用ID1,消费卡类别ID1|领用IDn,消费卡类别IDn|...
    varData = Split(strShareInvoice, "|")

    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(6)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            .TextMatrix(lngRow, .ColIndex("消费卡类别")) = Nvl(rsTemp!使用类别)
            .Cell(flexcpData, lngRow, .ColIndex("消费卡类别")) = Val(Nvl(rsTemp!使用类别ID))
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("消费卡类别"))) Then
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用消费卡列表", False, False
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用消费卡列表", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBill, Me.Name, "共用消费卡列表", False, False
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsBill
        Select Case Col
        Case .ColIndex("选择")
            If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                For i = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, Row, .ColIndex("消费卡类别"))) = Val(.Cell(flexcpData, i, .ColIndex("消费卡类别"))) _
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
