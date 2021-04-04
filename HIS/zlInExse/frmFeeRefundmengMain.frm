VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFeeRefundmentMain 
   BorderStyle     =   0  'None
   Caption         =   "frmFeeRefundmengMain"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   960
      ScaleHeight     =   1725
      ScaleWidth      =   4290
      TabIndex        =   7
      Top             =   3390
      Width           =   4290
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   780
         ScaleHeight     =   375
         ScaleWidth      =   3405
         TabIndex        =   8
         Top             =   930
         Width           =   3405
         Begin VB.TextBox txtSum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   1185
         End
         Begin VB.ComboBox cboStyle 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   615
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label lblBack 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   11
            Top             =   60
            Width           =   480
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   735
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   11160
         _cx             =   19685
         _cy             =   1296
         Appearance      =   0
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundmengMain.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         ExplorerBar     =   3
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
   Begin VB.PictureBox picBalanceStyle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -60
      ScaleHeight     =   1950
      ScaleWidth      =   3045
      TabIndex        =   4
      Top             =   975
      Width           =   3045
      Begin VSFlex8Ctl.VSFlexGrid vsBalanceStyle 
         Height          =   1290
         Left            =   0
         TabIndex        =   5
         Top             =   135
         Width           =   2565
         _cx             =   4524
         _cy             =   2275
         Appearance      =   3
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundmengMain.frx":00CB
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "当前转出合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   1605
         Width           =   1665
      End
   End
   Begin VB.PictureBox picInvoice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   5340
      ScaleHeight     =   3165
      ScaleWidth      =   2535
      TabIndex        =   2
      Top             =   2490
      Width           =   2535
      Begin VSFlex8Ctl.VSFlexGrid vsfInvoice 
         Height          =   1605
         Left            =   0
         TabIndex        =   3
         Top             =   45
         Width           =   2055
         _cx             =   3625
         _cy             =   2831
         Appearance      =   0
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeRefundmengMain.frx":0139
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   101
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
   Begin VB.PictureBox picFee 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   3870
      ScaleHeight     =   2145
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   315
      Width           =   2010
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   1470
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5490
         _cx             =   9684
         _cy             =   2593
         Appearance      =   0
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1500
      Top             =   630
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFeeRefundmentMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmFeeDetail As frmFeeDetail
Private mstrStyle As String, mlngModule As Long, mstrPrivs As String
Private mrsFeeList As ADODB.Recordset, mrsInfo As ADODB.Recordset, mrsBalance As ADODB.Recordset
Private mintType As Integer, mbln立即销帐 As Boolean, mbln门诊转住院先审核 As Boolean
Private mblnSel As Boolean, mbln药房单位 As Boolean, mint收费清单 As Integer
Private mstrFindFpNo As String, mstrFindNO As String, mlng病人ID As Long
Private mlngShareUseID As Long, mrsBalanceDup As ADODB.Recordset
Private mobjSquare As Object
Private mstrThreeSwapBalance As String
Private mstrThreeSwapCardType As String
Private mstrThreeSwapMoney As String
Private Enum mObjPancel
    Pan_BalanceInfo = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Invoice = 5
End Enum

Public Sub InitMe(ByVal lngModule As Long, ByVal strPrivs As String, ByVal intTYPE As Integer)
    '-------------------------------------------------------------------------------------------------
    '功能:程序入口,初始化
    '入参:
    '       lngModule-模块号
    '       strPrivs-权限串
    '编制:刘尔旋
    '日期:2014-06-18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mintType = intTYPE
    If mobjSquare Is Nothing Then Set mobjSquare = gobjSquare.objSquareCard
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
    mbln门诊转住院先审核 = IIf(Val(zlDatabase.GetPara("门诊转住院先审核", glngSys, 1143, 0)) = 1, True, False)
    mint收费清单 = 0: mbln药房单位 = False
    If mintType = 1 Then
        mint收费清单 = Val(zlDatabase.GetPara("收费清单打印方式", glngSys, 1121))   '门诊收费
        mbln药房单位 = zlDatabase.GetPara("药品单位", glngSys, 1121) = "1"
        mlngShareUseID = Val(zlDatabase.GetPara("共用收费票据批次", glngSys, mlngModule, "0"))
    Else
        mlngShareUseID = 0
    End If
    mblnSel = False
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_BalanceInfo
        Item.Handle = picBalanceStyle.hWnd
    Case Pan_Bill
        Item.Handle = picFee.hWnd
    Case Pan_List
        Item.Handle = mfrmFeeDetail.hWnd
    Case Pan_Balance
        Item.Handle = picBalance.hWnd
    Case Pan_Invoice
        Item.Handle = picInvoice.hWnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next

    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mintType = 1, "退费列表", "销帐列表"), True
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, IIf(mintType = 1, "历史退费列表", "历史销帐列表"), True
    zl_vsGrid_Para_Save mlngModule, vsBalanceStyle, Me.Caption, IIf(mintType = 1, "退费结算信息", "销帐结算信息"), , True
    zl_vsGrid_Para_Save mlngModule, vsfInvoice, Me.Caption, IIf(mintType = 1, "退费发票列表", "销帐发票列表"), True
    
    Unload mfrmFeeDetail
    Set mfrmFeeDetail = Nothing
    Set mrsFeeList = Nothing
    Set mrsInfo = Nothing
    Set mrsBalance = Nothing
End Sub

Private Sub Form_Load()
    Call InitPanel
    Call LoadStyle
    Call SetHeader
End Sub

Private Function InitBlanceData(ByVal strBalance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化结算数据
    '入参:strBalance-指定的结算序号,以逗号分离:'0001,0002
    '出参:
    '返回:
    '编制:刘尔旋
    '日期:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Err = 0: On Error GoTo errHandle
    If mintType = 2 Then
        InitBlanceData = True
        Exit Function
    End If
    If strBalance = "" Then InitBlanceData = True: Exit Function
    
    strSql = _
    "Select a.结算方式, Nvl(b.性质, 1) As 性质, b.应付款, a.金额" & vbNewLine & _
    "From (Select Decode(a.记录性质, 3, a.结算方式, Null) As 结算方式, Sum(a.冲预交) As 金额" & vbNewLine & _
    "       From 病人预交记录 A" & vbNewLine & _
    "       Where a.结帐id In (Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.结帐id" & vbNewLine & _
    "                        From 门诊费用记录 C, 门诊费用记录 D, (Select Distinct 结帐ID From 病人预交记录 I,Table(f_Str2list([1])) J Where I.结算序号=J.Column_Value) E" & vbNewLine & _
    "                        Where c.结帐id = e.结帐id And c.No = d.No And Mod(d.记录性质, 10) = 1) And a.记录性质 In (1, 11, 3) And a.病人id=[2] And" & vbNewLine & _
    "             Nvl(a.冲预交, 0) <> 0" & vbNewLine & _
    "" & vbNewLine & _
    "       Group By Decode(a.记录性质, 3, a.结算方式, Null)) A, 结算方式 B" & vbNewLine & _
    "Where a.结算方式 = b.名称(+)"
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Replace(strBalance, "'", ""), mlng病人ID)
    Set mrsBalanceDup = mrsBalance
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetPicBack(ByVal strBalance As String) As Boolean
    'vsBalance.Width = picBalance.Width - 4000
    'picBack.Left = vsBalance.Width + vsBalance.Left + 30
    picBack.Visible = True
    SetPicBack = True
End Function

Private Sub SetBlanceShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结算方式
    '入参:blnAllSel-选择所有的单据
    '编制:刘兴洪
    '日期:2011-02-23 14:54:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str结算 As String
    Dim bln全选 As Boolean, bln未选 As Boolean, intCol As Integer
    Dim strFilter As String, bln退款 As Boolean, rsTmp As ADODB.Recordset
    Dim strSelNos As String, strNO As String, strSql As String
    If mintType = 2 Then Exit Sub
    With vsFee
        bln全选 = True: bln未选 = True
        For lngRow = 1 To .Rows - 1
            strBalance = Trim(.TextMatrix(lngRow, .ColIndex("结算序号")))
            If .TextMatrix(lngRow, .ColIndex("选择")) = "√" And Val(strBalance) <> 0 Then
                If InStr(1, strSelNos & ",", "," & strBalance & ",") = 0 Then
                    strSelNos = strSelNos & "," & strBalance
                    bln未选 = False
                End If
            End If
             If InStr(1, strSelNos & ",", "," & strBalance & ",") = 0 Then bln全选 = False
        Next
    End With
    If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    bln退款 = False
    
    '显示所有选择的单据的结算方式之和
    If Not mrsBalance Is Nothing Then
        If bln全选 Or bln未选 Then
            mrsBalance.Filter = ""
            If bln全选 Then bln退款 = True
        Else
            strFilter = Replace(strSelNos, ",", "' Or 结算序号='")
            strFilter = " 结算序号=" & strFilter & ""
            'mrsBalance.Filter = strFilter
            bln退款 = True
        End If
        If SetPicBack(strSelNos) = True Then
            txtSum.Text = InitPatialBalance(strSelNos)
        Else
            Call InitBlanceData(strSelNos)
        End If
        mrsBalance.Sort = "性质,应付款,结算方式"
        mrsBalanceDup.Sort = "性质,应付款,结算方式"
        vsBalance.Redraw = flexRDNone
        vsBalance.Clear 1
        vsBalance.Cols = 1
        If Not mrsBalanceDup.EOF Then
            For i = 1 To mrsBalanceDup.RecordCount
                If Val(NVL(mrsBalanceDup!金额)) <> 0 Then
                    If NVL(mrsBalanceDup!结算方式, "冲预交") <> strBalance Then
                        strBalance = NVL(mrsBalanceDup!结算方式, "冲预交")
                        vsBalance.Cols = vsBalance.Cols + 2
                        vsBalance.ColAlignment(vsBalance.Cols - 2) = 7
                        vsBalance.ColAlignment(vsBalance.Cols - 1) = 1
                    End If
                    If mrsBalanceDup!性质 <> 1 Then
                        vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '粗体
                        vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue
                    ElseIf bln退款 Then
                        vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '粗体
                        vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue  '红色:退款
                    End If
                    vsBalance.TextMatrix(0, vsBalance.Cols - 2) = strBalance & ":"
                    vsBalance.TextMatrix(0, vsBalance.Cols - 1) = _
                        Val(vsBalance.TextMatrix(0, vsBalance.Cols - 1)) + NVL(mrsBalanceDup!金额, 0)
                End If
                mrsBalanceDup.MoveNext
            Next
        End If
        intCol = 0
        strBalance = ""
        If Not mrsBalance.EOF Then
            For i = 1 To mrsBalance.RecordCount
                If Val(NVL(mrsBalance!金额)) <> 0 And Val(NVL(mrsBalance!性质)) <> 9 Then
                    If NVL(mrsBalance!结算方式, "冲预交") <> strBalance Then
                        strBalance = NVL(mrsBalance!结算方式, "冲预交")
                        intCol = intCol + 2
                        vsBalance.ColAlignment(intCol - 1) = 7
                        vsBalance.ColAlignment(intCol) = 1
                    End If
                    If mrsBalance!性质 <> 1 Then
                        vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '粗体
                        vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = vbRed '红色
                    ElseIf bln退款 Then
                        vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '粗体
                        vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = vbRed '红色:退款
                    End If
                    vsBalance.TextMatrix(1, intCol - 1) = strBalance & ":"
                    vsBalance.TextMatrix(1, intCol) = _
                    Val(vsBalance.TextMatrix(1, intCol)) + NVL(mrsBalance!金额, 0)
                End If
                mrsBalance.MoveNext
            Next
        End If
        If strSelNos = "" Then
            For i = 1 To vsBalance.Cols - 1
                vsBalance.TextMatrix(1, i) = ""
            Next i
        End If
        
        Call vsBalance.AutoSize(0, vsBalance.Cols - 1)
        vsBalance.Row = vsBalance.FixedRows
        If vsBalance.Cols <> 1 Then vsBalance.Col = vsBalance.FixedCols
        'vsBalance.TextMatrix(0, 0) = "收款结算"
        vsBalance.Redraw = flexRDDirect
    End If
End Sub

Private Function InitPatialBalance(ByVal strBalance As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部分退费的结算数据
    '入参:strBalance-指定的结算序号,以逗号分离:'0001,0002
    '出参:
    '返回:
    '编制:刘尔旋
    '日期:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, dblSum As Double, i As Integer
    Err = 0: On Error GoTo errHandle
    If mintType = 2 Then
        InitPatialBalance = 0
        Exit Function
    End If
    If strBalance = "" Then InitPatialBalance = 0: Exit Function
    
    Call InitBlanceData(strBalance)
    Do While Not mrsBalance.EOF
        dblSum = dblSum + Val(NVL(mrsBalance!金额))
        mrsBalance.MoveNext
    Loop
    '全退记录(预交款)
    strSql = _
    "Select /*+ RULE*/" & vbNewLine & _
    " a.结算方式, Nvl(b.性质, 1) As 性质, b.应付款, a.金额" & vbNewLine & _
    "From (Select Decode(a.记录性质, 3, a.结算方式, Null) As 结算方式, Sum(a.冲预交) As 金额" & vbNewLine & _
    "       From 病人预交记录 A," & vbNewLine & _
    "            (Select /*+ rule */" & vbNewLine & _
    "              Distinct d.结帐id" & vbNewLine & _
    "              From 门诊费用记录 C, 门诊费用记录 D," & vbNewLine & _
    "                   (Select Distinct 结帐id" & vbNewLine & _
    "                     From 病人预交记录 I, Table(f_Str2list([1])) J" & vbNewLine & _
    "                     Where i.结算序号 = j.Column_Value) E" & vbNewLine & _
    "              Where c.结帐id = e.结帐id And c.No = d.No And Mod(d.记录性质, 10) = 1 And Not Exists" & vbNewLine & _
    "               (Select 1" & vbNewLine & _
    "                     From 门诊费用记录" & vbNewLine & _
    "                     Where 结帐id In (Select Max(结帐id)" & vbNewLine & _
    "                                    From 门诊费用记录" & vbNewLine & _
    "                                    Where NO In ((Select Distinct k.No" & vbNewLine & _
    "                                                 From 门诊费用记录 K, 病人预交记录 L" & vbNewLine & _
    "                                                 Where l.结算序号 In (Select Column_Value From Table(f_Str2list([1]))) And" & vbNewLine & _
    "                                                       k.结帐id = l.结帐id)) And Mod(记录性质, 10) = 1) And Mod(记录性质, 10) = 1 And" & vbNewLine & _
    "                           记录状态 = 2)) K" & vbNewLine & _
    "       Where a.结帐id = k.结帐id And a.记录性质 In (1, 11) And a.病人id = [2] And Nvl(a.冲预交, 0) <> 0" & vbNewLine & _
    "       Group By Decode(a.记录性质, 3, a.结算方式, Null)) A, 结算方式 B" & vbNewLine & _
    "Where a.结算方式 = b.名称(+) "

    '全退记录(消费卡)
    strSql = strSql & " Union " & _
    "Select a.结算方式, Nvl(b.性质, 1) As 性质, b.应付款, a.金额" & vbNewLine & _
    "From (Select Decode(a.记录性质, 3, a.结算方式, Null) As 结算方式, Sum(a.冲预交) As 金额" & vbNewLine & _
    "       From 病人预交记录 A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.结帐id" & vbNewLine & _
    "                        From 门诊费用记录 C, 门诊费用记录 D, (Select Distinct 结帐ID From 病人预交记录 I,Table(f_Str2list([1])) J Where I.结算序号=J.Column_Value) E" & vbNewLine & _
    "                        Where c.结帐id = e.结帐id And c.No = d.No And Mod(d.记录性质, 10) = 1) K" & _
    "       Where a.结帐id=K.结帐id  And a.记录性质 = 3 And a.病人id=[2] And Nvl(a.冲预交, 0) <> 0" & vbNewLine & _
    "       Group By Decode(a.记录性质, 3, a.结算方式, Null)) A, 结算方式 B" & vbNewLine & _
    "Where a.结算方式 = b.名称 And B.性质 = 8"
    '全退记录(不能退现的三方账户)
    strSql = strSql & " Union " & _
    "Select a.结算方式, Nvl(b.性质, 1) As 性质, b.应付款, a.金额" & vbNewLine & _
    "From (Select Decode(a.记录性质, 3, a.结算方式, Null) As 结算方式, Sum(a.冲预交) As 金额" & vbNewLine & _
    "       From 病人预交记录 A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.结帐id" & vbNewLine & _
    "                        From 门诊费用记录 C, 门诊费用记录 D, (Select Distinct 结帐ID From 病人预交记录 I,Table(f_Str2list([1])) J Where I.结算序号=J.Column_Value) E" & vbNewLine & _
    "                        Where c.结帐id = e.结帐id And c.No = d.No And Mod(d.记录性质, 10) = 1) K" & _
    "       Where a.结帐id=K.结帐id  And a.记录性质 = 3 And a.病人id=[2] And Nvl(a.冲预交, 0) <> 0" & vbNewLine & _
    "         And Exists (Select 1 From 医疗卡类别 Where ID=A.卡类别ID And 是否退现=0)" & _
    "       Group By Decode(a.记录性质, 3, a.结算方式, Null)) A, 结算方式 B" & vbNewLine & _
    "Where a.结算方式 = b.名称 And B.性质 = 7"
    '医保结算
    strSql = strSql & " Union " & _
    "Select a.结算方式, Nvl(b.性质, 1) As 性质, b.应付款, a.金额" & vbNewLine & _
    "From (Select Decode(a.记录性质, 3, a.结算方式, Null) As 结算方式, Sum(a.冲预交) As 金额" & vbNewLine & _
    "       From 病人预交记录 A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.结帐id" & vbNewLine & _
    "                        From 门诊费用记录 C, 门诊费用记录 D, (Select Distinct 结帐ID From 病人预交记录 I,Table(f_Str2list([1])) J Where I.结算序号=J.Column_Value) E" & vbNewLine & _
    "                        Where c.结帐id = e.结帐id And c.No = d.No And Mod(d.记录性质, 10) = 1) K" & _
    "       Where a.结帐id =K.结帐id  And a.记录性质 In (1, 11, 3) And a.病人id=[2] And" & vbNewLine & _
    "             Nvl(a.冲预交, 0) <> 0" & vbNewLine & _
    "" & vbNewLine & _
    "       Group By Decode(a.记录性质, 3, a.结算方式, Null)) A, 结算方式 B" & vbNewLine & _
    "Where a.结算方式 = b.名称 And B.性质 In (3,4)"
    '误差费
    strSql = strSql & " Union " & _
    "Select a.结算方式, Nvl(b.性质, 1) As 性质, b.应付款, a.金额" & vbNewLine & _
    "From (Select Decode(a.记录性质, 3, a.结算方式, Null) As 结算方式, Sum(a.冲预交) As 金额" & vbNewLine & _
    "       From 病人预交记录 A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.结帐id" & vbNewLine & _
    "                        From 门诊费用记录 C, 门诊费用记录 D, (Select Distinct 结帐ID From 病人预交记录 I,Table(f_Str2list([1])) J Where I.结算序号=J.Column_Value) E" & vbNewLine & _
    "                        Where c.结帐id = e.结帐id And c.No = d.No And Mod(d.记录性质, 10) = 1) K" & _
    "       Where a.结帐id =K.结帐id  And a.记录性质 = 3 And a.病人id=[2] And" & vbNewLine & _
    "             Nvl(a.冲预交, 0) <> 0" & vbNewLine & _
    "" & vbNewLine & _
    "       Group By Decode(a.记录性质, 3, a.结算方式, Null)) A, 结算方式 B" & vbNewLine & _
    "Where a.结算方式 = b.名称 And B.性质 = 9"
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Replace(strBalance, "'", ""), mlng病人ID)
    Do While Not mrsBalance.EOF
        dblSum = dblSum - Val(NVL(mrsBalance!金额))
        mrsBalance.MoveNext
    Loop
    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    
    InitPatialBalance = Format(dblSum, "0.00")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetFpToBIllNOs(ByVal strFpNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的发票号,找出对应的单据号
    '返回:返回对应的单据号,用逗号分隔
    '编制:刘兴洪
    '日期:2011-02-25 10:50:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, strNos As String
    
    On Error GoTo errHandle
    
    strSql = "" & _
    "   Select distinct NO From 票据打印内容 A,票据使用明细 B " & _
    "   Where A.数据性质=1 and A.ID=B.打印ID and B.票种=1 And B.号码=[1]  " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strFpNo)
    strNos = ""
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & NVL(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetFpToBIllNOs = strNos
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub CalcSUMMony()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:状态栏信息更新
    '编制:刘尔旋
    '日期:2014-6-20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur金额 As Currency
    With vsFee
        cur金额 = 0
        For i = .FixedRows To .Rows - 1
            If vsFee.TextMatrix(i, .ColIndex("选择")) = "√" Then
                cur金额 = cur金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
        lblSum.Caption = "当前转出合计:" & Format(cur金额, "###0.00;-###0.00;0.00;0.00")
    End With
End Sub

Private Sub vsFee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFee
        Select Case Col
        Case .ColIndex("选择")
            Call SetBlanceShow
            Call CalcSUMMony
        Case Else
        End Select
    End With
End Sub

Private Sub vsFee_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strBalance As String, bytType As Byte, blnOld As Boolean
    
    If NewRow = OldRow Or NewRow < 1 Then Exit Sub
    With vsFee
        If mintType = 1 Then
            strBalance = Trim(.TextMatrix(NewRow, .ColIndex("结算序号")))
            If Val(strBalance) > 0 Then
                strBalance = Trim(.TextMatrix(NewRow, .ColIndex("结帐ID")))
                blnOld = True
            End If
        Else
            strBalance = Trim(.TextMatrix(NewRow, .ColIndex("首张单据")))
        End If
        If NewRow = 0 Or strBalance = "" Then
            mfrmFeeDetail.zlRefresh mintType, 0
        Else
            mfrmFeeDetail.zlRefresh mintType, strBalance, blnOld
        End If
        .ForeColorSel = vsFee.CellForeColor
    End With
    LoadInvoice mintType, NewRow
    LoadBalance mintType, NewRow
End Sub

Private Sub LoadBalance(ByVal intTYPE As Integer, ByVal NewRow As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'功能:读取当前选择记录的结算信息
'编制:刘尔旋
'日期:2014-6-20
'备注:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset, strBalance As String
    vsBalanceStyle.Clear 1
    vsBalanceStyle.Rows = 2
    If intTYPE = 1 Then
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("结算序号")))
        strSql = "" & _
            "Select 结算方式, Sum(冲预交) As 结算金额" & vbNewLine & _
            "From (Select a.结算方式, a.冲预交" & vbNewLine & _
            "       From 病人预交记录 A," & vbNewLine & _
            "            (Select Distinct 结帐id" & vbNewLine & _
            "              From 门诊费用记录" & vbNewLine & _
            "              Where Mod(记录性质, 10) = 1 And 记录状态 <> 0 And" & vbNewLine & _
            "                    NO In (Select Distinct NO" & vbNewLine & _
            "                           From 门诊费用记录 C, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) D" & vbNewLine & _
            "                           Where Mod(c.记录性质, 10) = 1 And c.记录状态 <> 0 And c.结帐id = d.结帐id)) B" & vbNewLine & _
            "       Where a.记录性质 = 3 And a.结帐id = b.结帐id)" & vbNewLine & _
            "Group By 结算方式" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select 结算方式, Sum(冲预交) As 结算金额" & vbNewLine & _
            "From (Select '预交款' As 结算方式, a.冲预交" & vbNewLine & _
            "       From 病人预交记录 A," & vbNewLine & _
            "            (Select Distinct 结帐id" & vbNewLine & _
            "              From 门诊费用记录" & vbNewLine & _
            "              Where Mod(记录性质, 10) = 1 And 记录状态 <> 0 And" & vbNewLine & _
            "                    NO In (Select Distinct NO" & vbNewLine & _
            "                           From 门诊费用记录 C, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) D" & vbNewLine & _
            "                           Where Mod(c.记录性质, 10) = 1 And c.记录状态 <> 0 And c.结帐id = d.结帐id)) B" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And a.结帐id = b.结帐id)" & vbNewLine & _
            "Group By 结算方式"

        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strBalance))
        With vsBalanceStyle
            Do While Not rsTmp.EOF
                If Val(NVL(rsTmp!结算金额)) <> 0 Then
                    .TextMatrix(.Rows - 1, 0) = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("首张单据")))
                    .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!结算方式)
                    .TextMatrix(.Rows - 1, 2) = Format(NVL(rsTmp!结算金额), "0.00")
                    .Rows = .Rows + 1
                End If
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    Else
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("首张单据")))
        strSql = "Select a.结算方式, Sum(a.冲预交) As 结算金额" & vbNewLine & _
                "From 病人预交记录 A, 门诊费用记录 B" & vbNewLine & _
                "Where Mod(a.记录性质, 10) = 2 And a.结帐id = b.结帐id And Mod(b.记录性质, 10) = 2 And b.记录状态 <> 0 And b.No = [1]" & vbNewLine & _
                "Group By 结算方式" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select '预交款' As 结算方式, Sum(a.冲预交) As 结算金额" & vbNewLine & _
                "From 病人预交记录 A, 门诊费用记录 B" & vbNewLine & _
                "Where Mod(a.记录性质, 10) = 2 And a.结帐id = b.结帐id And Mod(b.记录性质, 10) = 2 And b.记录状态 <> 0 And b.No = [1]"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalance)
        With vsBalanceStyle
            Do While Not rsTmp.EOF
                If Val(NVL(rsTmp!结算金额)) <> 0 Then
                    .TextMatrix(.Rows - 1, 0) = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("首张单据")))
                    .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!结算方式)
                    .TextMatrix(.Rows - 1, 2) = Format(NVL(rsTmp!结算金额), "0.00")
                    .Rows = .Rows + 1
                End If
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    End If
End Sub

Private Sub LoadInvoice(ByVal bytType As Byte, ByVal NewRow As Long)
'-----------------------------------------------------------------------------------------------------------------------
'功能:读取当前选择记录的票据信息
'编制:刘尔旋
'日期:2014-6-20
'备注:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset, strBalance As String
    vsfInvoice.Clear 1
    vsfInvoice.Rows = 2
    If bytType = 1 Then
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("结算序号")))
        strSql = "Select Distinct D.号码" & vbNewLine & _
                " From 票据打印内容 C,票据使用明细 D," & _
                " (Select Distinct A.NO From 门诊费用记录 A,病人预交记录 B Where A.结帐ID=B.结帐ID And Mod(A.记录性质, 10) = 1 And B.结算序号= [1]) E" & vbNewLine & _
                " Where E.No=C.No(+) And C.数据性质(+)=1 And C.ID = D.打印ID(+) " & _
                " And Not Exists (Select 1 From 票据使用明细 Where 票种 = d.票种 And 号码 = d.号码 And 性质 = 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strBalance))
        With vsfInvoice
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, 0) = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("首张单据")))
                .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!号码)
                .Rows = .Rows + 1
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    Else
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("首张单据")))
        strSql = " Select Distinct E.号码,C.NO" & vbNewLine & _
                 " From 门诊费用记录 A, 门诊费用记录 B,病人结帐记录 C,票据打印内容 D,票据使用明细 E" & vbNewLine & _
                 " Where Mod(a.记录性质, 10) = 2 And a.No = [1] And b.结帐id = a.结帐ID And C.ID=B.结帐ID And C.No=D.No(+) And D.数据性质(+)=3 And D.ID=E.打印ID(+) " & _
                 " And Not Exists (Select 1 From 票据使用明细 Where 票种 = e.票种 And 号码 = e.号码 And 性质 = 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalance)
        With vsfInvoice
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, 0) = NVL(rsTmp!NO)
                .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!号码)
                .Rows = .Rows + 1
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    End If
End Sub

Private Function LoadStyle() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    cboStyle.Clear
    On Error GoTo errH
    Set rsTmp = Get结算方式("收费", "1,2")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,", "," & rsTmp!性质 & ",") > 0 And Val(NVL(rsTmp!应付款)) = 0 Then
            cboStyle.AddItem rsTmp!名称
            cboStyle.ItemData(cboStyle.NewIndex) = rsTmp!性质
            If rsTmp!缺省 = 1 And cboStyle.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboStyle.hWnd, cboStyle.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboStyle.ListIndex = -1 And cboStyle.ListCount > 0 Then Call zlControl.CboSetIndex(cboStyle.hWnd, 0)
    txtSum.ForeColor = vbRed
    strSql = "" & _
            " Select B.编码,B.名称,Nvl(B.缺省标志,0) as 缺省,Nvl(B.性质,1) as 性质,Nvl(B.应付款,0) as 应付款" & _
            " From 结算方式应用 A,结算方式 B" & _
            " Where A.应用场合=[1] And B.名称=A.结算方式 " & _
            " And B.性质<>8 " & _
            " Order by 性质,lpad(编码,3,' ')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "收费")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,7,", "," & rsTmp!性质 & ",") > 0 Then
            mstrStyle = mstrStyle & rsTmp!名称 & ":"
        End If
        rsTmp.MoveNext
    Next
    LoadStyle = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SelAllNO()
    Dim i As Long
    With vsFee
        If .Rows = 2 And .TextMatrix(1, .ColIndex("结算序号")) = "" Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("首张单据")) <> "" Then
                .TextMatrix(i, .ColIndex("选择")) = "√"
            End If
        Next
        Call CheckInsure
        Call SetBlanceShow
        Call CalcSUMMony
        mblnSel = True
    End With
End Sub

Private Sub CheckInsure()
    Dim i As Integer, intInsure As Integer, blnSelect As Boolean
    With vsFee
        For i = 1 To .Rows - 1
            intInsure = Val(.TextMatrix(i, .ColIndex("险类")))
            blnSelect = .TextMatrix(i, .ColIndex("选择")) <> ""
            If intInsure > 0 And blnSelect Then
                If gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, intInsure) = False Then
                    .TextMatrix(i, .ColIndex("选择")) = ""
                End If
            End If
        Next i
    End With
End Sub

Private Sub picBalanceStyle_Resize()
    On Error Resume Next
    With vsBalanceStyle
        .Top = 0
        .Left = 0
        .Width = picBalanceStyle.Width
        .Height = picBalanceStyle.Height - lblSum.Height - 60
    End With
    With lblSum
        .Top = vsBalanceStyle.Height
        .Left = 15
    End With
End Sub

Private Sub picBalance_Resize()
    On Error Resume Next
    With vsBalance
        .Top = 0
        .Left = 0
        .Width = picBalance.Width
        .Height = picBalance.Height - picBack.Height - 30
    End With
    
    With picBack
        .Left = picBalance.Width - 3500
        .Top = vsBalance.Top + vsBalance.Height
    End With
End Sub

Private Sub picFee_Resize()
    With vsFee
        .Top = 0
        .Left = 0
        .Width = picFee.Width
        .Height = picFee.Height
    End With
End Sub

Private Sub picInvoice_Resize()
    With vsfInvoice
        .Top = 0
        .Left = 0
        .Width = picInvoice.Width
        .Height = picInvoice.Height
    End With
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    With vsFee
        If .DataSource Is Nothing Then
            strHead = "选择,4,500|类别,4,850|单据,4,800|医保,4,500|首张单据,4,850|首张发票,4,1100|开单人,4,800|应收金额,7,850|实收金额,7,850|发生时间,4,1850|结算序号,4,0|险类,4,0"
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
            Next
            .Rows = 2
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        '选择,4,500|类别,4,850|医保,4,500|单据号,4,850|票据号,4,1100|开单人,4,800|应收金额,7,850|实收金额,7,850|发生时间,4,1850|结帐ID,4,0|险类,4,0
        For i = 0 To .Cols - 1
             .FixedAlignment(i) = flexAlignCenterCenter
             .ColAlignment(i) = flexAlignLeftCenter
             .ColKey(i) = Trim(.TextMatrix(0, i))
             Select Case .ColKey(i)
             Case "选择", "类别", "单据", "医保", "首张单据", "首张发票"
                .ColAlignment(i) = flexAlignCenterCenter
             Case "应收金额", "实收金额"
                .ColAlignment(i) = flexAlignRightCenter
             End Select
             If .ColKey(i) Like "*ID" Or .ColKey(i) = "险类" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
             End If
        Next
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, "门诊转住院列表", True
        .RowHeight(0) = 320
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置一行的选择状态
    '       如果是多张单据中的一张,则还需同时设置多张中的其它单据
    '编制:刘兴洪
    '日期:2011-02-21 16:10:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim str单据 As String
    
    With vsFee
        If .TextMatrix(lngRow, .ColIndex("选择")) <> IIf(blnSelect, "√", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("类别")))
            If intInsure > 0 And blnSelect And str单据 = "收费" Then
                strNO = .TextMatrix(lngRow, .ColIndex("首张单据"))
                If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, intInsure) Then
                    frmFeeRefundment.stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持门诊结算作废,此行不允许选择转入!"
                    .TextMatrix(lngRow, .ColIndex("选择")) = ""
                    Exit Function
                Else
                    '再判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                    'strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, intInsure, strBalanceType) Then
                                frmFeeRefundment.stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                .TextMatrix(lngRow, .ColIndex("选择")) = ""
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("选择")) = IIf(blnSelect, "√", "")
        End If
    End With
    SetRowSelected = True
End Function

Public Sub ClsAllNO()
   Dim i As Long
    With vsFee
        If .Rows = 2 And .TextMatrix(1, .ColIndex("结算序号")) = "" Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("首张单据")) <> "" Then
                .TextMatrix(i, .ColIndex("选择")) = ""
            End If
        Next
        Call SetBlanceShow
        mblnSel = False
        Call CalcSUMMony
    End With
End Sub

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域设置
    '编制:刘尔旋
    '日期:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panTop As Pane, panBottom As Pane, panRight As Pane
    If mfrmFeeDetail Is Nothing Then Set mfrmFeeDetail = New frmFeeDetail
    Call mfrmFeeDetail.ShowMe(lblBack.Font, mlngModule, mstrPrivs, 1, 0)
    Load mfrmFeeDetail
    
    Set panThis = dkpMain.CreatePane(mObjPancel.Pan_Bill, 250, 580, DockTopOf, Nothing)
    panThis.Title = "门诊转住院列表"
    panThis.Tag = mObjPancel.Pan_Bill
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picFee.hWnd
    
    Set panRight = dkpMain.CreatePane(mObjPancel.Pan_Invoice, 1500 / Screen.TwipsPerPixelX, 300, DockRightOf, panThis)
    panRight.Title = "发票信息"
    panRight.Tag = mObjPancel.Pan_Invoice
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picInvoice.hWnd
    
    Set panRight = dkpMain.CreatePane(mObjPancel.Pan_BalanceInfo, 1500 / Screen.TwipsPerPixelX, 580, DockBottomOf, panRight)
    panRight.Title = "收款结算"
    panRight.Tag = mObjPancel.Pan_Balance
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picBalanceStyle.hWnd
    
    Set panThis = dkpMain.CreatePane(mObjPancel.Pan_List, 250, 580, DockBottomOf, panThis)
    panThis.Title = "单据明细列表"
    panThis.Tag = mObjPancel.Pan_List
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = mfrmFeeDetail.hWnd
    
    
    Set panThis = dkpMain.CreatePane(mObjPancel.Pan_Balance, 250, 580, DockBottomOf, Nothing)
    panThis.Title = "结算信息"
    panThis.Tag = mObjPancel.Pan_BalanceInfo
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picBalance.hWnd
    panThis.MaxTrackSize.Height = 75
    panThis.MinTrackSize.Height = 75
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Set dkpMain.PaintManager.CaptionFont = vsFee.Font
    
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub

Public Function ReadListData(ByVal strFindNo As String, ByVal strFindFpNo As String, _
                            rsInfo As ADODB.Recordset, Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要销帐的明细数据
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSql As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String, strTable1 As String
    Dim strALLNOs As String
    mstrFindNO = strFindNo
    mstrFindFpNo = strFindFpNo
    Set mrsInfo = rsInfo
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(NVL(mrsInfo!病人ID))
    End If
    
    If mstrFindNO <> "" Then
        If mintType = 1 Then
            strNos = Replace(GetMultiNOs(mstrFindNO), "'", "")
            strSql = "Select 病人ID From 门诊费用记录 Where MOD(记录性质,10)=1 And NO=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNos)
            lng病人ID = Val(NVL(rsTemp!病人ID))
            strWhere = "  And A.病人ID=[1]"
        Else
            strNos = mstrFindNO
            strTable1 = ",Table( f_Str2list([2])) J "
            strWhere = "  And A.NO=J.Column_Value"
        End If
    ElseIf mstrFindFpNo <> "" And mintType = 1 Then
        strNos = zlGetFpToBIllNOs(mstrFindFpNo)
        If strNos = "" Then
            MsgBox "未找到对应发票号的单据,请检查!"
            Exit Function
        End If
        strSql = "Select 病人ID From 门诊费用记录 Where MOD(记录性质,10)=1 And NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNos)
        lng病人ID = Val(NVL(rsTemp!病人ID))
        strWhere = "  And A.病人ID=[1]"
    Else
        strTable1 = ""
        strWhere = "  And A.病人ID=[1]"
    End If
    mblnSel = False
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "正在读取单据数据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    If mintType = 1 Then
        strSql = "" & _
            "Select /*+ rule */" & vbNewLine & _
            " '√' As 选择, '收费' As 类别, Decode(Max(b.险类), Null, '', '√') As 医保, Min(a.No) As 首张单据, Min(a.实际票号) As 首张发票, a.费别," & vbNewLine & _
            " LTrim(To_Char(Sum(a.应收金额), '9999999990.00')) As 应收金额, LTrim(To_Char(Sum(a.实收金额), '9999999990.00')) As 实收金额," & vbNewLine & _
            " a.操作员姓名 As 操作员, Null As 开单人, Max(b.险类) As 险类, Nvl(e.结算序号, e.结帐id) As 结算序号, Min(a.登记时间) As 登记时间, Null as 结帐id " & vbNewLine & _
            "From 门诊费用记录 A, 保险结算记录 B," & vbNewLine & _
            "     (Select Distinct Nvl(d.结帐id, c.结帐id) As 结帐id, d.结算序号 As 结算序号" & vbNewLine & _
            "       From 病人预交记录 C, 病人预交记录 D" & vbNewLine & _
            "       Where d.病人id = [1] And c.病人id = [1] And c.结算序号 = d.结算序号(+) And d.结算序号(+) < 0 ) E" & vbNewLine & _
            "Where a.结帐id = b.记录id(+) And b.性质(+) = 1 And Nvl(b.序号(+),1)=1 And Mod(a.记录性质, 10) = 1 And a.病人id = [1] And Exists" & vbNewLine & _
            " (Select 1 From 门诊费用记录 Where 结帐id = e.结帐id And 记录状态 In (1, 3)) And" & vbNewLine & _
            "      a.No In (Select Distinct x.No" & vbNewLine & _
            "               From 门诊费用记录 X, 门诊费用记录 Y" & vbNewLine & _
            "               Where y.结帐id = e.结帐id And x.No = y.No And Mod(x.记录性质, 10) = 1" & vbNewLine & _
            "               Group By x.No, x.序号" & vbNewLine & _
            "               Having Sum(Nvl(x.付数, 1) * x.数次) <> 0) And Not Exists" & vbNewLine & _
            " (Select 1 From 门诊费用记录 Where NO = a.No And Mod(记录性质, 10) = 1 And Nvl(费用状态, 0) = 1) And Exists" & vbNewLine & _
            " (Select 1 From 门诊费用记录 Where 记录性质 = 1 And 记录状态 In (1, 3) And 结帐id = e.结帐id) And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From 费用审核记录 E, 门诊费用记录 F" & vbNewLine & _
            "       Where e.记录状态 = 1 And f.Id = e.费用id And f.No = a.No And Mod(f.记录性质, 10) = 1)" & vbNewLine & _
            "Group By a.费别, Nvl(e.结算序号, e.结帐id), a.操作员姓名"

        strSql = strSql & " Union " & _
            " Select  '√' As 选择, '收费' As 类别,Decode(a.险类, Null, '', '√') As 医保,a.No as 首张单据,a.实际票号 as 首张发票,a.费别, " & _
            "   LTrim(To_Char(a.应收金额, '9999999990.00')) As 应收金额, LTrim(To_Char(a.实收金额, '9999999990.00')) As 实收金额, " & _
            "   a.操作员姓名 As 操作员, Null As 开单人, a.险类 As 险类, Nvl(c.结算序号,c.结帐id) As 结算序号,a.登记时间 As 登记时间, c.结帐id " & vbNewLine & _
            " From (Select Max(险类) as 险类, Decode(Max(险类), 0, '', '√') As 医保, Min(Decode(价格父号, Null, ID, 0)) As ID, " & vbNewLine & _
            "           NO, 实际票号, Avg(Nvl(付数, 1)) As 付数, Sum(数次) 数次, Sum(应收金额) As 应收金额, " & vbNewLine & _
            "           Sum(实收金额) As 实收金额, 开单人, Min(登记时间) As 登记时间, " & vbNewLine & _
            "           Min(操作员姓名) As 操作员姓名, 费别 " & vbNewLine & _
            "       From (Select Row_Number() Over(Partition By a.ID Order By m.序号) As Rn,a.Id,Nvl(M.险类,0) as 险类, " & _
            "               A.价格父号, A.NO, A.实际票号, A.付数,A.数次,A.应收金额,A.实收金额, A.开单人, A.登记时间, " & vbNewLine & _
            "               a.操作员姓名, a.费别 " & vbNewLine & _
            "             From 门诊费用记录 A, 保险结算记录 M, 费用审核记录 Q " & vbNewLine & _
            "             Where A.记录性质 = 1 And A.病人ID= [1] " & _
            "                   And A.记录状态 <> 0 And A.结帐id = M.记录id(+) " & vbNewLine & _
            "                   And  M.性质(+) = 1 And A.ID = Q.费用id(+) And Nvl(a.附加标志,0) <> 9 " & vbNewLine & _
            "                   And a.Id In (Select b.Id " & vbNewLine & _
            "                        From 门诊费用记录 B, 门诊费用记录 C, 费用审核记录 D" & vbNewLine & _
            "                        Where c.Id = d.费用id And d.记录状态 = 1 And b.No = c.No))" & vbNewLine & _
            "       Where Rn < 2" & _
            "    Group By NO, 实际票号, 开单人, 费别 " & _
            "    Having Sum(数次) <> 0) A, 门诊费用记录 B, 病人预交记录 C " & _
            " Where a.Id = b.Id And b.结帐ID=c.结帐ID And Nvl(C.结算序号,1) > 0"
    Else
        '记帐单
        strSql = "" & _
            " Select /*+ rule */   '√' as 选择,'记帐' as 类别,Decode(NULL,Null,'','√') as 医保, A.NO As 首张单据, A.实际票号 As 首张发票, 0 As 结算序号,A.费别," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.应收金额), '999999999" & gstrDec & "')) As 应收金额," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.实收金额), '999999999" & gstrDec & "')) As 实收金额,A.操作员姓名 As 操作员,A.开单人," & vbNewLine & _
            "       Max(A.登记时间) As 登记时间,0 AS 险类 " & vbNewLine & _
            " From 门诊费用记录 A" & vbNewLine & _
            " Where A.记录性质 =2 And A.记录状态 <> 0 " & strWhere & vbNewLine & _
            "           And Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.记录性质=A.记录性质 And K.附加标志 <> 9 Group By K.序号 Having Sum(K.数次) <> 0) " & vbNewLine & _
            "      And Exists (Select 1 From 费用审核记录 E,门诊费用记录 F Where E.记录状态=1 And F.ID=E.费用ID And F.NO=A.NO And MOD(F.记录性质,10)=2)  " & _
            "Group By A.NO, A.实际票号, A.开单人, A.操作员姓名,A.费别 " & vbNewLine
            
    End If
    
    strSql = strSql & " Order By 类别,类别, 首张发票 Desc, 首张单据 Desc"
    If mrsFeeList Is Nothing Or blnFilter = False Then
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, strNos)
    Else
        mrsFeeList.Filter = 0
    End If
    mlng病人ID = lng病人ID
    vsFee.Redraw = flexRDNone
    vsFee.Clear: vsFee.Cols = 0
    Set vsFee.DataSource = mrsFeeList
    If vsFee.Rows <= 1 Then vsFee.Rows = 2
    With vsFee
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",险类,编码,序号,从属父号,转出标志,收费类别,结算序号,结帐ID,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              ElseIf .ColKey(lngCol) Like "选择*" Then
                    .ColAlignment(lngCol) = flexAlignCenterCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, IIf(mintType = 1, "退费列表", "销帐列表"), True
        '画线
        Dim strNO As String, str单据 As String
        strALLNOs = ""
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("结算序号"))) _
                 And strNO <> "" Then
                '画出分隔线
                .Select lngRow, .FixedCols, lngRow, .Cols - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("结算序号")) = .TextMatrix(lngRow, .ColIndex("结算序号"))
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("结算序号")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("类别")))
            strALLNOs = strALLNOs & "," & strNO
        Next
        .Editable = flexEDNone
    End With
    
    If strALLNOs <> "" Then strALLNOs = Mid(strALLNOs, 2)
    If blnFilter = False Then zlCommFun.StopFlash
    vsFee.Redraw = flexRDBuffered
    '加载结算方式
    Call CheckInsure
    Call InitBlanceData(strALLNOs)
    Call SetBlanceShow
    Call CalcSUMMony
    Call frmFeeRefundment.StatusShowBillSum
    Call vsFee_AfterRowColChange(0, 0, 1, 0)
    Call picBalance_Resize
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsFee.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function

Public Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐或退费
    '返回:退费或销帐成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-23 11:21:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lng险类 As Long, lng结帐ID As Long
    Dim strOutNos As String, strTemp As String, strDelDate As String
    Dim m As Long, i As Long, blnHaveData As Boolean, blnPrintList As Boolean '是否打印清单
    Dim cllDelNO As Collection, strDelNOs As String, lngRow As Long, strNO As String
    Dim lng病人ID As Long, rsTmp As ADODB.Recordset, blnOld As Boolean, strFirstNo As String
    
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    blnPrintList = False
    If InStr(mstrPrivs, ";打印清单;") > 0 And mintType = 1 Then
        Select Case mint收费清单    '0-不打印,1-要打印,2-选择是否打印
        Case 2
             If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
             End If
        Case 1
            blnPrintList = True
        End Select
    End If
    
    With vsFee
        If .Rows <= 1 Then Exit Function
        If .Cols <= 1 Then Exit Function
        Set cllDelNO = New Collection
        strTemp = ""
        For lngRow = 1 To .Rows - 1
            '销帐单据
            If mintType = 1 Then
                strNO = Trim(.TextMatrix(lngRow, .ColIndex("结算序号")))
                If CheckBillExistReplenishData(0, Val(strNO)) Then
                    MsgBox "选择的退费单据存在补充结算记录，无法进行退费！", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                strNO = Trim(.TextMatrix(lngRow, .ColIndex("首张单据")))
            End If
            If .TextMatrix(lngRow, .ColIndex("选择")) <> "" _
                And strNO <> "" And InStr(1, "," & strTemp & ",", "," & strNO & ",") = 0 Then
                lng险类 = Val(.TextMatrix(lngRow, .ColIndex("险类")))
                If mintType = 1 Then
                    If Val(strNO) > 0 Then
                        blnOld = True
                        strOutNos = strNO
                        strFirstNo = NVL(.TextMatrix(lngRow, .ColIndex("首张单据")))
                        lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结帐ID")))
                        cllDelNO.Add Array(strFirstNo, strFirstNo, lng险类, lng结帐ID, True, strFirstNo)
                        strTemp = strTemp & "," & strNO & "," & strOutNos
                    Else
                        lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结算序号")))
                        strFirstNo = NVL(.TextMatrix(lngRow, .ColIndex("首张单据")))
                        strOutNos = strNO
                        
                        If strOutNos <> "" Then
                            '检查该张单是否存在
                            blnHaveData = False
                            For i = 1 To cllDelNO.Count
                                If cllDelNO(i)(0) = strNO Then
                                    blnHaveData = True: Exit For
                                End If
                                If InStr(1, "," & cllDelNO(i)(1) & ",", "," & strNO & ",") > 0 Then
                                    blnHaveData = True: Exit For
                                End If
                            Next
                            If blnHaveData = False Then
                                '加入销帐单据
                                cllDelNO.Add Array(strNO, strOutNos, lng险类, lng结帐ID, False, strFirstNo)
                            End If
                            strTemp = strTemp & "," & strNO & "," & strOutNos
                        End If
                    End If
                Else
                    lng结帐ID = Val(.TextMatrix(lngRow, .ColIndex("结算序号")))
                    strOutNos = strNO
                    
                    If strOutNos <> "" Then
                        '检查该张单是否存在
                        blnHaveData = False
                        For i = 1 To cllDelNO.Count
                            If cllDelNO(i)(0) = strNO Then
                                blnHaveData = True: Exit For
                            End If
                            If InStr(1, "," & cllDelNO(i)(1) & ",", "," & strNO & ",") > 0 Then
                                blnHaveData = True: Exit For
                            End If
                        Next
                        If blnHaveData = False Then
                            '加入销帐单据
                            cllDelNO.Add Array(strNO, strOutNos, lng险类, lng结帐ID)
                        End If
                        strTemp = strTemp & "," & strNO & "," & strOutNos



                    End If

                End If
            End If
        Next
    End With
    '执行具体销帐或退费操作
    If cllDelNO.Count = 0 Then
        MsgBox "注意:" & vbCrLf & "    没有选择一张需要进行退费或销帐的单据,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '退费
    strDelNOs = ""
    If mintType = 2 Then
        If ExecuteWirteOff(strDelDate, cllDelNO) = False Then Exit Function
    Else
        For i = 1 To cllDelNO.Count
            If ExecuteDelBill(strDelDate, IIf(cllDelNO(i)(1) <> "", cllDelNO(i)(1), cllDelNO(i)(0)), Val(cllDelNO(i)(2)), Val(cllDelNO(i)(2)), cllDelNO(i)(4), cllDelNO(i)(5)) = False Then
                    Exit Function
            End If
            strDelNOs = strDelNOs & "," & cllDelNO(i)(5)
        Next
    End If
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    '打印费用清单
    If blnPrintList And mintType = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & "'" & Replace(strDelNOs, ",", "','") & "'", "药品单位=" & IIf(mbln药房单位, 1, 0), 2)
    End If
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteWirteOff(strDELDae As String, ByVal cllDel As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行门诊记帐销帐
    '编制:刘兴洪
    '日期:2011-02-25 10:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSql As String
    Dim cllPro As Collection
    Set cllPro = New Collection
    For i = 1 To cllDel.Count
        'Zl_门诊转住院_记帐转出
        strSql = "Zl_门诊转住院_记帐转出("
        '  No_In         住院费用记录.NO%Type,
        strSql = strSql & "'" & cllDel(i)(0) & "',"
        '  操作员编号_In 住院费用记录.操作员编号%Type,
        strSql = strSql & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
        strSql = strSql & "'" & UserInfo.姓名 & "',"
        '  退费时间_In   住院费用记录.发生时间%Type
        strSql = strSql & "To_Date('" & strDELDae & "','yyyy-mm-dd hh24:mi:ss'),"
        '   门诊销帐_In   Number := 0
        '   --门诊销帐_In:0-门诊转住院立即销帐;1-门诊记帐退费模式
        strSql = strSql & "1)"
        zlAddArray cllPro, strSql
    Next
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ExecuteWirteOff = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceSet() As ADODB.Recordset
'功能：返回一个结算记录集对象
    Dim rsTmp As New ADODB.Recordset
       
    rsTmp.Fields.Append "单据序号", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "结算方式", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "结算金额", adCurrency, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set GetBalanceSet = rsTmp
End Function

Private Function ExecuteDelBill(ByVal strDelDate As String, ByVal strNos As String, intInsure As Integer, _
                                ByVal lng结帐ID As Long, Optional ByVal blnOld As Boolean, Optional ByVal strFirstNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关退费操作
    '入参:strNos-单据号:可以是多单据
    '       lngInsure-险类
    '返回:执行成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-24 15:35:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, k As Long, varTemp  As Variant, strAllBalance      As String, strBalance As String
    Dim bln医保接口打印票据 As Boolean, bln多单据一次结算 As Boolean, blnYB结算作废 As Boolean, bln退费后打印回单 As Boolean
    Dim lng领用ID As Long, cllPro As Collection, blnTrans As Boolean, lng冲销ID As Long, str交易流水号 As String, str交易说明 As String
    Dim lng结帐ID1 As Long, varBalance As Variant, strAdvance As String, strInvoice As String
    Dim strSql As String, j As Long, blnTransMedicare As Boolean, rsTmp As ADODB.Recordset
    Dim str结算方式 As String, cur结算金额 As Currency, cur可分配额 As Currency, cur误差金额 As Currency, cur余额 As Currency, cur退款合计 As Currency
    Dim strDelNOs As String, lng病人ID As Long, blnExecuteThreeSwap As Boolean
    
    If intInsure <> 0 Then
        bln医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, , intInsure, CStr(lng结帐ID))
        bln多单据一次结算 = gclsInsure.GetCapability(support多单据一次结算, , intInsure)
        blnYB结算作废 = gclsInsure.GetCapability(support门诊结算作废, , intInsure)
        If blnYB结算作废 = False Then
            MsgBox "注意:" & vbCrLf & "   单据号为" & strNos & "的单据,不支持医保结算作废,请检查"
            Exit Function
        End If
        bln退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, , intInsure)
    End If
    
    If intInsure <> 0 And bln医保接口打印票据 Then
        Dim strUserType As String
        Dim lngShareUseID As Long
        If mrsInfo Is Nothing Then
            lng病人ID = mlng病人ID
        ElseIf mrsInfo.State <> 1 Then
            lng病人ID = mlng病人ID
        Else
            lng病人ID = Val(NVL(mrsInfo!病人ID))
        End If
        strUserType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
        lngShareUseID = zl_GetInvoiceShareID(1121, strUserType)
         
        lng领用ID = GetInvoiceGroupID(1, 1, lng领用ID, lngShareUseID)
        Select Case lng领用ID
            Case -1
                MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Exit Function
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Exit Function
        End Select
        strInvoice = GetNextBill(lng领用ID)
    End If
    
    '获取结帐ID
    Err = 0: On Error GoTo errHandle
    Set cllPro = New Collection
    varTemp = Split(strNos, ",")
    For i = 0 To UBound(varTemp)
            'Zl_门诊转住院_收费转出
            strSql = "Zl_门诊转住院_收费转出("
            '     结算序号_In   病人预交记录.结算序号%Type,
            strSql = strSql & IIf(blnOld, "Null,", "'" & varTemp(i) & "',")
            '     NO_In         门诊费用记录.NO%Type,
            strSql = strSql & IIf(blnOld, "'" & varTemp(i) & "',", "Null,")
            '     操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '     操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '     退费时间_In   住院费用记录.发生时间%Type,
            strSql = strSql & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '     门诊退费_In   Number := 0(门诊退费_In:0-门诊转住院立即销帐;1-门诊退费模式:为1时:入院科室id_In和主页ID_IN可以不传)
            strSql = strSql & "1,"
            '     入院科室id_In 住院费用记录.开单部门id%Type := Null,
            strSql = strSql & "Null,"
            '     主页id_In     住院费用记录.主页id%Type := Null
            strSql = strSql & "Null,"
            '     结算方式_In   病人预交记录.结算方式%Type := Null
            strSql = strSql & IIf(picBack.Visible, "'" & cboStyle.Text & "'", "Null") & ","
           
           strAllBalance = strAllBalance & "," & lng结帐ID
           cllPro.Add Array(strSql, lng结帐ID, varTemp(i), CStr(varTemp(i)))
    Next
    mstrThreeSwapBalance = ""
    mstrThreeSwapCardType = ""
    mstrThreeSwapMoney = ""
    
     If intInsure <> 0 And bln多单据一次结算 Then
        On Error GoTo errH: blnTrans = True
        
        gcnOracle.BeginTrans
            '从最后一张开始退
        For i = cllPro.Count To 1 Step -1
            blnExecuteThreeSwap = False
            lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
            Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & ")", Me.Caption)
            
            If ExecuteThreeSwap(Val(cllPro(i)(1)), lng冲销ID, str交易流水号, str交易说明) = True Then
                blnExecuteThreeSwap = True
            End If
            
            'Zl_门诊转住院_三方卡结算
            strSql = "Zl_门诊转住院_三方卡结算("
            '  结算序号_In   病人预交记录.结算序号%Type,
            strSql = strSql & IIf(blnOld, "Null,", "'" & cllPro(i)(2) & "',")
            '  No_In         住院费用记录.NO%Type,
            strSql = strSql & IIf(blnOld, "'" & cllPro(i)(2) & "',", "Null,")
            '  操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
            strSql = strSql & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  门诊退费_In   Number := 0,
            strSql = strSql & "" & 1 & ","
            '  入院科室id_In 病人预交记录.科室id%Type,
            strSql = strSql & "Null,"
            '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
            strSql = strSql & "Null,"
            '  三方退费_In   Number := 0,
            strSql = strSql & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
            '  结帐ID_In     住院费用记录.结帐id%Type)
            strSql = strSql & "" & lng冲销ID & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "三方卡结算")
        Next
        
        '先产生票据，医保接口才能取到
        If bln医保接口打印票据 Then
            strSql = "zl_门诊收费记录_RePrint('" & strFirstNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(strDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        strAdvance = strAllBalance
        If Not gclsInsure.ClinicDelSwap(Val(cllPro(cllPro.Count)(1)), , intInsure, strAdvance) Then
            GoTo errH
        Else
            blnTransMedicare = True
        End If
        
        If Not (strAdvance = strAllBalance Or strAdvance = "") Then
            '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2:金额...
            '先分摊到每张单据上
            Set rsTmp = GetBalanceSet
            varBalance = Split(strAdvance, "||")
            For i = 0 To UBound(varBalance)
                str结算方式 = Split(varBalance(i), "|")(0)
                cur结算金额 = -1 * Val(Split(varBalance(i), "|")(1))
                For k = 0 To UBound(varTemp)
                    cur可分配额 = Get实收金额(varTemp(k))
                    rsTmp.Filter = "单据序号=" & k
                    For j = 1 To rsTmp.RecordCount
                        cur可分配额 = cur可分配额 - rsTmp!结算金额
                        rsTmp.MoveNext
                    Next
                    If cur可分配额 > 0 Then
                        If cur可分配额 <= cur结算金额 Then
                            cur结算金额 = cur结算金额 - cur可分配额
                        Else
                            cur可分配额 = cur结算金额
                            cur结算金额 = 0
                        End If
                        rsTmp.AddNew
                        rsTmp!单据序号 = k
                        rsTmp!结算方式 = str结算方式
                        rsTmp!结算金额 = cur可分配额
                        rsTmp.Update
                        
                        If cur结算金额 = 0 Then Exit For
                    End If
                Next
            Next
            
            For k = 0 To UBound(varTemp)
                strBalance = ""
                cur误差金额 = 0
                cur余额 = Get实收金额(varTemp(k))
                
                rsTmp.Filter = "单据序号=" & k
                For i = 1 To rsTmp.RecordCount
                    strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!结算方式 & "|" & -1 * rsTmp!结算金额
                    cur余额 = cur余额 - rsTmp!结算金额
                    rsTmp.MoveNext
                Next

                '退为指定的结算方式，如果是现金，可能产生新的误差金额
                'If cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1 Then
                    cur结算金额 = Format(CentMoney(cur余额), "0.00")
                    cur误差金额 = cur结算金额 - cur余额
'                Else
'                    cur结算金额 = cur余额
'                End If
                cur退款合计 = cur退款合计 + cur结算金额
                lng结帐ID = GetDelBalanceID(varTemp(k))
                strSql = "zl_门诊收费结算_Update(" & lng结帐ID & ",'" & "现金" & "|" & -1 * cur结算金额 & "| ',0,'" & strBalance & "'," & -1 * cur误差金额 & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
     Else
         '从最后一张开始退
        For i = cllPro.Count To 1 Step -1
            gcnOracle.BeginTrans: On Error GoTo errH: blnTrans = True
            mstrThreeSwapBalance = ""
            mstrThreeSwapCardType = ""
            mstrThreeSwapMoney = ""
            blnExecuteThreeSwap = False
            
            lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
            Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng冲销ID & ")", Me.Caption)
            
            blnTransMedicare = False
            If intInsure <> 0 Then                    '处理医保接口
                  If blnYB结算作废 Then
                        strAdvance = cllPro.Count & "|" & i
                        If Not gclsInsure.ClinicDelSwap(CStr(cllPro(i)(1)), True, intInsure, strAdvance) Then
                            GoTo errH
                        Else
                            blnTransMedicare = True
                        End If
                    End If
            End If
            gcnOracle.CommitTrans: blnTrans = False
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
            
            If ExecuteThreeSwap(Val(cllPro(i)(2)), lng冲销ID, str交易流水号, str交易说明) = True Then
                blnExecuteThreeSwap = True
            End If
            
            'Zl_门诊转住院_三方卡结算
            strSql = "Zl_门诊转住院_三方卡结算("
            '  结算序号_In   病人预交记录.结算序号%Type,
            strSql = strSql & IIf(blnOld, "Null,", "'" & cllPro(i)(2) & "',")
            '  No_In         住院费用记录.NO%Type,
            strSql = strSql & IIf(blnOld, "'" & cllPro(i)(2) & "',", "Null,")
            '  操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  退费时间_In   住院费用记录.登记时间%Type, --多张单据退费时,每张单据的退费时间相同,都是系统当前时间
            strSql = strSql & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  门诊退费_In   Number := 0,
            strSql = strSql & "" & 1 & ","
            '  入院科室id_In 病人预交记录.科室id%Type,
            strSql = strSql & "Null,"
            '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
            strSql = strSql & "Null,"
            '  三方退费_In   Number := 0,
            strSql = strSql & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
            '  结帐ID_In     住院费用记录.结帐id%Type)
            strSql = strSql & "" & lng冲销ID & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "三方卡结算")
            
            strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & cllPro(i)(0)
        Next
     End If
    
    If intInsure <> 0 And bln退费后打印回单 And InStr(1, mstrPrivs, ";医保退费回单;") > 0 Then
        '问题:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO='" & strFirstNo & "'", 2)
    End If
    ExecuteDelBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTrans Then
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    
    If Err.Number <> 0 Then Call SaveErrLog
    
    '中断提示,不打印，重新退费后再打印或自己选择重打
    If strDelNOs <> "" Then
        MsgBox "单据[" & strNos & "]退费失败。但是，单据[" & strDelNOs & "]已成功退费。" & vbCrLf & _
            "单据未打印，请对执行失败的单据重新退费！", vbInformation, gstrSysName
    End If
    Exit Function
End Function

Private Function ExecuteThreeSwap(lngBalance As Long, lng冲销ID As Long, Optional ByRef str交易流水号 As String, Optional ByRef str交易说明 As String) As Boolean
    Dim strSql As String, rsTemp As ADODB.Recordset, strBalanceIDs As String, rsTotal As ADODB.Recordset
    Dim dblMoney As Double
    If mobjSquare Is Nothing Then Exit Function
    strSql = _
        " Select a.卡类别id, a.卡号, Min(a.结帐id) As 结帐id, Sum(a.冲预交) As 冲预交, Min(a.交易流水号) As 交易流水号, Min(a.交易说明) As 交易说明" & vbNewLine & _
        " From 病人预交记录 A, 结算方式 B," & vbNewLine & _
        "      (Select Distinct k.结帐id" & vbNewLine & _
        "        From 病人预交记录 I, 门诊费用记录 J, 门诊费用记录 K" & vbNewLine & _
        "        Where i.结算序号 = [1] And i.记录性质 = 3 And i.结帐id = j.结帐id And k.No = j.No And Mod(k.记录性质, 10) = 1) C, 医疗卡类别 D" & vbNewLine & _
        " Where a.结算方式 = b.名称 And b.性质 = 7 And a.结帐id = c.结帐id And a.卡类别id = d.id And d.是否退现 = 0 And a.校对标志 <> 1" & vbNewLine & _
        " Group By a.卡类别id, a.卡号" & vbNewLine & _
        " Having Sum(a.冲预交) <> 0"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalance)
    
'    strSQL = _
'        " Select  Sum(a.冲预交) As 冲预交 " & vbNewLine & _
'        " From 病人预交记录 A, 结算方式 B," & vbNewLine & _
'        "      (Select Distinct k.结帐id" & vbNewLine & _
'        "        From 病人预交记录 I, 门诊费用记录 J, 门诊费用记录 K" & vbNewLine & _
'        "        Where i.结算序号 = [1] And i.记录性质 = 3 And i.结帐id = j.结帐id And k.No = j.No And Mod(k.记录性质, 10) = 1) C" & vbNewLine & _
'        " Where a.结算方式 = b.名称 And a.结帐id = c.结帐id" & vbNewLine & _
'        " Having Sum(a.冲预交) <> 0"
'
'    Set rsTotal = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalance)
    
    If rsTemp.RecordCount = 0 Then Exit Function
'    If rsTotal.RecordCount = 0 Then Exit Function
    
'    If Val(NVL(rsTemp!冲预交)) > Val(NVL(rsTotal!冲预交)) Then
'        dblMoney = Val(NVL(rsTotal!冲预交))
'    Else
    dblMoney = Val(NVL(rsTemp!冲预交))
'    End If
    
    Do While Not rsTemp.EOF
        strBalanceIDs = "3|" & Val(NVL(rsTemp!结帐ID))
        If mobjSquare.zlReturnCheck(Me, mlngModule, Val(NVL(rsTemp!卡类别ID)), False, NVL(rsTemp!卡号), _
            strBalanceIDs, dblMoney, str交易流水号, str交易说明, "3|" & lng冲销ID) = False Then Exit Function
        If mobjSquare.zlReturnMoney(Me, mlngModule, Val(NVL(rsTemp!卡类别ID)), False, NVL(rsTemp!卡号), _
            strBalanceIDs, dblMoney, str交易流水号, str交易说明, "3|" & lng冲销ID) = False Then Exit Function
        mstrThreeSwapBalance = mstrThreeSwapBalance & "|" & lngBalance
        mstrThreeSwapCardType = mstrThreeSwapCardType & "|" & Val(NVL(rsTemp!卡类别ID))
        mstrThreeSwapMoney = mstrThreeSwapMoney & "|" & dblMoney
        rsTemp.MoveNext
    Loop
    ExecuteThreeSwap = True
End Function

Public Function Get实收金额(ByVal strNO As String) As Currency
    Dim i As Long, cur金额 As Currency
    With vsFee
        cur金额 = 0
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("结算序号")) = strNO Then
                cur金额 = cur金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
        Get实收金额 = cur金额
    End With
End Function

Public Sub ClearData()
    Dim i As Integer
    vsFee.Clear 1
    vsFee.Rows = 2
    vsfInvoice.Clear 1
    vsfInvoice.Rows = 2
    vsBalanceStyle.Clear 1
    vsBalanceStyle.Rows = 2
    For i = 1 To vsBalance.Cols - 1
        vsBalance.TextMatrix(0, i) = ""
        vsBalance.TextMatrix(1, i) = ""
    Next i
    txtSum.Text = 0
    picBack.Visible = False
End Sub

Private Sub vsFee_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsFee_DblClick()
    With vsFee
        If .TextMatrix(.Row, .ColIndex("类别")) = "" Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("选择")) = "" Then
            Call SetRowSelected(.Row, True)
        Else
            Call SetRowSelected(.Row, False)
        End If
    End With
    Call SetBlanceShow
    Call CalcSUMMony
End Sub

Private Sub vsFee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        With vsFee
            If .TextMatrix(.Row, .ColIndex("选择")) = "" Then
                Call SetRowSelected(.Row, True)
            Else
                Call SetRowSelected(.Row, False)
            End If
        End With
        Call SetBlanceShow
        Call CalcSUMMony
    End If
End Sub
