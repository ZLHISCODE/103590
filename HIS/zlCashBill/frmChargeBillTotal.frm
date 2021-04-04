VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeBillTotal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picReturnBill 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   5370
      ScaleHeight     =   2685
      ScaleWidth      =   3630
      TabIndex        =   7
      Top             =   2730
      Width           =   3630
      Begin VSFlex8Ctl.VSFlexGrid vsReturnBill 
         Height          =   1800
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmChargeBillTotal.frx":0000
         ScrollTrack     =   -1  'True
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
   Begin VB.PictureBox picBillInfor 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   570
      ScaleHeight     =   1470
      ScaleWidth      =   3630
      TabIndex        =   1
      Top             =   4365
      Width           =   3630
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   870
         Left            =   0
         TabIndex        =   5
         Top             =   90
         Width           =   1860
         _cx             =   3281
         _cy             =   1535
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillTotal.frx":007A
         ScrollTrack     =   -1  'True
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
   Begin VB.PictureBox picChargeInfor 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   450
      ScaleHeight     =   2685
      ScaleWidth      =   3630
      TabIndex        =   0
      Top             =   1170
      Width           =   3630
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1170
         TabIndex        =   4
         Top             =   2010
         Width           =   2280
      End
      Begin VSFlex8Ctl.VSFlexGrid vsChagre 
         Height          =   1800
         Left            =   315
         TabIndex        =   2
         Top             =   15
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillTotal.frx":00F4
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "收款合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   405
         TabIndex        =   3
         Top             =   2070
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRptPrint 
      Height          =   1560
      Left            =   765
      TabIndex        =   6
      Top             =   6105
      Visible         =   0   'False
      Width           =   3735
      _cx             =   6588
      _cy             =   2752
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChargeBillTotal.frx":016E
      ScrollTrack     =   -1  'True
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   45
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeBillTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long, mstrPrivs As String
Private Enum TotalType
    EM_收费员轧帐 = 1
    EM_小组收款 = 2
    EM_小组轧帐 = 3
    EM_财务收款 = 4
    EM_财务收款_非收费员 = 5
End Enum
'1-收费员轧帐；2-小组收款;3-小组轧帐;4-财务收款(针对收费员或财务组收款)或财务收款查询;5-财务收款(仅针对非收费员收款)。
Private mbytType As TotalType
Private mlngChargeRollingID As Long '轧帐ID或收款ID(根据mbytType)来决定
Private mdtStartDate As Date, mdtendDate As Date '轧帐的开始时间或结束时间
Private mblnDel As Boolean '是否作废记录
Public mrsList As ADODB.Recordset '收款相关记录
Public mrsListBill As ADODB.Recordset '票据相关记录
Private mrsBalance As ADODB.Recordset '操作员余额
Private mblnHideFilter As Boolean '是否隐藏过滤条件
Private mlngErrorRow As Long
Private mdblRemain As Double
Private mblnOlnyView As Boolean '仅能查看数据,不能编辑实际票号
Private Enum mPaneIndex
    EM_PN_ChargeTotal = 260102  '收款汇总
    EM_PN_BillTotal = 260103    '票据汇总
    EM_PN_BackFeeBill = 260104  '退费票据
    EM_PN_ReprintBill = 260105  '重打票据
End Enum
Private mlngCashRow As Long '现金所指定的行
Private mbytFontSize As Byte
Private mstrPersonName  As String '当前收费人员
Private mstrRollingType As String  '轧帐类别

Public Sub ClearData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
    '编制:刘兴洪
    '日期:2013-09-12 11:09:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call InitGrid
    mlngCashRow = 0
    txtTotal.Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
Public Function LoadChargeAndBillTotalData(ByVal frmMain As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As Byte, ByVal lngChargeRollingID As Long, _
      Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, _
      Optional blnOlnyView As Boolean = True, _
      Optional ByVal blnDel As Boolean = False, _
      Optional strPersonName As String = "", _
      Optional strRollingType As String) As Boolean
    '-------------------------------------------------------------------------------------------------
    '功能:收费员轧账接口
    '入参:frmMain-调用的主窗体
    '    lngModule-模块号
    '    strPrivs-权限串
    '　　bytType:1-收费员轧帐；2-小组收款;3-小组轧帐;
    '            4-财务收款(针对收费员或财务组收款)或财务收款查询;
    '            5-财务收款(仅针对非收费员收款)。
    '    lngChargeRollingID -收费员的轧帐ID
    '    dtStartDate-可选参数,开始轧帐时间,lngChargeRollIngID=0时，必须传入
    '    dtEndDate-可选参数，结束轧帐时间,lngChargeRollIngID=0时，必须传入
    '    blnOlnyView-仅能查看(不能编制票据号码)
    '    blnDel-是否作废记录
    '    strPersonName-指定的收费员(为空时,为当前操作员)
    '    strRollingType-轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡)
    '返回:数据加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-08-13 10:31:00
    '-------------------------------------------------------------------------------------------------
    mbytType = bytType:  mlngChargeRollingID = lngChargeRollingID
    mdtStartDate = dtStartDate: mdtendDate = dtEndDate
    mlngMode = lngModule: mstrPrivs = strPrivs
    mblnDel = blnDel
    mblnOlnyView = blnOlnyView
    vsChagre.Editable = IIf(mblnOlnyView, flexEDNone, flexEDKbdMouse)
    mstrPersonName = IIf(strPersonName = "", UserInfo.姓名, strPersonName)
    mstrRollingType = strRollingType
    
    If Not mblnOlnyView Then
        If Not zlStr.IsHavePrivs(mstrPrivs, "轧帐") Then vsChagre.Editable = flexEDNone
    End If
    LoadChargeAndBillTotalData = ReadChargeBillData
End Function
Private Function ReadChargeBillData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取收款及票据汇总数据
    '返回:数据获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-04 10:29:42
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mlngChargeRollingID = 0 And (mbytType = EM_收费员轧帐 Or mbytType = EM_财务收款_非收费员) Then
         ReadChargeBillData = LoadPersonChargeAndBill         '加载收费员轧帐记录或财务非收费员
    Else
         ReadChargeBillData = LoadChargeAndBillAndTotal          '加载相关的收款使用及票据汇总
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function GetHandInFee() As Double
    Dim dblTotal As Double
    If (mlngCashRow = 0 And mlngErrorRow = 0) Or (mlngCashRow > vsChagre.Rows And mlngErrorRow > vsChagre.Rows) Then Exit Function
    With vsChagre
        dblTotal = Val(.TextMatrix(mlngErrorRow, .ColIndex("金额"))) + Val(.TextMatrix(mlngCashRow, .ColIndex("金额")))
    End With
    GetHandInFee = dblTotal
End Function
 
Private Function LoadPersonChargeAndBill() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费员轧帐数据汇总
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-04 11:28:56
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bytType As Byte, rsBalanceMode As ADODB.Recordset
    Dim strWithTable As String, lngRow As Long, lngNo As Long
    Dim str现金 As String, str结算方式 As String, strTemp As String, int票种 As Integer
    Dim dblToTotal As Double, i As Integer, dblInsure As Double, intInsureRow As Integer
    Dim blnTempDelete As Boolean, strRollingType As String, strWhere As String
    Dim strCardUseRecord As String
    On Error GoTo errHandle
    
    blnTempDelete = False
    If mstrPersonName = "" Or mstrPersonName = "-" Then
        '只清除数据
        Call ClearData: LoadPersonChargeAndBill = True
        Exit Function
    End If
    
    bytType = 1: dblToTotal = 0
    Set rsBalanceMode = Get结算方式
    str现金 = "现金"
    rsBalanceMode.Filter = "性质=1"
    If Not rsBalanceMode.EOF Then
        str现金 = rsBalanceMode!名称
    End If
    
    Set mrsBalance = GetBalance(str现金)
    
    '1.加载收款数据汇总
    '预交款填NULL,2-结帐,3-收费,4-挂号,5-就诊卡,6-补充医保结算
    'mstrRollingType:轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
    strRollingType = "": strWhere = ""
    strSQL = ""
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",1,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",3,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",4,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",5,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",6,") > 0 Then
        If mstrRollingType <> "" Then
            strWhere = " And Instr([6],','|| Nvl(A.结算性质,0)||',')> 0"
        End If
        If InStr("," & mstrRollingType & ",", ",0,") > 0 Then
            strRollingType = ",2,3,4,5,6,"
        Else
            If InStr("," & mstrRollingType & ",", ",1,") > 0 Then
                strRollingType = "3,6,"
            End If
            If InStr("," & mstrRollingType & ",", ",3,") > 0 Then
                strRollingType = strRollingType & "2,"
            End If
            If InStr("," & mstrRollingType & ",", ",4,") > 0 Then
                strRollingType = strRollingType & "4,"
            End If
            If InStr("," & mstrRollingType & ",", ",5,") > 0 Then
                strRollingType = strRollingType & "5,"
            End If
            If InStr("," & mstrRollingType & ",", ",6,") > 0 Then
                strRollingType = strRollingType & "-,"
            End If
            If strRollingType <> "" Then strRollingType = "," & strRollingType
        End If

        strSQL = "" & _
        "   Select Decode(Nvl(a.结算性质, 0), 2, 2, 6, 9, 1) As 性质, a.结帐id, Decode(Mod(a.记录性质, 10), 1, '[冲预交款]', a.结算方式) As 结算方式," & vbNewLine & _
        "              Sum(a.冲预交) As 金额, Sum(Decode(Mod(a.记录性质, 10), 1, 1, 0) * a.冲预交) As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
        "       From 病人预交记录 A" & vbNewLine & _
        "       Where a.操作员姓名 || '' = [3] and a.记录性质<>1 " & strWhere & vbNewLine & _
        "       And a.收款时间 Between [4] And [5] " & vbNewLine & _
        "       And Not Exists(Select 1 From 门诊费用记录 B Where a.结帐id = b.结帐id And Nvl(b.费用状态, 0) = 1) " & vbNewLine & _
        "       And Not Exists(Select 1 From 病人结帐记录 B Where a.结帐id = b.Id And b.结算状态 Is Not Null)" & vbNewLine & _
        "       And Not Exists(Select 1 From 费用补充记录 B Where a.结帐id = b.结算id And Nvl(b.费用状态, 0) >= 1)" & vbNewLine & _
        "       Group By Decode(Nvl(a.结算性质, 0), 2, 2, 6, 9, 1), a.结帐id, Decode(Mod(a.记录性质, 10), 1, '[冲预交款]', a.结算方式)" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",2,") > 0 Then '预交款
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As 性质, ID As 结帐id, a.结算方式, a.金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
        "       From 病人预交记录 A" & vbNewLine & _
        "       Where 记录性质 = 1 And 操作员姓名 || '' = [3] And 收款时间 Between [4] And [5] And Nvl(结算性质,0) <> 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",21,") > 0 Then '门诊预交款
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As 性质, ID As 结帐id, a.结算方式, a.金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
        "       From 病人预交记录 A" & vbNewLine & _
        "       Where 记录性质 = 1 And Nvl(预交类别,0) = 1 And 操作员姓名 || '' = [3] And 收款时间 Between [4] And [5] And Nvl(结算性质,0) <> 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",22,") > 0 Then '住院预交款
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As 性质, ID As 结帐id, a.结算方式, a.金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
        "       From 病人预交记录 A" & vbNewLine & _
        "       Where 记录性质 = 1 And Nvl(预交类别,0) = 2  And 操作员姓名 || '' = [3] And 收款时间 Between [4] And [5] And Nvl(结算性质,0) <> 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",3,") > 0 Then '结帐产生预交
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As 性质, ID As 结帐id, a.结算方式, a.金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
        "       From 病人预交记录 A" & vbNewLine & _
        "       Where 记录性质 = 1 And 操作员姓名 || '' = [3] And 收款时间 Between [4] And [5] And Nvl(结算性质,0) = 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",6,") > 0 Then '消费卡充值
        strSQL = strSQL & _
            IIf(strSQL <> "", "Union All", "") & vbNewLine & _
            "Select 5 As 性质, a.结算Id As 结帐id, a.结算方式, a.实收金额 As 金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
            "From 病人卡结算记录 A, 病人卡结算记录 B" & vbNewLine & _
            "Where a.交易序号 = b.交易序号(+) And a.消费卡id = b.消费卡id(+)  " & vbNewLine & _
            "      And (a.记录性质 = 2 Or a.记录性质 = 3 And b.记录性质 = 2) And b.记录性质(+) = 2 " & vbNewLine & _
            "      And a.Id <> b.Id(+) And a.操作员姓名 || '' = [3] And a.登记时间 Between [4] And [5]" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select 6 As 性质, a.结算Id, a.结算方式, a.实收金额 As 金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
            "From 病人卡结算记录 A, 病人卡结算记录 B" & vbNewLine & _
            "Where a.交易序号 = b.交易序号(+) And a.消费卡id = b.消费卡id(+)  " & vbNewLine & _
            "      And a.Id <> b.Id(+) And (a.记录性质 = 1 Or a.记录性质 = 3 And b.记录性质 = 1) And b.记录性质(+) = 1 " & vbNewLine & _
            "      And a.操作员姓名 || '' = [3] And a.登记时间 Between [4] And [5]" & vbNewLine
    End If
    
    '借款及暂存金
    strSQL = strSQL & _
    IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
    "       Select 4 As 性质, a.Id As 结帐id, a.结算方式, Nvl(a.借款金额, 0) As 金额, 0 As 冲预交, Nvl(借款金额, 0) As 借款合计, 0 As 借出合计" & vbNewLine & _
    "       From 人员借款记录 A" & vbNewLine & _
    "       Where a.借款人 || '' = [3] And a.取消时间 Is Null " & vbNewLine & _
    "             And a.借出时间 Between [4] And [5]" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 4 As, a.Id As 结帐id, a.结算方式, -1 * Nvl(a.借款金额, 0) As 金额, 0 As 冲预交, 0 As 借款合计, Nvl(借款金额, 0) As 借出合计" & vbNewLine & _
    "       From 人员借款记录 A" & vbNewLine & _
    "       Where a.借出人 || '' = [3] And a.借出时间 Between [4] And [5] And a.取消时间 Is Null" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 7 As 性质, a.Id As 结帐id, '现金' As 结算方式, a.金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计" & vbNewLine & _
    "       From 人员暂存记录 A" & vbNewLine & _
    "       Where a.记录性质 = 2 And a.收回时间 Is Null And 收款员 || '' = [3] " & vbNewLine & _
    "             And a.登记时间 Between [4] And [5] " & vbNewLine


        
    strWithTable = "" & _
    " With 轧帐数据 as (" & vbNewLine & _
    "       Select 性质, 结帐id, 结算方式, Nvl(金额, 0) As 金额, Nvl(冲预交, 0) As 冲预交, Nvl(借款合计, 0) As 借款合计, Nvl(借出合计, 0) As 借出合计" & vbNewLine & _
    "       From ( " & strSQL & ") A" & vbNewLine & _
    "       Where Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y" & vbNewLine & _
    "                   Where y.记录id = a.结帐id And x.作废时间 Is Null And x.收款员 = [3] And x.Id = y.收缴id And y.性质 = a.性质) " & vbNewLine & _
    "                  )" & vbNewLine
 
    
    strSQL = strWithTable & vbNewLine & _
        "   Select -1 as 性质,0 as 结帐ID,结算方式,sum(nvl(金额,0)) as 金额, " & vbNewLine & _
        "           sum(nvl(冲预交,0)) as 冲预交,sum(nvl(借款合计,0)) as 借款合计,sum(nvl(借出合计,0)) as 借出合计   " & vbNewLine & _
        "   From 轧帐数据 " & _
        "   Group by 结算方式 " & _
        "   Union ALL" & _
        "   Select 性质,结帐ID,'-' as 结算方式,0 as 金额, " & vbNewLine & _
        "           0 as 冲预交,0 as 借款合计,0 as 借出合计   " & vbNewLine & _
        "   From 轧帐数据 " & vbNewLine & _
        "   Group by 性质,结帐ID " & vbNewLine & _
        "   "
    strSQL = "" & _
        "   Select  性质,  nvl(结帐ID,0) as 结帐ID,结算方式,金额,冲预交, 借款合计, 借出合计   " & vbNewLine & _
        "   From (" & strSQL & ")" & vbNewLine & _
        "   Order by 性质"
        
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtendDate, strRollingType)
    With vsChagre
        dblInsure = 0
        .Clear 1
        .Rows = 2: lngRow = 1
        mrsList.Filter = "性质=-1 And 金额 <> 0"
        mlngCashRow = 0
        
        Do While Not mrsList.EOF
            str结算方式 = Nvl(mrsList!结算方式)
            If str结算方式 = "[冲预交款]" Then
                lngNo = 1
                dblToTotal = dblToTotal + Val(Nvl(mrsList!金额))
            Else
                rsBalanceMode.Filter = "名称='" & str结算方式 & "'"
                If Not rsBalanceMode.EOF Then
                    '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
                    Select Case Val(Nvl(rsBalanceMode!性质))
                    Case 1 '现金结算方式
                        lngNo = 10
                        mlngCashRow = lngRow
                    Case 2  '其他非医保结算
                        lngNo = 11
                    Case 7 '一卡通结算
                        lngNo = 12
                    Case 8  '结算卡结算
                        lngNo = 14
                    Case 3 '个人账户
                        lngNo = 15
                    Case 4   '
                        lngNo = 16
                    Case 9
                        lngNo = 17
                        mlngErrorRow = lngRow
                    Case Else
                        lngNo = 18
                    End Select
                Else
                    lngNo = 13
                End If
                If lngNo = 15 Or lngNo = 16 Then dblInsure = dblInsure + Val(Nvl(mrsList!金额))
                dblToTotal = dblToTotal + Val(Nvl(mrsList!金额))
            End If
            
            If Not (lngNo = 15 Or lngNo = 16) Then
                .TextMatrix(lngRow, .ColIndex("序号")) = lngNo
                .TextMatrix(lngRow, .ColIndex("结算方式")) = str结算方式
                If lngNo = 17 Then
                    .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!金额)), "0.###########")
                Else
                    .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!金额)), "#,###0.00;-#,###0.00;0.00;-0.00")
                End If
                .RowData(lngRow) = Val(Nvl(mrsList!金额))
                If mlngChargeRollingID <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("结算号码")) = Nvl(mrsList!结算号)
                End If
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsList.MoveNext
        Loop
        mrsList.Filter = "性质=-1"
        If mrsList.RecordCount <> 0 Then
            mrsList.MoveFirst
            Do While Not mrsList.EOF
                If Val(Nvl(mrsList!借款合计)) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("序号")) = 2
                    .TextMatrix(lngRow, .ColIndex("结算方式")) = "[借入合计]"
                    .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!借款合计)), "#,###0.00;-#,###0.00;0.00;-0.00")
                    .RowData(lngRow) = Val(Nvl(mrsList!借款合计))
                    .Rows = .Rows + 1
                    lngRow = lngRow + 1
                End If
                If Val(Nvl(mrsList!借出合计)) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("序号")) = 3
                    .TextMatrix(lngRow, .ColIndex("结算方式")) = "[借出合计]"
                    .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!借出合计)), "#,###0.00;-#,###0.00;0.00;-0.00")
                    .RowData(lngRow) = Val(Nvl(mrsList!借出合计))
                    .Rows = .Rows + 1
                    lngRow = lngRow + 1
                End If
                mrsList.MoveNext
            Loop
        End If
        
        If .Rows > 2 Then .Rows = .Rows - 1: blnTempDelete = True
        .Cell(flexcpSort, 1, .ColIndex("序号"), .Rows - 1, .ColIndex("序号")) = flexSortNumericAscending
        If blnTempDelete = True Then .Rows = .Rows + 1
        intInsureRow = lngRow
        .TextMatrix(intInsureRow, .ColIndex("结算方式")) = "医保相关"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
        .IsSubtotal(intInsureRow) = True
        lngRow = lngRow + 1
        .Rows = .Rows + 1
        
        mrsList.Filter = ""
        mrsList.Filter = "性质=-1 And 金额 <> 0"
        Do While Not mrsList.EOF
            str结算方式 = Nvl(mrsList!结算方式)
            If str结算方式 = "[冲预交款]" Then
                lngNo = 1
            Else
                rsBalanceMode.Filter = "名称='" & str结算方式 & "'"
                If Not rsBalanceMode.EOF Then
                    '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
                    Select Case Val(Nvl(rsBalanceMode!性质))
                    Case 1 '现金结算方式
                        lngNo = 10
                        mlngCashRow = lngRow
                     Case 2  '其他非医保结算
                        lngNo = 11
                    Case 7 '一卡通结算
                        lngNo = 12
                    Case 8  '结算卡结算
                        lngNo = 14
                    Case 3 '个人账户
                        lngNo = 15
                    Case 4 '医保
                        lngNo = 16
                    Case 9
                        lngNo = 17
                        mlngErrorRow = lngRow
                    Case Else
                        lngNo = 18
                    End Select
                Else
                    lngNo = 13
                End If
            End If

            If lngNo = 15 Or lngNo = 16 Then
                .TextMatrix(lngRow, .ColIndex("序号")) = lngNo
                .TextMatrix(lngRow, .ColIndex("结算方式")) = str结算方式
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!金额)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(mrsList!金额))
                .RowOutlineLevel(lngRow) = 1
                If mlngChargeRollingID <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("结算号码")) = Nvl(mrsList!结算号)
                End If
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsList.MoveNext
        Loop
        
        .TextMatrix(intInsureRow, .ColIndex("金额")) = Format(dblInsure, "#,###0.00;-#,###0.00;0.00;-0.00")
        
        If .Rows > 2 Then .Rows = .Rows - 1
        
        If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "医保相关" Then
            .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("金额")) = ""
            .IsSubtotal(.Rows - 1) = False
        End If
        .Outline (0)
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
'        dblToTotal = Round(dblToTotal, 2)
        If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "" And .Rows > 2 Then .Rows = .Rows - 1
        txtTotal.Text = Format(dblToTotal, "0.#########") & "元"
        For i = 0 To .Rows - 1
            If .TextMatrix(i, .ColIndex("结算方式")) = str现金 Then mlngCashRow = i
        Next i
    End With
    
    '恢复列设置
    zl_vsGrid_Para_Restore mlngMode, vsChagre, Me.Name, "结算方式列表", False
    
    '2.加载票据使用相关信息
    '票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    'mstrRollingType:轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
    
    strRollingType = "": strWhere = ""
    If mstrRollingType <> "" Then
        If InStr("," & mstrRollingType & ",", ",0,") > 0 Then
            strWhere = ""
        Else
            '110414:李南春，2017/6/20，医疗卡使用门诊发票
            strWhere = ""
            If InStr("," & mstrRollingType & ",", ",1,") > 0 Then
                strWhere = " (Instr(',1,' , ','|| A.票种 || ',') > 0 And Not Exists (Select 1 From 票据使用明细 D,票据打印内容 E Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质 In (3,4,5))) "
            End If
            If InStr("," & mstrRollingType & ",", ",2,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',2,' , ','|| A.票种 || ',') > 0 And Not Exists (Select 1 From 票据使用明细 D,票据打印内容 E,病人预交记录 F Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=2 And E.NO=F.NO And F.记录性质=1 And Nvl(F.结算性质,0) = 12))"
            End If
            '问题号：118482,焦博,2017/12/19/轧帐类别为预交时，不应该判断 结算性质.且门诊转住院预交 或住院预交转门诊之后, 票据应该算入不算出
            If InStr("," & mstrRollingType & ",", ",21,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',2,' , ','|| A.票种 || ',') > 0 And Not Exists (Select 1 From 票据使用明细 D,票据打印内容 E,病人预交记录 F Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=2 And E.NO=F.NO And F.记录性质=1 And Nvl(F.结算性质,0) = 12) And Exists (Select 1 From 票据使用明细 D,票据打印内容 E,病人预交记录 F Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=2 And E.NO=F.NO And F.记录性质=1 And F.预交类别=1 And F.金额>0 ))"
            End If
            If InStr("," & mstrRollingType & ",", ",22,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',2,' , ','|| A.票种 || ',') > 0 And Not Exists (Select 1 From 票据使用明细 D,票据打印内容 E,病人预交记录 F Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=2 And E.NO=F.NO And F.记录性质=1 And Nvl(F.结算性质,0) = 12) And Exists (Select 1 From 票据使用明细 D,票据打印内容 E,病人预交记录 F Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=2 And E.NO=F.NO And F.记录性质=1 And F.预交类别=2 And F.金额>0 ))"
            End If
            If InStr("," & mstrRollingType & ",", ",3,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',1,2,3,' , ','|| A.票种 || ',') > 0 And Exists (Select 1 From 票据使用明细 D,票据打印内容 E Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质 = 3) And Not Exists (Select 1 From 票据使用明细 D,票据打印内容 E,病人预交记录 F Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=2 And E.NO=F.NO And F.记录性质=1 And Nvl(F.结算性质,0) <> 12))"
            End If
            If InStr("," & mstrRollingType & ",", ",4,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',1,4,' , ','|| A.票种 || ',') > 0 And Exists (Select 1 From 票据使用明细 D,票据打印内容 E Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=4))"
            End If
            If InStr("," & mstrRollingType & ",", ",5,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',1,5,' , ','|| A.票种 || ',') > 0 And Exists (Select 1 From 票据使用明细 D,票据打印内容 E Where D.ID=A.ID And D.打印ID=E.ID And E.数据性质=5))"
            End If
            If InStr("," & mstrRollingType & ",", ",6,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',-,' , ','|| A.票种 || ',') > 0)"
            End If
        End If
        '消费卡数据
        If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",6,") > 0 Then
            strCardUseRecord = _
            "Union All" & vbNewLine & _
            "Select 6 As 票种, a.原因, a.性质, a.卡号 As 号码, Zl_Incstr(a.卡号) As 下一号码, a.使用时间, Null As 打印id, b.批次" & vbNewLine & _
            "From 消费卡使用记录 A, 消费卡领用记录 B" & vbNewLine & _
            "Where a.领用id = b.Id(+) And a.使用人 || '' = [3] And a.使用时间 Between [4] And [5]" & vbNewLine & _
            "      And ((a.性质 = 2" & vbNewLine & _
            "               And Not Exists (Select 1" & vbNewLine & _
            "                   From 人员收缴票据 C, 人员收缴记录 D" & vbNewLine & _
            "                   Where c.收缴id = d.Id And Nvl(c.批次, '-') = Nvl(b.批次, '-') And c.性质 In (2, 3)" & vbNewLine & _
            "                       And Length(c.开始票号) = Length(a.卡号) And a.卡号 Between c.开始票号 And c.终止票号" & vbNewLine & _
            "                       And c.票种 = 6 And d.作废时间 Is Null))" & vbNewLine & _
            "           Or (a.性质 = 1" & vbNewLine & _
            "               And Not Exists (Select 1" & vbNewLine & _
            "                   From 人员收缴票据 C, 人员收缴记录 D" & vbNewLine & _
            "                   Where c.收缴id = d.Id And Nvl(c.批次, '-') = Nvl(b.批次, '-') And c.性质 = 1" & vbNewLine & _
            "                       And Length(c.开始票号) = Length(a.卡号) And a.卡号 Between c.开始票号 And c.终止票号" & vbNewLine & _
            "                       And c.票种 = 6 And d.作废时间 Is Null)))"
        End If
        
        If strWhere <> "" Then strWhere = "And (" & strWhere & ")"
        
        strWithTable = "" & _
        "    With 票据使用 As  ( " & _
        "           Select A.票种, A.原因,A.性质, A.号码, Zl_Incstr(A.号码) As 下一号码,A.使用时间,A.打印ID,B.批次 " & _
        "           From 票据使用明细 A,票据领用记录 B " & _
        "           Where A.使用人|| '' = [3] And A.领用id = B.id And A.使用时间 Between [4] and [5] " & _
        "                   And ((A.性质 = 2 And Not Exists(Select 1 " & _
        "                           From 人员收缴票据 C,人员收缴记录 D,票据使用明细 E,票据领用记录 F " & _
        "                           Where c.收缴id = d.Id And d.收款员 = [3] And e.打印id = a.打印id And e.领用id = f.Id " & _
        "                               And Nvl(f.批次,'-') = Nvl(c.批次,'-') And c.性质 In (2, 3) " & _
        "                               And d.作废时间 Is Null And a.票种 = c.票种 And Length(c.开始票号) = Length(a.号码) " & _
        "                               And a.号码 Between c.开始票号 And c.终止票号)) " & _
        "                       Or (a.性质 = 1 And Not Exists(Select 1 " & _
        "                           From 人员收缴票据 E,人员收缴记录 F,票据使用明细 G,票据领用记录 H " & _
        "                           Where e.收缴ID=f.ID And f.收款员 = [3] And g.打印id = a.打印id And g.领用id = h.Id " & _
        "                               And Nvl(h.批次,'-') = Nvl(e.批次,'-') And e.性质 = 1 And f.作废时间 Is Null And a.票种=e.票种 " & _
        "                               And Length(e.开始票号) = Length(a.号码) And a.号码 between e.开始票号 and e.终止票号)))" & _
                    strWhere & vbNewLine & _
                    strCardUseRecord & "), "
        
        strWithTable = strWithTable & _
        "           收回票据 as (  " & _
        "               Select Distinct 1 as 性质,y.票种,x.No, y.号码 " & _
        "               From 票据打印内容 X, 票据使用 Y " & _
        "               Where y.原因 = 2 AND y.性质=2 And x.Id = y.打印id AND Y.票种<>1 " & _
        "               Union all " & _
        "               Select Distinct 2 as 性质,y.票种,x.No, y.号码 " & _
        "               From 票据打印内容 X, 票据使用 Y " & _
        "               Where y.原因 = 4 AND y.性质=2 And x.Id = y.打印id AND Y.票种<>1 ), " & _
        "           收回票据_收费 as ( " & _
        "               Select 1 as 性质,票种,号码 From 票据使用  where 原因=2 And 性质=2 AND 票种=1 " & _
        "               Union all " & _
        "               Select 2 as 性质,票种,号码 From 票据使用  where 原因=4 And 性质=2 AND 票种=1 ), " & _
        "           收回票据金额 as ( " & _
        "               Select  a.性质, a.票种,A.号码,Sum(C.结帐金额) As 单据金额  " & _
        "               From 收回票据_收费 A,票据打印明细 B, 门诊费用记录 C " & _
        "               Where a.票种=b.票种 And a.号码=b.票号 and a.票种=1     " & _
        "                           And b.No =c.No And C.记录性质 = 1 And C.记录状态 In (3, 1) " & _
        "                           And Instr(',' || b.序号 || ',', ',' || Nvl(c.价格父号, c.序号) || ',') > 0 " & _
        "                Group by a.性质,a.票种,A.号码 "
        strWithTable = strWithTable & _
        "               Union all " & _
        "               Select a.性质,A.票种,A.号码,sum(C.冲预交) as 单据金额  " & _
        "               From 收回票据 A,病人结帐记录 B,病人预交记录 C " & _
        "               Where a.票种=3 And A.NO=B.NO And B.结算状态 Is Null and b.记录状态 in (1,3) and b.ID=C.结帐ID  " & _
        "               Group by A.性质,a.票种,a.号码 " & _
        "               Union all " & _
        "               Select a.性质,A.票种,A.号码,sum(B.金额) " & _
        "               From 收回票据 A,病人预交记录  B " & _
        "               Where a.票种=2 And a.no=b.No  and b.记录性质=1 and B.记录状态 in (1,3) " & _
        "               Group by A.性质,a.票种,a.号码 " & _
        "               Union all " & _
        "               Select a.性质,A.票种,a.号码,sum(b.结帐金额) as 单据金额  " & _
        "               From 收回票据 A,门诊费用记录 B " & _
        "               Where a.票种=4 And A.NO=B.NO and b.记录性质=4 and B.记录状态 in (1,3)  " & _
        "               Group by A.性质,a.票种,a.号码 " & _
        "               Union all " & _
        "               Select a.性质,A.票种,a.号码,sum(b.结帐金额) as 单据金额  " & _
        "               From 收回票据 A,住院费用记录 B " & _
        "               Where a.票种=5 And A.NO=B.NO and b.记录性质=5 and B.记录状态 in (1,3) and Nvl(B.记帐费用, 0) = 0 " & _
        "               Group by A.性质,a.票种,a.号码 "
        strWithTable = strWithTable & _
        "               Union all " & _
        "               Select a.性质, a.票种, a.号码, Sum(b.结帐金额) As 单据金额" & vbNewLine & _
        "               From 收回票据_收费 A, 门诊费用记录 B, 票据使用明细 C, 票据打印内容 D" & vbNewLine & _
        "               Where a.票种 = 1 And a.号码 = c.号码 And c.性质 = 2 And c.打印id = d.Id And d.数据性质 = 4 And d.No = b.No And b.记录性质 = 4 And b.记录状态 In (1, 3)" & vbNewLine & _
        "               Group By a.性质, a.票种, a.号码 " & _
        "               Union all " & _
        "               Select a.性质, a.票种, a.号码, Sum(b.结帐金额) As 单据金额" & vbNewLine & _
        "               From 收回票据_收费 A, 住院费用记录 B, 票据使用明细 C, 票据打印内容 D" & vbNewLine & _
        "               Where a.票种 = 1 And a.号码 = c.号码 And c.性质 = 2 And c.打印id = d.Id And d.数据性质 = 5 And d.No = b.No And b.记录性质 = 5 And b.记录状态 In (1, 3)" & vbNewLine & _
        "               Group By a.性质, a.票种, a.号码 ) "
    
        strSQL = "" & _
        "   Select /*+ Rule*/   票种,性质,张数,开始号码 as 开始票号,终止号码 as 终止票号,金额,使用时间  as 发生时间,批次 " & _
        "   FROM (   " & strWithTable & _
        "               Select 1 As 性质, a.票种,a.号码, a.号码 As 开始号码, b.号码 As 终止号码,count(*) as 张数,null as 使用时间,0 as 金额, a.批次 " & _
        "               From (Select Rownum As 行号, 票种, 号码, 批次 " & _
        "                           From (Select 票种, 号码, 批次 From 票据使用 where 原因 In (1,3,6)  Minus Select 票种, 下一号码, 批次 From 票据使用 where 原因 In (1,3,6))) A, " & _
        "                         (Select Rownum As 断号, 票种, Zl_Incstr_Pre(号码) As 号码,批次 " & _
        "                           From (Select 票种, 下一号码 As 号码,批次 From 票据使用  票据使用 where 原因 In (1,3,6)  Minus Select 票种, 号码,批次 From 票据使用  票据使用 where 原因 In (1,3,6))) B,"
        strSQL = strSQL & "" & _
        "                          ( Select distinct 票种,号码 from 票据使用) M " & _
        "               Where a.行号 = b.断号 And a.票种 = b.票种 and a.票种=M.票种 And M.号码 between a.号码 and b.号码 And Nvl(a.批次,0) = Nvl(b.批次,0)   " & _
        "               Group by a.票种,a.号码,b.号码,a.批次 " & _
        "               Union all " & _
        "               Select 2 As 性质, a.票种,a.号码, a.号码 As 开始号码, b.号码 As 终止号码,count(*) as 张数,m.使用时间 as 使用时间,sum(q.单据金额) as 单据金额, a.批次 " & _
        "               From (Select Rownum As 行号, 票种, 号码, 批次 " & _
        "               From (Select 票种,使用时间, 号码,批次 From 票据使用 where 原因=2 And 性质=2 Minus Select 票种,使用时间, 下一号码,批次 From 票据使用 where 原因=2 and 性质=2)) A, " & _
        "                          (Select Rownum As 断号, 票种, Zl_Incstr_Pre(号码) As 号码, 批次 " & _
        "                           From (Select 票种,使用时间, 下一号码 As 号码,批次 From 票据使用    where 原因=2 And 性质=2 Minus Select 票种,使用时间, 号码,批次 From 票据使用  where 原因=2 And 性质=2)) B, " & _
        "                           (select  票种,号码,Max(使用时间) as 使用时间 From 票据使用 Where 原因=2 and 性质=2 Group by 票种,号码)  M,收回票据金额 Q " & _
        "               Where a.行号 = b.断号 And a.票种 = b.票种 and a.票种=M.票种 And M.号码 between a.号码 and b.号码 And Nvl(a.批次,0) = Nvl(b.批次,0)  " & _
        "                           and m.票种=Q.票种(+) and m.号码=Q.号码(+) AND q.性质(+)=1  " & _
        "               group by a.票种,a.号码,b.号码,m.使用时间,a.批次 " & _
        "               union all  " & _
        "               Select 3 As 性质, a.票种,a.号码, a.号码 As 开始号码, b.号码 As 终止号码,count(*) as 张数,m.使用时间 as 使用时间,sum(q.单据金额) as 单据金额, a.批次 " & _
        "               From (  Select Rownum As 行号, 票种, 号码,批次 " & _
        "                           From (Select 票种,使用时间, 号码,批次 From 票据使用 where 原因=4 And 性质=2 Minus Select 票种,使用时间, 下一号码,批次 From 票据使用 where 原因=4 and 性质=2)) A, " & _
        "                           (  Select Rownum As 断号, 票种, Zl_Incstr_Pre(号码) As 号码, 批次 " & _
        "                               From (Select 票种,使用时间, 下一号码 As 号码,批次 From 票据使用    where 原因=4 And 性质=2 Minus Select 票种,使用时间, 号码,批次 From 票据使用    where 原因=4 And 性质=2)) B, " & _
        "                           (select  票种,号码,Max(使用时间) as 使用时间 From 票据使用 Where 原因=4 And 性质=2 Group by 票种,号码)  M,收回票据金额 Q " & _
        "               Where a.行号 = b.断号 And a.票种 = b.票种 and a.票种=M.票种 And M.号码 between a.号码 and b.号码 And Nvl(a.批次,0) = Nvl(b.批次,0)  " & _
        "                       and m.票种=Q.票种(+) and m.号码=Q.号码(+) AND q.性质(+)=2 " & _
        "               group by a.票种,a.号码,b.号码,m.使用时间,a.批次 ) " & _
        " ORDER BY 票种,性质,使用时间 ,开始号码"
    Else
        strSQL = "" & _
        "   Select 1 as 票种, 1 as 性质,0 as 张数,''  as 开始票号, '' as 终止票号,0 as 金额, sysdate as 发生时间, Null as 批次" & _
        "   From dual " & _
        "   Where 1=2"
    End If
    
    Set mrsListBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtendDate)
    
    With vsReturnBill
        .Clear 1
        .Rows = 2: lngRow = 1
        mrsListBill.Filter = "性质=2 Or 性质=3"
        Do While Not mrsListBill.EOF
            If .TextMatrix(lngRow - 1, .ColIndex("类别")) <> GetBillTypeName(mrsListBill!票种) Then
                .TextMatrix(lngRow, .ColIndex("类别")) = GetBillTypeName(mrsListBill!票种)
                .IsSubtotal(lngRow) = True
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
            .RowOutlineLevel(lngRow) = 1
            .TextMatrix(lngRow, .ColIndex("类别")) = GetBillTypeName(mrsListBill!票种)
            .TextMatrix(lngRow, .ColIndex("类型")) = Decode(mrsListBill!性质, _
                                                                2, Decode(mrsListBill!票种, 6, "收回", "退费"), _
                                                                3, Decode(mrsListBill!票种, 6, "换卡", "重打"), "其他")
            .TextMatrix(lngRow, .ColIndex("收回时间")) = Format(mrsListBill!发生时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsListBill!金额)), "#,###0.00;-#,###0.00;0.00;-0.00")
            If InStr(";收费;消费卡;", ";" & .TextMatrix(lngRow, .ColIndex("类别")) & ";") > 0 And Val(.TextMatrix(lngRow, .ColIndex("金额"))) = 0 Then
                .TextMatrix(lngRow, .ColIndex("金额")) = "-"
            End If
            If Nvl(mrsListBill!开始票号) = Nvl(mrsListBill!终止票号) Then
                .TextMatrix(lngRow, .ColIndex("票据号")) = Nvl(mrsListBill!开始票号)
            Else
                .TextMatrix(lngRow, .ColIndex("票据号")) = Nvl(mrsListBill!开始票号) & "-" & Nvl(mrsListBill!终止票号)
            End If
            .Rows = .Rows + 1: lngRow = lngRow + 1
            mrsListBill.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        If .TextMatrix(.Rows - 1, .ColIndex("类别")) = "" Then
            .IsSubtotal(.Rows - 1) = False
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        '恢复列设置
    zl_vsGrid_Para_Restore mlngMode, vsReturnBill, Me.Name, "收回票据列表", False
    End With
  
    With vsBill
        .Clear 1
        .Rows = 1: lngRow = 1: .Cols = 3
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        mrsListBill.Filter = 0:  strTemp = ""
        Do While Not mrsListBill.EOF
            int票种 = Val(Nvl(mrsListBill!票种))
            If InStr(1, strTemp & ",", "," & int票种 & ",") = 0 Then
                strTemp = strTemp & "," & int票种
                .Rows = .Rows + 2
                .TextMatrix(.Rows - 3, 0) = GetBillTypeName(mrsListBill!票种)
                .TextMatrix(.Rows - 2, 0) = GetBillTypeName(mrsListBill!票种)
                .TextMatrix(.Rows - 3, 1) = "收退票据"
                .TextMatrix(.Rows - 2, 1) = "票据范围"
                .Cell(flexcpData, .Rows - 3, 0, .Rows - 2, 0) = int票种
            End If
            mrsListBill.MoveNext
        Loop
        Dim lngBillTotal(0 To 2) As Long
        lngRow = 0
        For lngRow = 0 To .Rows - 1 Step 2
             int票种 = Val(.Cell(flexcpData, lngRow, 0))
             mrsListBill.Filter = "票种=" & int票种
             lngBillTotal(0) = 0: lngBillTotal(1) = 0: lngBillTotal(2) = 0
             Do While Not mrsListBill.EOF
                Select Case Val(Nvl(mrsListBill!性质))
                Case 1 '正常票据统计
                        lngBillTotal(0) = lngBillTotal(0) + Val(Nvl(mrsListBill!张数))
                        .TextMatrix(lngRow + 1, 2) = Trim(.TextMatrix(lngRow + 1, 2)) & IIf(Trim(.TextMatrix(lngRow + 1, 2)) = "", "", ";") & _
                            IIf(Trim(.TextMatrix(lngRow, 2)) = "", "", ";") & _
                            IIf(Nvl(mrsListBill!开始票号) = Nvl(mrsListBill!终止票号), _
                                Nvl(mrsListBill!开始票号), _
                                Nvl(mrsListBill!开始票号) & "-" & Nvl(mrsListBill!终止票号))
                Case 2
                        lngBillTotal(1) = lngBillTotal(1) + Val(Nvl(mrsListBill!张数))
                Case 3
                        lngBillTotal(2) = lngBillTotal(2) + Val(Nvl(mrsListBill!张数))
                End Select
                mrsListBill.MoveNext
             Loop
             If int票种 <> 0 Then
                .TextMatrix(lngRow, 2) = "使用:" & lngBillTotal(0) & "张; " & _
                    Decode(int票种, 6, "收回(退卡、回收和换卡):", "收回(退费和重打):") & lngBillTotal(1) + lngBillTotal(2) & "张"
            End If
        Next
        .Rows = .Rows - 1
    End With
    LoadPersonChargeAndBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Function GetBillTypeName(ByVal byt票种 As Byte) As String
    '根据票种获取类别名称
    On Error GoTo errHandle
    GetBillTypeName = Decode(byt票种, _
        1, "收费", 2, "预交", 3, "结帐", 4, "挂号", _
        5, "就诊卡", 6, "消费卡", "其他")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function LoadChargeAndBillAndTotal() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示收款及票据汇总信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-04 11:28:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bytType As Byte, rsBalanceMode As ADODB.Recordset
    Dim strWithTable As String, lngRow As Long, lngNo As Long
    Dim str结算方式 As String, strTemp As String, int票种 As Integer
    Dim strWhere As String, dblToTotal As Double, blnCYJ As Boolean
    Dim strTable As String, i As Integer, intInsureRow As Integer
    Dim objRecord As ReportRecord, dblInsure As Double
    Dim objItem As ReportRecordItem, rsList As ADODB.Recordset
    
    On Error GoTo errHandle
    '1.加载收款数据汇总
    strSQL = "" & _
    "   Select decode(nvl(M.性质,0),1,1,2,2,3,10,4,11,9,9,4) as 序号,  " & _
    "           b.结算方式,b.金额,b.结算号,b.余额," & _
    "           a.冲预交款 as 冲预交,A.借入合计 as 借款合计,A.借出合计 " & _
    "   From 人员收缴记录 A, 人员收缴明细 B,结算方式 M" & _
    "   Where a.Id = b.收缴id And a.ID=[1] and B.结算方式=M.名称(+) and nvl(金额,0)<>0 " & _
    "   Order by 序号,结算方式"
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtStartDate)
    
    strSQL = "Select a.冲预交款 As 冲预交, a.借入合计 As 借款合计, a.借出合计 From 人员收缴记录 A Where a.Id = [1]"
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
    With vsChagre
        .Clear 1
        .Rows = 2: lngRow = 1: blnCYJ = False
        If rsList.RecordCount <> 0 Then
            If Val(Nvl(rsList!冲预交)) <> 0 Then
                .TextMatrix(lngRow, .ColIndex("序号")) = 1
                .TextMatrix(lngRow, .ColIndex("结算方式")) = "[冲预交款]"
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsList!冲预交)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(rsList!冲预交))
                dblToTotal = dblToTotal + Val(Nvl(rsList!冲预交))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            If Val(Nvl(rsList!借款合计)) <> 0 Then
                .TextMatrix(lngRow, .ColIndex("序号")) = 2
                .TextMatrix(lngRow, .ColIndex("结算方式")) = "[借入合计]"
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsList!借款合计)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(rsList!借款合计))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            If Val(Nvl(rsList!借出合计)) <> 0 Then
                .TextMatrix(lngRow, .ColIndex("序号")) = 3
                .TextMatrix(lngRow, .ColIndex("结算方式")) = "[借出合计]"
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsList!借出合计)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(rsList!借出合计))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
        End If
        mlngCashRow = 0
        mrsList.Filter = "序号<>10 And 序号<>11"
        Do While Not mrsList.EOF
            str结算方式 = Nvl(mrsList!结算方式)
            dblToTotal = dblToTotal + Val(Nvl(mrsList!金额))
            .TextMatrix(lngRow, .ColIndex("序号")) = Val(Nvl(mrsList!序号)) + 10
            If Val(Nvl(mrsList!序号)) = 1 Then mlngCashRow = lngRow
            .TextMatrix(lngRow, .ColIndex("结算方式")) = str结算方式
            If Val(Nvl(mrsList!序号)) = 9 Then
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!金额)), "0.#########")
            Else
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!金额)), "#,###0.00;-#,###0.00;0.00;-0.00")
            End If
            .RowData(lngRow) = Val(Nvl(mrsList!金额))
            .TextMatrix(lngRow, .ColIndex("结算号码")) = Nvl(mrsList!结算号)
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            mrsList.MoveNext
        Loop
        mrsList.Filter = "序号=10 Or 序号=11"
        If mrsList.RecordCount <> 0 Then
            dblInsure = 0
            .TextMatrix(lngRow, .ColIndex("结算方式")) = "医保相关"
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
            intInsureRow = lngRow
            .IsSubtotal(intInsureRow) = True
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            Do While Not mrsList.EOF
                str结算方式 = Nvl(mrsList!结算方式)
                dblInsure = dblInsure + Val(Nvl(mrsList!金额))
                dblToTotal = dblToTotal + Val(Nvl(mrsList!金额))
                .TextMatrix(lngRow, .ColIndex("序号")) = Val(Nvl(mrsList!序号)) + 10
                .RowOutlineLevel(lngRow) = 1
                .TextMatrix(lngRow, .ColIndex("结算方式")) = str结算方式
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsList!金额)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(mrsList!金额))
                .TextMatrix(lngRow, .ColIndex("结算号码")) = Nvl(mrsList!结算号)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
                mrsList.MoveNext
            Loop
            .TextMatrix(intInsureRow, .ColIndex("金额")) = Format(dblInsure, "#,###0.00;-#,###0.00;0.00;-0.00")
            .Outline (0)
        End If
        If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "" And .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        txtTotal.Text = Format(dblToTotal, "0.#########") & "元"
    End With
    '恢复列设置
    zl_vsGrid_Para_Restore mlngMode, vsChagre, Me.Name, "结算方式列表", False
    
    '2.加载票据使用相关信息
        '加载历史数据
        '票种,性质,张数,开始号码 as 开始票号,终止号码 as 终止票号,金额,使用时间  as 发生时间
    strTable = ""
    If mblnDel Or mbytType = EM_收费员轧帐 Then
        If mbytType = EM_收费员轧帐 Then
            strWhere = " And  A.ID =[1] And a.记录性质=1"
        Else
            strWhere = " And  A.ID=C.记录ID And C.性质=8 And C.收缴ID=[1] And a.记录性质=1"
            strTable = ",人员收缴对照 C"
        End If
    Else
        If mbytType = EM_小组收款 Then
            strWhere = " And  A.小组收款ID =[1] And a.记录性质=1"
        ElseIf mbytType = EM_小组轧帐 Then
            strWhere = " And  A.小组轧账ID =[1] And a.记录性质=1"
        Else
            strWhere = " And  A.财务收款ID =[1] And a.记录性质=1"
        End If
    End If
    
    strSQL = "" & _
    "   Select b.票种,b.性质,b.票据张数 as 张数,b.开始票号,b.终止票号,b.金额,b.发生时间 " & _
    "   From 人员收缴记录 a,人员收缴票据 b " & strTable & _
    "   Where  a.id=b.收缴id  " & strWhere & _
    "   Order by 票种,性质,序号"
    Set mrsListBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtStartDate)
    
    With vsReturnBill
        .Clear 1
        .Rows = 2: lngRow = 1
        mrsListBill.Filter = "性质=2 Or 性质=3"
        Do While Not mrsListBill.EOF
            If .TextMatrix(lngRow - 1, .ColIndex("类别")) <> GetBillTypeName(mrsListBill!票种) Then
                .TextMatrix(lngRow, .ColIndex("类别")) = GetBillTypeName(mrsListBill!票种)
                .IsSubtotal(lngRow) = True
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
            .RowOutlineLevel(lngRow) = 1
            .TextMatrix(lngRow, .ColIndex("类别")) = GetBillTypeName(mrsListBill!票种)
            .TextMatrix(lngRow, .ColIndex("类型")) = Decode(mrsListBill!性质, _
                                                                2, Decode(mrsListBill!票种, 6, "收回", "退费"), _
                                                                3, Decode(mrsListBill!票种, 6, "换卡", "重打"), "其他")
            .TextMatrix(lngRow, .ColIndex("收回时间")) = Format(mrsListBill!发生时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(mrsListBill!金额)), "#,###0.00;-#,###0.00;0.00;-0.00")
            If InStr(";收费;消费卡;", ";" & .TextMatrix(lngRow, .ColIndex("类别")) & ";") > 0 And Val(.TextMatrix(lngRow, .ColIndex("金额"))) = 0 Then
                .TextMatrix(lngRow, .ColIndex("金额")) = "-"
            End If
            If Nvl(mrsListBill!开始票号) = Nvl(mrsListBill!终止票号) Then
                .TextMatrix(lngRow, .ColIndex("票据号")) = Nvl(mrsListBill!开始票号)
            Else
                .TextMatrix(lngRow, .ColIndex("票据号")) = Nvl(mrsListBill!开始票号) & "-" & Nvl(mrsListBill!终止票号)
            End If
            .Rows = .Rows + 1: lngRow = lngRow + 1
            mrsListBill.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        If .TextMatrix(.Rows - 1, .ColIndex("类别")) = "" Then
            .IsSubtotal(.Rows - 1) = False
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        '恢复列设置
        zl_vsGrid_Para_Restore mlngMode, vsReturnBill, Me.Name, "收回票据列表", False
    End With
    With vsBill
        .Clear 1
        .Rows = 1: lngRow = 0: .Cols = 3
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        mrsListBill.Filter = 0:  strTemp = ""
        Do While Not mrsListBill.EOF
            int票种 = Val(Nvl(mrsListBill!票种))
            If InStr(1, strTemp & ",", "," & int票种 & ",") = 0 Then
                strTemp = strTemp & "," & int票种
                .Rows = .Rows + 2
                .TextMatrix(.Rows - 3, 0) = GetBillTypeName(mrsListBill!票种)
                .TextMatrix(.Rows - 2, 0) = GetBillTypeName(mrsListBill!票种)
                .TextMatrix(.Rows - 3, 1) = "收退票据"
                .TextMatrix(.Rows - 2, 1) = "票据范围"
                .Cell(flexcpData, .Rows - 3, 0, .Rows - 2, 0) = int票种
            End If
            mrsListBill.MoveNext
        Loop
        Dim lngBillTotal(0 To 2) As Long
        lngRow = 0
        For lngRow = 0 To .Rows - 1 Step 2
            int票种 = Val(.Cell(flexcpData, lngRow, 0))
            mrsListBill.Filter = "票种=" & int票种
            lngBillTotal(0) = 0: lngBillTotal(1) = 0: lngBillTotal(2) = 0
            Do While Not mrsListBill.EOF
               Select Case Val(Nvl(mrsListBill!性质))
               Case 1 '正常票据统计
                    lngBillTotal(0) = lngBillTotal(0) + Val(Nvl(mrsListBill!张数))
                    .TextMatrix(lngRow + 1, 2) = Trim(.TextMatrix(lngRow + 1, 2)) & _
                        IIf(Trim(.TextMatrix(lngRow + 1, 2)) = "", "", ";") & _
                        IIf(Nvl(mrsListBill!开始票号) = Nvl(mrsListBill!终止票号), _
                            Nvl(mrsListBill!开始票号), _
                            Nvl(mrsListBill!开始票号) & "-" & Nvl(mrsListBill!终止票号))
               Case 2
                    lngBillTotal(1) = lngBillTotal(1) + Val(Nvl(mrsListBill!张数))
               Case 3
                    lngBillTotal(2) = lngBillTotal(2) + Val(Nvl(mrsListBill!张数))
               End Select
               mrsListBill.MoveNext
            Loop
            If int票种 <> 0 Then
                .TextMatrix(lngRow, 2) = "使用:" & lngBillTotal(0) & "张; " & _
                    Decode(int票种, 6, "收回(退卡、回收和换卡):", "收回(退费和重打):") & lngBillTotal(1) + lngBillTotal(2) & "张"
            End If
        Next
        .Rows = .Rows - 1
    End With
    
    LoadChargeAndBillAndTotal = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
 
Public Sub SetFontSize(ByVal bytSize As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省9号字)，1-大号(12号);>1: 为指定的字号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-03 18:05:20
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置字体大小
    '编制:刘兴洪
    '日期:2013-09-03 18:04:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Me.FontSize = mbytFontSize
    Set dkpMan.PaintManager.CaptionFont = Me.Font
    Set vsBill.Font = Me.Font
    Set vsChagre.Font = Me.Font
    Set vsReturnBill.Font = Me.Font
    txtTotal.Font = Me.Font
 End Sub
 
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化窗体
    '编制:刘兴洪
    '日期:2013-09-03 15:28:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtTotal.Text = "": txtTotal.Locked = True: txtTotal.Enabled = False
    txtTotal.FontBold = True: mbytFontSize = 9
    Call InitPanel
    Call InitGrid
    Call ReSetFontSize
 End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim objReturnPane As Pane
    Dim objChargePane As Pane
    Dim lngFilterHeight As Long, lngBillHeight As Long
    Dim lngBalanceHeight As Long
    
    lngFilterHeight = 810 / Screen.TwipsPerPixelY
    lngBillHeight = 1275 / Screen.TwipsPerPixelY
    lngBalanceHeight = (Me.ScaleHeight - 1275) \ Screen.TwipsPerPixelY - 205
    With dkpMan
        Set objChargePane = .CreatePane(mPaneIndex.EM_PN_ChargeTotal, 400, lngBalanceHeight, DockBottomOf, Nothing)
        objChargePane.Title = "收款信息"
        objChargePane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        objChargePane.Handle = picChargeInfor.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_ChargeTotal, 400, lngBillHeight, DockBottomOf, objChargePane)
        objPane.MinTrackSize.Height = lngBillHeight * 0.5
        objPane.Title = "票据使用信息"
        objPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        objPane.Handle = picBillInfor.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_ReprintBill, 400, lngBalanceHeight, DockRightOf, objChargePane)
        objPane.Title = "收回票据信息"
        objPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        objPane.Handle = picReturnBill.hWnd

'        Set objReturnPane = .CreatePane(mPaneIndex.EM_PN_BackFeeBill, 400, 400, DockRightOf, objChargePane)
'        objReturnPane.Title = "退费收回票据"
'        objReturnPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
'        objReturnPane.Handle = picDelFeeBill.hwnd
'
'        Set objPane = .CreatePane(mPaneIndex.EM_PN_ReprintBill, 400, 165, DockBottomOf, objReturnPane)
'        objPane.Title = "重打收回票据"
'        objPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
'        objPane.Handle = picRePrintBill.hwnd
        
      '  .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件信息
    '返回:
    '编制:刘兴洪
    '日期:2013-09-03 11:38:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objCol As ReportColumn
    
    '收回票据信息
    With vsReturnBill
        Set .Font = Me.Font
        .Clear 1
        .Cols = 5: .Rows = 2
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        .FixedRows = 1
        .TextMatrix(0, 0) = "类别"
        .TextMatrix(0, 1) = "类型"
        .TextMatrix(0, 2) = "收回时间"
        .TextMatrix(0, 3) = "金额"
        .TextMatrix(0, 4) = "票据号"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            If i = .ColIndex("金额") Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        .ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsReturnBill, Me.Name, "收回票据列表", False
    End With
    
    '收款汇总信息
    With vsChagre
        Set .Font = Me.Font
        .Clear 1
        .Cols = 4: .Rows = 2
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 1
        .SubtotalPosition = flexSTAbove
        .FixedRows = 1
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "结算方式"
        .TextMatrix(0, 2) = "金额"
        .TextMatrix(0, 3) = "结算号码"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            If i = .ColIndex("金额") Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .ColHidden(.ColIndex("序号")) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        .ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsChagre, Me.Name, "结算方式列表", False
    End With
    '标据使用信息
    With vsBill
        .Clear 1
        Set .Font = Me.Font
        .Cols = 3: .Rows = 1
        .FixedRows = 0: .FixedCols = 1
        .ColAlignment(2) = flexAlignLeftCenter
        .MergeCells = flexMergeFree
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        .ExtendLastCol = True
        For i = 0 To .Rows - 1
            .MergeRow(0) = True
        Next
    End With
 End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngCashRow = 0
End Sub

Private Sub picBillInfor_Resize()
    Err = 0: On Error Resume Next
    With picBillInfor
        vsBill.Left = .ScaleLeft
        vsBill.Top = .ScaleTop
        vsBill.Height = .ScaleHeight
        vsBill.Width = .ScaleWidth
    End With
End Sub
Private Sub picChargeInfor_Resize()
    Err = 0: On Error Resume Next
    With picChargeInfor
        vsChagre.Top = .ScaleTop
        vsChagre.Left = .ScaleLeft
        lblTotal.Left = .ScaleLeft
        txtTotal.Left = lblTotal.Left + lblTotal.Width + 10
        txtTotal.Width = .ScaleWidth - txtTotal.Left
        txtTotal.Top = .ScaleHeight - txtTotal.Height
        lblTotal.Top = txtTotal.Top + (txtTotal.Height - lblTotal.Height) \ 2
        vsChagre.Height = txtTotal.Top - vsChagre.Top - 50
        vsChagre.Width = .ScaleWidth
    End With
End Sub

Private Sub picReturnBill_Resize()
    Err = 0: On Error Resume Next
    With picReturnBill
        vsReturnBill.Left = .ScaleLeft
        vsReturnBill.Top = .ScaleTop
        vsReturnBill.Height = .ScaleHeight
        vsReturnBill.Width = .ScaleWidth
    End With
End Sub

 

Private Sub vsChagre_GotFocus()
    Call zl_VsGridGotFocus(vsChagre)
End Sub

Private Sub vsChagre_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsChagre, vsChagre.BackColor)
    'On Error Resume Next
    With vsChagre
        If .TextMatrix(.RowSel, .ColIndex("结算方式")) = "医保相关" Then .Cell(flexcpBackColor, .RowSel, 0, .RowSel, .Cols - 1) = &H80000016
    End With
End Sub
Private Sub vsChagre_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsChagre, Me.Name, "结算方式列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsChagre_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsChagre, OldRow, NewRow, OldCol, NewCol)
    With vsChagre
        If OldRow >= .Rows - 1 Then Exit Sub
        If .TextMatrix(OldRow, .ColIndex("结算方式")) = "医保相关" Then .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000016
    End With
End Sub
Private Sub vsChagre_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsChagre, Me.Name, "结算方式列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsChagre_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnOlnyView Then Cancel = True: Exit Sub
    With vsChagre
        Select Case Col
        Case .ColIndex("结算号码")
            If .TextMatrix(Row, .ColIndex("结算方式")) Like "*冲预交*" _
                Or .TextMatrix(Row, Col) Like "*借入合计*" _
                Or .TextMatrix(Row, Col) Like "*借出合计*" _
                Or .TextMatrix(Row, .ColIndex("结算方式")) = "医保相关" Then
                Cancel = True: Exit Sub
            End If
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub vsChagre_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsChagre, vsChagre.ColIndex("结算方式"), vsChagre.Cols - 1, False)
End Sub

Private Sub vsChagre_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsChagre, vsChagre.ColIndex("结算方式"), vsChagre.Cols - 1, False)
End Sub

Private Sub vsChagre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub vsChagre_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsChagre
        If Row <= 1 Then Exit Sub
            VsFlxGridCheckKeyPress vsChagre, Row, Col, KeyAscii, m文本式
            If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
    End With
End Sub
Private Sub vsChagre_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '数据验证
    With vsChagre
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("结算号码")
            If zlCommFun.ActualLen(strKey) > 10 Then
                MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(1, strKey, "'") > 0 Or InStr(1, strKey, "|") > 0 Or InStr(1, strKey, ",") > 0 Then
                MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsReturnBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsReturnBill, Me.Name, "收回票据列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsReturnBill_GotFocus()
    Call zl_VsGridGotFocus(vsReturnBill)
End Sub
Private Sub vsReturnBill_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsReturnBill, vsReturnBill.BackColor)
    'On Error Resume Next
    With vsReturnBill
        If .IsSubtotal(.RowSel) = True Then .Cell(flexcpBackColor, .RowSel, 0, .RowSel, .Cols - 1) = &H80000016
    End With
End Sub

Private Sub vsReturnBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsReturnBill, OldRow, NewRow, OldCol, NewCol)
    With vsReturnBill
        If OldRow >= .Rows - 1 Then Exit Sub
        If .IsSubtotal(OldRow) = True Then .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000016
    End With
End Sub

Private Function GetBalance(ByVal strCash As String) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取截止时间段的余额
    '入参:strCash-现金结算方式
    '返回:返回余额,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 10:40:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select 结算方式, Sum(金额) As 金额 " & _
    "   From (  Select 结算方式, 余额 As 金额  From 人员缴款余额  " & _
    "                Where 收款员 =[1] and 性质=1" & _
    "                Union All " & _
    "                Select a.结算方式, -1 * Sum(a.冲预交) As 金额" & vbNewLine & _
    "                From 病人预交记录 A" & vbNewLine & _
    "                Where Nvl(a.校对标志, 0) = 0 And 操作员姓名 || '' = [1] And Mod(a.记录性质, 10) <> 1 And" & vbNewLine & _
    "                   a.收款时间 > [2] And Not Exists" & vbNewLine & _
    "                   (Select 1 From 门诊费用记录 B Where a.结帐id = b.结帐id And Nvl(b.费用状态, 0) = 1) And Not Exists" & vbNewLine & _
    "                   (Select 1 From 病人结帐记录 B Where a.结帐id = b.Id And Nvl(b.结算状态, 0) <> 0) And Not Exists" & vbNewLine & _
    "                   (Select 1 From 费用补充记录 B Where a.结帐id = b.结算id And Nvl(b.费用状态, 0) >= 1)" & vbNewLine & _
    "                Group By a.结算方式" & _
    "                Union All " & _
    "                Select 结算方式, -1 * nvl(sum(金额),0) As 金额 " & _
    "                From 病人预交记录 A " & _
    "               Where 记录性质 = 1 And 操作员姓名 || '' =[1] And 收款时间 > [2]  " & _
    "               Group by 结算方式 "
    
    strSQL = strSQL & _
    "               Union All " & _
    "               Select a.结算方式, -1 * Nvl(Sum(a.实收金额), 0) As 金额 " & _
    "               From 病人卡结算记录 A " & _
    "               Where a.记录性质 In(1, 2, 3) And a.操作员姓名 || '' =[1] And a.登记时间 > [2]  " & _
    "               Group By 结算方式 " & _
    "               Union All " & _
    "               Select a.结算方式, -1 * Nvl(Sum(a.借款金额), 0) As 金额 " & _
    "               From 人员借款记录 A " & _
    "               Where a.借款人 || '' =[1] And a.取消时间 Is Null And a.借出时间 > [2]  " & _
    "               Group By 结算方式 " & _
    "               Union All " & _
    "               Select a.结算方式, Nvl(Sum(a.借款金额), 0) As 金额 " & _
    "               From 人员借款记录 A" & _
    "               Where a.借出人 || '' =[1] And a.借出时间 > [2]  And a.取消时间 Is Null " & _
    "               Group By a.结算方式 " & _
    "               Union All " & _
    "               Select '" & strCash & "' As 结算方式, -1 * a.金额  From 人员暂存记录 A " & _
    "               Where 收款员 || '' =[1] And 登记时间 > [2] ) " & _
    " Group By 结算方式  Having Sum(金额) <> 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrPersonName, mdtendDate)
    Set GetBalance = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 09:48:31
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, i As Long, strCaption As String
    
    strCaption = IIf(mbytType = EM_财务收款_非收费员, "收款", "轧帐")
    On Error GoTo errHandle
    If mrsList Is Nothing Then
        If MsgBox("不存在相关的" & strCaption & "数据,需重新提取" & strCaption & "数据,是否重新提取" & strCaption & "数据", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call LoadPersonChargeAndBill
        End If
        Exit Function
    End If
    mrsList.Filter = "性质=-1"
    mrsListBill.Filter = ""
    If mrsList.RecordCount = 0 And mrsListBill.RecordCount = 0 Then
        If MsgBox("不存在相关的" & strCaption & "数据,需重新提取" & strCaption & "数据,是否重新提取" & strCaption & "数据", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call LoadPersonChargeAndBill
        End If
        Exit Function
    End If
    
    mrsList.Filter = "性质>=1"
    If mrsList.RecordCount = 0 And mrsListBill.RecordCount = 0 Then
        If MsgBox("不存在相关的" & strCaption & "数据,需重新提取" & strCaption & "数据,是否重新提取" & strCaption & "数据", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call LoadPersonChargeAndBill
        End If
        Exit Function
    End If
    If CheckMzFeeChargeValied = False Then Exit Function
    With vsChagre
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("结算号码"))
            If zlCommFun.ActualLen(strTemp) > 10 Then
                MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("结算号码")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") > 0 Or InStr(1, strTemp, "|") > 0 Or InStr(1, strTemp, ",") > 0 Then
                MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("结算号码")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
        Next
    End With
    CheckValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Get收费对照(ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费对照
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 10:14:21
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    
    On Error GoTo errHandle
    Set cllData = New Collection
    
    '收缴对照
    mrsList.Filter = "性质>=1 and 结帐ID<>0"
    With mrsList
        .Sort = "性质,结帐id"
        If .RecordCount <> 0 Then .MoveFirst
        strTemp = ""
        Do While Not .EOF
            '性质, 结帐id, '' As 结算方式, 0 As 金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计
            If strTemp <> "" And zlCommFun.ActualLen(strTemp & !性质 & "," & !结帐id) >= 4000 Then
                '性质1,记录ID1|性质2,记录ID2|...|性质n,记录IDn
                strTemp = Mid(strTemp, 2)
                cllData.Add strTemp
                strTemp = ""
            End If
            strTemp = strTemp & "|" & !性质 & "," & !结帐id
            .MoveNext
        Loop
        If strTemp <> "" Then
            '性质1,记录ID1|性质2,记录ID2|...|性质n,记录IDn
            strTemp = Mid(strTemp, 2)
            cllData.Add strTemp
            strTemp = ""
        End If
    End With
    If cllData.Count = 0 Then Exit Function
    Get收费对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Get收缴明细(ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费对照
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 10:14:21
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strHaveBalance As String
    Dim strBalance As String, i As Long
    
    On Error GoTo errHandle
    Set cllData = New Collection
    
    With vsChagre
        strBalance = ""
        For i = 1 To .Rows - 1
            '结算方式1,结算金额1,结算号1,余额1|结算方式2,结算金额2,结算号2,余额2|...
            strTemp = .TextMatrix(i, .ColIndex("结算方式"))
            If Not (strTemp Like "*冲预交*" Or strTemp Like "*借入合计*" _
                        Or strTemp Like "*借出合计*" Or strTemp = "医保相关") And strTemp <> "" Then
                strHaveBalance = strHaveBalance & "," & strTemp
                If i = mlngCashRow Then
                    strTemp = strTemp & "," & Val(Replace(.RowData(i), ",", "")) - mdblRemain
                Else
                    strTemp = strTemp & "," & Val(Replace(.RowData(i), ",", ""))
                End If
                strTemp = strTemp & "," & Trim(.TextMatrix(i, .ColIndex("结算号码")))
                mrsBalance.Filter = "结算方式='" & .TextMatrix(i, .ColIndex("结算方式")) & "'"
                If mrsBalance.EOF Then
                    strTemp = strTemp & "," & 0
                Else
                    strTemp = strTemp & "," & Val(Nvl(mrsBalance!金额))
                End If
                If zlCommFun.ActualLen(strBalance & "|" & strTemp) > 4000 Then
                    strBalance = Mid(strBalance, 2)
                    cllData.Add strBalance
                    strBalance = ""
                End If
                strBalance = strBalance & "|" & strTemp
            End If
        Next
    End With
    mrsBalance.Filter = 0
    With mrsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strTemp = Nvl(!结算方式)
            If InStr(strHaveBalance & ",", "," & strTemp & ",") = 0 And strTemp <> "" Then
                strTemp = strTemp & "," & 0
                strTemp = strTemp & "," & ""
                strTemp = strTemp & "," & Val(Nvl(mrsBalance!金额))
                If zlCommFun.ActualLen(strBalance & "|" & strTemp) > 4000 Then
                    strBalance = Mid(strBalance, 2)
                    cllData.Add strBalance
                    strBalance = ""
                End If
                strBalance = strBalance & "|" & strTemp
            End If
            .MoveNext
        Loop
    End With
    If strBalance <> "" Then
        strBalance = Mid(strBalance, 2)
        cllData.Add strBalance
    End If
    Get收缴明细 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Function Get收缴票据(ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收缴票据
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-11 10:14:21
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strData As String, lngNo As Long
    Dim strPre As String
    On Error GoTo errHandle
    Set cllData = New Collection
    
    '收缴对照
    mrsListBill.Filter = "性质>=1"
    With mrsListBill
        If .RecordCount = 0 Then
            Get收缴票据 = True
            Exit Function
        End If
        .Sort = "票种,性质,开始票号"
        If .RecordCount <> 0 Then .MoveFirst
        strTemp = "": strPre = "": lngNo = 0
        strData = ""
        Do While Not .EOF
            '票种,性质,张数, 开始票号, 终止票号,金额, 发生时间, 批次
            If strPre <> Val(Nvl(!票种)) & "-" & Val(Nvl(!性质)) Then
                 strPre = Val(Nvl(!票种)) & "-" & Val(Nvl(!性质))
                 lngNo = 0
            End If
            lngNo = lngNo + 1
            strTemp = Val(Nvl(!票种))
            strTemp = strTemp & "," & Val(Nvl(!性质))
            strTemp = strTemp & "," & lngNo
            strTemp = strTemp & "," & Val(Nvl(!张数))
            strTemp = strTemp & "," & Nvl(!开始票号)
            strTemp = strTemp & "," & Nvl(!终止票号)
            strTemp = strTemp & "," & Val(Nvl(!金额))
            strTemp = strTemp & "," & Format(!发生时间, "yyyy-mm-dd HH:MM:SS")
            strTemp = strTemp & "," & Nvl(!批次)
            
            If strTemp <> "" And zlCommFun.ActualLen(strData & "|" & strTemp) >= 4000 Then
                '票种,性质,序号,票据张数,开始票号,终止票号,金额,发生时间,批次|票种,性质,序号,票据张数,开始票号,终止票号,金额,发生时间,批次|...
                strData = Mid(strData, 2)
                cllData.Add strData
                strData = ""
            End If
            strData = strData & "|" & strTemp
            .MoveNext
        Loop
        If strData <> "" Then
            '性质1,记录ID1|性质2,记录ID2|...|性质n,记录IDn
            strData = Mid(strData, 2)
            cllData.Add strData
            strData = ""
        End If
    End With
    
    Get收缴票据 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Public Function SaveData(ByVal strStartDate As String, ByVal strEndDate As String, _
    ByVal strMemo As String, ByVal lngDeptID As Long, _
    ByRef strNO As String, ByRef lngID As Long, _
    Optional ByVal dblRemain As Double = 0, Optional strRollingType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存轧帐数据
    '入参:strStartDate-开始轧帐时间
    '       strEndDate-终止轧帐时间
    '       strMemo-备注
    '       lngDeptID-收款部门ID
    '       strRollingType-轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡)
    '出参:strNo-NO
    '        lngID-轧帐ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-09 18:00:05
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, cllData(0 To 3) As Collection
    Dim strTemp As String, i As Long, strDate As String, strSQL1 As String
    Dim dbl预交 As Double, dbl借款合计 As Double, dbl借出合计 As Double
    Dim lng收款ID As Long, str收款NO As String, rsTemp As ADODB.Recordset
    Dim lng组长ID As Long
    
    On Error GoTo errHandle

    If CheckValied = False Then Exit Function
    
    If mbytType <> EM_财务收款_非收费员 Then
        strSQL = _
            "Select b.组id" & vbNewLine & _
            "From 财务组组长构成 A, 缴款成员组成 B, 财务缴款分组 C" & vbNewLine & _
            "Where a.组id = b.组id And b.组id = c.Id And b.成员id = [1]" & vbNewLine & _
            "      And (c.删除日期 > Sysdate Or c.删除日期 Is Null)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
        If Not rsTemp.EOF Then
            lng组长ID = frmChargeBillSel.ShowMe(Me, Val(Nvl(rsTemp!组ID)))
        Else
            lng组长ID = 0
        End If
    End If
    
    mrsList.Filter = "性质=-1"
    mdblRemain = dblRemain
    Do While Not mrsList.EOF
        dbl预交 = dbl预交 + Round(Val(Nvl(mrsList!冲预交)), 2)
        dbl借款合计 = dbl借款合计 + Round(Val(Nvl(mrsList!借款合计)), 2)
        dbl借出合计 = dbl借出合计 + Round(Val(Nvl(mrsList!借出合计)), 2)
        mrsList.MoveNext
    Loop
    If Get收费对照(cllData(0)) = False And Get收缴明细(cllData(1)) = False And Get收缴票据(cllData(2)) = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Set cllPro = New Collection
    lngID = zlDatabase.GetNextId("人员收缴记录")
    strNO = zlDatabase.GetNextNo(137)
    With vsChagre
        'Zl_收费员轧帐记录_Insert
        strSQL = "Zl_收费员轧帐记录_Insert("
        '  Id_In         In 人员收缴记录.Id%Type,
        strSQL = strSQL & "" & lngID & ","
        '  No_In         In 人员收缴记录.No%Type,
        strSQL = strSQL & "'" & strNO & "',"
        '  收款员_In     In 人员收缴记录.收款员%Type,
        strSQL = strSQL & "'" & mstrPersonName & "',"
        '  收款部门id_In In 人员收缴记录.收款部门id%Type,
        strSQL = strSQL & "Null,"
        '  组长id_In In     人员表.id%Type,
        strSQL = strSQL & ZVal(lng组长ID) & ","
        '  冲预交款_In   In 人员收缴记录.冲预交款%Type,
        strSQL = strSQL & "" & dbl预交 & ","
        '  借入合计_In   In 人员收缴记录.借入合计%Type,
        strSQL = strSQL & "" & dbl借款合计 & ","
        '  借出合计_In   In 人员收缴记录.借出合计%Type,
        strSQL = strSQL & "" & dbl借出合计 & ","
        '  摘要_In       In 人员收缴记录.摘要%Type,
        strSQL = strSQL & IIf(strMemo = "", "NULL", "'" & strMemo & "'") & ","
        '  开始时间_In   In 人员收缴记录.开始时间%Type,
        strSQL = strSQL & "to_Date('" & strStartDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  终止时间_In   In 人员收缴记录.终止时间%Type,
        strSQL = strSQL & "to_Date('" & strEndDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  登记人_In     In 人员收缴记录.登记人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  登记时间_In   In 人员收缴记录.登记时间%Type,
        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  收缴标志_In   In 人员收缴记录.收缴标志%Type,
        strSQL = strSQL & "NULL,"
        If cllData(0).Count = 0 Then
            '  收费对照_In   In Varchar2
            strSQL = strSQL & "NULL,"
            '  操作类别_In   In Integer := 0,
            '   0-保存轧帐记录和对照;1-只保存对照
            strSQL = strSQL & "0,"
            '  暂存金_In     In 人员暂存记录.金额%Type := 0,
            strSQL = strSQL & "0,"
            '  类别_In       In Varvhar2(100)
            strSQL = strSQL & "" & IIf(strRollingType = "", "0", "'" & strRollingType & "'") & ")"
            cllPro.Add strSQL
        Else
            For i = 1 To cllData(0).Count
                '  收费对照_In   In Varchar2
                strSQL1 = strSQL & "'" & cllData(0)(i) & "',"
                '  操作类别_In   In Integer := 0,
                '   0-保存轧帐记录和对照;1-只保存对照
                strSQL1 = strSQL1 & "" & IIf(i = 1, "0", "1") & ","
                '  暂存金_In     In 人员暂存记录.金额%Type := 0,
                strSQL1 = strSQL1 & "" & IIf(i = 1, dblRemain, 0) & ","
                '  类别_In       In Varvhar2(100)
                strSQL1 = strSQL1 & "" & IIf(strRollingType = "", "0", "'" & strRollingType & "'") & ")"
                cllPro.Add strSQL1
            Next
        End If
        
        '加载收缴明细
        For i = 1 To cllData(1).Count
            'Zl_收费员轧帐明细_Insert
            strSQL = "Zl_收费员轧帐明细_Insert("
            '收缴id_In   In 人员收缴明细.收缴id%Type,
            strSQL = strSQL & "" & lngID & ","
            '结算信息_In In Varchar2
            '       结算信息_IN:结算方式1,结算金额1,结算号1,余额1|结算方式2,结算金额2,结算号2,余额2|...
            strSQL = strSQL & "'" & cllData(1)(i) & "')"
            cllPro.Add strSQL
        Next
                
        '加载收缴票据
        For i = 1 To cllData(2).Count
            'Zl_收费员轧帐明细_Insert
            strSQL = "Zl_收费员轧帐票据_Insert("
            '收缴id_In   In 人员收缴明细.收缴id%Type,
            strSQL = strSQL & "" & lngID & ","
            ' 票据信息_In Varchar2
            '       格式:票种,性质,序号,票据张数,开始票号,终止票号,金额,发生时间|票种,性质,序号,票据张数,开始票号,终止票号,金额,发生时间|...
            '  --           票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
            '  --           性质:1-正常票据;2-退费收回票据;3-重打收回票据
            '  --            发生时间:yyyy-mm-dd hh24:mi:ss
            strSQL = strSQL & "'" & cllData(2)(i) & "')"
            cllPro.Add strSQL
        Next
    End With
    For i = 0 To 2
        Set cllData(i) = Nothing
    Next
    If mbytType = EM_财务收款_非收费员 Then
        '还需要产生收费员的收款记录
        lng收款ID = zlDatabase.GetNextId("人员收缴记录")
        str收款NO = zlDatabase.GetNextNo(140)
        'Zl_非收费员收款记录_Insert
        strSQL = "Zl_非收费员收款记录_Insert("
        '  Id_In         In 人员收缴记录.Id%Type,
        strSQL = strSQL & "" & lng收款ID & ","
        '  No_In         In 人员收缴记录.No%Type,
        strSQL = strSQL & "'" & str收款NO & "',"
        '  收款部门id_In In 人员收缴记录.收款部门id%Type,
        strSQL = strSQL & "" & "Null,"
        '  摘要_In       In 人员收缴记录.摘要%Type,
        strSQL = strSQL & IIf(strMemo = "", "NULL", "'" & strMemo & "'") & ","
        '  登记人_In     In 人员收缴记录.登记人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  登记时间_In   In 人员收缴记录.登记时间%Type,
        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  轧帐id_In     In 人员收缴记录.Id%Type
        strSQL = strSQL & "" & lngID & ")"
        cllPro.Add strSQL
    End If
    '执行过程
    On Error GoTo ErrCommit:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrCommit:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlPrint(ByVal bytMode As Byte, _
    Optional strDeptName As String = "", Optional strMemo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '       strDeptName-收款部门名称(收费员轧帐时转入)
    '       strMemo-备注(收费员轧帐时转入)
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo ErrHand:
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    If mbytType = EM_收费员轧帐 Or mbytType = EM_财务收款_非收费员 Then
        objPrint.Title.Text = gstr单位名称 & "收费员收款及票据汇总"
        If mlngChargeRollingID = 0 Then
            Set objRow = New zlTabAppRow
            objRow.Add "收费员：" & mstrPersonName
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            'objRow.Add "收款部门:" & strDeptName
            objRow.Add IIf(mbytType <> EM_财务收款_非收费员, "轧帐时间：", "收款时间") & Format(mdtStartDate, "yyyy-mm-dd HH:MM:SS") & "至" & Format(mdtendDate, "yyyy-mm-dd HH:MM:SS")
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add IIf(mbytType <> EM_财务收款_非收费员, "轧帐说明：", "收款时间") & strMemo
            objPrint.BelowAppRows.Add objRow
        Else
            strSQL = "" & _
            "   Select /*+ rule */a.Id,a.No ,a.收款员, a.开始时间, a.终止时间, a.登记时间 ,  " & _
            "         b.名称 As 收款部门, a.摘要,M.组名称 as 财务组  " & _
            "  From 人员收缴记录 A, 部门表 B,财务缴款分组 M" & _
            "  Where a.收款部门id = b.Id(+) and a.缴款组ID=M.ID(+) And a.ID=[1]  " & _
            "  Order by 登记时间,NO desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
            If rsTemp.EOF Then Exit Sub
            Set objRow = New zlTabAppRow
            objRow.Add "收费员：" & rsTemp!收款员
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add "轧帐单号：" & Nvl(rsTemp!NO)
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add "轧帐时间：" & Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            'objRow.Add "收款部门:" & rsTemp!收款部门
            objRow.Add "轧帐时间：" & Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM:SS") & "至" & Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS")
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add "轧帐说明:" & rsTemp!摘要
            objPrint.BelowAppRows.Add objRow
        End If
    ElseIf mbytType = EM_小组收款 Or mbytType = EM_小组轧帐 Then
        strSQL = "" & _
        "   Select /*+ rule */a.Id,a.No , a.开始时间, a.终止时间, a.登记时间 ,  " & _
        "         b.名称 As 收款部门, a.摘要,M.组名称 as 财务组  " & _
        "  From 人员收缴记录 A, 部门表 B,财务缴款分组 M" & _
        "  Where a.收款部门id = b.Id(+) and a.缴款组ID=M.ID(+) And a.ID=[1]  " & _
        "  Order by 登记时间,NO desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
        If rsTemp.EOF Then Exit Sub
        objPrint.Title.Text = gstr单位名称 & "财务组收款及票据汇总"
        Set objRow = New zlTabAppRow
        objRow.Add "小组负责人：" & UserInfo.姓名
        'objRow.Add "收款部门:" & Nvl(rsTemp!收款部门)
        objRow.Add "财务组:" & Nvl(rsTemp!财务组)
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add IIf(mbytType = EM_小组轧帐, "轧帐单号:", "收款单号：") & Nvl(rsTemp!NO)
        If mbytType = EM_小组轧帐 Then
            objRow.Add "轧帐时间：" & Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM:SS") & "至" & Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS")
        Else
            objRow.Add "收款时间：" & Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
        End If
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add IIf(mbytType = EM_小组轧帐, "轧帐说明:", "收款说明:") & Nvl(rsTemp!备注)
        objPrint.BelowAppRows.Add objRow
    ElseIf mbytType = EM_财务收款 Then
        objPrint.Title.Text = gstr单位名称 & "财务收款及票据汇总"
       strSQL = "" & _
        "   Select /*+ rule */a.Id,a.No , a.开始时间, a.终止时间,  a.登记时间 ,  " & _
        "         b.名称 As 收款部门, a.摘要 " & _
        "  From 人员收缴记录 A, 部门表 B " & _
        "  Where a.收款部门id = b.Id(+) and a.缴款组ID=M.ID(+) And a.ID=[1]  " & _
        "  Order by 登记时间,NO desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
        If rsTemp.EOF Then Exit Sub
        objPrint.Title.Text = gstr单位名称 & "财务组收款及票据汇总"
        Set objRow = New zlTabAppRow
        objRow.Add "小组负责人：" & UserInfo.姓名
        'objRow.Add "收款部门:" & Nvl(rsTemp!收款部门)
        objRow.Add "财务组:" & Nvl(rsTemp!财务组)
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "收款单号：" & Nvl(rsTemp!NO)
        objRow.Add "收款时间：" & Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "收款说明:" & Nvl(rsTemp!备注)
        objPrint.BelowAppRows.Add objRow
    Else
        Exit Sub
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    '组装数据
    With vsRptPrint
        .Clear: .Redraw = flexRDNone
        .Rows = 1: .Cols = 5: lngRow = 0: .FixedRows = 0
        '组装收款记录
        .TextMatrix(lngRow, 0) = "收款信息"
        .TextMatrix(lngRow, 1) = "收款信息"
        .TextMatrix(lngRow, 2) = "收款信息"
        .TextMatrix(lngRow, 3) = "收款信息"
        .TextMatrix(lngRow, 4) = "收款信息"
        
        lngRow = lngRow + 1: .Rows = .Rows + 1
        .TextMatrix(lngRow, 0) = "结算方式"
        .TextMatrix(lngRow, 1) = "结算方式"
        .TextMatrix(lngRow, 2) = "结算方式"
        .TextMatrix(lngRow, 3) = "金额"
        .TextMatrix(lngRow, 4) = "结算号"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = Me.BackColor
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = 4
        
        lngRow = lngRow + 1: .Rows = .Rows + 1
        blnFind = False
        For i = 1 To vsChagre.Rows - 1
            strTemp = Trim(vsChagre.TextMatrix(i, vsChagre.ColIndex("结算方式")))
            If strTemp <> "" And strTemp <> "医保相关" Then
                .TextMatrix(lngRow, 0) = strTemp
                .TextMatrix(lngRow, 1) = strTemp
                .TextMatrix(lngRow, 2) = strTemp
                .TextMatrix(lngRow, 3) = Trim(vsChagre.TextMatrix(i, vsChagre.ColIndex("金额")))
                .TextMatrix(lngRow, 4) = Trim(vsChagre.TextMatrix(i, vsChagre.ColIndex("结算号码")))
                blnFind = True
                .Rows = .Rows + 1: lngRow = lngRow + 1
            End If
        Next
        '合计信息
        .TextMatrix(lngRow, 0) = "合计"
        .Cell(flexcpAlignment, lngRow, 0, lngRow, 0) = 4
        .TextMatrix(lngRow, 1) = txtTotal.Text
        .TextMatrix(lngRow, 2) = txtTotal.Text
        .TextMatrix(lngRow, 3) = txtTotal.Text
        .TextMatrix(lngRow, 4) = txtTotal.Text
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        '票据使用信息
        .TextMatrix(lngRow, 0) = "票据使用信息"
        .TextMatrix(lngRow, 1) = "票据使用信息"
        .TextMatrix(lngRow, 2) = "票据使用信息"
        .TextMatrix(lngRow, 3) = "票据使用信息"
        .TextMatrix(lngRow, 4) = "票据使用信息"
        
        lngRow = lngRow + 1: .Rows = .Rows + 1
        blnFind = False
        For i = 0 To vsBill.Rows - 1
            strTemp = Trim(vsBill.TextMatrix(i, 0))
            If strTemp <> "" Then
                .TextMatrix(lngRow, 0) = strTemp
                .TextMatrix(lngRow, 1) = Trim(vsBill.TextMatrix(i, 1))
                .Cell(flexcpAlignment, lngRow, 0, lngRow, 1) = 4
                .TextMatrix(lngRow, 2) = Trim(vsBill.TextMatrix(i, 2))
                .TextMatrix(lngRow, 3) = Trim(vsBill.TextMatrix(i, 2))
                .TextMatrix(lngRow, 4) = Trim(vsBill.TextMatrix(i, 2))
                blnFind = True
                .Rows = .Rows + 1: lngRow = lngRow + 1
            End If
        Next
        If blnFind = False Then
            .TextMatrix(lngRow, 0) = Space(1)
            .TextMatrix(lngRow, 1) = Space(2)
            .TextMatrix(lngRow, 2) = Space(3)
            .TextMatrix(lngRow, 3) = Space(3)
            .TextMatrix(lngRow, 4) = Space(3)
            .Rows = .Rows + 1: lngRow = lngRow + 1
        End If
        blnFind = False
        '退费回收信息
        .TextMatrix(lngRow, 0) = "退费回收票据"
        .TextMatrix(lngRow, 1) = "退费回收票据"
        .TextMatrix(lngRow, 2) = "退费回收票据"
        .TextMatrix(lngRow, 3) = "退费回收票据"
        .TextMatrix(lngRow, 4) = "退费回收票据"
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        .TextMatrix(lngRow, 0) = "类别"
        .TextMatrix(lngRow, 1) = "退费时间"
        .TextMatrix(lngRow, 2) = "退费时间"
        .TextMatrix(lngRow, 3) = "退费金额"
        .TextMatrix(lngRow, 4) = "票据号"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = Me.BackColor
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = 4
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        For i = 1 To vsReturnBill.Rows - 1
            If vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("类型")) = "退费" Then
                strTemp = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("类别")))
                If strTemp <> "" Then
                    .TextMatrix(lngRow, 0) = strTemp
                    .Cell(flexcpAlignment, lngRow, 0, lngRow, 0) = 4
                    .TextMatrix(lngRow, 1) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("收回时间")))
                    .TextMatrix(lngRow, 2) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("收回时间")))
                    .TextMatrix(lngRow, 3) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("金额")))
                    .TextMatrix(lngRow, 4) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("票据号")))
                    blnFind = True
                    .Rows = .Rows + 1: lngRow = lngRow + 1
                    blnFind = True
                End If
            End If
        Next
        If blnFind = False Then
            .TextMatrix(lngRow, 0) = Space(1)
            .TextMatrix(lngRow, 1) = Space(2)
            .TextMatrix(lngRow, 2) = Space(2)
            .TextMatrix(lngRow, 3) = Space(3)
            .TextMatrix(lngRow, 4) = Space(4)
            .Rows = .Rows + 1: lngRow = lngRow + 1
        End If
        blnFind = False
        '重打回收信息
        .TextMatrix(lngRow, 0) = "重打回收信息"
        .TextMatrix(lngRow, 1) = "重打回收信息"
        .TextMatrix(lngRow, 2) = "重打回收信息"
        .TextMatrix(lngRow, 3) = "重打回收信息"
        .TextMatrix(lngRow, 4) = "重打回收信息"
        lngRow = lngRow + 1: .Rows = .Rows + 1
       .TextMatrix(lngRow, 0) = "类别"
        .TextMatrix(lngRow, 1) = "重打时间"
        .TextMatrix(lngRow, 2) = "重打时间"
        .TextMatrix(lngRow, 3) = "重打金额"
        .TextMatrix(lngRow, 4) = "票据号"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = Me.BackColor
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = 4
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        For i = 1 To vsReturnBill.Rows - 1
            If vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("类型")) = "重打" Then
                strTemp = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("类别")))
                If strTemp <> "" Then
                    .TextMatrix(lngRow, 0) = strTemp
                    .Cell(flexcpAlignment, lngRow, 0, lngRow, 0) = 4
                    .TextMatrix(lngRow, 1) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("收回时间")))
                    .TextMatrix(lngRow, 2) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("收回时间")))
                    .TextMatrix(lngRow, 3) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("金额")))
                    .TextMatrix(lngRow, 4) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("票据号")))
                    .Rows = .Rows + 1: lngRow = lngRow + 1
                    blnFind = True
                End If
            End If
        Next
        If blnFind = False Then
            .TextMatrix(lngRow, 0) = Space(1)
            .TextMatrix(lngRow, 1) = Space(2)
            .TextMatrix(lngRow, 2) = Space(2)
            .TextMatrix(lngRow, 3) = Space(3)
            .TextMatrix(lngRow, 4) = Space(4)
            .Rows = .Rows + 1: lngRow = lngRow + 1
        End If
        .Rows = .Rows - 1
       ' .AutoSizeMode = flexAutoSizeColWidth
        '.AutoSize 0, .Cols - 1
        For i = 0 To .Rows - 1
            .MergeRow(i) = True
            .RowHeight(i) = 350
        Next
        For i = 0 To .Cols - 1
            If i = 0 Then .ColWidth(i) = 800
            If i = 1 Then .ColWidth(i) = 1000
            If i = 2 Then .ColWidth(i) = 800
            If i = 3 Then .ColWidth(i) = 1400
            If i = 4 Then .ColWidth(i) = 5000
            .MergeCol(i) = True
        Next
        .MergeCells = flexMergeRestrictRows
        .Redraw = flexRDDirect
    End With
    
    Set objPrint.Body = vsRptPrint
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
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Property Get GetCashMoney() As Double
    '获取现金金额
    With vsChagre
        If mlngCashRow < 1 Or mlngCashRow > .Rows - 1 Then GetCashMoney = 0: Exit Property
        GetCashMoney = Val(Replace(.TextMatrix(mlngCashRow, .ColIndex("金额")), ",", ""))
    End With
End Property

Private Function CheckMzFeeChargeValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊收费的费用合计与结算合计是否合法
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-22 10:27:31
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, cllData As Collection
    '       cllData -Array(性质, NO, 记录状态, 结算金额, 冲预交)
    '                     性质=1(结算金额不正确;2.异常单据)
    On Error GoTo errHandle
    
    If mrsList Is Nothing Then Exit Function
    If mrsList.State <> 1 Then Exit Function
    mrsList.Filter = "性质=1"
    If mrsList.RecordCount = 0 Then GoTo GoSucces:
    Set cllData = New Collection
    With mrsList
        .Sort = "结帐id": .MoveFirst
        strTemp = ""
        Do While Not .EOF
            '性质, 结帐id, '' As 结算方式, 0 As 金额, 0 As 冲预交, 0 As 借款合计, 0 As 借出合计
            If strTemp <> "" And zlCommFun.ActualLen(strTemp & "," & !结帐id) >= 4000 Then
                '结帐ID1,结帐ID2,...
                strTemp = Mid(strTemp, 2)
                Call CheckMzFeeValied(strTemp, cllData)
                strTemp = ""
            End If
            strTemp = strTemp & "," & !结帐id
            .MoveNext
        Loop
        If strTemp <> "" Then
            '结帐ID1,结帐ID2,...
            strTemp = Mid(strTemp, 2)
            Call CheckMzFeeValied(strTemp, cllData)
            strTemp = ""
        End If
    End With
    'cllData:Array(性质, NO, 记录状态, 结算金额, 冲预交)
    If cllData.Count <> 0 Then
        If frmErrInfor.ShowErrInfor(Me, cllData) = False Then Exit Function
    End If
GoSucces:
    mrsList.Filter = 0
    CheckMzFeeChargeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckMzFeeValied(ByVal strIDs As String, ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定结帐ID的门诊费用与预交是否合法
    '入参:strIDs-结帐IDs,格式为:结帐ID1,结帐ID2,...
    '出参:cllData-array(性质,NO,记录状态,结算金额,冲预交)
    '                     性质=1(结算金额不正确;2.异常单据)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-22 10:57:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = " " & _
    "   Select * " & _
    "   From (With c_异常 As (Select /*+cardinality(j,10)*/ a.结帐id, Max(a.记录状态) As 记录状态, Sum(a.结帐金额) As 结算金额 " & _
    "                       From 门诊费用记录 A, Table(f_Num2list([1])) J " & _
    "                       Where a.结帐id = j.Column_Value And MOD(a.记录性质,10) = 1 " & _
    "                       Group By a.结帐id " & _
    "                       Having Nvl(Sum(a.结帐金额), 0) <> (Select Nvl(Sum(冲预交), 0) " & _
    "                                                     From 病人预交记录 M " & _
    "                                                     Where a.结帐id = m.结帐id)) " & _
    "          Select a.结帐id, b.No, Max(a.记录状态) As 记录状态, Max(a.结算金额) As 结算金额, Sum(m.冲预交) As 冲预交 " & _
    "          From c_异常 A, 门诊费用记录 B," & _
    "               (Select B1.结帐id, Nvl(Sum(B1.冲预交), 0) As 冲预交 " & _
    "                 From 病人预交记录 B1, c_异常 C " & _
    "                 Where B1.结帐id = c.结帐id " & _
    "                 Group By B1.结帐id) M " & _
    "          Where a.结帐id = m.结帐id(+) And a.结帐id = b.结帐id(+) " & _
    "          Group By a.结帐id, b.No) " & _
    "          Order By NO "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
    With rsTemp
        strSQL = ""
        Do While Not .EOF
            '性质,NO,记录状态,结算金额,冲预交
            cllData.Add Array(1, Nvl(!NO), Val(Nvl(!记录状态)), Val(Nvl(!结算金额)), Val(Nvl(!冲预交)))
            .MoveNext
        Loop
    End With
    CheckMzFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
