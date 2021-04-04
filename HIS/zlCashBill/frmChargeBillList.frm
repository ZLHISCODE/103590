VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeBillList 
   Caption         =   "收款明细数据"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChargeBillList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picConList 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   13065
      TabIndex        =   9
      Top             =   510
      Width           =   13065
      Begin VB.Frame fraSplit 
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   13575
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "单据号:"
         Height          =   210
         Left            =   75
         TabIndex        =   12
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "时间范围:"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   150
         Width           =   945
      End
   End
   Begin VB.PictureBox picFeeList 
      BorderStyle     =   0  'None
      Height          =   7140
      Left            =   9960
      ScaleHeight     =   7140
      ScaleWidth      =   3015
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
         Height          =   1800
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   2475
         _cx             =   4366
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillList.frx":0502
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
   Begin VB.PictureBox picList 
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
      Height          =   3030
      Left            =   585
      ScaleHeight     =   3030
      ScaleWidth      =   9525
      TabIndex        =   4
      Top             =   1470
      Width           =   9525
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   1800
         Left            =   330
         TabIndex        =   5
         Top             =   645
         Width           =   8505
         _cx             =   15002
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
         FormatString    =   $"frmChargeBillList.frx":057C
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
   Begin VB.PictureBox picBalance 
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
      Height          =   2925
      Left            =   930
      ScaleHeight     =   2925
      ScaleWidth      =   4485
      TabIndex        =   2
      Top             =   3900
      Width           =   4485
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   870
         Left            =   150
         TabIndex        =   3
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
         FormatString    =   $"frmChargeBillList.frx":05F6
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
   Begin VB.PictureBox picBillList 
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
      Left            =   5790
      ScaleHeight     =   2685
      ScaleWidth      =   3630
      TabIndex        =   0
      Top             =   4080
      Width           =   3630
      Begin VSFlex8Ctl.VSFlexGrid vsBillList 
         Height          =   1800
         Left            =   585
         TabIndex        =   1
         Top             =   480
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
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillList.frx":0670
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   7935
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   635
      SimpleText      =   $"frmChargeBillList.frx":06EA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeBillList.frx":0731
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16272
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "刘兴洪"
            TextSave        =   "刘兴洪"
            Object.ToolTipText     =   "当前操作员:刘兴洪"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmChargeBillList.frx":0FC5
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeBillList"
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

Private mstrCashBalance As String '现金结算方式
Private mbytType As TotalType '1-收费员轧帐；2-小组收款;3-小组轧帐;4-财务收款
Private mstrChargeRollingID As String '轧帐ID或收款ID(根据mbytType)来决定,多个时,用逗号分隔,如:123,23,11
Private mdtStartDate As Date, mdtendDate As Date '轧帐的开始时间或结束时间
Private mblnDel As Boolean '是否作废记录
Private Enum mPaneIndex
    EM_PN_ConList = 260101  '条件
    EM_PN_LIST = 260102  '收款汇总
    EM_PN_BALANCE = 260103    '结算方式
    EM_PN_BILL = 260104  '退费票据
    EM_PN_FeeLIST = 260105  '费用汇总
End Enum
Private mbytFontSize As Byte
Private mbyt票据分配规则 As Byte    '票据分配规则:0-根据实际打印分配票号;1-根据系统预定规则分配;2-根据用户自定义规则分配
Private mblnNotBrush As Boolean '不刷新数据
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mstrPersonName As String
Private mstrRollingType As String '轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
Private mblnFirst As Boolean

Public Sub ShowMe(ByVal frmMain As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As Byte, ByVal strChargeRollingID As String, _
      Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, _
      Optional ByVal blnDel As Boolean = False, Optional strPersonName As String, _
      Optional strRollingType As String)
    '-------------------------------------------------------------------------------------------------
    '功能:程序入口,显示指定轧帐或收款记录的明细数据
    '入参:frmMain-调用的主窗体
    '    lngModule-模块号
    '    strPrivs-权限串
    '　　bytType:1-收费员轧帐；2-小组收款;3-小组轧帐;4-财务收款。
    '    strChargeRollingID -轧帐ID或收款ID(多个时,用逗号分隔,如:123,23,11)
    '    dtStartDate-可选参数,开始轧帐时间,strChargeRollingID=0时，必须传入
    '    dtEndDate-可选参数，结束轧帐时间,strChargeRollingID=0时，必须传入
    '    blnDel-是否作废记录
    '    strPersonName-指定的收费员(为空时,为当前操作员)
    '    strRollingType-轧帐类别,bytType=1时有效分别为:
    '               0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡
    '编制:刘兴洪
    '日期:2013-09-16 10:08:39
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytType = bytType:  mstrChargeRollingID = strChargeRollingID
    mdtStartDate = dtStartDate: mdtendDate = dtEndDate: mblnDel = blnDel
    mstrPersonName = IIf(strPersonName = "", UserInfo.姓名, strPersonName)
    mstrRollingType = strRollingType: mblnFirst = True
    Call InitFace: Call zlDefCommandBars
    Err = 0: On Error Resume Next
    Call zlRefresh
    Me.Show 1, frmMain
End Sub
Private Function ReadListData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取相关的明细数据
    '返回:数据获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-16 10:12:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Call LoadFilterRange
    If mstrChargeRollingID = "" And (mbytType = EM_收费员轧帐 Or mbytType = EM_财务收款_非收费员) Then
         ReadListData = LoadPersonList         '加载收费员轧帐明细数据
         Call LoadFeeData(True)   '加载费用信息
    Else
         ReadListData = LoadList          '加载相关的明细数据
    End If
    Call picBalance_Resize
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Function LoadPersonList() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费员当前轧帐的明细数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-16 10:13:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bytType As Byte
    Dim strWithTable As String
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim lngRow As Long, lngNo As Long
    Dim mstrCashBalance As String, str结算方式 As String, strTemp As String, int票种 As Integer
    Dim dblToTotal As Double, strWhere As String, strRollingType As String
    
    On Error GoTo errHandle
    mblnNotBrush = True
    bytType = 1: dblToTotal = 0
    '就诊卡
    strWhere = " And instr([4],','||A.结算性质||',')>0 "
    If mstrRollingType <> "" And InStr("," & mstrRollingType & ",", ",0,") = 0 Then
        If Get轧帐结算性质(mstrRollingType, strRollingType) = False Then
            Call InitGrid: LoadPersonList = False
            Exit Function
        End If
    Else
        strRollingType = ",2,3,4,5,6,"
    End If
    '预交款填NULL,2-结帐,3-收费,4-挂号,5-就诊卡,6-补充医保结算
    'mstrRollingType:轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
    '收缴对照:1-收费(含挂号),2-结帐,3-预交,4-借款;5-消费卡充值;6--消费卡面值;7-暂存金(本次增加)；８－收款或轧帐作废对照（本次增加）；9-二次结算记录
    strWithTable = "" & _
    "   With 轧帐数据 as (" & _
    "   Select   A.结算性质,A.结帐ID,sum(A.冲预交) as 冲预交" & vbCrLf & _
    "   From 病人预交记录 A" & vbCrLf & _
    "   Where a.操作员姓名 || '' = [1]  and a.记录性质<>1 " & strWhere & vbCrLf & _
    "        And A.收款时间 Between  [2] And [3] " & vbCrLf & _
    "   Group by  A.结算性质,A.结帐ID)" & vbCrLf

    strSQL = ""
    '记录性质:1,'收费',4,'挂号',5,'就诊卡',6,'结帐',7,'预交款',8,'消费卡充值',9,'消费卡面额',10,'借款',11,'借出',12,'暂存',13,'二次结算'
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",5,") > 0 Then  '就诊卡
        strSQL = "" & _
        "Select 5 AS 记录性质,A.结帐ID as 记录ID,A.NO,A.记录状态,max(A.姓名) AS 病人姓名,max(A.性别) AS 性别,max(A.年龄) AS 年龄,max(decode(A.门诊标志,2,-1*NULL,A.标识号)) AS 门诊号,max(decode(A.门诊标志,2,标识号,null)) AS 住院号, " & vbCrLf & _
        "     Max(A.操作员姓名) AS 操作员姓名, NULL AS 结算方式, " & _
        "     sum(A.结帐金额) AS 金额合计,Max(登记时间) AS 收费时间 " & vbCrLf & _
        "   From 住院费用记录 A,轧帐数据 B " & vbCrLf & _
        "   Where A.结帐ID=B.结帐ID And B.结算性质=5 " & vbCrLf & _
        "       And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.结帐id And x.收款员 = [1] And x.Id = y.收缴ID And y.性质 = 1 And x.作废时间 Is Null) " & vbCrLf & _
        "   GROUP BY A.NO,A.结帐ID,A.记录状态 " & vbCrLf
    End If
    

    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",1,") > 0 Or InStr("," & mstrRollingType & ",", ",4,") > 0 Then '就收费,挂号诊卡,补结算
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 记录性质, Min(记录id) As 记录id, NO, Max(记录状态) As 记录状态, 病人姓名, 性别, 年龄, 门诊号, 住院号, 操作员姓名, 结算方式, Sum(金额合计) As 金额合计, 收费时间 From ( " & _
        "   Select  MOD(A.记录性质,10) As 记录性质,A.结帐ID as 记录ID,A.NO,A.记录状态,max(A.姓名) AS 病人姓名,max(A.性别) AS 性别,max(A.年龄) AS 年龄,max(decode(A.门诊标志,2,-1*NULL,A.标识号)) AS 门诊号,max(decode(A.门诊标志,2,标识号,null)) AS 住院号, " & vbCrLf & _
        "     Max(A.操作员姓名) AS 操作员姓名, NULL AS 结算方式, " & vbCrLf & _
        "     sum(A.结帐金额) AS 金额合计,Max(A.登记时间) AS 收费时间 " & vbCrLf & _
        "   From 门诊费用记录 A,轧帐数据 B " & vbCrLf & _
        "   Where A.结帐ID=B.结帐ID And B.结算性质 in (3,4) And Nvl(A.费用状态, 0)=0 " & _
        "       And Not Exists(Select y.记录id  From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.结帐id And x.收款员 = [1] And x.Id = y.收缴ID And y.性质 = 1 And x.作废时间 Is Null)     " & _
        "   GROUP BY A.记录性质,A.结帐ID,A.NO,A.记录状态 ) " & _
        "   Group By 记录性质, NO, 病人姓名, 性别, 年龄, 门诊号, 住院号, 操作员姓名, 结算方式, 收费时间 " & _
        "   Union ALL" & vbCrLf & _
        "   Select 13 as 记录性质, a.结算id As 记录id, a.No, a.记录状态, Max(c.姓名) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, c.门诊号 As 门诊号," & vbNewLine & _
        "        c.住院号 As 住院号, Max(a.操作员姓名) As 操作员姓名, Null As 结算方式, Sum(B.冲预交) As 金额合计, Max(a.登记时间) As 收费时间" & vbNewLine & _
        "   From 费用补充记录 A, 轧帐数据 B,病人信息 C" & vbNewLine & _
        "   Where A.结算ID=B.结帐ID and A.病人ID=C.病人ID And B.结算性质=6 And a.记录性质 = 1 And Nvl(A.费用状态, 0)=0  " & _
        "           And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.结算id  And x.Id = y.收缴id And y.性质 = 9 And x.作废时间 Is Null)" & vbNewLine & _
        "   Group By a.记录性质, a.结算id, a.No, a.记录状态, c.门诊号, c.住院号"
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",3,") > 0 Then '结帐
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 6 AS 记录性质,A.ID as 记录ID,A.NO,A.记录状态,max(C.姓名) AS 病人姓名,max(C.性别) AS 性别,max(C.年龄) AS 年龄, " & vbCrLf & _
        "     max(c.门诊号) AS 门诊号,max(C.住院号) AS 住院号, " & vbCrLf & _
        "     Max(A.操作员姓名) AS 操作员姓名,NULL AS 结算方式, " & vbCrLf & _
        "     sum(b.冲预交) AS 金额合计,Max(A.收费时间) AS 收费时间 " & vbCrLf & _
        "   From 病人结帐记录 A,轧帐数据 B,病人信息 C " & vbCrLf & _
        "   Where A.ID=B.结帐ID  And B.结算性质=2 And A.病人ID=C.病人ID(+) And nvl(A.结算状态,0)=0 " & vbCrLf & _
        "         And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴ID And y.性质 = 2 And x.作废时间 Is Null) " & vbCrLf & _
        "   Group by A.NO,A.ID,A.记录状态 " & vbCrLf
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",2,") > 0 Then '预交款(充值)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As 记录性质,A.ID as 记录ID, a.No, a.记录状态, Max(nvl(M.姓名,c.姓名)) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, Max(c.门诊号) As 门诊号, " & vbCrLf & _
        "      Max(Decode(m.住院号, Null, c.住院号, m.住院号)) As 住院号, Max(a.操作员姓名) As 操作员姓名,Max(结算方式) AS 结算方式, Sum(a.金额) As 金额合计, Max(a.收款时间) As 收费时间 " & vbCrLf & _
        "   From 病人预交记录 A, 病人信息 C, 病案主页 M " & vbCrLf & _
        "   Where a.病人id = c.病人id(+) And a.病人id = m.病人id(+) And a.主页id = m.主页id(+)  " & vbCrLf & _
        "     And a.记录性质 = 1 And Nvl(a.结算性质,0) <> 12 And a.操作员姓名 || '' = [1]  " & vbCrLf & _
        "     And a.收款时间  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴id And y.性质 = 3 And x.作废时间 Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.记录状态 "
    End If
    
    If InStr("," & mstrRollingType & ",", ",21,") > 0 Then '门诊预交款(充值)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As 记录性质,A.ID as 记录ID, a.No, a.记录状态, Max(nvl(M.姓名,c.姓名)) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, Max(c.门诊号) As 门诊号, " & vbCrLf & _
        "      Max(Decode(m.住院号, Null, c.住院号, m.住院号)) As 住院号, Max(a.操作员姓名) As 操作员姓名,Max(结算方式) AS 结算方式, Sum(a.金额) As 金额合计, Max(a.收款时间) As 收费时间 " & vbCrLf & _
        "   From 病人预交记录 A, 病人信息 C, 病案主页 M " & vbCrLf & _
        "   Where a.病人id = c.病人id(+) And a.病人id = m.病人id(+) And a.主页id = m.主页id(+)  " & vbCrLf & _
        "     And a.记录性质 = 1 And Nvl(a.预交类别,0) = 1 And Nvl(a.结算性质,0) <> 12 And a.操作员姓名 || '' = [1]  " & vbCrLf & _
        "     And a.收款时间  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴id And y.性质 = 3 And x.作废时间 Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.记录状态 "
    End If
    
    If InStr("," & mstrRollingType & ",", ",22,") > 0 Then '住院预交款(充值)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As 记录性质,A.ID as 记录ID, a.No, a.记录状态, Max(nvl(M.姓名,c.姓名)) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, Max(c.门诊号) As 门诊号, " & vbCrLf & _
        "      Max(Decode(m.住院号, Null, c.住院号, m.住院号)) As 住院号, Max(a.操作员姓名) As 操作员姓名,Max(结算方式) AS 结算方式, Sum(a.金额) As 金额合计, Max(a.收款时间) As 收费时间 " & vbCrLf & _
        "   From 病人预交记录 A, 病人信息 C, 病案主页 M " & vbCrLf & _
        "   Where a.病人id = c.病人id(+) And a.病人id = m.病人id(+) And a.主页id = m.主页id(+)  " & vbCrLf & _
        "     And a.记录性质 = 1 And Nvl(a.预交类别,0) = 2 And Nvl(a.结算性质,0) <> 12 And a.操作员姓名 || '' = [1]  " & vbCrLf & _
        "     And a.收款时间  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴id And y.性质 = 3 And x.作废时间 Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.记录状态 "
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",3,") > 0 Then '结帐(充值)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 7 As 记录性质,A.ID as 记录ID, a.No, a.记录状态, Max(nvl(M.姓名,c.姓名)) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, Max(c.门诊号) As 门诊号, " & vbCrLf & _
        "      Max(Decode(m.住院号, Null, c.住院号, m.住院号)) As 住院号, Max(a.操作员姓名) As 操作员姓名,Max(结算方式) AS 结算方式, Sum(a.金额) As 金额合计, Max(a.收款时间) As 收费时间 " & vbCrLf & _
        "   From 病人预交记录 A, 病人信息 C, 病案主页 M " & vbCrLf & _
        "   Where a.病人id = c.病人id(+) And a.病人id = m.病人id(+) And a.主页id = m.主页id(+)  " & vbCrLf & _
        "     And a.记录性质 = 1 And Nvl(a.结算性质,0) = 12  And a.操作员姓名 || '' = [1]  " & vbCrLf & _
        "     And a.收款时间  between [2] And [3]  " & vbCrLf & _
        "     And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴id And y.性质 = 3 And x.作废时间 Is Null) " & vbCrLf & _
        "   Group By a.No,A.ID, a.记录状态 "
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",6,") > 0 Then '消费卡(充值,面值)
        strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
        "   Select 8 as 记录性质,A.结算ID as 记录ID, a.卡号 As no,a.记录状态,'' as 病人姓名,'' as 性别,'' as 年龄,null 门诊号,NULL as 住院号, " & vbCrLf & _
        "        A.操作员姓名 AS 操作员姓名,a.结算方式, a.实收金额  As 金额合计, A.登记时间 AS 收费时间  " & vbCrLf & _
        "   From 病人卡结算记录 A, 病人卡结算记录 B " & vbCrLf & _
        "   Where a.交易序号 = b.交易序号(+) And a.消费卡id = b.消费卡id(+) And (a.记录性质 = 2 Or a.记录性质 = 3 And b.记录性质 = 2) And b.记录性质(+) = 2 " & vbCrLf & _
        "         And a.Id <> b.Id(+) And a.操作员姓名 || '' = [1] And a.登记时间 Between [2] And [3]  " & vbCrLf & _
        "         And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.结算Id And x.收款员 = [1] And x.Id = y.收缴ID And y.性质 = 5 And x.作废时间 Is Null)" & vbCrLf & _
        "   Union All " & vbCrLf & _
        "   Select 9 As 记录性质,A.结算ID as 记录ID, a.卡号 As NO, a.记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, " & vbCrLf & _
        "       Null As 住院号, a.操作员姓名, a.结算方式, a.实收金额 As 金额合计, a.登记时间 " & vbCrLf & _
        "   From 病人卡结算记录 A, 病人卡结算记录 B" & vbCrLf & _
        "   Where a.交易序号 = b.交易序号(+) And a.消费卡id = b.消费卡id(+) And (a.记录性质 = 1 Or a.记录性质 = 3 And b.记录性质 = 1) And b.记录性质(+) = 1 " & vbCrLf & _
        "         And a.Id <> b.Id(+) And a.操作员姓名 || '' = [1] And a.登记时间 Between [2] And [3] " & vbCrLf & _
        "         And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.结算Id And x.收款员 = [1] And x.Id = y.收缴id And y.性质 = 6 And x.作废时间 Is Null) " & vbCrLf
    End If
    
    '借款和借出
    strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
    "   Select 10 As 记录性质,A.ID as 记录ID, ltrIm(to_Char(a.ID)) As NO, 1 As 记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, " & _
    "     Null As 住院号, a.借款人 As 操作员姓名, a.结算方式, a.借款金额 As 金额, a.借出时间 " & _
    "   From 人员借款记录 A " & _
    "   Where a.借款人 || '' = [1]   And a.取消时间 Is Null  " & _
    "     And a.借出时间 Between [2] And [3] " & _
    "     And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴ID And y.性质 = 4 And x.作废时间 Is Null)  " & _
    "        Union All " & _
    "   Select 11 As 记录性质,A.ID as 记录ID,ltrIm( to_Char(a.ID)) As NO, 1 As 记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, " & _
    "          Null As 住院号, a.借出人 As 操作员姓名, a.结算方式, a.借款金额 As 金额, a.借出时间 " & _
    "   From 人员借款记录 A " & _
    "   Where a.借出人 || '' = [1]   And a.取消时间 Is Null  " & _
    "     And a.借出时间 Between [2] And [3] " & _
    "     And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴ID And y.性质 = 4 And x.作废时间 Is Null)  "
    
    '暂存金
    strSQL = strSQL & IIf(strSQL <> "", vbCrLf & " Union ALL ", "") & vbCrLf & _
    "  Select 12 As 记录性质,A.ID as 记录ID, a.No As NO, Decode(收回时间, Null, 1, 2) As 记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, Null As 住院号, " & vbCrLf & _
    "      a.登记人 As 操作员姓名, '" & mstrCashBalance & "' As 结算方式, a.金额, a.登记时间 " & _
    "   From 人员暂存记录 A " & vbCrLf & _
    "   Where 收款员 || '' = [1] And 登记时间 > [2] And  登记时间 <= [3]  " & vbCrLf & _
    "     And A.记录性质=2 And A.收回时间 Is Null" & vbCrLf & _
    "     And Not Exists (Select y.记录id  From 人员收缴记录 X, 人员收缴对照 Y  Where y.记录id = a.Id And x.收款员 = [1] And x.Id = y.收缴id And y.性质 = 7 And x.作废时间 Is Null)" & vbCrLf
    
    If InStr(strSQL, "轧帐数据") > 0 Then
        strSQL = strWithTable & strSQL
    End If
    
    strSQL = "" & _
    "   SELECT /*+ rule */ 记录性质 As 性质,记录ID,decode(记录性质,1,'收费',4,'挂号',5,'就诊卡',6,'结帐',7,'预交款',8,'消费卡充值',9,'消费卡面额',10,'借款',11,'借出',12,'暂存',13,'补结算','') AS 类别, " & _
    "        NO,记录状态,病人姓名,性别,年龄,门诊号,住院号,操作员姓名,结算方式," & _
    "       Trim(to_char(金额合计,'99999999990.00')) As 金额合计,to_char(收费时间,'yyyy-mm-dd hh24:mi:ss')  as 收费时间 " & _
    "   FROM ( " & strSQL & " ) " & _
    "   ORDER BY 记录性质,收费时间 DESC,NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrPersonName, mdtStartDate, mdtendDate, strRollingType)
    
    With vsList
        .Clear 1: .Rows = 2: .FixedRows = 1
        If Not rsTemp.EOF Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            Select Case .ColKey(i)
            Case "性质", "记录状态", "记录ID"
                .ColWidth(i) = 0: .ColHidden(i) = True
            Case "NO", "类别", "性别", "年龄", "门诊号", "住院号", "结算方式", "收费时间"
                .ColAlignment(i) = flexAlignCenterCenter
            Case "金额合计"
                .ColAlignment(i) = flexAlignRightCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        If rsTemp.RecordCount <> 0 Then
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, .ColIndex("金额合计"), "#######0.00", &HFFC0C0, vbBlack, True, "合计"
            For i = 0 To .ColIndex("金额合计") - 1
                .TextMatrix(.Rows - 1, i) = "合计"
            Next
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .MergeRow(.Rows - 1) = True
            .MergeCells = flexMergeRestrictRows
           '问题号:110535,焦博,2017/08/10,用颜色区分退费记录和被退费记录
            For i = 1 To .Rows - 1
                Select Case .TextMatrix(i, .ColIndex("记录状态"))
                Case 1
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                Case 2
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                Case 3
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                End Select
            Next
        End If
    End With
    zl_vsGrid_Para_Restore mlngMode, vsList, Me.Name, "明细信息列表", False
    Call LoadDetailData
    mblnNotBrush = False
    LoadPersonList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Function

Private Sub LoadDetailData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载明细数据
    '编制:刘兴洪
    '日期:2013-09-16 14:43:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, bytType As Byte, intRecordSta As Integer
    Dim lng记录ID As Long, rsTemp As ADODB.Recordset, lngRow As Long
    Dim blnNOMoved As Boolean, byt票种 As Byte, strSQL As String
    
    On Error GoTo errHandle
    With vsList
        If .Row < 1 Or .Col < 0 Then GoTo GoClear:
        strNO = Trim(.TextMatrix(.Row, .ColIndex("NO")))
        bytType = Val(.TextMatrix(.Row, .ColIndex("性质")))
        intRecordSta = Val(.TextMatrix(.Row, .ColIndex("记录状态")))
        lng记录ID = Val(.TextMatrix(.Row, .ColIndex("记录ID")))
        If strNO = "" Or lng记录ID = 0 Then
            If ShowOnlyFactList Then
                vsBalance.Clear 1: vsBalance.Rows = 2
                Exit Sub
            End If
            GoTo GoClear:
        End If
    End With
    '加载结算方式信息
    'decode(记录性质,1,'收费',4,'挂号',5,'就诊卡',6,'结帐',7,'预交款',8,'消费卡充值',9,'消费卡面额',10,'借款',11,'借出',12,'暂存',13,'二次结算')
    Select Case bytType
    Case 1, 4, 5, 6, 7, 13
        '收费,挂号,就诊卡,结帐
        If bytType = 1 Then
            blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNO, , "1")
            byt票种 = 1
        ElseIf bytType = 4 Then
            blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNO, , "4")
            byt票种 = 4
        ElseIf bytType = 5 Then
            blnNOMoved = zlDatabase.NOMoved("住院费用记录", strNO, , "5")
            byt票种 = 5
        ElseIf bytType = 6 Then
            blnNOMoved = zlDatabase.NOMoved("病人结帐记录", strNO)
            byt票种 = 3
        ElseIf bytType = 7 Then
            blnNOMoved = zlDatabase.NOMoved("病人预交记录", strNO, , "1")
            byt票种 = 2
        ElseIf bytType = 13 Then
            blnNOMoved = zlDatabase.NOMoved("费用补充记录", strNO, , "1")
            byt票种 = 1
        Else
            blnNOMoved = False
        End If
        If byt票种 = 2 Then
            strSQL = " " & _
             " Select  A.结算方式, a.金额, A.结算号码, A.卡号, A.交易流水号, A.交易说明 " & _
             " From " & IIf(blnNOMoved, "H", "") & "病人预交记录 A,结算方式 B " & _
             " Where a.ID=[1] And a.结算方式=B.名称(+) " & _
             " Order by  decode(nvl(B.性质,0),1,1,2,2,3,10,4,11,4) ,A.结算方式"
        Else
            If bytType = 1 Then
                strSQL = "Select Decode(Mod(记录性质, 10), 1, '[冲预交款]', A.结算方式) As 结算方式," & vbNewLine & _
                "                   A.冲预交 As 金额, A.结算号码, A.卡号,A. 交易流水号,A.交易说明" & vbNewLine & _
                "              From 病人预交记录 A,结算方式 B" & vbNewLine & _
                "              Where a.结算序号=[2] And a.结算方式=B.名称(+)" & vbNewLine & _
                "              Order by decode(Mod(记录性质, 10), 0,decode(nvl(B.性质,0),1,1,2,2,3,10,4,11,4)) ,结算方式"
            Else
                strSQL = " " & _
                 " Select Decode(Mod(记录性质, 10), 1, '[冲预交款]', A.结算方式) As 结算方式,  " & _
                 "      A.冲预交 As 金额, A.结算号码, A.卡号,A. 交易流水号,A.交易说明 " & _
                 " From " & IIf(blnNOMoved, "H", "") & "病人预交记录 A,结算方式 B " & _
                 " Where a.结帐ID=[1] And a.结算方式=B.名称(+) " & _
                 " Order by decode(Mod(记录性质, 10), 0,decode(nvl(B.性质,0),1,1,2,2,3,10,4,11,4)) ,结算方式"
            End If
        End If
        strSQL = "" & _
        "   Select decode(nvl(b.性质,0),0,0,3 ,12,4,12,  b.性质 ) as 排序,A.结算方式,sum(A.金额) as 金额, " & _
        "          max(a.结算号码) as 结算号码,max(a.卡号) as 卡号,max(a.交易流水号) as 交易流水号,max(a.交易说明) as 交易说明" & _
        "   From (" & strSQL & ") A,结算方式 B " & _
        "   Where A.结算方式=b.名称(+)" & _
        "   Group by A.结算方式,decode(nvl(b.性质,0),0,0,3 ,12,4,12,  b.性质 )" & _
        "   Order by 排序,结算方式"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng记录ID, -1 * lng记录ID)
        With vsBalance
            .Clear 1
            .Rows = rsTemp.RecordCount + IIf(rsTemp.RecordCount = 0, 2, 1)
            lngRow = 1
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsTemp!金额)), "###0.00;-###0.00; ;")
                .TextMatrix(lngRow, .ColIndex("结算号码")) = Nvl(rsTemp!结算号码)
                .TextMatrix(lngRow, .ColIndex("卡号")) = Nvl(rsTemp!卡号)
                .TextMatrix(lngRow, .ColIndex("交易流水号")) = Nvl(rsTemp!交易流水号)
                .TextMatrix(lngRow, .ColIndex("交易说明")) = Nvl(rsTemp!交易说明)
                lngRow = lngRow + 1
                rsTemp.MoveNext
            Loop
            .AutoSizeMode = flexAutoSizeColWidth
            Call .AutoSize(0, .Cols - 1)
            zl_vsGrid_Para_Restore mlngMode, vsBalance, Me.Name, "结算信息列表", False
        End With
        '票据使用情况
        '票种:1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
        Call ShowFactList(byt票种, strNO, blnNOMoved)
    Case 8, 9, 10, 11, 12
        ' 8,'消费卡充值',9,'消费卡面额',10,'借款',11,'借出',12,'暂存'
        With vsBalance
            .Clear 1: .Rows = 2
            lngRow = 1
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("结算方式")))
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("金额合计"))), "###0.00;-###0.00; ;")
            .TextMatrix(lngRow, .ColIndex("结算号码")) = ""
            .TextMatrix(lngRow, .ColIndex("卡号")) = ""
            .TextMatrix(lngRow, .ColIndex("交易流水号")) = ""
            .TextMatrix(lngRow, .ColIndex("交易说明")) = ""
            .AutoSizeMode = flexAutoSizeColWidth
            Call .AutoSize(0, .Cols - 1)
            zl_vsGrid_Para_Restore mlngMode, vsBalance, Me.Name, "结算信息列表", False
        End With
        vsBillList.Clear 1: vsBillList.Rows = 2
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
GoClear:
    With vsList
        vsBalance.Clear 1: vsBalance.Rows = 2
        vsBillList.Clear 1: vsBillList.Rows = 2
    End With
 End Sub
 
Private Sub ShowFactList(ByVal byt票种 As Byte, ByVal strNO As String, _
    ByVal blnNOMoved As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示发票信息
    '入参:strNO-单据号
    '       byt票种-票种(1-收费,2-预交,3-结帐,4-挂号,5-就诊卡)
    '       blnNOMoved-是否历史表空间
    '编制:刘兴洪
    '日期:2013-09-16 15:14:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim blnIsHaveData As Boolean, lngRow As Long
    
    On Error GoTo errH
    
    blnIsHaveData = False
    If byt票种 = 1 And mbyt票据分配规则 <> 0 Then
        '票种,按新票据格式处理
        strSQL = _
        " Select distinct B.ID,B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票发出') as 使用原因," & _
        " To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
        " From " & IIf(blnNOMoved, "H", "") & "票据打印明细 A," & _
                IIf(blnNOMoved, "H", "") & "票据使用明细 B " & _
        " Where A.票种=1 And A.票号=B.号码 " & _
        "             And B.票种=1 And A.NO=[1]" & _
        " Order by ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTemp.RecordCount <> 0 Then GoTo GoGridData
    End If
    
    If byt票种 = 3 Then
        strSQL = _
        "   Select B.ID, B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票发出') as 使用原因," & _
        "           To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
        "   From " & IIf(blnNOMoved, "H", "") & "票据打印内容 A," & _
                    IIf(blnNOMoved, "H", "") & "票据使用明细 B" & _
        " Where A.数据性质=[2]  And A.ID=B.打印ID" & _
        "           And B.票种 In (1,3)  And A.NO=[1]" & _
        " Order by ID"
    Else
        If byt票种 = 4 Then
            strSQL = _
            "   Select B.ID, B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票发出') as 使用原因," & _
            "           To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
            "   From " & IIf(blnNOMoved, "H", "") & "票据打印内容 A," & _
                        IIf(blnNOMoved, "H", "") & "票据使用明细 B" & _
            " Where A.数据性质=[2]  And A.ID=B.打印ID" & _
            "           And B.票种 In (1,4)  And A.NO=[1]" & _
            " Order by ID"
        '110414:李南春，2017/6/20，医疗卡使用门诊发票
        ElseIf byt票种 = 5 Then
            strSQL = _
            "   Select B.ID, B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票发出') as 使用原因," & _
            "           To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
            "   From " & IIf(blnNOMoved, "H", "") & "票据打印内容 A," & _
                        IIf(blnNOMoved, "H", "") & "票据使用明细 B" & _
            " Where A.数据性质=[2]  And A.ID=B.打印ID" & _
            "           And B.票种 In (1,5)  And A.NO=[1]" & _
            " Order by ID"
        Else
            strSQL = _
            "   Select B.ID, B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票发出') as 使用原因," & _
            "           To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
            "   From " & IIf(blnNOMoved, "H", "") & "票据打印内容 A," & _
                        IIf(blnNOMoved, "H", "") & "票据使用明细 B" & _
            " Where A.数据性质=[2]  And A.ID=B.打印ID" & _
            "           And B.票种=[2]  And A.NO=[1]" & _
            " Order by ID"
        End If
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, byt票种)
GoGridData:
    With vsBillList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2: lngRow = 1
        Do While Not rsTemp.EOF
            'strHead = "票据号,使用原因,使用时间,使用人"
            .TextMatrix(lngRow, .ColIndex("票据号")) = Nvl(rsTemp!票据号)
            .TextMatrix(lngRow, .ColIndex("使用原因")) = Nvl(rsTemp!使用原因)
            .TextMatrix(lngRow, .ColIndex("使用时间")) = Format(rsTemp!使用时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("使用人")) = Nvl(rsTemp!使用人)
            .Rows = .Rows + 1: lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
    End With
    '恢复列设置
    zl_vsGrid_Para_Restore mlngMode, vsBillList, Me.Name, "票据明细列表", False
    vsBillList.Redraw = flexRDBuffered
    Exit Sub
errH:
    vsBillList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadFeeData(Optional ByVal blnRollingData As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载费目信息
    '入参:blnRollingData-是否获取轧帐数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-03-05 14:03:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String
    Dim strWithTable As String, strWhere As String
    Dim strRollingType As String, i As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If blnRollingData Then
        '预交款填NULL,2-结帐,3-收费,4-挂号,5-就诊卡,6-补充医保结算
        'mstrRollingType:轧帐类别(0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡)
        '收缴对照:1-收费(含挂号),2-结帐,3-预交,4-借款;5-消费卡充值;6--消费卡面值;7-暂存金(本次增加)；８－收款或轧帐作废对照（本次增加）；9-二次结算记录
        If InStr(",2,6,", "," & mstrRollingType & ",") > 0 Then
            '预交款和消费卡无收据费目
            vsFeeList.Clear 1: vsFeeList.Rows = 2
            vsFeeList.Row = 1: vsFeeList.Col = 0
            LoadFeeData = True
            Exit Function
        End If
        
        strWhere = " And instr([5],','||A.结算性质||',')>0 "
        If mstrRollingType <> "" And InStr("," & mstrRollingType & ",", ",0,") = 0 Then
            If Get轧帐结算性质(mstrRollingType, strRollingType) = False Then
                vsFeeList.Clear 1: vsFeeList.Rows = 2
                vsFeeList.Row = 1: vsFeeList.Col = 0
                LoadFeeData = True
                Exit Function
            End If
        Else
            strRollingType = ",2,3,4,5,"
        End If
        strWithTable = "" & _
        "   Select distinct A.结算性质, " & vbCrLf & _
        "          Decode(nvl(A.结算性质,0),2,2,3,1,4,1,5,1,0) as 性质,A.结帐ID as 记录ID" & vbCrLf & _
        "   From 病人预交记录 A " & vbCrLf & _
        "   Where a.操作员姓名 || '' = [2]  and a.记录性质<>1" & strWhere & vbCrLf & _
        "       And A.收款时间 Between  [3] And [4]  " & vbCrLf & _
        "       And Not Exists(Select 1 From 门诊费用记录 B Where a.结帐id = b.结帐id And Nvl(b.费用状态, 0) = 1) " & vbNewLine & _
        "       And Not Exists(Select 1 From 病人结帐记录 B Where a.结帐id = b.Id And b.结算状态 Is Not Null)" & vbNewLine & _
        "       And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y " & _
        "                       Where y.记录id = a.结帐id" & _
        "                               And x.Id = y.收缴ID And Decode(nvl(A.结算性质,0),2,2,3,1,4,1,5,1,0)=y.性质  " & _
        "                               And x.作废时间 Is Null) " & vbCrLf
        '不能包含补结算的,这样会造成重复统计且统计的金额也不正确
'        If mstrRollingType = 1 Or mstrRollingType = 0 Then  '收费和所有类别,包含二次补结算数据
'            strWithTable = strWithTable & _
'            "   Union ALL " & _
'            "   Select distinct A.结算性质,9 as 性质,B.收费结帐ID   as 记录ID" & vbCrLf & _
'            "   From 病人预交记录 A,费用补充记录 B " & vbCrLf & _
'            "   Where a.结帐ID=B.结算ID And nvl(B.费用状态,0)=0 And A.结算性质=6 " & vbCrLf & _
'            "       And a.操作员姓名 || '' = [2] " & vbCrLf & _
'            "       And A.收款时间 Between  [3] And [4]  " & vbCrLf & _
'            "       And Not Exists (Select y.记录id From 人员收缴记录 X, 人员收缴对照 Y Where y.记录id = a.结帐id  And x.Id = y.收缴ID And y.性质 = 1 And x.作废时间 Is Null) " & vbCrLf
'        End If
        strWithTable = "With c_对照信息 As  ( " & strWithTable & " )" & vbCrLf
    Else
        strTable = ""
        If mblnDel Or mbytType = EM_收费员轧帐 Then
            'And mblnDel = False
            If mbytType = EM_收费员轧帐 Then
                strWhere = " And  A.ID =J.Column_Value And a.记录性质=1"
            Else
                strWhere = " And  A.ID=C.记录ID And C.性质=8 And C.收缴ID=J.Column_Value And a.记录性质=1"
                strTable = ",人员收缴对照 C"
            End If
        Else
            If mbytType = EM_小组收款 Then
                strWhere = " And  A.小组收款ID =J.Column_Value And a.记录性质=1"
            ElseIf mbytType = EM_小组轧帐 Then
                strWhere = " And  A.小组轧账ID =J.Column_Value And a.记录性质=1"
            Else
                strWhere = " And  A.财务收款ID =J.Column_Value And a.记录性质=1"
            End If
        End If
         '人员收缴对照.性质:1-收费(含挂号),2-结帐,3-预交,4-借款;5-消费卡充值;6--消费卡面值;7-暂存金(本次增加)；８－收款或轧帐作废对照（本次增加）；9-二次结算记录
        strWithTable = "" & _
        "   With c_对照信息 As  ( " & _
        "           Select /*+cardinality(j,10)*/ b.收缴id, b.性质, b.记录id" & _
        "           From 人员收缴记录 A,人员收缴对照 B,Table( f_Num2list([1])) J " & strTable & _
        "           Where a.Id = b.收缴id And b.性质 in (1,2)" & strWhere & ")" & vbCrLf
        
        '不能包含补结算的,这样会造成重复统计且统计的金额也不正确
'        strWithTable = strWithTable & _
'        "           Union ALL " & vbCrLf & _
'        "           Select b.收缴id, b.性质,b1.收费结帐ID" & vbCrLf & _
'        "           From 人员收缴记录 A, 人员收缴对照 B,费用补充记录 B1, " & vbCrLf & _
'        "                Table( f_Num2list([1])) J " & strTable & vbCrLf & _
'        "           Where a.Id = b.收缴id And b.性质=9 And B.记录ID=b1.结算ID " & strWhere & ") "
    End If
    
    strSQL = strWithTable & vbCrLf & _
    "   Select A.收据费目,sum(A.结帐金额) AS 结帐金额" & _
    "   From 住院费用记录 A,c_对照信息 Q1 " & _
    "   Where A.结帐ID=Q1.记录ID and Q1.性质 in (1,2,9) " & _
    "   GROUP BY A.收据费目 " & _
    "   Union ALL  " & _
    "   Select A.收据费目,sum(A.结帐金额) AS 结帐金额" & _
    "   From 门诊费用记录 A,c_对照信息 Q1 " & _
    "   Where  A.结帐ID=Q1.记录ID  and Q1.性质 in (1,2,9)" & _
    "   GROUP BY A.收据费目 "
    
    strSQL = "" & _
    "   SELECT A.收据费目,ltrim(to_char(sum(A.结帐金额),'99999990.00')) AS 结帐金额 " & _
    "   FROM ( " & strSQL & " ) a " & _
    "   GROUP BY A.收据费目 " & _
    "   ORDER BY 收据费目"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID, mstrPersonName, mdtStartDate, mdtendDate, strRollingType)
    mblnNotBrush = True
    With vsFeeList
        .Clear 1
        .Rows = 2
        .FixedCols = 0: .FixedRows = 1
        If Not rsTemp.EOF Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            Select Case .ColKey(i)
            Case "结帐金额"
                  .ColAlignment(i) = flexAlignRightCenter
            Case Else
                  .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .SubtotalPosition = flexSTBelow
        If rsTemp.RecordCount <> 0 Then
            .Subtotal flexSTSum, -1, .ColIndex("结帐金额"), "#######0.00", &HFFC0C0, vbBlack, True, "合计"
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    zl_vsGrid_Para_Restore mlngMode, vsFeeList, Me.Name, "收据费目列表", False
    If mblnFirst Then Call reSetFeeListPancelWidth
    
    mblnNotBrush = False
    LoadFeeData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function ShowOnlyFactList() As Boolean
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim strStartDate As String, strEndDate As String
    
    On Error GoTo errH
    strSQL = "Select 开始时间,终止时间 From 人员收缴记录 Where ID= [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID)
    If rsTemp.RecordCount <> 0 Then
        strStartDate = Nvl(rsTemp!开始时间)
        strEndDate = Nvl(rsTemp!终止时间)
    Else
        Exit Function
    End If
    
    If strStartDate = "" Or strEndDate = "" Then Exit Function
    
    strSQL = "Select Distinct a.Id, a.号码 As 票据号, Decode(a.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回',6,'红票发出') As 使用原因," & vbNewLine & _
            "                To_Char(a.使用时间, 'MM-DD HH24:MI') As 使用时间, a.使用人" & vbNewLine & _
            "From 票据使用明细 A, (Select 票种, 性质, 开始票号, 终止票号 From 人员收缴票据 Where 收缴id = [1]) B" & vbNewLine & _
            "Where a.票种 = b.票种 And a.号码 Between b.开始票号 And b.终止票号 And a.性质 = Decode(b.性质, 1, 1, 2, 2, 3, 2)" & vbNewLine & _
            "      And a.使用时间 Between To_date([2],'yyyy-mm-dd hh24:mi:ss') And To_date([3],'yyyy-mm-dd hh24:mi:ss') " & _
            "Order By ID"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID, strStartDate, strEndDate)

    With vsBillList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2: lngRow = 1
        Do While Not rsTemp.EOF
            ShowOnlyFactList = True
            'strHead = "票据号,使用原因,使用时间,使用人"
            .TextMatrix(lngRow, .ColIndex("票据号")) = Nvl(rsTemp!票据号)
            .TextMatrix(lngRow, .ColIndex("使用原因")) = Nvl(rsTemp!使用原因)
            .TextMatrix(lngRow, .ColIndex("使用时间")) = Format(rsTemp!使用时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("使用人")) = Nvl(rsTemp!使用人)
            .Rows = .Rows + 1: lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
    End With
    '恢复列设置
    zl_vsGrid_Para_Restore mlngMode, vsBillList, Me.Name, "票据明细列表", False
    vsBillList.Redraw = flexRDBuffered
    Exit Function
errH:
    vsBillList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadList() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示收款及票据汇总信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-04 11:28:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWithTable As String, strWhere As String, i As Long
    Dim strTable As String
    strTable = ""
    If mblnDel Or mbytType = EM_收费员轧帐 Or mbytType = EM_财务收款_非收费员 Then
        If mbytType = EM_收费员轧帐 Or mbytType = EM_财务收款_非收费员 Then
            strWhere = " And  A.ID =J.Column_Value And a.记录性质=1"
        Else
            strWhere = " And  A.ID=C.记录ID And C.性质=8 And C.收缴ID=J.Column_Value And a.记录性质=1"
            strTable = ",人员收缴对照 C"
        End If
    Else
        If mbytType = EM_小组收款 Then
            strWhere = " And  A.小组收款ID =J.Column_Value And a.记录性质=1"
        ElseIf mbytType = EM_小组轧帐 Then
            strWhere = " And  A.小组轧账ID =J.Column_Value And a.记录性质=1"
        Else
            strWhere = " And  A.财务收款ID =J.Column_Value And a.记录性质=1"
        End If
    End If
    
    strWithTable = "" & _
    "   With c_对照信息 As  ( " & _
    "           Select /*+cardinality(j,10)*/ b.收缴id, b.性质, b.记录id  " & _
    "           From 人员收缴记录 A, 人员收缴对照 B,Table( f_Num2list([1])) J " & strTable & _
    "           Where a.Id = b.收缴id  " & strWhere & ") "
    
    strSQL = "" & _
         "   Select 5 AS 记录性质,A.结帐ID as 记录ID,A.NO,A.记录状态,max(A.姓名) AS 病人姓名,max(A.性别) AS 性别,max(A.年龄) AS 年龄,max(decode(A.门诊标志,2,-1*NULL,A.标识号)) AS 门诊号,max(decode(A.门诊标志,2,标识号,null)) AS 住院号, " & _
         "     Max(A.操作员姓名) AS 操作员姓名, NULL AS 结算方式, " & _
         "     sum(结帐金额) AS 金额合计,Max(登记时间) AS 收费时间 " & _
         "   From 住院费用记录 A,c_对照信息 Q1 " & _
         "   Where Nvl(A.记帐费用, 0) = 0 And A.记录状态 <> 0 AND a.记录性质=5 And A.结帐ID=Q1.记录ID and Q1.性质=1 " & _
         "   GROUP BY A.NO,A.结帐ID,A.记录状态 " & _
         "   Union ALL  " & _
         "  Select 记录性质, Min(记录id) As 记录id, NO, Max(记录状态) As 记录状态, 病人姓名, 性别, 年龄, 门诊号, 住院号, 操作员姓名, 结算方式, Sum(金额合计) As 金额合计, 收费时间 From (" & _
         "   Select  mod(a.记录性质,10) as 记录性质,A.结帐ID as 记录ID,A.NO,a.记录状态,max(A.姓名) AS 病人姓名,max(A.性别) AS 性别,max(A.年龄) AS 年龄,max(decode(A.门诊标志,2,-1*NULL,A.标识号)) AS 门诊号,max(decode(A.门诊标志,2,标识号,null)) AS 住院号, " & _
         "     Max(A.操作员姓名) AS 操作员姓名, NULL AS 结算方式, " & _
         "     sum(结帐金额) AS 金额合计,Max(登记时间) AS 收费时间 " & _
         "   From 门诊费用记录 A,c_对照信息 Q1 " & _
         "   Where (MOD(A.记录性质,10)=1 OR A.记录性质=4) AND Nvl(A.记帐费用, 0) = 0 And A.记录状态 <> 0 And Nvl(A.费用状态, 0) <> 1  " & _
         "        And A.结帐ID=Q1.记录ID and Q1.性质=1 " & _
         "   GROUP BY mod(a.记录性质,10),A.结帐ID,A.NO,a.记录状态) " & _
         "  Group By 记录性质, NO, 病人姓名, 性别, 年龄, 门诊号, 住院号, 操作员姓名, 结算方式, 收费时间 "
         
         strSQL = strSQL & _
         "   Union All " & _
         "   Select 6 AS 记录性质,A.ID as 记录ID,A.NO,A.记录状态,max(C.姓名) AS 病人姓名,max(C.性别) AS 性别,max(C.年龄) AS 年龄, " & _
         "     max(c.门诊号) AS 门诊号,max(C.住院号) AS 住院号, " & _
         "     Max(A.操作员姓名) AS 操作员姓名,NULL AS 结算方式, " & _
         "     sum(b.冲预交) AS 金额合计,Max(A.收费时间) AS 收费时间  " & _
         "   From 病人结帐记录 A,病人预交记录 B,病人信息 C,c_对照信息 Q1" & _
        "   Where A.ID=B.结帐ID And A.结算状态 Is Null and A.病人ID=C.病人ID(+)  " & _
         "         And A.ID=Q1.记录ID and Q1.性质=2  " & _
         "   group by A.NO,A.ID,A.记录状态 "
        
        strSQL = strSQL & _
         "   Union All " & _
         "   Select 7 As 记录性质,A.ID as 记录ID, a.No, a.记录状态, Max(nvl(M.姓名,c.姓名)) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, Max(c.门诊号) As 门诊号, " & _
         "      Max(Decode(m.住院号, Null, c.住院号, m.住院号)) As 住院号, Max(a.操作员姓名) As 操作员姓名,Max(结算方式) AS 结算方式, Sum(a.金额) As 金额合计, Max(a.收款时间) As 收费时间 " & _
         "   From 病人预交记录 A, 病人信息 C, 病案主页 M,c_对照信息 Q1 " & _
         "   Where a.病人id = c.病人id(+) And a.病人id = m.病人id(+) And a.主页id = m.主页id(+)  " & _
         "     And a.记录性质 = 1  And A.ID=Q1.记录ID and Q1.性质=3  " & _
         "   Group By a.No,A.ID, a.记录状态 " & _
         "   Union All " & _
         "   Select 8 as 记录性质,A.结算ID as 记录ID,a.卡号 As no,a.记录状态,'' as 病人姓名,'' as 性别,'' as 年龄,null 门诊号,NULL as 住院号, " & _
         "        A.操作员姓名,a.结算方式, a.实收金额  As 金额, A.登记时间 AS 收费时间  " & _
         "   From 病人卡结算记录 A,c_对照信息 Q1 " & _
         "   Where a.记录性质 In (2, 3) And A.结算ID=Q1.记录ID and Q1.性质=5   " & _
         "   Union All " & _
         "   Select 9 As 记录性质,A.结算ID as 记录ID, a.卡号 As NO, a.记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, " & _
         "       Null As 住院号, a.操作员姓名, a.结算方式, a.实收金额 As 金额, a.登记时间 " & _
         "   From 病人卡结算记录 A,c_对照信息 Q1 " & _
         "   Where a.记录性质 In (1, 3) And A.结算ID=Q1.记录ID and Q1.性质=6    "
         
        strSQL = strSQL & _
        "   Union All" & _
        "   Select 10 As 记录性质,A.ID as 记录ID, ltrIm(to_Char(a.ID)) As NO, 1 As 记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, " & _
        "     Null As 住院号, a.借款人 As 操作员姓名, a.结算方式, a.借款金额 As 金额, a.借出时间 " & _
        "   From 人员借款记录 A,c_对照信息 Q1,人员收缴记录 M " & _
        "   Where A.ID=Q1.记录ID and Q1.性质=4 and Q1.收缴ID=M.ID and M.收款员||''=a.借款人 " & _
        "   Union All " & _
        "   Select 11 As 记录性质, A.ID as 记录ID,ltrIm( to_Char(a.ID)) As NO, 1 As 记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, " & _
        "     Null As 住院号, a.借出人 As 操作员姓名, a.结算方式, a.借款金额 As 金额, a.借出时间 " & _
        "   From 人员借款记录 A,c_对照信息 Q1,人员收缴记录 M " & _
        "   Where   A.ID=Q1.记录ID and Q1.性质=4 and Q1.收缴ID=M.ID and M.收款员||''=a.借出人  " & _
        "   Union All " & _
        "   Select 12 As 记录性质, A.ID as 记录ID,a.No As NO, Decode(收回时间, Null, 1, 2) As 记录状态, '' As 病人姓名, '' As 性别, '' As 年龄, Null 门诊号, Null As 住院号, " & _
        "      a.登记人 As 操作员姓名, '现金' As 结算方式, a.金额, a.登记时间 " & _
        "   From 人员暂存记录 A,c_对照信息 Q1 " & _
        "   Where A.ID=Q1.记录ID and Q1.性质=7 and a.记录性质=2 And a.收回时间 Is Null"
        
        strSQL = strSQL & _
        "   Union ALL" & vbCrLf & _
        "   Select 13 as 记录性质, a.结算id As 记录id, a.No, a.记录状态, Max(c.姓名) As 病人姓名, Max(c.性别) As 性别, Max(c.年龄) As 年龄, c.门诊号 As 门诊号," & vbNewLine & _
        "        c.住院号 As 住院号, Max(a.操作员姓名) As 操作员姓名, Null As 结算方式, Sum(M.冲预交) As 金额合计, Max(a.登记时间) As 收费时间" & vbNewLine & _
        "   From (Select a.结算ID,A.NO,A.记录状态,Max(a.病人ID) as 病人ID,max(a.操作员姓名) as 操作员姓名,max(a.登记时间) as 登记时间 " & _
        "         From 费用补充记录 A, c_对照信息 B " & _
        "         Where A.结算ID=B.记录ID and B.性质=9  " & _
        "         Group by a.结算ID,A.NO,A.记录状态) A,病人预交记录 M,病人信息 C" & vbNewLine & _
        "   Where  A.结算ID=M.结帐ID And A.病人ID=C.病人ID " & _
        "   Group By  a.结算id, a.No, a.记录状态, c.门诊号, c.住院号"
        
    
    strSQL = strWithTable & vbCrLf & strSQL
    strSQL = "" & _
    "   SELECT /*+ rule */ 记录性质 as 性质,记录ID,decode(记录性质,1,'收费',4,'挂号',5,'就诊卡',6,'结帐',7,'预交款',8,'消费卡充值',9,'消费卡面额',10,'借款',11,'借出',12,'暂存',13,'补结算','') AS 类别, " & _
    "       NO,记录状态,病人姓名,性别,年龄,门诊号,住院号,操作员姓名,结算方式," & _
    "       Trim(to_char(金额合计,'99999999990.00')) As 金额合计,to_char(收费时间,'yyyy-mm-dd hh24:mi:ss') as 收费时间 " & _
    "   FROM ( " & strSQL & " ) " & _
    "   ORDER BY 记录性质,收费时间 DESC,NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID)
    mblnNotBrush = True
    With vsList
        .Clear 1
        .Rows = 2
        .FixedCols = 0: .FixedRows = 1
        If Not rsTemp.EOF Then
            Set .DataSource = rsTemp
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            Select Case .ColKey(i)
            Case "性质", "记录状态", "记录ID"
                  .ColWidth(i) = 0: .ColHidden(i) = True
            Case "NO", "类别", "性别", "年龄", "门诊号", "住院号", "结算方式", "收费时间"
                  .ColAlignment(i) = flexAlignCenterCenter
            Case "金额合计"
                  .ColAlignment(i) = flexAlignRightCenter
            Case Else
                  .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        If rsTemp.RecordCount <> 0 Then
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, .ColIndex("金额合计"), "#######0.00", &HFFC0C0, vbBlack, True, "合计"
            For i = 0 To .ColIndex("金额合计") - 1
                .TextMatrix(.Rows - 1, i) = "合计"
            Next
            .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
            .MergeRow(.Rows - 1) = True
            .MergeCells = flexMergeRestrictRows
            '问题号:110535,焦博,2017/09/04,用颜色区分退费记录和被退费记录
             For i = 1 To .Rows - 1
                Select Case .TextMatrix(i, .ColIndex("记录状态"))
                Case 1
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                Case 2
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                Case 3
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                End Select
            Next
        End If
        
    End With
    zl_vsGrid_Para_Restore mlngMode, vsList, Me.Name, "明细信息列表", False
    mblnNotBrush = False
    Call LoadDetailData
    Call LoadFeeData '加载收据费目信息
    LoadList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
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
    Set vsList.Font = Me.Font
    Set vsBalance.Font = Me.Font
    Set vsBillList.Font = Me.Font
 End Sub
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化窗体
    '编制:刘兴洪
    '日期:2013-09-03 15:28:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    mbytFontSize = 9
    mstrCashBalance = "现金"
    Set rsTemp = Get结算方式
    rsTemp.Filter = "性质=1"
    If Not rsTemp.EOF Then mstrCashBalance = rsTemp!名称
    strTmp = Trim(zlDatabase.GetPara("票据分配规则", glngSys, 1121, "0||0;0;0;0;0"))
    mbyt票据分配规则 = Val(Split(strTmp & "||", "||")(0))
    stbThis.Panels(3).Text = UserInfo.姓名
    stbThis.Panels(3).ToolTipText = "当前操作员:" & UserInfo.姓名
    
    Call InitPanel
    Call InitGrid
    Call ReSetFontSize
 End Sub
Private Sub LoadFilterRange()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载条件范围
    '编制:刘兴洪
    '日期:2013-09-16 17:18:08
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNO As String, strTime As String
    lblNO.Visible = mstrChargeRollingID <> ""
    Select Case mbytType
     Case EM_收费员轧帐, EM_财务收款_非收费员
        If mstrChargeRollingID = "" Then
            lblRange.Caption = "时间范围:" & Format(mdtStartDate, "yyyy-mm-dd HH:MM:SS") & "至" & Format(mdtendDate, "yyyy-mm-dd HH:MM:SS")
            Exit Sub
        End If
        
         strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.开始时间,'yyyy-mm-dd hh24:mi:ss')||'至'||to_char(A.终止时间,'yyyy-mm-dd hh24:mi:ss') as 时间范围 " & _
        "   From 人员收缴记录 A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
     Case EM_小组收款
        strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 时间范围 " & _
        "   From 人员收缴记录 A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
     Case EM_小组轧帐
        strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.开始时间,'yyyy-mm-dd hh24:mi:ss')||'至'||to_char(A.终止时间,'yyyy-mm-dd hh24:mi:ss') as 时间范围 " & _
        "   From 人员收缴记录 A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
     Case EM_财务收款
        strSQL = "" & _
        "   Select /*+cardinality(j,10)*/ A.NO,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 时间范围 " & _
        "   From 人员收缴记录 A ,Table( f_Num2list([1])) J" & _
        "   Where A.ID= J.Column_Value" & _
        "   Order by A.NO "
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrChargeRollingID)
    With rsTemp
        strNO = "": strTime = ""
        Do While Not .EOF
            strNO = strNO & ";" & rsTemp!NO
            strTime = strTime & ";" & rsTemp!时间范围
            .MoveNext
        Loop
        lblNO.Caption = "单据号:"
        If strNO <> "" Then lblNO.Caption = lblNO.Caption & Mid(strNO, 2)
        lblRange.Caption = "时间范围:"
        If strTime <> "" Then lblRange.Caption = lblRange.Caption & Mid(strTime, 2)
    End With
End Sub
Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2013-09-16 16:47:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, objConPan As Pane
    Dim objFeePan As Pane
    Dim lngDetailHeight As Long '缺省明细高度
    Dim lngTemp As Long
    lngDetailHeight = 2925 / Screen.TwipsPerPixelX
    lngTemp = picConList.Height \ Screen.TwipsPerPixelY
    With dkpMan
        Set objConPan = .CreatePane(mPaneIndex.EM_PN_ConList, 400, 400, DockBottomOf, Nothing)
        objConPan.Title = "轧帐条件信息"
        objConPan.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        objConPan.Handle = picConList.hWnd
        objConPan.MaxTrackSize.Height = lngTemp
        objConPan.MinTrackSize.Height = lngTemp
        objConPan.Tag = mPaneIndex.EM_PN_ConList
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_LIST, 600, 400, DockBottomOf, objConPan)
        If mbytType = EM_收费员轧帐 Then
            objPane.Title = "收费员轧帐明细"
        ElseIf mbytType = EM_小组轧帐 Then
            objPane.Title = "小组轧帐明细"
        ElseIf mbytType = EM_小组收款 Then
            objPane.Title = "小组收款明细"
        Else
            objPane.Title = "财务收款明细"
        End If
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        objConPan.Tag = mPaneIndex.EM_PN_LIST
        
        Set objFeePan = .CreatePane(mPaneIndex.EM_PN_FeeLIST, 160, 400, DockRightOf, objPane)
        objFeePan.Title = "费目信息"
        objFeePan.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objFeePan.Handle = picFeeList.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_BALANCE, 400, lngDetailHeight, DockBottomOf, objPane)
        objPane.MinTrackSize.Height = lngDetailHeight
        objPane.Title = "收款信息"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_BILL, 400, 400, DockRightOf, objPane)
        objPane.Title = "票据信息"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBillList.hWnd
        
        
        .SetCommandBars Me.cbsThis
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
    Dim i As Long, strHead As String, strArr As Variant
    '收款汇总信息
    strHead = "性质,记录ID,类别,NO,记录状态,病人姓名,性别,性别,年龄,门诊号,住院号,操作员姓名,结算方式,金额合计,收费时间"
    strArr = Split(strHead, ",")
    With vsList
        Set .Font = Me.Font
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "性质", "记录状态", "记录ID"
                .ColWidth(i) = 0: .ColHidden(i) = True
          Case "NO", "类别", "性别", "年龄", "门诊号", "住院号", "结算方式", "收费时间"
                .ColAlignment(i) = flexAlignCenterCenter
          Case "金额合计"
                .ColAlignment(i) = flexAlignRightCenter
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsList, Me.Name, "明细信息列表", False
    End With
    
    
    '结算信息信息
    strHead = "结算方式,金额,结算号码,卡号,交易流水号,交易说明"
    strArr = Split(strHead, ",")
    With vsBalance
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "结算方式", "卡号"
                .ColAlignment(i) = flexAlignCenterCenter
          Case "金额"
                .ColAlignment(i) = flexAlignRightCenter
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsBalance, Me.Name, "结算信息列表", False
    End With
   '票据信息
   strHead = "票据号,使用原因,使用时间,使用人"
    strArr = Split(strHead, ",")
    With vsBillList
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "使用时间", "使用人"
                .ColAlignment(i) = flexAlignCenterCenter
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsBillList, Me.Name, "票据明细列表", False
    End With
    
   '票据信息
   strHead = "收据费目,结帐金额"
    strArr = Split(strHead, ",")
    With vsFeeList
        .Cols = UBound(strArr) + 1: .Rows = 2
        .FixedRows = 1
        For i = 0 To UBound(strArr)
            .TextMatrix(0, i) = strArr(i): .ColKey(i) = UCase(strArr(i))
          Select Case .ColKey(i)
          Case "实收金额"
                .ColAlignment(i) = flexAlignRightBottom
          Case Else
                .ColAlignment(i) = flexAlignLeftCenter
          End Select
          .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        '.ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsFeeList, Me.Name, "收据费目列表", False
    End With
 End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    End Select
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    reSetFeeListPancelWidth True
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
    
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case EM_PN_ConList  '条件
        Item.Handle = picConList.hWnd
    Case EM_PN_LIST '收款汇总
        Item.Handle = picList.hWnd
    Case EM_PN_BALANCE  '结算方式
        Item.Handle = picBalance.hWnd
    Case EM_PN_BILL  '退费票据
        Item.Handle = picBillList.hWnd
    Case EM_PN_FeeLIST   '费用汇总
        Item.Handle = picFeeList.hWnd
    End Select
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
 Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.Height = .ScaleHeight
        vsBalance.Width = .ScaleWidth
    End With
End Sub

 

Private Sub picConList_Resize()
    Err = 0: On Error Resume Next
    With picConList
        fraSplit.Top = .ScaleTop
        fraSplit.Left = .ScaleLeft
        fraSplit.Width = .ScaleWidth
        lblNO.Top = .ScaleHeight - lblNO.Height - 50
        lblRange.Top = lblNO.Top
        lblNO.Left = .ScaleLeft + 50
        lblRange.Left = IIf(mstrChargeRollingID <> "", lblNO.Left + lblNO.Width * 2 + 50, lblNO.Left)
    End With
End Sub
Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        vsFeeList.Left = .ScaleLeft
        vsFeeList.Top = .ScaleTop
        vsFeeList.Width = .ScaleWidth
        vsFeeList.Height = .ScaleHeight - vsFeeList.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        vsList.Top = .ScaleTop
        vsList.Left = .ScaleLeft
        vsList.Height = .ScaleHeight - vsList.Top
        vsList.Width = .ScaleWidth
    End With
End Sub
 
Private Sub picBillList_Resize()
    Err = 0: On Error Resume Next
    With picBillList
        vsBillList.Left = .ScaleLeft
        vsBillList.Top = .ScaleTop
        vsBillList.Height = .ScaleHeight
        vsBillList.Width = .ScaleWidth
    End With
End Sub

Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsFeeList, Me.Name, "收据费目列表", False
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsFeeList, Me.Name, "收据费目列表", False
   
End Sub
Private Sub vsList_GotFocus()
    Call zl_VsGridGotFocus(vsList)
End Sub
Private Sub vsList_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsList)
End Sub
Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsList, Me.Name, "明细信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsList, OldRow, NewRow, OldCol, NewCol)
    If OldRow = NewRow Or mblnNotBrush Then Exit Sub
    Call LoadDetailData
End Sub
Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsList, Me.Name, "明细信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
 
Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsBalance, Me.Name, "结算信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsBalance, Me.Name, "结算信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBillList_GotFocus()
    Call zl_VsGridGotFocus(vsBillList)
End Sub
Private Sub vsBillList_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBillList)
End Sub
Private Sub vsBillList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsBillList, Me.Name, "票据明细列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBillList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBillList, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBillList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsBillList, Me.Name, "票据明细列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim blnFeeList As Boolean
    
    Err = 0: On Error GoTo ErrHand:
    '输出轧帐信息
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    If Me.ActiveControl Is vsFeeList Then
        objPrint.Title.Text = gstr单位名称 & "费目汇总表"
    Else
        objPrint.Title.Text = gstr单位名称 & "收款清册"
    End If
    Set objRow = New zlTabAppRow
    If lblNO.Visible Then
        objRow.Add "" & lblNO.Caption
    End If
    If lblRange.Visible Then
        objRow.Add "" & lblRange.Caption
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = IIf(Me.ActiveControl Is vsFeeList, vsFeeList, vsList)
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
Public Function zlDefCommandBars() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-16 16:56:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    For Each mcbrControl In mcbrToolBar.Controls
          If mcbrControl.ID <> conMenu_COMBOX_INTERFACE Then
            mcbrControl.Style = xtpButtonIconAndCaption
          End If
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
     Select Case Control.ID
        Case conMenu_File_Exit: Unload Me: '退出(&X)
        Case conMenu_File_PrintSet: Call zlPrintSet '打印设置
        Case conMenu_File_Preview: Call zlPrint(2)  '预览(&V)
        Case conMenu_File_Print: Call zlPrint(1) '打印(&P)
        Case conMenu_File_Excel: Call zlPrint(3)  '输出到&Excel…
        Case conMenu_View_Refresh: zlRefresh '刷新(&R)
        Case conMenu_View_StatusBar '状态栏(&S)
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub zlRefresh()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '编制:刘兴洪
    '日期:2013-09-16 17:08:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ReadListData
End Sub
Private Sub reSetFeeListPancelWidth(Optional blnSetMaxWidth As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置费用的宽度
    '入参:blnSetMaxWidth-设置最大宽度
    '返回:返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-03-06 12:07:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim objPan As Pane
    Set objPan = dkpMan.FindPane(mPaneIndex.EM_PN_FeeLIST)
    If objPan Is Nothing Then Exit Sub
    
    If blnSetMaxWidth Then
        dkpMan.RecalcLayout
        sngWidth = (Me.ScaleWidth \ Screen.TwipsPerPixelY) * Round(1 / 3, 4)
        If sngWidth < 200 Then sngWidth = 200
        objPan.MaxTrackSize.Width = sngWidth
       ' dkpMan.RecalcLayout
        Exit Sub
    End If
    
    sngWidth = GetFeeListMaxWidth \ Screen.TwipsPerPixelY
    objPan.MaxTrackSize.Width = sngWidth
    dkpMan.RecalcLayout
End Sub
Private Function GetFeeListMaxWidth() As Single
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费用列表的最大宽度
    '返回:返回最大宽度
    '编制:刘兴洪
    '日期:2015-03-06 11:47:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, i As Long
    
    With vsFeeList
        sngWidth = 0
        For i = 0 To .Cols - 1
            sngWidth = sngWidth + .ColWidth(i) + 70
        Next
    End With
    GetFeeListMaxWidth = sngWidth
End Function

