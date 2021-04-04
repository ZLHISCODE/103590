VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDepositNormal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picDeposit 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1815
      ScaleWidth      =   4335
      TabIndex        =   3
      Top             =   240
      Width           =   4335
      Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
         Height          =   1305
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
   Begin VB.PictureBox picBalanceInfo 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   5520
      ScaleHeight     =   2055
      ScaleWidth      =   2295
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
      Begin VSFlex8Ctl.VSFlexGrid vsBalanceInfor 
         Height          =   1305
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
   Begin VB.PictureBox picInvoice 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   2880
      ScaleHeight     =   2055
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
      Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
         Height          =   1305
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
   Begin VB.PictureBox picEinvoice 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   240
      ScaleHeight     =   2055
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsEInvoice 
         Height          =   1305
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2055
         _cx             =   3625
         _cy             =   2302
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   8760
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDepositNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mblnDateMoved As Boolean, mblnNOMoved As Boolean
Private mfrmMain As Object, mlngModule As Long
Private mstrPrivs As String, mstrEInvoicePrivs As String
Private mobjEInvoice As clsEinvoice  '电子票据部件
Private mblnNotRefresh As Boolean  '不刷新数据
Private mrsList As ADODB.Recordset  '单据列表
Private mlngTopRow As Long, mlngCurRow As Long '记录上次选择的行
Private mint类别 As Integer '0-查询所有;1-查询门诊预交;2-查询住院预交;3-查询押金记录
Public mblnGo As Boolean
Private mlngGo As Long
Private mcllFilter As Collection '过滤条件集合
Public Event SelectDeposit(ByVal blnUse As Boolean, ByVal bln押金 As Boolean, ByVal lng预交ID As Long, ByVal int记录状态 As Integer, _
                                         ByVal bln电子票据 As Boolean, ByVal lng电子票据ID As Long, ByVal bln换开 As Boolean, ByVal str票据号 As String, _
                                         ByVal bln退预交 As Boolean, ByVal int附加标志 As Integer, ByVal lng原始ID As Long)  '选择预交列表某一行
Public Event ShowStatus(ByVal strMessage As String)     '显示状态栏
Public Event FilterDeposit()                                             '弹出过滤窗口
Public Event PopupMenu()                                             '弹出菜单
Public Event ViewGo()                                                    '定位
Public Event MoneyEnum()                                             '现金点钞
Public Event RollingCurtain()                                           '收费轧帐
Public Event FileLocalSet()                                               '本机参数设置
Public Event FilePrint()                                                    '打印
Public Event EditDeposit()                                               '缴预交
Public Event EidtBalanceDel()                                          '余额退款
Public Event ReadPati(ByVal strName As String)              '按病人id过滤后加载病人姓名

Public Sub zlInit(ByVal frmMain As Object, ByVal objEInvoice As clsEinvoice, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strEInvoicePrivs As String, _
                         ByVal int类别 As Integer, ByVal cllFilter As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:objEInvoice-电子票据处理件
    '     strPrivs-预交权限串
    '     strEInvoicePrivs-电子票据操作权限串
    '     int类别-:0-所有;1-门诊预交;2-住院预交;3-押金
    '出参:
    '编制:焦博
    '日期:2020-06-29 17:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mint类别 = int类别
    Set mcllFilter = cllFilter
    mstrPrivs = strPrivs: mstrEInvoicePrivs = strEInvoicePrivs
    Set mobjEInvoice = objEInvoice: mlngModule = lngModule
    Call InitDepositGrid
    Call InitInvoiceGrid
    Call InitEinvoiceGrid
    Call InitBalanceGrid
    Call InitdkpMain
End Sub

Public Function zlRefrshListData(ByVal cllFilter As Collection, ByVal blnDateMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '入参:cllFilter-过滤条件
    '     格式:array(过滤名称,值1,值2,..),名称
    '返回:成功返回true,否则返回False
    '日期:2012-06-12 14:43:06
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcllFilter = cllFilter
    mblnDateMoved = blnDateMoved
    On Error GoTo errHandle
    Call ShowBills(cllFilter)
    zlCommFun.StopFlash
    zlRefrshListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal cllFilter As Collection)
    '功能:按条件读取单据列表(过滤功能)
    '入参：cllFilter-过滤条件
    Dim strWhere As String, strSQL As String, strSQLYJ As String, strYJwhere As String
    Dim lng病人ID As Long, lng主页ID As Long, lng门诊号 As Long, lng住院号 As Long
    Dim str姓名 As String, str收款人 As String
    Dim str开始单号 As String, str结束单号   As String, str结算号码 As String, str押金类别 As String
    Dim str开始时间 As String, str结束时间 As String, str开始票号 As String, str结束票号  As String
    Dim i As Integer, intCount As Integer, int类型 As Integer
    Dim varData As Variant, dbl金额 As Double, blnOnly押金 As Boolean
    Dim strTempSQL As String, strTmp As String, strNO As String, strNoTmp As String

    On Error GoTo errH
    
    strWhere = ""
    'cllFilter-过滤条件
    For i = 1 To cllFilter.Count
        varData = cllFilter(i)
        Select Case varData(0)
        Case "病人ID"
            lng病人ID = Val(varData(1))
            If lng病人ID <> 0 Then strWhere = strWhere & " And A.病人ID=[1]"
        Case "主页ID"
            lng主页ID = Val(varData(1))
            If lng主页ID <> 0 Then strWhere = strWhere & " And A.主页id=[2]"
        Case "门诊号"
            lng门诊号 = Val(varData(1))
            If lng门诊号 <> 0 Then strWhere = strWhere & " And b.门诊号=[3]"
        Case "住院号"
            lng住院号 = Val(varData(1))
            If lng住院号 <> 0 Then strWhere = strWhere & " And b.病人ID In (Select 病人ID From 病案主页 where 住院号=[4])"
        Case "开始单号"
            str开始单号 = Trim(varData(1))
        Case "结束单号"
            str结束单号 = Trim(varData(1))
        Case "姓名"
            str姓名 = Trim(varData(1))
            If str姓名 <> "" Then strWhere = strWhere & " And Upper(A.姓名) like [11]"
        Case "结算号码"
            str结算号码 = Trim(varData(1))
            If str结算号码 <> "" Then strWhere = strWhere & " And A.str结算号码  = [12]"
        Case "收款人"
            str收款人 = Trim(varData(1))
            If str收款人 <> "" Then strWhere = strWhere & " And A.操作员姓名  = [13]"
        Case "开始时间"
            str开始时间 = Trim(varData(1))
        Case "结束时间"
            str结束时间 = Trim(varData(1))
        Case "开始票号"
            str开始票号 = Trim(varData(1))
        Case "结束票号"
            str结束票号 = Trim(varData(1))
        Case "押金类别"
            str押金类别 = Trim(varData(1))
            If str押金类别 <> "" Then
                blnOnly押金 = True
                strYJwhere = strYJwhere & " And A.押金类别=[14]"
            End If
        Case "查询记录"
            '0-所有;1-正常缴预交;2-仅退款和原预交记录
             int类型 = Val(varData(1))
             If int类型 = 1 Then
                strWhere = strWhere & " and 记录状态=1"
             ElseIf int类型 = 2 Then
                strWhere = strWhere & " And 记录状态<>1"
             End If
        End Select
    Next
    
    If str开始时间 <> "" Then strWhere = strWhere & " And A.收款时间  between  [5] and [6]"
                
    If str开始单号 <> "" And str结束单号 <> "" Then
         strWhere = strWhere & " And a.NO between [7]  and [8]"
    ElseIf str开始单号 <> "" Or str结束单号 <> "" Then
         strWhere = strWhere & IIf(str开始单号 <> "", " And a.NO=[7]", " And a.NO=[8]")
    End If
    
    strSQL = ""
    If str开始票号 <> "" And str结束票号 <> "" Then
        strSQL = "Select A.NO  From 票据打印内容 A,票据使用明细 B Where A.数据性质=2 And A.ID=B.打印ID And B.票种=2 And B.性质=1   And B.号码  between [9] and [10] "
    ElseIf str开始票号 <> "" Or str结束票号 <> "" Then
        strSQL = "Select A.NO  From 票据打印内容 A,票据使用明细 B Where A.数据性质=2 And A.ID=B.打印ID And B.票种=2 And B.性质=1   And B.号码" & IIf(str开始票号 <> "", " = [9] ", " =[10]")
    End If
    If strSQL <> "" Then strWhere = strWhere & " And A.NO in (" & strSQL & ")"
            
    If strWhere = "" Then
        mblnNotRefresh = True
        vsDeposit.Clear 1: vsDeposit.Rows = 2:
        Call RefrshDataDetial
        mblnNotRefresh = False
        Exit Sub
    End If
    strWhere = strWhere & " And Nvl(A.校对标志,0) =0  "
    
    strSQL = ""
    strSQL = _
        "   Select a.附加标志, a.Id As 原始id,'' As 退款NO,a.id,A.NO as NO,A.实际票号 as 票据号,A.操作员姓名 as 操作员," & _
        "           To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS') as 操作时间," & _
        "           A.病人ID,A.门诊号,A.住院号,A.姓名,A.性别,A.年龄,D.名称 as 科室," & _
        "           To_Char(Sum(A.金额),'9999999990.00') as 金额," & _
        "           A.结算方式,A.结算号码,A.摘要,A.记录状态,A.付款方式名称, " & _
        "           Decode(nvl(A.预交类别,2),1,'门诊预交', '住院预交') as 预交类别, nvl(A.预交类别,0) as 预交类别ID, " & _
        "           NULL as 押金类型,NULL as 押金类别,a.交易流水号,a.交易说明,a.预交电子票据" & _
        " From  病人预交记录 A,部门表 D " & _
        " Where A.科室ID=D.ID(+)  And a.记录性质 = 1 " & strWhere & _
                      IIf(mint类别 = 0 Or mint类别 = 3, "", "  And   A.预交类别=" & mint类别) & _
        " Group by a.id,A.NO,A.记录状态,A.实际票号 ,Nvl(A.预交类别, 0),Decode(nvl(A.预交类别,2),1,'门诊预交', '住院预交'),A.操作员姓名," & _
        "           To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS'),A.病人ID,A.门诊号,A.住院号,A.姓名,A.年龄," & _
        "           A.性别 , D.名称, A.结算方式,A.结算号码, A.摘要,A.付款方式名称,a.交易流水号,a.交易说明,a.预交电子票据,a.附加标志 "

     strSQLYJ = _
              "   Select 0 as 附加标志,a.Id As 原始id,'' As 退款NO,a.id,A.NO as NO,A.实际票号 as 票据号,A.操作员姓名 as 操作员," & _
              "           To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS') as 操作时间," & _
              "           A.病人ID,A.门诊号,A.住院号,A.姓名,A.性别,A.年龄,D.名称 as 科室," & _
              "           To_Char(Sum(A.金额),'9999999990.00') as 金额," & _
              "           A.结算方式,A.结算号码,A.摘要,A.记录状态,A.付款方式名称, " & _
              "           NULL as 预交类别,NULL as 预交类别ID,Decode(nvl(A.是否门诊,0),1,'门诊押金', '住院押金') as 押金类型,A.押金类别,a.交易流水号,a.交易说明,0 as 预交电子票据 " & _
              " From 病人押金记录  A,部门表 D " & _
              " Where A.科室ID=D.ID(+) " & strWhere & strYJwhere & _
              " Group by a.id,A.NO,A.记录状态,A.实际票号 ,A.押金类别,Decode(nvl(A.是否门诊,0),1,'门诊押金', '住院押金'),A.操作员姓名," & _
              "           To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS'),A.病人ID,A.门诊号,A.住院号,A.姓名,A.年龄," & _
              "           A.性别 , D.名称, A.结算方式,A.结算号码, A.摘要,A.付款方式名称,a.交易流水号,a.交易说明"
              
    'mint类别:0-查询所有;1-查询门诊预交;2-查询住院预交;3-查询押金记录
    If mint类别 <> 3 Then
             strSQL = strSQL & " Union ALL " & _
              "   Select 11 as 附加标志,e.Id As 原始ID,A.no as 退款NO,a.id,J.NO as NO,A.实际票号 as 票据号,A.操作员姓名 as 操作员," & _
              "           To_Char(J.收款时间,'YYYY-MM-DD HH24:MI:SS') as 操作时间," & _
              "           A.病人ID,A.门诊号,A.住院号,A.姓名,A.性别,A.年龄,D.名称 as 科室," & _
              "           To_Char(Sum(-1*A.冲预交),'9999999990.00') as 金额," & _
              "           A.结算方式,A.结算号码,A.摘要,A.记录状态,A.付款方式名称, " & _
              "           Decode(nvl(A.预交类别,2),1,'门诊预交', '住院预交') as 预交类别, nvl(A.预交类别,0) as 预交类别ID, " & _
              "           NULL as 押金类型,NULL as 押金类别,a.交易流水号,a.交易说明,e.预交电子票据" & _
              " From   ( Select Distinct a.结帐id, a.NO, a.收款时间,a.预交电子票据 From  病人预交记录 A  " & _
              "                Where  a.记录性质 = 1 " & strWhere & "and nvl(附加标志,0)>=1 )  J," & _
               "                   病人预交记录 A,部门表 D,病人预交记录 E " & _
              " Where  J.结帐ID=A.结帐ID And  A.科室ID=D.ID(+) And A.记录性质=11 " & _
              "            And Nvl(a.校对标志, 0) = 0 And a.冲预交 > 0  And a.no=e.no And e.记录性质=1 And e.记录状态=1 " & _
              " Group by J.no ,e.Id,a.id,A.NO,A.记录状态,A.实际票号 ,Nvl(A.预交类别, 0),Decode(nvl(A.预交类别,2),1,'门诊预交', '住院预交'),A.操作员姓名," & _
              "           To_Char(J.收款时间,'YYYY-MM-DD HH24:MI:SS'),A.病人ID,A.门诊号,A.住院号,A.住院号,A.姓名,A.年龄," & _
              "           A.性别 , D.名称, A.结算方式,A.结算号码, A.摘要,A.付款方式名称,a.交易流水号,a.交易说明,e.预交电子票据,a.附加标志 "
    End If
       
    If mint类别 = 0 Then
        If blnOnly押金 Then '仅读取押金
            strSQL = strSQLYJ
        Else
            strSQL = strSQL & " Union all " & strSQLYJ
        End If
    ElseIf mint类别 = 3 Then
        strSQL = strSQLYJ
    End If
    If mblnDateMoved Then
        strTempSQL = Replace(Replace(strSQL, "病人预交记录", "H病人预交记录"), "病人押金记录", "H病人押金记录")
        strTempSQL = Replace(Replace(strSQL, "票据打印内容", "H票据打印内容"), "票据使用明细", "H票据使用明细")
        strSQL = strSQL & " Union ALL " & vbCrLf & strTempSQL
    End If
      
    strSQL = "Select a.附加标志, a.原始id,a.退款no, a.Id, a.NO, a.票据号, a.操作员, a.操作时间, a.病人id, a.门诊号, a.住院号, a.姓名, a.性别, a.年龄, a.科室, a.金额, a.结算方式, a.结算号码," & vbNewLine & _
                 "  a.摘要, a.记录状态, a.付款方式名称, a.预交类别, a.预交类别id,  a.押金类型, a.押金类别, a.交易流水号, a.交易说明, a.预交电子票据 " & _
                 " From (" & strSQL & ") A Order by a.操作时间 desc,a.NO,a.退款NO desc"

    Set mrsList = New ADODB.Recordset
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, lng门诊号, lng住院号, CDate(str开始时间), CDate(str结束时间), str开始单号, str结束单号, str开始票号, str结束票号, _
                                                                    str姓名 & "%", str结算号码, str收款人, str押金类别)
    mblnNotRefresh = True
    vsDeposit.Clear
    vsDeposit.Rows = 2
    If mrsList.EOF Then
        Call InitDepositGrid
        strTmp = "当前设置没有过滤出任何单据"
        RaiseEvent ShowStatus(strTmp)
        RaiseEvent SelectDeposit(False, False, 0, 0, False, 0, False, "", False, 0, 0)
    Else
        RaiseEvent ReadPati(Nvl(mrsList!姓名))
        Call InitDepositGrid
        vsDeposit.ForeColorSel = vsDeposit.CellForeColor
        mrsList.MoveFirst: dbl金额 = 0
        vsDeposit.Rows = mrsList.RecordCount + 1
        With vsDeposit
            .OutlineBar = flexOutlineBarSymbolsLeaf
            .Subtotal flexSTClear
            .MultiTotals = True
            .SubtotalPosition = flexSTAbove
            .OutlineCol = .ColIndex("单据号")
            For i = 1 To mrsList.RecordCount
                .TextMatrix(i, .ColIndex("退款NO")) = Nvl(mrsList!退款NO)
                .TextMatrix(i, .ColIndex("NO")) = Nvl(mrsList!NO)
                .TextMatrix(i, .ColIndex("原始ID")) = Val(Nvl(mrsList!原始ID))
                .TextMatrix(i, .ColIndex("ID")) = Val(Nvl(mrsList!ID))
                .TextMatrix(i, .ColIndex("单据号")) = IIf(Nvl(mrsList!退款NO) = "", Nvl(mrsList!NO), Nvl(mrsList!退款NO))
                .TextMatrix(i, .ColIndex("票据号")) = Nvl(mrsList!票据号)
                .TextMatrix(i, .ColIndex("操作员")) = Nvl(mrsList!操作员)
                .TextMatrix(i, .ColIndex("操作时间")) = Nvl(mrsList!操作时间)
                .TextMatrix(i, .ColIndex("病人ID")) = Val(Nvl(mrsList!病人ID))
                .TextMatrix(i, .ColIndex("门诊号")) = Nvl(mrsList!门诊号)
                .TextMatrix(i, .ColIndex("住院号")) = Nvl(mrsList!住院号)
                .TextMatrix(i, .ColIndex("姓名")) = Nvl(mrsList!姓名)
                .TextMatrix(i, .ColIndex("性别")) = Nvl(mrsList!性别)
                .TextMatrix(i, .ColIndex("年龄")) = Nvl(mrsList!年龄)
                .TextMatrix(i, .ColIndex("科室")) = Nvl(mrsList!科室)
                .TextMatrix(i, .ColIndex("金额")) = Nvl(mrsList!金额)
                If IsNumeric(.TextMatrix(i, .ColIndex("金额"))) Then .TextMatrix(i, .ColIndex("金额")) = Format(.TextMatrix(i, .ColIndex("金额")), "0.00")
                .TextMatrix(i, .ColIndex("结算方式")) = Nvl(mrsList!结算方式)
                .TextMatrix(i, .ColIndex("结算号码")) = Nvl(mrsList!结算号码)
                .TextMatrix(i, .ColIndex("摘要")) = Nvl(mrsList!摘要)
                .TextMatrix(i, .ColIndex("记录状态")) = Nvl(mrsList!记录状态)
                If .TextMatrix(i, .ColIndex("记录状态")) = "2" Then
                    .Cell(flexcpForeColor, i, 0, i, .ColIndex("电子票据")) = &HFF&
                ElseIf .TextMatrix(i, .ColIndex("记录状态")) = "3" Then
                    .Cell(flexcpForeColor, i, 0, i, .ColIndex("电子票据")) = &HFF0000
                End If
                .TextMatrix(i, .ColIndex("医疗付款方式")) = Nvl(mrsList!付款方式名称)
                .TextMatrix(i, .ColIndex("预交类别")) = Nvl(mrsList!预交类别)
                .TextMatrix(i, .ColIndex("预交类别ID")) = Val(Nvl(mrsList!预交类别ID))
                .TextMatrix(i, .ColIndex("押金类型")) = Nvl(mrsList!押金类型)
                .TextMatrix(i, .ColIndex("押金类别")) = Nvl(mrsList!押金类别)
                .TextMatrix(i, .ColIndex("交易流水号")) = Nvl(mrsList!交易流水号)
                .TextMatrix(i, .ColIndex("交易说明")) = Nvl(mrsList!交易说明)
                .TextMatrix(i, .ColIndex("电子票据")) = IIf(Val(Nvl(mrsList!预交电子票据)) = 1, "√", "")
                .TextMatrix(i, .ColIndex("附加标志")) = Val(Nvl(mrsList!附加标志))
                .IsSubtotal(i) = True
                If Nvl(mrsList!退款NO) = "" And Val(Nvl(mrsList!附加标志)) >= 1 Then
                    strNoTmp = Nvl(mrsList!NO)
                    .Cell(flexcpBackColor, i, 0, i, .ColIndex("电子票据")) = &HC0C0FF
                    .RowOutlineLevel(i) = 1
                Else
                    If Nvl(mrsList!NO) = strNoTmp And Val(Nvl(mrsList!附加标志)) = 11 Then
                        .RowOutlineLevel(i) = 2
                        intCount = intCount + 1
                    Else
                        .RowOutlineLevel(i) = 1
                    End If
                End If
                If Val(Nvl(mrsList!附加标志)) <> 11 Then
                    dbl金额 = dbl金额 + Val(Nvl(mrsList!金额))
                End If
                mrsList.MoveNext
            Next
            .Outline 1
        End With
        mrsList.MoveFirst
        strTmp = "共 " & mrsList.RecordCount - intCount & " 张单据,合计:" & Format(dbl金额, "0.00")
        RaiseEvent ShowStatus(strTmp)
    End If
    mblnNotRefresh = False
    '恢复上次行
    If mlngCurRow = 0 Then mlngCurRow = 1
    If mlngTopRow = 0 Then mlngTopRow = 1
    If mlngCurRow <= vsDeposit.Rows - 1 Then
        vsDeposit.Row = mlngCurRow
    Else
        vsDeposit.Row = vsDeposit.Rows - 1
    End If
    If mlngTopRow <= vsDeposit.Rows - 1 Then
        vsDeposit.TopRow = mlngTopRow
    Else
        vsDeposit.TopRow = vsDeposit.Row
    End If
    Call RefrshDataDetial   '加载明细数据
    Me.Refresh
    Exit Sub
errH:
    mblnNotRefresh = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefrshDataDetial()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新数据
    '编制:刘兴洪
    '日期:2020-04-28 10:19:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng预交ID As Long, lng电子票据ID As Long, lng原始ID As Long
    Dim strNO As String, str票据号 As String
    Dim int记录状态 As Integer, int附加标志 As Integer
    Dim bln是否押金 As Boolean, bln预交电子票据 As Boolean, bln换开 As Boolean, bln退预交 As Boolean
    
    On Error GoTo errHandle
    
     With vsDeposit
        If Not (.Row <= 0 Or .TextMatrix(.Row, .ColIndex("NO")) = "") Then
            strNO = .TextMatrix(.Row, .ColIndex("NO"))
            lng预交ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            bln是否押金 = .TextMatrix(.Row, .ColIndex("押金类型")) <> ""
            bln预交电子票据 = Nvl(.TextMatrix(.Row, .ColIndex("电子票据"))) = "√"
            int记录状态 = Val(Nvl(.TextMatrix(.Row, .ColIndex("记录状态"))))
            str票据号 = .TextMatrix(.Row, .ColIndex("票据号"))
            bln退预交 = Val(.TextMatrix(.Row, .ColIndex("附加标志"))) = 11
            int附加标志 = Val(.TextMatrix(.Row, .ColIndex("附加标志")))
            lng原始ID = Val(.TextMatrix(.Row, .ColIndex("原始ID")))
            mlngGo = .Row: mlngCurRow = .Row: mlngTopRow = .TopRow
        End If
    End With
    If strNO = "" Then Exit Sub
    
    If mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("病人预交记录", strNO, , "1", Me.Caption)
    Else
        mblnNOMoved = False
    End If
    
    '加载明细
    Call LoadEInvoiceData(strNO)
    Call LoadInvoiceData(strNO)
    Call LoadBalanceInfor(lng预交ID, bln是否押金)

    RaiseEvent SelectDeposit(True, bln是否押金, lng预交ID, int记录状态, bln预交电子票据, lng电子票据ID, bln换开, str票据号, bln退预交, int附加标志, lng原始ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitDepositGrid()
    Dim strHead As String
    Dim i As Integer
    
    strHead = "退款NO,1,0|原始ID,1,0|NO,1,0|ID,1,0|单据号,1,1200|票据号,4,1050|操作员,1,850|操作时间,4,1850|病人ID,1,750|门诊号,1,750|住院号,1,750|姓名,1,800|性别,4,500|" & _
              "年龄,4,500|科室,1,850|金额,7,850|结算方式,1,850|结算号码,1,1500|摘要,1,1500|记录状态,1,0|医疗付款方式,1,1500|预交类别,4,800|" & _
              "预交类别ID,1,0|押金类型,4,800|押金类别,1,900|交易流水号,4,1000|交易说明,4,1500|电子票据,4,900|附加标志,1,0"
    With vsDeposit
        mblnNotRefresh = True
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
        Next
        If Not Visible Then Call RestoreFlexState(vsDeposit, App.ProductName & "\" & Me.Name)
        .ColHidden(.ColIndex("预交类别ID")) = True
        .ColHidden(.ColIndex("退款NO")) = True
        .ColHidden(.ColIndex("原始ID")) = True
        .ColHidden(.ColIndex("NO")) = True
        .ColHidden(.ColIndex("ID")) = True
        .ColHidden(.ColIndex("附加标志")) = True
        .ColHidden(.ColIndex("记录状态")) = True
        If mint类别 = 1 Or mint类别 = 2 Then .ColHidden(.ColIndex("押金类型")) = True: .ColHidden(.ColIndex("押金类别")) = True:
        .RowHeight(0) = 320
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        .Col = 0: .ColSel = .Cols - 1
        Call vsDeposit_EnterCell
        mblnNotRefresh = False
        zl_vsGrid_Para_Restore mlngModule, vsDeposit, Me.Name, "预交信息列表", False
        If .Rows > 1 Then .Row = 1
        .Redraw = True
    End With

End Sub

Private Sub LoadInvoiceData(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取票据使用明细
    '入参:strNO-病人信息集
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-04-01 18:55:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    On Error GoTo errH
    If strNO = "" Then
        vsInvoice.Redraw = flexRDNone
        vsInvoice.Rows = 2
        vsInvoice.Clear 1
        vsInvoice.Redraw = flexRDBuffered: Exit Sub
    End If
    
    strSQL = _
    " Select b.Id, b.号码 As 票据号," & vbNewLine & _
    " Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出',7,'红票收回') As 使用原因," & vbNewLine & _
    "    To_Char(b.使用时间, 'MM-DD HH24:MI') As 使用时间, b.使用人" & vbNewLine & _
    " From 票据打印内容 A, 票据使用明细 B" & vbNewLine & _
    " Where a.数据性质 = 2 And a.Id = b.打印id And a.No = [1] and B.票种=2" & vbNewLine & _
    " Order By ID"
    
    mblnNOMoved = zlDatabase.NOMoved("病人预交记录", strNO, , 1)
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "票据打印内容", "H票据打印内容")
        strSQL = Replace(strSQL, "票据使用明细", "H票据使用明细")
    End If
    
    Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    dkpMain.Panes(3).Closed = rsInvoice.EOF
    If rsInvoice.EOF Then
        vsInvoice.Rows = 2
        vsInvoice.Clear 1
        Exit Sub
    End If
    
    With vsInvoice
        .Clear 1
        .Rows = 2
        If rsInvoice.RecordCount <> 0 Then rsInvoice.MoveFirst
        .Rows = IIf(rsInvoice.RecordCount = 0, 1, rsInvoice.RecordCount) + 1
        i = 1
        Do While Not rsInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsInvoice!ID)
            .TextMatrix(i, .ColIndex("票据号")) = Nvl(rsInvoice!票据号)
            .TextMatrix(i, .ColIndex("使用原因")) = Nvl(rsInvoice!使用原因)
            .TextMatrix(i, .ColIndex("使用时间")) = Nvl(rsInvoice!使用时间)
            .TextMatrix(i, .ColIndex("使用人")) = Nvl(rsInvoice!使用人)
            i = i + 1
            rsInvoice.MoveNext
        Loop
    End With
    vsInvoice.Redraw = flexRDBuffered
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadEInvoiceData(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载电子发票信息
    '编制:刘兴洪
    '日期:2020-03-25 17:13:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSQL As String
    Dim rsEInvoice As ADODB.Recordset
    Dim lng原预交ID As Long
    On Error GoTo errHandle

    vsEInvoice.Tag = ""
    vsEInvoice.Clear 1: vsEInvoice.Rows = 2
     
    If strNO = "" Or mobjEInvoice Is Nothing Then Exit Sub
    Call mobjEInvoice.zlIsStartEinvoicFromNO(strNO, lng原预交ID)
    
    If Not mobjEInvoice.zlGetEInvoiceInforFromBalanceID(lng原预交ID, rsEInvoice, 2, 0) Then Exit Sub
    dkpMain.Panes(4).Closed = rsEInvoice.EOF
    If rsEInvoice.EOF Then Exit Sub
    
    With vsEInvoice
        If rsEInvoice.RecordCount <> 0 Then rsEInvoice.MoveFirst
        i = 1
        Do While Not rsEInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsEInvoice!ID)
            .TextMatrix(i, .ColIndex("记录状态")) = Nvl(rsEInvoice!记录状态)
            .TextMatrix(i, .ColIndex("结算ID")) = Nvl(rsEInvoice!结算ID)
            .TextMatrix(i, .ColIndex("发票代码")) = Nvl(rsEInvoice!代码)
            .TextMatrix(i, .ColIndex("发票号码")) = Nvl(rsEInvoice!号码)
            .TextMatrix(i, .ColIndex("票据金额")) = Format(Nvl(rsEInvoice!票据金额), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("生成时间")) = Format(rsEInvoice!生成时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("换开纸质发票")) = IIf(Val(Nvl(rsEInvoice!是否换开)) = 1, "已换开", "未换开")
            .TextMatrix(i, .ColIndex("纸质发票号")) = Nvl(rsEInvoice!纸质发票号)
            .TextMatrix(i, .ColIndex("备注")) = Nvl(rsEInvoice!备注)
            .TextMatrix(i, .ColIndex("操作员姓名")) = Nvl(rsEInvoice!操作员姓名)
            If Val(Nvl(rsEInvoice!记录状态)) = 1 Then
                 .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = Me.ForeColor
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = IIf(Val(Nvl(rsEInvoice!记录状态)) = 2, vbRed, vbBlue)
            End If
            i = i + 1: .Rows = .Rows + 1
            rsEInvoice.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mfrmMain.mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyF5
            Call ShowBills(mcllFilter)
        Case vbKeyF9
            RaiseEvent MoneyEnum
        Case vbKeyF11
            RaiseEvent RollingCurtain
        Case vbKeyF12
            RaiseEvent FileLocalSet
        Case vbKeyF
            If Shift = vbCtrlMask Then RaiseEvent FilterDeposit
        Case vbKeyG
            If Shift = vbCtrlMask Then RaiseEvent ViewGo
        Case vbKeyP
            If Shift = vbCtrlMask Then RaiseEvent FilePrint
        Case vbKeyEscape
            mblnGo = False
        Case vbKeyC
            If mfrmMain.mnuEidt_CationMoney_Del.Enabled And mfrmMain.mnuEidt_CationMoney_Del.Visible Then Call ExcuteCationMoney_Del
        Case vbKeyA
            If Shift = vbCtrlMask Then
                If mfrmMain.mnuEdit_Deposit.Enabled And mfrmMain.mnuEdit_Deposit.Visible Then RaiseEvent EditDeposit
            End If
        Case vbKeyR
            If Shift = vbCtrlMask Then RaiseEvent EidtBalanceDel
        Case vbKeyDelete
            If Shift = vbShiftMask Then
                If mfrmMain.mnuEdit_Del.Enabled And mfrmMain.mnuEdit_Del.Visible Then Call ExcuteMoney_Del
            End If
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hwnd, Me.Name
    End Select
End Sub

Private Sub vsDeposit_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsDeposit, Me.Name, "预交信息列表", False
End Sub

Private Sub vsDeposit_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsDeposit, Me.Name, "预交信息列表", False
End Sub

Private Sub vsDeposit_DblClick()
    If vsDeposit.MouseRow = 0 Then Exit Sub
    If mfrmMain.mnuEdit_View.Enabled Then Call ExcuteViewDepositNO
End Sub

Private Sub vsDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call ExcuteMoney_Del
    Else
        Call Form_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub vsDeposit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then RaiseEvent PopupMenu
End Sub

Private Sub vsDeposit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsDeposit.MouseRow = 0 Then
        vsDeposit.MousePointer = 99
    Else
        vsDeposit.MousePointer = 0
    End If
End Sub

Private Sub vsDeposit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = vsDeposit.MouseCol
    
    If Button = 1 And vsDeposit.MousePointer = 99 And lngCol > 0 Then
        If vsDeposit.TextMatrix(0, lngCol) = "" Then Exit Sub
        If vsDeposit.TextMatrix(1, vsDeposit.ColIndex("NO")) = "" Then Exit Sub
    End If
End Sub

Private Sub picBalanceInfo_Resize()
    Err = 0: On Error Resume Next
    With vsBalanceInfor
        .Top = picBalanceInfo.ScaleTop
        .Left = picBalanceInfo.ScaleLeft
        .Height = picBalanceInfo.ScaleHeight
        .Width = picBalanceInfo.ScaleWidth
    End With
End Sub

Private Sub picDeposit_Resize()
    Err = 0: On Error Resume Next
    With picDeposit
        vsDeposit.Left = 0
        vsDeposit.Top = 0
        vsDeposit.Height = .ScaleHeight
        vsDeposit.Width = .ScaleWidth
    End With
End Sub

Private Sub picEinvoice_Resize()
    Err = 0: On Error Resume Next
    With vsEInvoice
        .Top = picEinvoice.ScaleTop
        .Left = picEinvoice.ScaleLeft
        .Height = picEinvoice.ScaleHeight
        .Width = picEinvoice.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Set mobjEInvoice = Nothing
    mint类别 = 0
    mblnNOMoved = False: mblnDateMoved = False
    mstrPrivs = "": mstrEInvoicePrivs = ""
    mblnNotRefresh = False
    Set mrsList = Nothing
    mlngTopRow = 0: mlngCurRow = 0
End Sub

Private Sub vsBalanceInfor_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalanceInfor, Me.Name, "结算相关信息列表", False
End Sub

Private Sub vsBalanceInfor_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalanceInfor, Me.Name, "结算相关信息列表", False
End Sub

Private Sub vsEInvoice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsEInvoice, Me.Name, "电子票据信息列表", False
End Sub

Private Sub vsEInvoice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsEInvoice, Me.Name, "电子票据信息列表", False
End Sub

Private Sub vsInvoice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoice, Me.Name, "发票信息列表", False
End Sub

Private Sub vsInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsInvoice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsInvoice, Me.Name, "发票信息列表", False
End Sub

Private Sub vsInvoice_GotFocus()
    zl_VsGridGotFocus vsInvoice, &HFFC0C0
End Sub

Private Sub vsInvoice_LostFocus()
    zl_VsGridLOSTFOCUS vsInvoice, , vsInvoice.Cell(flexcpForeColor, vsInvoice.Row, vsInvoice.Col)
End Sub
Private Sub InitInvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化发票网格控件
    '编制:刘兴洪
    '日期:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 5
        .TextMatrix(0, i) = "ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "票据号": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "使用原因": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "使用时间": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "使用人": .ColWidth(i) = 1000: i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1200
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "ID"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "票据号"
                .ColAlignment(i) = flexAlignCenterCenter
            End Select
        Next
        
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore mlngModule, vsInvoice, Me.Name, "发票信息列表", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub InitEinvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化电子发票网格控件
    '编制:刘兴洪
    '日期:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsEInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 11
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "记录状态": i = i + 1
        .TextMatrix(0, i) = "结算ID": i = i + 1
        .TextMatrix(0, i) = "发票代码": i = i + 1
        .TextMatrix(0, i) = "发票号码": i = i + 1
        .TextMatrix(0, i) = "票据金额": i = i + 1
        .TextMatrix(0, i) = "生成时间": i = i + 1
        .TextMatrix(0, i) = "换开纸质发票": i = i + 1
        .TextMatrix(0, i) = "纸质发票号": i = i + 1
        .TextMatrix(0, i) = "备注": i = i + 1
        .TextMatrix(0, i) = "操作员姓名": i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1200
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "记录状态"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "备注"
                .ColWidth(i) = 2000
            Case "操作员姓名"
                 .ColWidth(i) = 1000
            Case "票据金额"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore mlngModule, vsEInvoice, Me.Name, "电子票据信息列表", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub InitBalanceGrid()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|结算方式,1,0|名称,1,0|金额,1,0|项目,1,1200|内容,1,2000|交易流水号,1,0 "
    
    With vsBalanceInfor
        .HighLight = flexHighlightWithFocus
        .Redraw = flexRDNone
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "ID" Or .ColKey(i) = "交易流水号" Or .ColKey(i) = "结算方式" Or .ColKey(i) = "名称" Or .ColKey(i) = "金额" Or .ColKey(i) = "位置" Then .ColHidden(i) = True
        Next
        If .Rows < 2 Then .Rows = 2
        .RowHeightMin = 350
        '.Row = 1: .Col = 0: .ColSel = .COLS - 1
         .Redraw = flexRDBuffered
        If .TextMatrix(1, 0) = "" Then Exit Sub

        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        .Subtotal flexSTNone, .ColIndex("ID"), .ColIndex("项目"), gstrDec, &H8000000F
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("项目")
        .OutlineCol = .ColIndex("项目")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("项目")) = strTemp

                strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("结算方式"))
                strTemp = strTemp & "(" & Format(.Cell(flexcpTextDisplay, i + 1, .ColIndex("金额")), gstrDec) & ")"
                If .Cell(flexcpTextDisplay, i + 1, .ColIndex("交易流水号")) <> "" Then
                   strTemp = strTemp & Space(1) & "交易流水号:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("交易流水号"))
                End If
                
                .MergeRow(i) = True
                .MergeCells = flexMergeRestrictRows
                .Cell(flexcpAlignment, i, .ColIndex("项目"), i, .ColIndex("项目")) = 1
                
                For j = 0 To .Cols - 1
                   If j <= .ColIndex("内容") Then
                       If j >= .ColIndex("项目") Then
                           .Cell(flexcpText, i, j) = strTemp
                           .Cell(flexcpFontBold, i, j) = False
                       End If
                   End If
                Next
            End If
        Next
        Call .AutoSize(.ColIndex("项目"))
        For j = 0 To .Cols - 1
            .MergeCol(j) = True
        Next
        zl_vsGrid_Para_Restore mlngModule, vsBalanceInfor, Me.Name, "结算相关信息列表", False
    End With
End Sub

Private Sub LoadBalanceInfor(ByVal lng预交ID As Long, ByVal bln是否押金 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取三方结算交易明细
    '入参:lng预交ID-病人预交记录.ID
    '       bln是否押金-是否为押金
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-04-01 18:55:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsInfo As ADODB.Recordset
    
    On Error GoTo errH
    
    If lng预交ID = 0 Then
        vsBalanceInfor.Clear 2: vsBalanceInfor.Rows = 2
        Exit Sub
    End If
    If bln是否押金 Then
        strSQL = _
            "Select b.交易id || '_' || b.原预交id As ID, a.结算方式, Max(c.名称) As 名称, Sum(A.金额) As 金额," & vbNewLine & _
            "       b.交易项目, b.交易内容, Max(a.交易流水号) As 交易流水号 " & vbNewLine & _
            "From 病人押金记录 A, 三方结算交易 B, 医疗卡类别 C" & vbNewLine & _
            "Where a.Id = b.交易id And a.卡类别id = c.Id(+) And a.ID = [1] " & vbNewLine & _
            "      And Nvl(b.性质,0) = 2 " & vbNewLine & _
            "Group By b.交易id, b.原预交id, a.结算方式, b.交易项目, b.交易内容" & vbNewLine & _
            "Order By ID"
    Else
        strSQL = _
            "Select b.交易id || '_' || b.原预交id As ID, a.结算方式, Max(c.名称) As 名称, Sum(Nvl(-1 * f.金额, a.冲预交)) As 金额," & vbNewLine & _
            "       b.交易项目, b.交易内容, Max(Nvl(f.交易流水号, a.交易流水号)) As 交易流水号 " & vbNewLine & _
            "From 病人预交记录 A, 三方结算交易 B, 医疗卡类别 C, 病人预交记录 E, 三方退款信息 F" & vbNewLine & _
            "Where a.Id = b.交易id And a.卡类别id = c.Id(+) And a.ID = [1] " & vbNewLine & _
            "      And b.原预交id = e.Id(+) And e.id = f.记录id(+) And f.结帐id(+) =  [1] And Nvl(b.性质,0) = 0 " & vbNewLine & _
            "Group By b.交易id, b.原预交id, a.结算方式, b.交易项目, b.交易内容" & vbNewLine & _
            "Order By ID"
    End If
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = Replace(strSQL, "病人押金记录", "H病人押金记录")
        strSQL = Replace(strSQL, "三方结算交易", "H三方结算交易")
    End If
    
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng预交ID)
    dkpMain.Panes(2).Closed = rsInfo.EOF
    If rsInfo.EOF Then
        vsBalanceInfor.Rows = 2
        vsBalanceInfor.Clear 1
        Exit Sub
    End If
    Set vsBalanceInfor.DataSource = rsInfo
    Call InitBalanceGrid
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picInvoice_Resize()
    Err = 0: On Error Resume Next
    With vsInvoice
        .Top = picInvoice.ScaleTop
        .Left = picInvoice.ScaleLeft
        .Height = picInvoice.ScaleHeight
        .Width = picInvoice.ScaleWidth
    End With
End Sub

Private Sub vsBalanceInfor_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsBalanceInfor, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBalanceInfor_GotFocus()
    zl_VsGridGotFocus vsBalanceInfor, &HFFC0C0
End Sub

Private Sub vsBalanceInfor_LostFocus()
    zl_VsGridLOSTFOCUS vsBalanceInfor, , vsBalanceInfor.Cell(flexcpForeColor, vsBalanceInfor.Row, vsBalanceInfor.Col)
End Sub

Private Sub vsDeposit_AfterSort(ByVal Col As Long, Order As Integer)
    If vsDeposit.Row <= 0 Or vsDeposit.Col <= 0 Then Exit Sub
    vsDeposit.ForeColorSel = vsDeposit.CellForeColor
End Sub

Private Sub vsDeposit_EnterCell()
    If mblnNotRefresh Then Exit Sub
    vsDeposit.ForeColorSel = vsDeposit.CellForeColor
    Call RefrshDataDetial
End Sub

Private Sub InitdkpMain()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:初始化dkpMain控件
    '编制:焦博
    '日期:2020-04-26
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 1500, 3000, DockTopOf, Nothing)
        objPanel.Handle = picDeposit.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        
        Set objPanel = .CreatePane(2, 500, 1500, DockBottomOf, objPanel)
        objPanel.Title = "结算关联信息"
        objPanel.Handle = picBalanceInfo.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        Set objPanel = .CreatePane(3, 500, 1500, DockRightOf, objPanel)
        objPanel.Title = "预交票据信息"
        objPanel.Handle = picInvoice.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

        
        Set objPanel = .CreatePane(4, 500, 1500, DockRightOf, objPanel)
        objPanel.Title = "电子票据信息"
        objPanel.Handle = picEinvoice.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

        .Panes(2).Closed = True
        .Panes(3).Closed = True
        .Panes(4).Closed = True
        
        .Options.HideClient = True

    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SeekBill(blnHead As Boolean)
    Dim i As Long, strTmp As String
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    strTmp = "正在定位满足条件的单据,按ESC终止 ..."
    RaiseEvent ShowStatus(strTmp)
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To vsDeposit.Rows - 1
        
        '比较条件
        blnFill = True
        With frmDepositFind
            If .txtNO.Text <> "" Then
                blnFill = blnFill And vsDeposit.TextMatrix(i, vsDeposit.ColIndex("NO")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And vsDeposit.TextMatrix(i, vsDeposit.ColIndex("票据号")) = .txtFact.Text
            End If
            If .cbo操作员.ListIndex > 0 Then
                blnFill = blnFill And vsDeposit.TextMatrix(i, vsDeposit.ColIndex("操作员")) = zlCommFun.GetNeedName(.cbo操作员.Text)
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If IsNumeric(.txt住院号.Text) Then
                blnFill = blnFill And Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("住院号"))) = Val(.txt住院号.Text)
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            vsDeposit.Row = i: vsDeposit.TopRow = i
            vsDeposit.Col = 0: vsDeposit.ColSel = vsDeposit.Cols - 1
            strTmp = "找到一条记录"
            RaiseEvent ShowStatus(strTmp)
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            strTmp = "用户取消定位操作"
            RaiseEvent ShowStatus(strTmp)
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    strTmp = "已定位到清单尾部"
    RaiseEvent ShowStatus(strTmp)
    Screen.MousePointer = 0
End Sub

Public Sub ExcuteViewDepositNO()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查看单据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnViewCancel  As Boolean
    Dim strNO As String, str操作员 As String, byt预交类型 As Byte
    Dim int记录状态 As Integer, blnNOMoved As Boolean
    Dim bln是否押金 As Boolean
    
    With vsDeposit
        strNO = .TextMatrix(.Row, .ColIndex("单据号"))
        str操作员 = .TextMatrix(.Row, .ColIndex("操作员"))
        byt预交类型 = Val(.TextMatrix(.Row, .ColIndex("预交类别ID")))
        int记录状态 = Val(.TextMatrix(.Row, .ColIndex("记录状态")))
        blnViewCancel = int记录状态 = 2
        bln是否押金 = .TextMatrix(.Row, .ColIndex("押金类型")) <> ""
    End With

    '是否已转入后备数据表中
    If mblnNOMoved Then
        blnNOMoved = zlDatabase.NOMoved("病人预交记录", strNO, , "1")
    End If
    
    If strNO = "" Then MsgBox "当前没有记录可以查阅！", vbExclamation, gstrSysName: Exit Sub
    '显示单据内容
    If bln是否押金 Then
        Call frmCautionMoney.zlShowEdit(Me, 1, mstrPrivs, mlngModule, strNO, blnViewCancel, blnNOMoved)
    Else
        Call frmDeposit.zlShowEdit(Me, 0, 1, mobjEInvoice, mstrPrivs, mlngModule, byt预交类型, strNO, blnViewCancel, blnNOMoved)
    End If
End Sub

 Public Function ExcuteMoney_Del() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:预交退款或押金退款
    '编制:焦博
    '日期:2020-06-22 11:17:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, str操作员 As String
    Dim byt预交类型 As Byte, bln押金 As Boolean
    
    With vsDeposit
        strNO = .TextMatrix(.Row, .ColIndex("单据号"))
        str操作员 = .TextMatrix(.Row, .ColIndex("操作员"))
        byt预交类型 = Val(.TextMatrix(.Row, .ColIndex("预交类别ID")))
        bln押金 = .TextMatrix(.Row, .ColIndex("押金类型")) <> ""
    End With
    If strNO = "" Then
        MsgBox "当前没有记录可以退款！", vbExclamation, gstrSysName
        Exit Function
    End If
        
    '单据权限
    If Not BillOperCheck(6, str操作员, _
        CDate(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("操作时间"))), "退款") Then Exit Function
    
    If Val(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("金额"))) < 0 Then
        MsgBox "该缴款记录金额为负,表示退款,不能执行该操作！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 6, Me.Caption) Then Exit Function
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If bln押金 Then
         If InStr(1, mstrPrivs, ";押金退款;") = 0 Then
            MsgBox "你没有权限进行押金退款操作！", vbInformation, gstrSysName
            Exit Function
        End If
        On Error Resume Next
        Err.Clear
        ExcuteMoney_Del = frmCautionMoney.zlShowEdit(Me, 2, mstrPrivs, mlngModule, strNO)
        If ExcuteMoney_Del Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Call ShowBills(mcllFilter)
        End If
        Exit Function
    End If
    
    If ChekcDepositDelPrivs(strNO) = False Then Exit Function

    On Error Resume Next
    Err.Clear
    ExcuteMoney_Del = frmDeposit.zlShowEdit(Me, 0, 2, mobjEInvoice, mstrPrivs, mlngModule, byt预交类型, strNO)
    If ExcuteMoney_Del Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Call ShowBills(mcllFilter)
    End If
End Function

Private Function ChekcDepositDelPrivs(ByVal strNO As String, Optional ShowMsgbox As Boolean = True, _
                                                           Optional blnErr As Boolean = False) As Boolean
    'blnErr-是否异常单据
    '功能:检查预交退款(包括异常退款)权限
    If Trim(strNO) = "" Then Exit Function
    If is代收款(strNO) Then
         If InStr(mstrPrivs, "代收款退款") = 0 Then
            If ShowMsgbox Then MsgBox "你没有权限进行代收款退款操作！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf InStr(mstrPrivs, "预交退款") = 0 Then
        If ShowMsgbox Then MsgBox "你没有权限进行预交退款操作！", vbInformation, gstrSysName
        Exit Function
    Else
        If blnErr Then ChekcDepositDelPrivs = True: Exit Function
        If HaveSpare(strNO) = 0 And InStr(mstrPrivs, "预交结清退款") = 0 Then
            If ShowMsgbox Then MsgBox "该病人已没有预交余额,你没有权限作废这张单据！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If HaveBalance(strNO) <> 0 Then
            If ShowMsgbox Then MsgBox "该笔预交已经被病人使用,你不能作废这张单据！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    ChekcDepositDelPrivs = True
End Function

Public Function ExcuteCationMoney_Del() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:押金退款
    '编制:焦博
    '日期:2020-06-22 11:17:33
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strNO As String, str操作员 As String
    With vsDeposit
        strNO = .TextMatrix(.Row, .ColIndex("单据号"))
        str操作员 = .TextMatrix(.Row, .ColIndex("操作员"))
    End With
    
    If strNO = "" Then
        MsgBox "当前没有记录可以退款！", vbExclamation, gstrSysName
        Exit Function
    End If
        
    '单据权限
    If Not BillOperCheck(6, str操作员, _
        CDate(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("操作时间"))), "退款") Then Exit Function
    
    If Val(vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("金额"))) < 0 Then
        MsgBox "该缴款记录金额为负,表示退款,不能执行该操作！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 6, Me.Caption) Then Exit Function
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    ExcuteCationMoney_Del = frmCautionMoney.zlShowEdit(Me, 2, mstrPrivs, mlngModule, strNO)
End Function


