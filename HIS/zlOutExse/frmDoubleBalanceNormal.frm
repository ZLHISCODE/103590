VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDoubleBalanceNormal 
   BorderStyle     =   0  'None
   Caption         =   "frmDoubleBalanceNormal"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBalanceInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   5655
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfBalanceInfo 
         Height          =   1845
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
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
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   7455
      ScaleHeight     =   1065
      ScaleWidth      =   2550
      TabIndex        =   7
      Top             =   4335
      Width           =   2550
      Begin VSFlex8Ctl.VSFlexGrid vsfBalance 
         Height          =   1845
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
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
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   7515
      ScaleHeight     =   945
      ScaleWidth      =   2415
      TabIndex        =   6
      Top             =   3135
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfInvoice 
         Height          =   1845
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
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
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   4395
      ScaleHeight     =   2115
      ScaleWidth      =   2745
      TabIndex        =   2
      Top             =   3420
      Width           =   2745
      Begin XtremeSuiteControls.TabControl tabInfo 
         Height          =   2010
         Left            =   -705
         TabIndex        =   3
         Top             =   -375
         Width           =   2820
         _Version        =   589884
         _ExtentX        =   4974
         _ExtentY        =   3545
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   1065
      ScaleHeight     =   2805
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   3345
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1845
         Left            =   300
         TabIndex        =   5
         Top             =   330
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
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
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   3015
      ScaleHeight     =   2640
      ScaleWidth      =   3120
      TabIndex        =   0
      Top             =   540
      Width           =   3120
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1800
         Left            =   510
         TabIndex        =   4
         Top             =   270
         Width           =   2550
         _cx             =   4498
         _cy             =   3175
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDoubleBalanceNormal.frx":0000
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
      Left            =   1725
      Top             =   1710
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDoubleBalanceNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmFilter As New frmDoubleBalanceFilter
Public mblnNOMoved As Boolean
Private mblnPrinting As Boolean

Public Sub MakeFilter(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String)
    Call mfrmFilter.InitFilter(Me, lngModul, strPrivs)
End Sub

Public Sub ReadData(ByVal intType As Integer, ByVal strPrivs As String, Optional ByVal lngPatiID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取保险补充计算记录
    '编制:刘尔旋
    '入参:intType-读取记录的方式，0为使用过滤条件读取，1为使用IDKIND条件读取
    '日期:2014-9-11
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMain As ADODB.Recordset, strTable As String
    Dim strFilter As String, strInvoice As String, strSQLtmp As String
    Dim DatBegin As Date, DatEnd As Date, blnMoved As Boolean
    Dim str红票已打印 As String
    
    On Error GoTo ErrHand
    str红票已打印 = ",Nvl((Select 1" & vbNewLine & _
            "            From 票据打印内容 M, 票据使用明细 N" & vbNewLine & _
            "            Where m.Id = n.打印id And n.性质 = 1 And n.原因 = 6 And m.No = a.no And Rownum < 2), 0) As 红票已打印" & vbNewLine
    If intType = 0 Then
        If mfrmFilter.mblnInit = True Then
            With mfrmFilter
                DatBegin = .dtpBegin.Value
                DatEnd = .dtpEnd.Value
                blnMoved = zlDatabase.DateMoved(IIf(DatBegin < DatEnd, DatBegin, DatEnd))
                strFilter = ""
                strFilter = strFilter & IIf(.txt姓名.Text = "", "", " And C.姓名=[3] ")
                strFilter = strFilter & IIf(.cbo操作员.Text = "所有收费员", "", " And A.操作员姓名=[4] ")
                strFilter = strFilter & IIf(.txt门诊号.Text = "", "", " And C.门诊号=[5] ")
                strFilter = strFilter & IIf(.txt住院号.Text = "", "", " And C.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [6]) ")
                If .txtNOBegin.Text <> "" Then
                    If .txtNoEnd.Text <> "" Then
                        strFilter = strFilter & " And A.NO Between [7] And [8] "
                    Else
                        strFilter = strFilter & " And A.NO=[7] "
                    End If
                End If
                strInvoice = ""
                If (.txtFactBegin.Text <> "" And .txtFactEnd.Text <> "") Or (.txtFactBegin.Text <> "" And .txtFactEnd.Text = "") Then
                    '无需根据票据号判断,直接根据单据的登记时间判断
                    strSQLtmp = IIf(.txtFactEnd.Text = "", " =[9] ", " Between [9] And [10] ")
                    strInvoice = "Select A.NO" & _
                    " From 票据打印内容 A,票据使用明细 B" & _
                    " Where A.数据性质=1 And A.ID=B.打印ID And B.票种=1 And B.性质=1" & _
                    " And B.号码 " & strSQLtmp
                End If
                If strInvoice <> "" Then strFilter = strFilter & " And A.NO In (" & strInvoice & ") "
                'strFilter = strFilter & IIf(.chkDelRecord, " And A.记录状态 <> 0 ", " ")
                '病人来源
                If .opt病人(0).Value Then '门诊
                    strFilter = strFilter & " And  b.门诊标志 in (1,4)"
                ElseIf .opt病人(1).Value Then '住院
                    strFilter = strFilter & " And  b.门诊标志 =2"
                Else '门诊和住院
                End If
                If blnMoved Then
                    strTable = zlGetFullFieldsTable("费用补充记录", 2, "", True, "A", True)
                Else
                    strTable = "费用补充记录 A"
                End If
                strSQL = " Select A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, Sum(B.结帐金额), A.操作员姓名, A.登记时间, A.结算序号, Max(A.记录状态) As 退费标志,A.病人ID,A.实际票号" & str红票已打印 & _
                         " From " & strTable & ", 门诊费用记录 B, 病人信息 C " & _
                         " Where A.登记时间 Between [1] And [2] And A.病人ID=C.病人ID And A.收费结帐ID=B.结帐ID And Nvl(A.费用状态,0)=0 " & _
                         "      And A.记录状态 In (1,3) And Not Exists (Select 1 From 费用补充记录 Where 结算序号=A.结算序号 And 记录状态=2) " & strFilter & _
                         " Group By A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, A.操作员姓名, A.登记时间, A.结算序号,A.病人ID,A.实际票号"
                         
                If .chkDelRecord.Value = 1 Then
                    strSQL = strSQL & " Union " & _
                        "   Select NO, 附加标志, 姓名, 性别, 年龄, -1 * Sum(实收金额) As 结算金额, 操作员姓名, 登记时间, 结算序号, 2 As 退费标志, 病人ID, 实际票号" & str红票已打印 & _
                        "   From (Select Distinct a.No, Decode(Nvl(a.附加标志, 0), 1, '挂号', '收费') As 附加标志, b.姓名, b.性别, b.年龄, b.实收金额," & _
                        "                a.操作员姓名, a.登记时间, a.结算序号, b.No As 单据号, b.结帐id, a.病人id, a.实际票号" & _
                        "          From  " & strTable & ", 门诊费用记录 B, 病人预交记录 D, 病人信息 C" & _
                        "          Where a.登记时间 Between [1] And [2] And a.结算序号 = d.结算序号 And b.结帐id = d.结帐id And a.病人id = c.病人id" & _
                        "                And Nvl(a.费用状态, 0) = 0 And a.记录状态 = 2 " & strFilter & ") A" & _
                        "   Group By NO, 附加标志, 姓名, 性别, 年龄, 操作员姓名, 登记时间, 结算序号,病人ID,实际票号"
                End If
                strSQL = strSQL & " Order By 登记时间 Desc"
                Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, DatBegin, DatEnd, _
                                                    .txt姓名.Text, zlStr.NeedName(.cbo操作员.Text), .txt门诊号.Text, .txt住院号.Text, _
                                                    .txtNOBegin.Text, .txtNoEnd.Text, .txtFactBegin.Text, .txtFactEnd.Text)
                Set vsfMain.DataSource = rsMain
                Call SetMain
            End With
        Else
            strSQL = " Select A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, Sum(B.结帐金额), A.操作员姓名, A.登记时间, A.结算序号, Max(A.记录状态) As 退费标志, A.病人ID, A.实际票号" & str红票已打印 & _
                     " From 费用补充记录 A, 门诊费用记录 B " & _
                     " Where Trunc(A.登记时间)=Trunc(Sysdate) And A.收费结帐ID=B.结帐ID And A.记录状态 In (1,3) And Nvl(A.费用状态,0)=0 " & _
                     "      And Not Exists (Select 1 From 费用补充记录 Where 结算序号=A.结算序号 And 记录状态=2) " & _
                     " And A.操作员姓名=[1] " & _
                     " Group By A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, A.操作员姓名, A.登记时间, A.结算序号,A.病人ID,A.实际票号" & _
                     " Order By 登记时间 Desc"
            Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
            Set vsfMain.DataSource = rsMain
            Call SetMain
        End If
    End If
    If intType = 1 Then
        '使用IDKIND条件读取
        strSQL = " Select A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, Sum(B.结帐金额), A.操作员姓名, A.登记时间, A.结算序号, Max(A.记录状态) As 退费标志,A.病人ID,A.实际票号" & str红票已打印 & _
                 " From 费用补充记录 A, 门诊费用记录 B " & _
                 " Where A.病人ID= [1] And A.收费结帐ID=B.结帐ID And A.记录状态 In (1,3) And Nvl(A.费用状态,0)=0 " & _
                 "      And Not Exists (Select 1 From 费用补充记录 Where 结算序号=A.结算序号 And 记录状态=2) " & _
                 IIf(InStr(strPrivs, "所有操作员") > 0, "", " And A.操作员姓名=[2] ") & _
                 " Group By A.No, Decode(Nvl(A.附加标志,0),1,'挂号','收费'), B.姓名, B.性别, B.年龄, A.操作员姓名, A.登记时间, A.结算序号,A.病人ID,A.实际票号"
        strSQL = strSQL & " Union " & _
                        "   Select NO, 附加标志, 姓名, 性别, 年龄, -1 * Sum(实收金额) As 结算金额, 操作员姓名, 登记时间, 结算序号, 2 As 退费标志, 病人ID, 实际票号" & str红票已打印 & _
                        "   From (Select Distinct a.No, Decode(Nvl(a.附加标志, 0), 1, '挂号', '收费') As 附加标志, b.姓名, b.性别, b.年龄, b.实收金额," & _
                        "                a.操作员姓名, a.登记时间, a.结算序号, b.No As 单据号, b.结帐id, a.病人id,a.实际票号" & _
                        "          From 费用补充记录 A, 门诊费用记录 B, 病人预交记录 D, 病人信息 C" & _
                        "          Where A.病人ID= [1] And a.结算序号 = d.结算序号 And b.结帐id = d.结帐id And a.病人id = c.病人id" & _
                        "                And Nvl(a.费用状态, 0) = 0 And a.记录状态 = 2" & _
                        IIf(InStr(strPrivs, "所有操作员") > 0, "", " And A.操作员姓名=[2]") & ") A" & _
                        "   Group By NO, 附加标志, 姓名, 性别, 年龄, 操作员姓名, 登记时间, 结算序号, 病人ID, 实际票号" & _
                        " Order By 登记时间 Desc"
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, UserInfo.姓名)
        Set vsfMain.DataSource = rsMain
        Call SetMain
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建DOCKINGPANEL控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 2000, 4000, DockTopOf)
        objPanel.Handle = picMain.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption

        Set objPanel = .CreatePane(2, 1700, 2000, DockBottomOf, objPanel)
        objPanel.Handle = picDetail.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption

        Set objPanel = .CreatePane(3, 1000, 2000, DockRightOf, objPanel)
        objPanel.Handle = picInfo.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetInvoiceList()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "ID,1,0|票据号,4,1000|使用原因,4,1000|使用时间,4,1200|使用人,1,1000"
    With vsfInvoice
        .Redraw = flexRDNone
        
        If .Rows = 1 Then .Rows = 2
        
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            If Not Visible Then
                .ColWidth(i) = Split(varData(i), ",")(2)
                If .ColWidth(i) = 0 Then .ColHidden(i) = True
            End If
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfInvoice, App.ProductName & "\" & Me.Name)
        
        .HighLight = flexHighlightWithFocus
        .RowHeight(-1) = 300: .RowHeight(0) = 350

        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetMain()
    Dim i As Long
    With vsfMain
        zl_vsGrid_Para_Restore 1124, vsfMain, Me.Caption, "费用信息列表", True
        .RowHeight(0) = 350
        If .Rows = 1 Then .Rows = 2
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("退费标志"))) = 2 Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
            ElseIf Val(.TextMatrix(i, .ColIndex("退费标志"))) = 3 Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
            Else
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
            End If
            .TextMatrix(i, .ColIndex("结算金额")) = Format(.TextMatrix(i, .ColIndex("结算金额")), gstrDec)
            .TextMatrix(i, .ColIndex("结算时间")) = Format(.TextMatrix(i, .ColIndex("结算时间")), "yyyy-mm-dd hh:mm:ss")
            .RowHeight(i) = 300
        Next i
        If .Rows >= 2 Then .Select 1, 1
    End With
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "单据号,1,0|序号,1,0|开单科室,1,0|开单人,1,0|费别,1,0|类别,4,800|名称,1,2000|商品名,1,2000|" & _
            "规格,1,1200|单位,4,500|数量,7,800|单价,7,1000|应收金额,7,1000|实收金额,7,1000|执行科室,4,1000|" & _
            "类型,4,1000|说明,1,1800|记录状态,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        'Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
        If .TextMatrix(1, .ColIndex("单据号")) <> "" Then Call DetailSplitGroup
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                .RowHeight(i) = 300
            End If
        Next i
        
        If gTy_System_Para.byt药品名称显示 = 0 Then
            .ColHidden(.ColIndex("名称")) = False
            .ColHidden(.ColIndex("商品名")) = True
        End If
        If gTy_System_Para.byt药品名称显示 = 1 Then
            .ColHidden(.ColIndex("名称")) = True
            .ColHidden(.ColIndex("商品名")) = False
        End If
        If gTy_System_Para.byt药品名称显示 = 2 Then
            .ColHidden(.ColIndex("名称")) = False
            .ColHidden(.ColIndex("商品名")) = False
        End If
    End With
End Sub

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对费用列表信息进行分组显示
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsfDetail
        For i = 0 To .COLS - 1
            If i < .ColIndex("类别") And i > .ColIndex("说明") Then
                .ColHidden(i) = True
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("实收金额"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("应收金额"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("类别")
        .OutlineCol = .ColIndex("类别")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("类别")) = strTemp

                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("单据号"))
                 strTemp = strTemp & Space(2) & "费别:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("费别"))
                 strTemp = strTemp & Space(2) & "开单部门:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单科室"))
                 strTemp = strTemp & Space(2) & "开单人:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单人"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("类别"), i, .ColIndex("类别")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                 For j = 0 To .COLS - 1
                    If j < .ColIndex("应收金额") Then
                        If j >= .ColIndex("类别") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    ElseIf .ColIndex("实收金额") = j Then
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("应收金额") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("单价")) = Format(Val(.TextMatrix(i, .ColIndex("单价"))), gstrDec)
                .TextMatrix(i, .ColIndex("数量")) = Format(Val(.TextMatrix(i, .ColIndex("数量"))), gstrDec)
                .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(.TextMatrix(i, .ColIndex("应收金额"))), gstrDec)
                .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(.TextMatrix(i, .ColIndex("实收金额"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("类别"))
        Call .AutoSize(.ColIndex("单价"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("应收金额") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call SetDockingPanel
    Call SetTab
    Call SetMain
    Call SetInvoiceList
    Call SetDetail
    Call SetExtendInfo
    Call SetBalanceList
End Sub

Private Sub SetBalanceList()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Long
    Dim varData As Variant
    
    strHead = "结算方式,4,1000|金额,7,1000|结算号码,4,1000|摘要,1,1200|卡号,1,1000|交易流水号,1,1000|交易说明,1,1200|性质,1,0"
    
    With vsfBalance
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("结算方式")) Like "*误差*" Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                strTemp = Val(.TextMatrix(i, .ColIndex("金额")))
                If InStr(strTemp, ".") = 0 Then
                    strAcc = "0.00"
                Else
                    strTemp = Split(strTemp, ".")(1)
                    strAcc = "0."
                    If Len(strTemp) < 2 Then
                        strAcc = "0.00"
                    Else
                        For j = 1 To Len(strTemp)
                            strAcc = strAcc & "0"
                        Next j
                    End If
                End If
                .TextMatrix(i, .ColIndex("金额")) = Format(.TextMatrix(i, .ColIndex("金额")), strAcc)
            Else
                If .TextMatrix(i, .ColIndex("金额")) <> "" Then .TextMatrix(i, .ColIndex("金额")) = Format(.TextMatrix(i, .ColIndex("金额")), "0.00")
            End If
            .RowHeight(i) = 300
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        
        .Redraw = True
    End With
End Sub

Private Sub SetExtendInfo()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|结算方式,1,0|名称,1,0|金额,1,0|项目,1,1200|内容,1,2000|交易流水号,1,0"
    
    With vsfBalanceInfo
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
        Next
        If .Rows < 2 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        .RowHeight(0) = 350
        
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
                 
                 For j = 0 To .COLS - 1
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
        For j = 0 To .COLS - 1
            .MergeCol(j) = True
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetTab()
    On Error GoTo errHandle
    With tabInfo
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        'Set .PaintManager.Font = txtSendFeeNO.Font
        .InsertItem 1, "票据信息", picInvoice.hWnd, 0
        .InsertItem 2, "结算信息", picBalance.hWnd, 0
        .InsertItem 3, "结算关联信息", picBalanceInfo.hWnd, 0
        .Item(0).Selected = True
        .Item(2).Visible = False
        .PaintManager.BoldSelected = True
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save 1124, vsfMain, Me.Caption, "费用信息列表", True
End Sub

Private Sub PicDetail_Resize()
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = picDetail.Height
        .Width = picDetail.Width
    End With
End Sub

Private Sub picInfo_Resize()
    With tabInfo
        .Top = 0
        .Left = 0
        .Width = picInfo.Width
        .Height = picInfo.Height
    End With
End Sub

Private Sub picInvoice_Resize()
    With vsfInvoice
        .Top = 0
        .Left = 0
        .Height = picInvoice.Height
        .Width = picInvoice.Width
    End With
End Sub

Private Sub picBalance_Resize()
    With vsfBalance
        .Top = 0
        .Left = 0
        .Height = picBalance.Height
        .Width = picBalance.Width
    End With
End Sub

Private Sub picBalanceInfo_Resize()
    With vsfBalanceInfo
        .Top = 0
        .Left = 0
        .Height = picBalanceInfo.Height
        .Width = picBalanceInfo.Width
    End With
End Sub

Private Sub picMain_Resize()
    With vsfMain
        .Top = 0
        .Left = 0
        .Height = picMain.Height
        .Width = picMain.Width
    End With
End Sub

Private Sub vsfBalance_GotFocus()
    SetActiveList vsfBalance
End Sub

Private Sub vsfBalanceInfo_GotFocus()
    SetActiveList vsfBalanceInfo
End Sub

Private Sub vsfDetail_GotFocus()
    SetActiveList vsfDetail
End Sub

Private Sub vsfInvoice_GotFocus()
    SetActiveList vsfInvoice
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结算序号")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结算序号"))), _
                    vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("类型")) = "挂号")
    Call ReadInVoice(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结算序号"))))
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结算序号"))))
    Call ReadBalanceInfo(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结算序号"))))
End Sub

Private Sub ReadBalance(ByVal lngBalanceID As Long)
    Dim strSQL As String, rsBalance As ADODB.Recordset
    
    strSQL = _
    "Select Decode(Mod(a.记录性质, 10), 1, '冲预存款', Nvl(a.结算方式, '未结金额')) As 结算方式, Sum(a.冲预交) As 冲预交," & vbNewLine & _
    "       Decode(Mod(Max(a.记录性质), 10), 1, '', Max(a.结算号码)) As 结算号码," & vbNewLine & _
    "       Decode(Mod(Max(a.记录性质), 10), 1, '', Max(a.摘要)) As 摘要," & vbNewLine & _
    "       Decode(Mod(Max(a.记录性质), 10), 1, '', Max(a.卡号)) As 卡号," & vbNewLine & _
    "       Decode(Mod(Max(a.记录性质), 10), 1, '', Max(a.交易流水号)) As 交易流水号," & vbNewLine & _
    "       Decode(Mod(Max(a.记录性质), 10), 1, '', Max(a.交易说明)) As 交易说明" & vbNewLine & _
    "From 病人预交记录 A" & vbNewLine & _
    "Where a.结算序号 = [1]" & vbNewLine & _
    "Group By Decode(Mod(a.记录性质, 10), 1, '冲预存款', Nvl(a.结算方式, '未结金额'))"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    
    vsfBalance.Redraw = False
    vsfBalance.Clear
    vsfBalance.Rows = 2
    If Not rsBalance.EOF Then
        Set vsfBalance.DataSource = rsBalance
    End If
    Call SetBalanceList
    vsfBalance.Redraw = True
End Sub

Private Sub ReadBalanceInfo(ByVal lngBalanceID As Long)
    Dim strSQL As String, rsInfo As ADODB.Recordset
    
    strSQL = _
        "Select b.交易id || '_' || b.原预交id As ID, a.结算方式, Max(c.名称) As 名称, Sum(Nvl(-1 * f.金额, a.冲预交)) As 金额, b.交易项目, b.交易内容," & vbNewLine & _
        "       Max(Nvl(f.交易流水号, a.交易流水号)) As 交易流水号" & vbNewLine & _
        "From 病人预交记录 A, 三方结算交易 B, 医疗卡类别 C, 病人预交记录 E, 三方退款信息 F" & vbNewLine & _
        "Where a.Id = b.交易id And a.卡类别id = c.Id(+) And a.结算序号 = [1]" & vbNewLine & _
        "      And b.原预交id = e.Id(+) And e.结帐id = f.记录id(+) And f.结帐id(+) = [1]" & vbNewLine & _
        "Group By b.交易id, b.原预交id, a.结算方式, b.交易项目, b.交易内容" & vbNewLine & _
        "Order By ID"
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = Replace(strSQL, "三方结算交易", "H三方结算交易")
        strSQL = Replace(strSQL, "三方退款信息", "H三方退款信息")
    End If
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    
    Set vsfBalanceInfo.DataSource = rsInfo
    If rsInfo.RecordCount = 0 Then
        '没有第三方交易记录时，隐藏分页
        tabInfo.Item(2).Visible = False
        If tabInfo.Selected.Index = 2 Then tabInfo.Item(0).Selected = True
    Else
        tabInfo.Item(2).Visible = True
    End If
    Call SetExtendInfo
End Sub

Private Sub ReadDetail(ByVal lngBalanceID As Long, ByVal bln挂号补充 As Boolean)
    Dim strSQL As String, rsDetail As ADODB.Recordset
    Dim blnDel As Boolean
    mblnNOMoved = zlDatabase.NOMoved("费用补充记录", vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("结算单号")))
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("退费标志"))) = 2
    If blnDel Then
        strSQL = _
                " Select NO As 单据号, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, " & _
                "       Sum(数量) As 数量, 单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, 执行科室, Max(类型) As 类型, Max(说明),Max(状态), Min(退费状态)" & vbNewLine & _
                " From (Select a.结帐ID,D1.名称 as 开单科室,A.开单人,a.No,C.名称 as 类别,Nvl(E.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格," & _
                        IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
                "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                        IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
                "       a.费别,To_Char(Sum(A.标准单价)" & _
                        IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
                "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型,Max(Decode(A.记录状态,2,'第'||ABS(A.执行状态)||'次退费',Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行'))) As 说明," & _
                "       Max(A.记录状态) As 状态,Min(A.记录状态) As 退费状态, Nvl(a.价格父号, a.序号) As 序号" & _
                " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 D1,收费项目别名 E,收费项目别名 E1,药品规格 X," & _
                "       (Select Distinct 结帐ID From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 Where 结算序号= [1]) F" & _
                " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
                "       And Mod(A.记录性质,10)=[2] And A.结帐ID = F.结帐ID " & _
                "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                "       And A.收费细目ID=E1.收费细目ID(+) And A.开单部门ID=D1.ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
                " Group by a.结帐id, D1.名称, a.开单人, a.费别,a.No,Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称),E1.名称 , B.规格,A.计算单位,D.名称," & _
                "       Nvl(A.费用类型,B.费用类型),X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1) )" & _
                " Group By NO, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, 单价, 执行科室 Having Sum(数量) <> 0" & _
                " Order By 单据号, 序号"
    Else
        strSQL = _
                " Select NO As 单据号, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, " & _
                "       Sum(数量) As 数量, 单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, 执行科室, Max(类型) As 类型, Max(说明),Max(状态), Min(退费状态)" & vbNewLine & _
                " From (Select a.结帐ID,D1.名称 as 开单科室,A.开单人,a.No,C.名称 as 类别,Nvl(E.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格," & _
                        IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
                "       To_Char(Avg(Nvl(A.付数,1)*A.数次)" & _
                        IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
                "       a.费别,To_Char(Sum(A.标准单价)" & _
                        IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
                "       To_Char(Sum(A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
                "       To_Char(Sum(A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
                "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型,Max(Decode(A.记录状态,2,'第'||ABS(A.执行状态)||'次退费',Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行'))) As 说明," & _
                "       Max(A.记录状态) As 状态,Min(A.记录状态) As 退费状态, Nvl(a.价格父号, a.序号) As 序号" & _
                " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 D1,收费项目别名 E,收费项目别名 E1,药品规格 X," & _
                "       (Select Distinct 收费结帐ID As 结帐ID From " & IIf(mblnNOMoved, "H", "") & "费用补充记录 Where 结算序号= [1]) F" & _
                " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
                "       And Mod(A.记录性质,10)=[2] And A.结帐ID = F.结帐ID " & _
                "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                "       And A.收费细目ID=E1.收费细目ID(+) And A.开单部门ID=D1.ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
                " Group by a.结帐id, D1.名称, a.开单人, a.费别,a.No,Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称),E1.名称 , B.规格,A.计算单位,D.名称," & _
                "       Nvl(A.费用类型,B.费用类型),X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1) )" & _
                " Group By NO, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, 单价, 执行科室 Having Sum(数量) <> 0" & _
                " Order By 单据号, 序号"
    End If
    Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID, IIf(bln挂号补充, 4, 1))
    vsfDetail.Redraw = False
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    vsfDetail.Redraw = True
End Sub

Private Sub ReadInVoice(ByVal lngBalanceID As Long)
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    
    strSQL = _
    " Select Distinct B.ID, B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票发出') as 使用原因," & _
    " To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
    " From 票据打印内容 A,票据使用明细 B," & _
            "(Select Distinct NO From 费用补充记录 Where 结算序号= [1]) C" & _
    " Where A.数据性质=1 And A.ID=B.打印ID" & _
    " And B.票种=1 And A.NO=C.NO" & _
    " Order by ID"
    
    Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    Set vsfInvoice.DataSource = rsInvoice
    Call SetInvoiceList
End Sub

Private Sub vsfMain_DblClick()
    Call frmReplenishTheBalanceManage.ViewBalance(0)
End Sub

Private Sub vsfMain_GotFocus()
    SetActiveList vsfMain
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is vsfMain Then
        vsfMain.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfBalance Then
        vsfBalance.BackColorSel = &HC0C0C0
        vsfMain.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfBalanceInfo Then
        vsfBalanceInfo.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfDetail Then
        vsfDetail.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfInvoice Then
        vsfInvoice.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub vsfMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    With vsfMain
        'If .TextMatrix(1, .ColIndex("结算序号")) = "" Then Exit Sub
        If Button = 2 Then
            If Y <= 300 Then
                Exit Sub
            End If
'            intRow = Y \ 300
'            If intRow <= .Rows - 1 Then
'                If .Enabled And .Visible Then .SetFocus
'                .Select intRow, 0
'            End If
            Call frmReplenishTheBalanceManage.ShowPopup
        End If
    End With
End Sub

Public Property Get zlGetFeeState() As Integer
    '------------------------------------------------------------
    '功能：获取当前选中记录行的退费标志
    '编制：冉俊明
    '时间：2014-12-11
    '返回：0-无记录,1-收费记录,2-退费记录,3-已被退费的收费记录
    '------------------------------------------------------------
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("结算序号")) = "" Then Exit Property
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("退费标志")) = "" Then
        zlGetFeeState = 0
    Else
        zlGetFeeState = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("退费标志")))
    End If
End Property

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    Dim i As Long, lngCurrentRow As Long
    Dim objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    
    With vsfMain
        If .Rows = 1 Then Exit Sub
        If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("结算序号"))) = 0 Then Exit Sub
    End With
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = "保险补充结算正常结算记录清单"
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    '由于打印控件不能识别列隐藏属性
    With vsfMain
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .COLS - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Then
                .ColWidth(i) = 0
            End If
        Next
    End With

    Err = 0: On Error GoTo ErrHand:
    mblnPrinting = True
    lngCurrentRow = vsfMain.Row
    Set objPrint.Body = vsfMain
    If bytFunc = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    
    '恢复
    With vsfMain
        For i = 0 To .COLS - 1
            If .ColHidden(i) = True Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    vsfMain.Row = lngCurrentRow
    mblnPrinting = False
    Exit Sub
ErrHand:
    mblnPrinting = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
