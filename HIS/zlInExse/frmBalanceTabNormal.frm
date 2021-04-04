VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmBalanceTabNormal 
   BorderStyle     =   0  'None
   Caption         =   "frmBalanceTabNormal"
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
      TabIndex        =   11
      Top             =   5655
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfBalanceInfo 
         Height          =   1845
         Left            =   0
         TabIndex        =   5
         Top             =   15
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
         ForeColorSel    =   -2147483640
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
      TabIndex        =   10
      Top             =   4335
      Width           =   2550
      Begin VSFlex8Ctl.VSFlexGrid vsfBalance 
         Height          =   1845
         Left            =   0
         TabIndex        =   4
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
         ForeColorSel    =   -2147483640
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
      TabIndex        =   9
      Top             =   3135
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfInvoice 
         Height          =   1845
         Left            =   0
         TabIndex        =   3
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
         ForeColorSel    =   -2147483640
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
      TabIndex        =   7
      Top             =   3420
      Width           =   2745
      Begin XtremeSuiteControls.TabControl tabInfo 
         Height          =   2010
         Left            =   -705
         TabIndex        =   8
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3345
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1845
         Left            =   300
         TabIndex        =   2
         Top             =   300
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
         ForeColorSel    =   -2147483640
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
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   3015
      ScaleHeight     =   2640
      ScaleWidth      =   3120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   3120
      Begin VB.PictureBox picMainControl 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   540
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   12
         Top             =   300
         Width           =   210
         Begin VB.Image imgMainControl 
            Height          =   195
            Left            =   0
            Picture         =   "frmBalanceTabNormal.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1800
         Left            =   510
         TabIndex        =   1
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
         BackColorSel    =   16772055
         ForeColorSel    =   8
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBalanceTabNormal.frx":054E
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
Attribute VB_Name = "frmBalanceTabNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnPrint As Boolean
Private mfrmFilter As New frmBalanceFilter
Private mblnNOMoved As Boolean

Public Sub MakeFilter(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String)
    Call mfrmFilter.InitFilter(Me, lngModul, strPrivs)
End Sub

Public Sub ReadData(ByVal intTYPE As Integer, ByVal strPrivs As String, Optional ByVal lngPatiID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取结帐记录
    '编制:刘尔旋
    '入参:intType-读取记录的方式，0为使用过滤条件读取，1为使用IDKIND条件读取
    '日期:2015-01-06
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMain As ADODB.Recordset, strTable As String
    Dim strFilter As String, strInvoice As String, strSQLtmp As String
    Dim DatBegin As Date, DatEnd As Date, blnMoved As Boolean, strSource As String
    Dim i As Integer, str来源 As String, strUpgrade As String
    On Error GoTo ErrHand
    If intTYPE = 0 Then
        If mfrmFilter.mblnInit = True Then
            With mfrmFilter
                DatBegin = .dtpBegin.Value
                DatEnd = .dtpEnd.Value
                blnMoved = zlDatabase.DateMoved(IIf(DatBegin < DatEnd, DatBegin, DatEnd))
                strFilter = " And A.收费时间 Between [1] And [2] "
                strFilter = strFilter & IIf(.txt姓名.Text = "", "", " And C.姓名=[3] ")
                strFilter = strFilter & IIf(.cbo操作员.Text = "所有结帐人", "", " And A.操作员姓名=[4] ")
                strFilter = strFilter & IIf(.txt门诊号.Text = "", "", " And C.门诊号=[5] ")
                strFilter = strFilter & IIf(.txt住院号.Text = "", "", " And C.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[6]) ")
                If Not (.chkType(0).Value = 1 And .chkType(1).Value = 1) Then
                    If .chkType(0).Value = 1 Then
                        strFilter = strFilter & " And A.记录状态 In (1,3) "
                    Else
                        strFilter = strFilter & " And A.记录状态 = 2 "
                    End If
                End If
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
                    If blnMoved Then
                        strInvoice = "" & _
                         "(  Select A.NO" & _
                         "   From 票据打印内容 A,票据使用明细 B" & _
                         "   Where A.数据性质=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.打印ID And B.票种=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.性质=1" & _
                         "         And B.号码 " & strSQLtmp & ")  Union All" & _
                         " (Select A.NO " & _
                         " From H票据打印内容 A,H票据使用明细 B" & _
                         " Where A.数据性质=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.打印ID And B.票种=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.性质=1" & _
                         " And B.号码 " & strSQLtmp & ")"
                    Else
                        strInvoice = "Select A.NO" & _
                        " From 票据打印内容 A,票据使用明细 B" & _
                        " Where A.数据性质=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.打印ID And B.票种=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.性质=1" & _
                        " And B.号码 " & strSQLtmp
                    End If
                End If
                If strInvoice <> "" Then strFilter = strFilter & " And A.NO In (" & strInvoice & ") "
                
                For i = 0 To .chkFeeOrigin.Count - 1
                    strSource = strSource & IIf(.chkFeeOrigin(i).Value = 1, 1, 0) '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
                Next
                If strSource = "" Then strSource = "0100"
                str来源 = ""
                For i = 1 To Len(strSource)
                    If Mid(strSource, i, 1) = 1 Then
                        str来源 = str来源 & "," & Choose(i, 1, 2, 4, 3)  '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
                    End If
                Next
                If str来源 <> "" Then str来源 = Mid(str来源, 2)
                If str来源 = "" Then str来源 = "-1"
                
                strTable = "" & _
                "   Select A.ID ,1 as 住院标志,0 as 门诊标志,A.NO,A.实际票号,A.病人ID," & _
                "           B.病人ID as 费用病人ID,Nvl(D.性别,C.性别) as 性别,Nvl(D.年龄,C.年龄) as 年龄,A.开始日期,A.结束日期,Max(A.记录状态) As 记录状态,Sum(B.结帐金额) As 结帐金额," & _
                "           A.操作员姓名,A.收费时间,A.中途结帐,A.原因 as 合约单位,A.结帐类型,A.主页ID,Max(Decode(a.结帐金额,Null,1,0)) As 需要升级" & _
                "   From 病人结帐记录 A,住院费用记录 B,病人信息 C,病案主页 D " & _
                "   Where A.ID=B.结帐ID and  B.病人ID=C.病人ID And A.病人ID =D.病人ID(+) And A.主页ID = D.主页ID(+) And (A.结算状态 = 2 Or A.结算状态 Is Null) " & _
                        IIf(strSource = "1111", "", " And Instr(',' || [11] || ',',',' || Nvl(B.门诊标志,0) || ',') > 0 ") & strFilter & _
                "   Group By A.ID,A.NO,A.实际票号,A.病人ID,B.病人ID,Nvl(D.性别,C.性别),Nvl(D.年龄,C.年龄),A.开始日期,A.结束日期,A.操作员姓名,A.收费时间,A.中途结帐,A.原因,A.结帐类型,A.主页ID "

                Select Case strSource
                Case "1010", "1000", "0010"  '门诊
                    strTable = Replace(strTable, "住院费用记录", "门诊费用记录")
                    strTable = Replace(strTable, "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
                Case "0101", "0001", "0100" '住院
                    '已经存在
                Case Else '门诊和住院
                    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
                End Select
                
                If blnMoved Then
                    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(Replace(strTable, "病人结帐记录", "H病人结帐记录"), "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
                End If
                
                strTable = "Select ID, Decode(Max(住院标志), 1, Decode(Max(门诊标志), 1, 3, 2), 1) As 标志, NO, 实际票号, 病人id, 费用病人id, 性别, 年龄, 开始日期, 结束日期," & vbNewLine & _
                            "              记录状态, Sum(结帐金额) As 结帐金额, 操作员姓名, 收费时间, 中途结帐, 合约单位, 结帐类型, 主页id, 需要升级" & vbNewLine & _
                            "       From (" & strTable & ") " & _
                            "       Group By ID, NO, 实际票号, 病人id, 费用病人id, 性别, 年龄, 开始日期, 结束日期, 记录状态, 操作员姓名, 收费时间, 中途结帐, 合约单位, 结帐类型, 主页id, 需要升级 "
                
                strSql = _
                " Select A.ID 结帐ID,标志,decode(A.结帐类型,1,'门诊结帐',2,'住院结帐','') As 结帐类型 ,Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√') as 医保,A.NO as 单据号,A.实际票号 as 票据号," & _
                "        Decode(A.病人ID,Null,' ',A.病人ID) 病人ID,Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)) 门诊号,Decode(A.病人ID,Null,' ',Decode(A.主页ID,Null,C.住院号,P.住院号)) 住院号," & _
                "        Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名) 姓名,Decode(A.病人ID,Null,' ',A.性别) 性别," & _
                "        Decode(A.病人ID,Null,' ',A.年龄) 年龄,Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别)) as 费别," & _
                "        To_Char(A.开始日期,'YYYY-MM-DD') as 开始日期,To_Char(A.结束日期,'YYYY-MM-DD') as 结束日期," & _
                "        To_Char(Decode(A.记录状态,2,-1,1) *A.结帐金额,'999999999" & gstrDec & "') as 结帐金额," & _
                "        A.操作员姓名 as 操作员,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(Nvl(A.中途结帐,0),1,'√',' ') 中途结帐,A.记录状态 as 记录状态,A.需要升级" & _
                " From ( " & strTable & ") A,病人信息 C,病案主页 P,合约单位 Q,人员表 N" & _
                " Where  A.费用病人ID=C.病人ID And A.操作员姓名=N.姓名 " & _
                "        And A.费用病人ID=P.病人ID(+) And Nvl(A.主页ID,0)=P.主页ID(+) And C.合同单位ID=Q.ID(+)" & _
                "       And (N.站点='" & gstrNodeNo & "' Or N.站点 is Null)" & vbNewLine
                
                strSql = strSql & " Order by 收费时间 Desc,单据号 Desc"
                
                Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, DatBegin, DatEnd, _
                                                    .txt姓名.Text, zlStr.NeedName(.cbo操作员.Text), .txt门诊号.Text, .txt住院号.Text, _
                                                    .txtNOBegin.Text, .txtNoEnd.Text, .txtFactBegin.Text, .txtFactEnd.Text, str来源)
                Do While Not rsMain.EOF
                    If Val(NVL(rsMain!需要升级)) = 1 Then
                        strUpgrade = "Zl_病人结帐记录_Upgrade(" & rsMain!结帐ID & ")"
                        zlDatabase.ExecuteProcedure strUpgrade, Me.Caption
                    End If
                    rsMain.MoveNext
                Loop
                If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
                Set vsfMain.DataSource = rsMain
                
                Call SetMain
            End With
        Else
            DatBegin = Format(zlDatabase.Currentdate, "YYYY-MM-DD 00:00:00")
            DatEnd = Format(zlDatabase.Currentdate, "YYYY-MM-DD 23:59:59")
            strFilter = " And A.收费时间 Between [1] And [2] "
            strFilter = strFilter & " And A.操作员姓名=[3] "
            strSource = "1111"
            str来源 = ""
            For i = 1 To Len(strSource)
                If Mid(strSource, i, 1) = 1 Then
                    str来源 = str来源 & "," & Choose(i, 1, 2, 4, 3)  '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
                End If
            Next
            If str来源 <> "" Then str来源 = Mid(str来源, 2)
            If str来源 = "" Then str来源 = "-1"
            
            strTable = "" & _
            "   Select " & IIf(strSource = "1111", "", " /*+cardinality(L1,10)*/ ") & " A.ID ,1 as 住院标志,0 as 门诊标志,A.NO,A.实际票号,A.病人ID,B.病人ID as 费用病人ID,Nvl(D.性别,C.性别) as 性别,Nvl(D.年龄,C.年龄) as 年龄,A.开始日期,A.结束日期,Max(A.记录状态) As 记录状态,Sum(B.结帐金额) As 结帐金额,A.操作员姓名,A.收费时间,A.中途结帐,A.原因 as 合约单位,A.结帐类型,A.主页ID,Max(Decode(a.结帐金额,Null,1,0)) As 需要升级 " & _
            "   From 病人结帐记录 A,住院费用记录 B,病人信息 C,病案主页 D " & _
                    IIf(strSource = "1111", "", ",Table(Cast(f_Num2list('" & str来源 & "') As Zltools.t_Numlist)) L1") & _
            "   Where A.ID=B.结帐ID and  B.病人ID=C.病人ID And A.病人ID =D.病人ID(+) And A.主页ID = D.主页ID(+) And (A.结算状态 = 2 Or A.结算状态 Is Null) " & _
                    IIf(strSource = "1111", "", " And nvl(B.门诊标志,0)=L1.Column_Value ") & strFilter & _
            "   Group By A.ID,A.NO,A.实际票号,A.病人ID,B.病人ID,Nvl(D.性别,C.性别),Nvl(D.年龄,C.年龄),A.开始日期,A.结束日期,A.操作员姓名,A.收费时间,A.中途结帐,A.原因,A.结帐类型,A.主页ID "
                    
            Select Case strSource
            Case "1010", "1000", "0010"  '门诊
                strTable = Replace(strTable, "住院费用记录", "门诊费用记录")
                strTable = Replace(strTable, "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
            Case "0101", "0001", "0100" '住院
                '已经存在
            Case Else '门诊和住院
                strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
            End Select
            
            strTable = "Select ID, Decode(Max(住院标志), 1, Decode(Max(门诊标志), 1, 3, 2), 1) As 标志, NO, 实际票号, 病人id, 费用病人id, 性别, 年龄, 开始日期, 结束日期," & vbNewLine & _
                            "              记录状态, Sum(结帐金额) As 结帐金额, 操作员姓名, 收费时间, 中途结帐, 合约单位, 结帐类型, 主页id, 需要升级" & vbNewLine & _
                            "       From (" & strTable & ") " & _
                            "       Group By ID, NO, 实际票号, 病人id, 费用病人id, 性别, 年龄, 开始日期, 结束日期, 记录状态, 操作员姓名, 收费时间, 中途结帐, 合约单位, 结帐类型, 主页id, 需要升级 "
                
            strSql = _
                " Select A.ID 结帐ID,标志,decode(A.结帐类型,1,'门诊结帐',2,'住院结帐','') As 结帐类型 ,Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√') as 医保,A.NO as 单据号,A.实际票号 as 票据号," & _
                "        Decode(A.病人ID,Null,' ',A.病人ID) 病人ID,Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)) 门诊号,Decode(A.病人ID,Null,' ',Decode(A.主页ID,Null,C.住院号,P.住院号)) 住院号," & _
                "        Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名) 姓名,Decode(A.病人ID,Null,' ',A.性别) 性别," & _
                "        Decode(A.病人ID,Null,' ',A.年龄) 年龄,Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别)) as 费别," & _
                "        To_Char(A.开始日期,'YYYY-MM-DD') as 开始日期,To_Char(A.结束日期,'YYYY-MM-DD') as 结束日期," & _
                "        To_Char(Decode(A.记录状态,2,-1,1) *A.结帐金额,'999999999" & gstrDec & "') as 结帐金额," & _
                "        A.操作员姓名 as 操作员,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(Nvl(A.中途结帐,0),1,'√',' ') 中途结帐,A.记录状态 as 记录状态,A.需要升级" & _
                " From ( " & strTable & ") A,病人信息 C,病案主页 P,合约单位 Q,人员表 N" & _
                " Where  A.费用病人ID=C.病人ID And A.操作员姓名=N.姓名 " & _
                "        And A.费用病人ID=P.病人ID(+) And Nvl(A.主页ID,0)=P.主页ID(+) And C.合同单位ID=Q.ID(+)" & _
                "       And (N.站点='" & gstrNodeNo & "' Or N.站点 is Null)" & vbNewLine
            strSql = strSql & " Order by 收费时间 Desc,单据号 Desc"
            
            Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, DatBegin, DatEnd, UserInfo.姓名)
            Do While Not rsMain.EOF
                If Val(NVL(rsMain!需要升级)) = 1 Then
                    strUpgrade = "Zl_病人结帐记录_Upgrade(" & rsMain!结帐ID & ")"
                    zlDatabase.ExecuteProcedure strUpgrade, Me.Caption
                End If
                rsMain.MoveNext
            Loop
            If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
            Set vsfMain.DataSource = rsMain
            Call SetMain
        End If
    End If
    
    If intTYPE = 1 Then
        strFilter = " And C.病人ID = [1]  "
        strSource = "1111"
        str来源 = ""
        For i = 1 To Len(strSource)
            If Mid(strSource, i, 1) = 1 Then
                str来源 = str来源 & "," & Choose(i, 1, 2, 4, 3)  '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
            End If
        Next
        If str来源 <> "" Then str来源 = Mid(str来源, 2)
        If str来源 = "" Then str来源 = "-1"
        
        strTable = "" & _
        "   Select " & IIf(strSource = "1111", "", " /*+cardinality(L1,10)*/ ") & " A.ID ,1 as 住院标志,0 as 门诊标志,A.NO,A.实际票号,A.病人ID,B.病人ID as 费用病人ID,Nvl(D.性别,C.性别) as 性别,Nvl(D.年龄,C.年龄) as 年龄,A.开始日期,A.结束日期,A.记录状态,B.结帐金额,A.操作员姓名,A.收费时间,A.中途结帐,A.原因 as 合约单位,A.结帐类型,A.主页ID,Decode(a.结帐金额,Null,1,0) As 需要升级 " & _
        "   From 病人结帐记录 A,住院费用记录 B,病人信息 C,病案主页 D " & _
                IIf(strSource = "1111", "", ",Table(Cast(f_Num2list([11]) As Zltools.t_Numlist)) L1") & _
        "   Where A.ID=B.结帐ID and  B.病人ID=C.病人ID And A.病人ID =D.病人ID(+) And A.主页ID = D.主页ID(+) And Nvl(A.结算状态,2) = 2 " & _
                IIf(strSource = "1111", "", " And nvl(B.门诊标志,0)=L1.Column_Value ") & strFilter
                
        Select Case strSource
        Case "1010", "1000", "0010"  '门诊
            strTable = Replace(strTable, "住院费用记录", "门诊费用记录")
            strTable = Replace(strTable, "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
        Case "0101", "0001", "0100" '住院
            '已经存在
        Case Else '门诊和住院
            strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
        End Select
        strTable = "Select ID, Decode(Max(住院标志), 1, Decode(Max(门诊标志), 1, 3, 2), 1) As 标志, NO, 实际票号, 病人id, 费用病人id, 性别, 年龄, 开始日期, 结束日期," & vbNewLine & _
                            "              记录状态, Sum(结帐金额) As 结帐金额, 操作员姓名, 收费时间, 中途结帐, 合约单位, 结帐类型, 主页id, 需要升级" & vbNewLine & _
                            "       From (" & strTable & ") " & _
                            "       Group By ID, NO, 实际票号, 病人id, 费用病人id, 性别, 年龄, 开始日期, 结束日期, 记录状态, 操作员姓名, 收费时间, 中途结帐, 合约单位, 结帐类型, 主页id, 需要升级 "
                
        '使用IDKIND条件读取
        strSql = _
                " Select A.ID 结帐ID,标志,decode(A.结帐类型,1,'门诊结帐',2,'住院结帐','') As 结帐类型 ,Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√') as 医保,A.NO as 单据号,A.实际票号 as 票据号," & _
                "        Decode(A.病人ID,Null,' ',A.病人ID) 病人ID,Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)) 门诊号,Decode(A.病人ID,Null,' ',Decode(A.主页ID,Null,C.住院号,P.住院号)) 住院号," & _
                "        Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名) 姓名,Decode(A.病人ID,Null,' ',A.性别) 性别," & _
                "        Decode(A.病人ID,Null,' ',A.年龄) 年龄,Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别)) as 费别," & _
                "        To_Char(A.开始日期,'YYYY-MM-DD') as 开始日期,To_Char(A.结束日期,'YYYY-MM-DD') as 结束日期," & _
                "        To_Char(Sum(Decode(A.记录状态,2,-1,1) *A.结帐金额),'999999999" & gstrDec & "') as 结帐金额," & _
                "        A.操作员姓名 as 操作员,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(Nvl(A.中途结帐,0),1,'√',' ') 中途结帐,Max(A.记录状态) as 记录状态,A.需要升级" & _
                " From ( " & strTable & ") A,病人信息 C,病案主页 P,合约单位 Q,人员表 N" & _
                " Where  A.费用病人ID=C.病人ID And A.操作员姓名=N.姓名 " & _
                "        And A.费用病人ID=P.病人ID(+) And Nvl(A.主页ID,0)=P.主页ID(+) And C.合同单位ID=Q.ID(+)" & _
                "       And (N.站点='" & gstrNodeNo & "' Or N.站点 is Null)" & vbNewLine & _
                " Group by A.ID,标志,Decode(a.结帐类型, 1, '门诊结帐', 2, '住院结帐', ''),Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√'),A.NO,A.实际票号,Decode(A.病人ID,Null,' ',A.病人ID),Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)),Decode(A.病人ID,Null,' ',Decode(A.主页ID,Null,C.住院号,P.住院号))," & _
                "           Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名),Decode(A.病人ID,Null,' ',A.性别),Decode(A.病人ID,Null,' ',A.年龄),Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别))," & _
                "           To_Char(A.开始日期,'YYYY-MM-DD'),To_Char(A.结束日期,'YYYY-MM-DD')," & _
                "           A.需要升级,A.操作员姓名,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS'),Decode(Nvl(A.中途结帐,0),1,'√',' ')"
        strSql = strSql & " Order by 收费时间 Desc,单据号 Desc"
        
        Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatiID)
        Do While Not rsMain.EOF
            If Val(NVL(rsMain!需要升级)) = 1 Then
                strUpgrade = "Zl_病人结帐记录_Upgrade(" & rsMain!结帐ID & ")"
                zlDatabase.ExecuteProcedure strUpgrade, Me.Caption
            End If
            rsMain.MoveNext
        Loop
        If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
        Set vsfMain.DataSource = rsMain
        Call SetMain
    End If
    
    If Not rsMain Is Nothing Then
        If rsMain.RecordCount <> 0 Then
            frmManageBalance.stbThis.Panels(2).Text = "当前共有" & rsMain.RecordCount & "条结账记录,合计:" & Format(GetTotal, gstrDec) & "元"
        Else
            frmManageBalance.stbThis.Panels(2).Text = ""
        End If
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
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfInvoice, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        .Row = 1: .Col = 0: .ColSel = .Cols - 1

        .Redraw = True
    End With
End Sub

Private Function GetTotal() As Double
    Dim dblTotal As Double
    Dim i As Integer
    With vsfMain
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("记录状态"))) <> 2 Then
                dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("结帐金额")))
            Else
                dblTotal = dblTotal - Val(.TextMatrix(i, .ColIndex("结帐金额")))
            End If
        Next i
    End With
    GetTotal = dblTotal
End Function

Private Sub SetMain()
    Dim i As Long, strHead As String
    Dim dblTotal As Double
    
    strHead = "结帐ID,1,0|标志,1,0|  结帐类型,4,1100|医保,4,500|单据号,4,850|票据号,4,850|病人ID,1,750|门诊号,1,750|住院号,1,750|姓名,4,800|性别,4,500|年龄,4,500|费别,4,750|开始日期,4,1000|结束日期,4,1000|结帐金额,7,850|操作员,4,800|收费时间,4,1850|中途结帐,4,800|记录状态,1,0"
    vsfBalance.Clear 1
    vsfInvoice.Clear 1
    vsfBalanceInfo.Clear 1
    vsfDetail.Clear 1
    With vsfMain
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            If .TextMatrix(0, i) = "结帐ID" Then .ColHidden(i) = True
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "结帐ID" Or .ColKey(i) = "标志" Or .ColKey(i) = "记录状态" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "单据号" Or .ColKey(i) = "收费时间" Or .ColKey(i) = "  结帐类型" Then .ColData(i) = "-1|1"
        Next
        
        zl_vsGrid_Para_Restore 1137, vsfMain, Me.Name, "结帐信息列表", False
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 2 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            ElseIf Val(.TextMatrix(i, .ColIndex("记录状态"))) = 3 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
            End If
            .RowHeight(i) = 300
        Next i
        
        .Select 1, 1

        .Redraw = True
    End With
    Call vsfMain_GotFocus
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "类型,4,750|单据号,4,850|开单科室,1,850|项目,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|费目,1,850|婴儿费,4,650|结帐金额,7,850|费用时间,1,1850"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vsfMain.Cell(flexcpForeColor, vsfMain.Row, 1, vsfMain.Row, 1)
        Next i
        
        .Redraw = True
    End With
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
    
    strHead = "类型,4,800|单据号,4,1000|金额,7,1000|结算方式,1,1200|结算号码,1,1000"
    
    With vsfBalance
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vsfMain.Cell(flexcpForeColor, vsfMain.Row, 1, vsfMain.Row, 1)
            .TextMatrix(i, .ColIndex("金额")) = Formatex(Val(.TextMatrix(i, .ColIndex("金额"))), 6, , , 2)
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        .Redraw = True
    End With
End Sub

Private Sub SetExtendInfo()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|结算方式,1,0|名称,1,0|金额,1,0|项目,1,1200|内容,1,2000|交易流水号,1,0"
    
    With vsfBalanceInfo
        .HighLight = flexHighlightWithFocus
        .Redraw = False
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
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        
        .RowHeight(0) = 350
        '.Row = 1: .Col = 0: .ColSel = .COLS - 1
        .Redraw = True
        
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
    mblnPrint = False
    zl_vsGrid_Para_Save 1137, vsfMain, Me.Name, "结帐信息列表", False
End Sub

Private Sub picDetail_Resize()
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
    With picMainControl
        .Top = 90
        .Left = 75
    End With
End Sub

Private Sub imgMainControl_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picMainControl.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picMainControl.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfMain, lngLeft, lngTop, picMainControl.Height)
    zl_vsGrid_Para_Save 1137, vsfMain, Me.Name, "结帐信息列表", False
End Sub

Private Sub picMainControl_Click()
    Call imgMainControl_Click
End Sub

Private Sub vsfBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfBalance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfBalance_GotFocus()
    zl_VsGridGotFocus vsfBalance, &HFFC0C0
End Sub

Private Sub vsfBalance_LostFocus()
    zl_VsGridLOSTFOCUS vsfBalance, , vsfBalance.Cell(flexcpForeColor, vsfBalance.Row, vsfBalance.Col)
End Sub

Private Sub vsfBalanceInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfBalanceInfo, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfBalanceInfo_GotFocus()
    zl_VsGridGotFocus vsfBalanceInfo, &HFFC0C0
End Sub

Private Sub vsfBalanceInfo_LostFocus()
    zl_VsGridLOSTFOCUS vsfBalanceInfo
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfDetail, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfDetail_GotFocus()
    zl_VsGridGotFocus vsfDetail, &HFFC0C0
End Sub

Private Sub vsfDetail_LostFocus()
    zl_VsGridLOSTFOCUS vsfDetail, , vsfDetail.Cell(flexcpForeColor, vsfDetail.Row, vsfDetail.Col)
End Sub

Private Sub vsfInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfInvoice_GotFocus()
    zl_VsGridGotFocus vsfInvoice, &HFFC0C0
End Sub

Private Sub vsfInvoice_LostFocus()
    zl_VsGridLOSTFOCUS vsfInvoice, , vsfInvoice.Cell(flexcpForeColor, vsfInvoice.Row, vsfInvoice.Col)
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If NewRow <> 0 And OldRow <> 0 Then zl_VsGridRowChange vsfMain, OldRow, NewRow, OldCol, NewCol
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID"))))
    Call ReadInVoice(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("单据号")))
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID"))))
    Call ReadBalanceInfo(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID"))))
End Sub

Private Sub ReadBalance(ByVal lngBalanceID As Long)
    Dim strSql As String, rsBalance As ADODB.Recordset, blnDel As Boolean
    
    If mblnPrint Then Exit Sub
    
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("记录状态"))) = 2
    strSql = _
            "Select Decode(Substr(记录性质,Length(记录性质),1),1,'冲预交',2,'补款') as 类型," & _
            " NO as 单据号," & IIf(blnDel, "-1*", "") & "冲预交 as 金额," & _
            " 结算方式,结算号码 From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 Where 结帐ID=[1] And 冲预交 <> 0 " & _
            " Order by 类型 Desc,NO Desc,结算方式"

    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
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
    Dim strSql As String, rsInfo As ADODB.Recordset
    
    If mblnPrint Then Exit Sub
    
    strSql = _
        "Select b.交易id || '_' || b.原预交id As ID, a.结算方式, Max(c.名称) As 名称, Sum(Nvl(-1 * f.金额, a.冲预交)) As 金额," & vbNewLine & _
        "       b.交易项目, b.交易内容, Max(Nvl(f.交易流水号, a.交易流水号)) As 交易流水号" & vbNewLine & _
        "From 病人预交记录 A, 三方结算交易 B, 医疗卡类别 C, 病人预交记录 E, 三方退款信息 F" & vbNewLine & _
        "Where a.Id = b.交易id And a.卡类别id = c.Id(+) And a.结帐id = [1] And a.记录性质 <> 1" & vbNewLine & _
        "      And b.原预交id = e.Id(+) And e.id = f.记录id(+) And f.结帐id(+) =  [1]" & vbNewLine & _
        "Group By b.交易id, b.原预交id, a.结算方式, b.交易项目, b.交易内容" & vbNewLine & _
        "Order By ID"
    Set rsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
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

Private Sub ReadDetail(ByVal lngBalanceID As Long)
    Dim strSql As String, rsDetail As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim blnDel As Boolean, strDec As String, int来源 As Integer
    
    If mblnPrint Then Exit Sub
    
    mblnNOMoved = zlDatabase.NOMoved("病人结帐记录", vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("单据号")))
    int来源 = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("标志")))
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("记录状态"))) = 2
    strDec = gstrDec
    If lngBalanceID <> 0 Then
        Select Case int来源
        Case 1 '门诊
            strSql = "Select Max(Length(Abs(结帐金额) - Trunc(Abs(结帐金额))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 Where 结帐ID=[1]"
        Case 2 '住院
            strSql = "Select Max(Length(Abs(结帐金额) - Trunc(Abs(结帐金额))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 Where 结帐ID=[1]"
        Case Else
            
            strSql = "Select Length(Abs(结帐金额) - Trunc(Abs(结帐金额)))  as  declen From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 Where 结帐ID=[1] Union ALL " & _
                     "Select Length(Abs(结帐金额) - Trunc(Abs(结帐金额)))   as  declen  From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 Where 结帐ID=[1]"
            strSql = "Select Max(declen)-1 as declen  From ( " & strSql & ")"
        End Select
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        If rsTmp.RecordCount > 0 Then
            If Len(strDec) < Len("0." & String(rsTmp!declen, "0")) Then
                strDec = "0." & String(rsTmp!declen, "0")
            End If
        End If
    End If
    
    Select Case int来源
    Case 1  '门诊
        strSql = " (Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,Sum(结帐金额) As 结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A where A.结帐ID=[1] Group By 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,收据费目,婴儿费,发生时间) A "
        'strSQL = IIf(mblnNOMoved, "H", "") & "门诊费用记录 A "
    Case 2  '住院
        strSql = IIf(mblnNOMoved, "H", "") & "住院费用记录 A"
    Case Else '门诊和住院
        strSql = " (Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A where A.结帐ID=[1] Union ALL " & _
                   " Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A where A.结帐ID=[1] )  A"
    End Select
    
    strSql = _
    "   Select Decode(门诊标志,1,'门诊',4,'门诊',Decode(Nvl(A.主页ID,0),0,'','第'||Nvl(A.主页ID,0)||'次')) As 类型," & _
    "         A.NO as 单据号,Nvl(B.名称,'未知') as 开单科室,Nvl(E.名称,D.名称) as 项目," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & _
    "       A.收据费目 as 费目,Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费," & _
    "       To_Char(" & IIf(blnDel, "-1*", "") & "A.结帐金额,'999999999" & strDec & "') as 结帐金额," & _
    "       To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 费用时间" & _
    " From " & strSql & ",部门表 B,收费项目目录 D,收费项目别名 E" & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.开单部门ID=B.ID And A.收费细目ID=D.ID" & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
    "       And A.结帐ID=[1]" & _
    " Order by 类型 Desc,费用时间 Desc,单据号 Desc,A.序号"
    Set rsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    
End Sub

Private Sub ReadInVoice(ByVal strNO As String)
    Dim strSql As String, rsInvoice As ADODB.Recordset
    
    If mblnPrint Then Exit Sub
    
    strSql = _
    " Select b.Id, b.号码 As 票据号," & vbNewLine & _
    " Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因," & vbNewLine & _
    "    To_Char(b.使用时间, 'MM-DD HH24:MI') As 使用时间, b.使用人" & vbNewLine & _
    " From 票据打印内容 A, 票据使用明细 B" & vbNewLine & _
    " Where a.数据性质 = 3 And a.Id = b.打印id And a.No = [1]" & vbNewLine & _
    " Order By ID"

    Set rsInvoice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)

    vsfInvoice.Redraw = False
    vsfInvoice.Clear 1
    vsfInvoice.Rows = 2
    If Not rsInvoice.EOF Then
        Set vsfInvoice.DataSource = rsInvoice
    End If
    Call SetInvoiceList
    vsfInvoice.Redraw = True
End Sub

Private Sub vsfMain_DblClick()
    Call frmManageBalance.ViewBalance(0)
End Sub

Private Sub vsfMain_GotFocus()
    If vsfMain.Row = 0 Then Exit Sub
    zl_VsGridGotFocus vsfMain, &HFFC0C0
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

Private Sub vsfMain_LostFocus()
    zl_VsGridLOSTFOCUS vsfMain, , vsfMain.Cell(flexcpForeColor, vsfMain.Row, vsfMain.Col)
End Sub

Private Sub vsfMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    With vsfMain
        If Button = 2 Then
            If Y <= 300 Then
                Exit Sub
            End If
            Call frmManageBalance.ShowPopup
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
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("结帐ID")) = "" Then Exit Property
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("记录状态")) = "" Then
        zlGetFeeState = 0
    Else
        zlGetFeeState = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("记录状态")))
    End If
End Property

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘尔旋
    '日期:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    lngRow = vsfMain.Row
    Set vsBill = vsfMain: strTittle = GetUnitName & "病人结帐记录信息"
    mblnPrint = True
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If vsBill Is Nothing Then Exit Sub
    '由于打印控件不能识别列隐藏属性
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
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
    With vsBill
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    
    mblnPrint = False
    vsfMain.Select lngRow, 1
    Exit Sub
ErrHand:
    mblnPrint = False
    vsfMain.Select lngRow, 1
    If ErrCenter = 1 Then Resume
End Sub
