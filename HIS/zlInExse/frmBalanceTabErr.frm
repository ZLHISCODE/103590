VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmBalanceTabErr 
   BorderStyle     =   0  'None
   Caption         =   "frmBalanceTabErr"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   6570
      ScaleHeight     =   2805
      ScaleWidth      =   2715
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4245
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfBalance 
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
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   885
      ScaleHeight     =   2805
      ScaleWidth      =   2715
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4515
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1845
         Left            =   300
         TabIndex        =   3
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
      Left            =   4425
      ScaleHeight     =   2640
      ScaleWidth      =   3120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   510
      Width           =   3120
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1830
         Left            =   510
         TabIndex        =   1
         Top             =   270
         Width           =   1800
         _cx             =   3175
         _cy             =   3228
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBalanceTabErr.frx":0000
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
      Left            =   975
      Top             =   540
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBalanceTabErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnPrint As Boolean
Private mblnNOMoved As Boolean
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
        Set objPanel = .CreatePane(1, 2000, 2000, DockTopOf)
        objPanel.Handle = picMain.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        Set objPanel = .CreatePane(2, 1700, 1000, DockBottomOf, objPanel)
        objPanel.Handle = picDetail.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        Set objPanel = .CreatePane(3, 1000, 1000, DockRightOf, objPanel)
        objPanel.Handle = picBalance.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        .Options.HideClient = True
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is vsfMain Then
        vsfMain.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfBalance Then
        vsfBalance.BackColorSel = &HC0C0C0
        vsfMain.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfDetail Then
        vsfDetail.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnPrint = False
End Sub

Private Sub vsfBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfBalance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfBalance_GotFocus()
    zl_VsGridGotFocus vsfBalance, &HFFC0C0
End Sub

Private Sub vsfBalance_LostFocus()
    zl_VsGridLOSTFOCUS vsfBalance
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfDetail, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfDetail_GotFocus()
    zl_VsGridGotFocus vsfDetail, &HFFC0C0
End Sub

Private Sub vsfDetail_LostFocus()
    zl_VsGridLOSTFOCUS vsfDetail
End Sub

Private Sub vsfMain_GotFocus()
    zl_VsGridGotFocus vsfMain, &HFFC0C0
End Sub

Private Sub vsfMain_LostFocus()
    zl_VsGridLOSTFOCUS vsfMain
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

Public Sub ReadData()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取异常结算记录
    '编制:刘尔旋
    '日期:2015-01-06
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMain As ADODB.Recordset, strFilter As String, strTable As String
    Dim dtStartDate As Date, dtEndDate As Date, blnAll As Boolean
    Select Case frmManageBalance.cboDate.ListIndex
        Case 0 '所有异常
            dtStartDate = CDate(Format("1900-01-01", "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format("3000-01-01", "yyyy-mm-dd") & " 23:59:59")
        Case 1 '今日
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 '前一天至今日
            dtStartDate = CDate(Format(DateAdd("d", -1, frmManageBalance.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3 '前二天至今日
            dtStartDate = CDate(Format(DateAdd("d", -2, frmManageBalance.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  '前一周至今日
            dtStartDate = CDate(Format(DateAdd("d", -7, frmManageBalance.dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 5  '本月
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm") & "-01 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case Else
            dtStartDate = CDate(Format(frmManageBalance.dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(frmManageBalance.dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    strFilter = " And A.记录状态 In (1,3) And A.收费时间 Between [1] And [2] And A.操作员姓名 = [3] "
    strTable = "" & _
            "   Select A.ID ,1 as 住院标志,0 as 门诊标志,A.NO,A.实际票号,A.病人ID,B.病人ID as 费用病人ID,Nvl(D.性别,C.性别) as 性别,Nvl(D.年龄,C.年龄) as 年龄 ,A.开始日期,A.结束日期,Max(A.记录状态) As 记录状态,Sum(B.结帐金额) As 结帐金额,A.操作员姓名,A.收费时间,A.中途结帐,A.原因 as 合约单位,A.结帐类型 " & _
            "   From 病人结帐记录 A,住院费用记录 B,病人信息 C,病案主页 D " & _
            "   Where A.ID=B.结帐ID and  B.病人ID=C.病人ID And A.病人ID =D.病人ID(+) And A.主页ID = D.主页ID(+) And A.结算状态 = 1 And Not Exists (Select 1 From 病人结帐记录 Where NO = a.No And 记录状态 = 2) " & strFilter & _
            "   Group By A.ID ,A.NO,A.实际票号,A.病人ID,B.病人ID,Nvl(D.性别,C.性别),Nvl(D.年龄,C.年龄),A.开始日期,A.结束日期,A.操作员姓名,A.收费时间,A.中途结帐,A.原因,A.结帐类型 "

    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
    
    strSql = _
            " Select A.ID 结帐ID,decode(住院标志,1,decode(门诊标志,1,3,2),1) as 标志,decode(A.结帐类型,1,'门诊结帐',2,'住院结帐','') As 结帐类型 ,Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√') as 医保,A.NO as 单据号,A.实际票号 as 票据号," & _
            "        Decode(A.病人ID,Null,' ',A.病人ID) 病人ID,Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)) 门诊号,Decode(A.病人ID,Null,' ',C.住院号) 住院号," & _
            "        Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名) 姓名,Decode(A.病人ID,Null,' ',A.性别) 性别," & _
            "        Decode(A.病人ID,Null,' ',A.年龄) 年龄,Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别)) as 费别," & _
            "        To_Char(A.开始日期,'YYYY-MM-DD') as 开始日期,To_Char(A.结束日期,'YYYY-MM-DD') as 结束日期," & _
            "        To_Char(Decode(A.记录状态,2,-1,1) *A.结帐金额,'999999999" & gstrDec & "') as 结帐金额," & _
            "        A.操作员姓名 as 操作员,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(Nvl(A.中途结帐,0),1,'√',' ') 中途结帐,A.记录状态 as 记录状态" & _
            " From ( " & strTable & ") A,病人信息 C,病案主页 P,合约单位 Q,人员表 N" & _
            " Where  A.费用病人ID=C.病人ID And A.操作员姓名=N.姓名 " & _
            "        And C.病人ID=P.病人ID(+) And Nvl(C.主页ID,0)=P.主页ID(+) And C.合同单位ID=Q.ID(+)" & _
            "       And (N.站点='" & gstrNodeNo & "' Or N.站点 is Null)" & vbNewLine
            
    strSql = strSql & " Order by 收费时间 Desc,单据号 Desc"
    
    Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    Set vsfMain.DataSource = rsMain
    If rsMain.RecordCount <> 0 Then
        frmManageBalance.tabMain.Item(1).Caption = "异常结算记录(" & rsMain.RecordCount & ")"
        frmManageBalance.stbThis.Panels(2).Text = "当前共有" & rsMain.RecordCount & "条异常结算记录,合计:" & Format(GetTotal, gstrDec) & "元"
    Else
        frmManageBalance.tabMain.Item(1).Caption = "异常结算记录"
        frmManageBalance.stbThis.Panels(2).Text = ""
    End If
    Call SetMain
End Sub

Private Sub ReadBalance(Optional ByVal lngBalanceID As Long)
    Dim strSql As String, i As Long, rsBalance As ADODB.Recordset
    
    If mblnPrint Then Exit Sub
    
    strSql = _
        " Select Nvl(A.结算方式,'未结金额') As 结算方式,Sum(A.冲预交) As 冲预交," & _
        "       Decode(Nvl(A.校对标志,0),0,'√',2,'√','×') As 标志,Nvl(B.性质,0) As 性质 " & _
        " From 病人预交记录 A,结算方式 B " & _
        " Where A.结帐ID = [1] And A.结算方式=B.名称(+)" & _
        " Group By Nvl(A.结算方式,'未结金额'),Nvl(A.校对标志,0),Nvl(B.性质,0)" & _
        " Having Sum(A.冲预交) <> 0 Order By 性质"
    
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsfBalance.Redraw = False
    vsfBalance.Clear
    vsfBalance.Rows = 2
    If Not rsBalance.EOF Then
        Set vsfBalance.DataSource = rsBalance
    End If
    Call SetBalance
    vsfBalance.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReadDetail(ByVal lngBalanceID As Long)
    Dim strSql As String, rsDetail As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim blnDel As Boolean, strDec As String, int来源 As Integer
    
    If mblnPrint Then Exit Sub
    
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
        strSql = " (Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A where A.结帐ID=[1] ) A "
        'strSQL = IIf(mblnNOMoved, "H", "") & "门诊费用记录 A "
    Case 2  '住院
        strSql = IIf(mblnNOMoved, "H", "") & "住院费用记录 A"
    Case Else '门诊和住院
        strSql = " (Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A where A.结帐ID=[1] Union ALL " & _
                   " Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A where A.结帐ID=[1] )  A"
    End Select
    
    strSql = _
    "   Select Decode(门诊标志,1,'门诊',4,'门诊','第'||Nvl(A.主页ID,0)||'次') as 住院," & _
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
    " Order by 住院 Desc,费用时间 Desc,单据号 Desc,A.序号"
    Set rsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "住院,4,750|单据号,4,850|开单科室,1,850|项目,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|费目,1,850|婴儿费,4,650|结帐金额,7,850|费用时间,1,1850"
    
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
        Next i
        
        .Redraw = True
    End With
End Sub

Private Sub SetBalance()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Long
    Dim varData As Variant
    
    strHead = "结算方式,4,1200|结算金额,7,1000|结算是否成功,4,1200|性质,1,0"
    
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
            .TextMatrix(i, .ColIndex("结算金额")) = Formatex(Val(.TextMatrix(i, .ColIndex("结算金额"))), 6, , , 2)
            .RowHeight(i) = 300
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1

        .Redraw = True
    End With
End Sub

Private Function GetTotal() As Double
    Dim dblTotal As Double
    Dim i As Integer
    With vsfMain
        For i = 1 To .Rows - 1
            dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("结帐金额")))
        Next i
    End With
    GetTotal = dblTotal
End Function

Private Sub SetMain()
    Dim i As Long, strHead As String
    Dim dblTotal As Double
    
    strHead = "结帐ID,1,0|标志,1,0|结帐类型,4,800|医保,4,500|单据号,4,850|票据号,4,850|病人ID,1,750|门诊号,1,750|住院号,1,750|姓名,4,800|性别,4,500|年龄,4,500|费别,4,750|开始日期,4,1000|结束日期,4,1000|结帐金额,7,850|操作员,4,800|收费时间,4,1850|中途结帐,4,800|记录状态,1,0"
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    vsfBalance.Clear 1
    vsfBalance.Rows = 2
    With vsfMain
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            If .TextMatrix(0, i) = "结帐ID" Then
                .ColHidden(i) = True
            Else
                .ColHidden(i) = False
            End If
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "结帐ID" Or .ColKey(i) = "标志" Or .ColKey(i) = "记录状态" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "单据号" Or .ColKey(i) = "收费时间" Then .ColData(i) = "1|0"
        Next
        
        .RowHeight(0) = 350
        If .Rows = 1 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .TextMatrix(i, .ColIndex("结帐金额")) = Format(.TextMatrix(i, .ColIndex("结帐金额")), gstrDec)
        Next i
        
        If .Rows >= 2 Then .Select 1, 1
        If .Enabled And .Visible Then .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Call SetDockingPanel
    Call SetMain
    Call SetBalance
    Call SetDetail
End Sub

Private Sub picDetail_Resize()
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = picDetail.Height
        .Width = picDetail.Width
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If OldRow <> 0 And NewRow <> 0 Then zl_VsGridRowChange vsfMain, OldRow, NewRow, OldCol, NewCol
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID"))))
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("结帐ID"))))
End Sub

Private Sub picBalance_Resize()
    With vsfBalance
        .Top = 0
        .Left = 0
        .Height = picBalance.Height
        .Width = picBalance.Width
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

Private Sub vsfMain_DblClick()
    Call frmManageBalance.ViewBalance(1)
End Sub

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
    Set vsBill = vsfMain: strTittle = GetUnitName & "异常结帐记录信息"
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
