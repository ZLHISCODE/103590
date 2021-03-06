VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStuffInSel 
   Caption         =   "备货入库单选择"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "frmStuffInSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11775
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   11775
      TabIndex        =   12
      Top             =   6600
      Width           =   11775
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "全选(&A)"
         Height          =   380
         Left            =   105
         TabIndex        =   15
         ToolTipText     =   "快键:CTRL+A"
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdALLCls 
         Caption         =   "全清(&S)"
         Height          =   380
         Left            =   1365
         TabIndex        =   14
         ToolTipText     =   "快键:CTRL+C"
         Top             =   165
         Width           =   1250
      End
      Begin VB.Frame fraBottomSplit 
         Height          =   30
         Left            =   -210
         TabIndex        =   13
         Top             =   0
         Width           =   12405
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   380
         Left            =   8865
         TabIndex        =   7
         Top             =   165
         Width           =   1250
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   380
         Left            =   10125
         TabIndex        =   8
         Top             =   165
         Width           =   1250
      End
   End
   Begin VB.PictureBox picSeach 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   11775
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   11775
      Begin VB.Frame fraSearch 
         Height          =   960
         Left            =   90
         TabIndex        =   10
         Top             =   -15
         Width           =   11235
         Begin VB.CommandButton cmdSel 
            Caption         =   "过滤(&F)"
            Height          =   380
            Left            =   8130
            TabIndex        =   5
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txtNO 
            Height          =   330
            Left            =   870
            TabIndex        =   1
            Top             =   355
            Width           =   1770
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   345
            Left            =   4830
            TabIndex        =   3
            Top             =   348
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   180748291
            CurrentDate     =   40528
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   345
            Left            =   6525
            TabIndex        =   4
            Top             =   348
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   180748291
            CurrentDate     =   40528
         End
         Begin VB.CheckBox chk入库 
            Caption         =   "按入库日期查找"
            Height          =   330
            Left            =   3195
            TabIndex        =   2
            Top             =   355
            Width           =   1680
         End
         Begin VB.Label lbl入库 
            AutoSize        =   -1  'True
            Caption         =   "入库单号"
            Height          =   180
            Left            =   90
            TabIndex        =   0
            Top             =   430
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "至"
            Height          =   180
            Left            =   6255
            TabIndex        =   11
            Top             =   430
            Width           =   180
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5490
      Left            =   90
      TabIndex        =   6
      Top             =   1065
      Width           =   11355
      _cx             =   20029
      _cy             =   9684
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffInSel.frx":06EA
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
      ExplorerBar     =   3
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   16
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmStuffInSel.frx":0717
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmStuffInSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mblnOK As Boolean
Private mrsSel As ADODB.Recordset, mlng虚拟库房ID As Long
Public Function zlSelect(ByVal frmMain As Form, ByVal lngMoudle As Long, _
    ByVal strPrivs As String, ByVal lng虚拟库房ID As Long, ByRef rsReturnSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择入口
    '入参:frmMain-调用的窗口
    '出参:rsReturnSel-返回被选择的结果集
    '返回:如果选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-16 10:28:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngModule = lngMoudle: mstrPrivs = strPrivs: mblnOK = False
    Set mrsSel = Nothing: mlng虚拟库房ID = lng虚拟库房ID
    Me.Show 1, frmMain
    Set rsReturnSel = mrsSel
    zlSelect = mblnOK
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub chk入库_Click()
        dtpStart.Enabled = chk入库.Value = 1
        dtpEnd.Enabled = chk入库.Value = 1
End Sub

Private Sub chk入库_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdALLCls_Click()
    Dim i As Long
    With vsItem
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 0
            .TextMatrix(i, .ColIndex("本次记帐数量")) = ""
        Next
    End With
End Sub

Private Sub cmdAllSel_Click()
    Dim i As Long
    With vsItem
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 1
            If Val(.TextMatrix(i, .ColIndex("本次记帐数量"))) = 0 Then
                .TextMatrix(i, .ColIndex("本次记帐数量")) = .TextMatrix(i, .ColIndex("可用数量"))
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If zlBuliedRec = False Then Exit Sub
    mblnOK = True
    Unload Me:
End Sub

Private Sub cmdSel_Click()
    Call FillData(False)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyA
            If Shift = vbCtrlMask Then cmdAllSel_Click
        Case vbKeyC
            If Shift = vbCtrlMask Then cmdALLCls_Click
        End Select
End Sub

Private Sub Form_Load()
    RestoreWinState Me, Me.Name
    dtpEnd.MaxDate = gobjDatabase.Currentdate
    dtpEnd.Value = dtpEnd.MaxDate
    dtpEnd.minDate = dtpEnd.MaxDate - 2 * 365
    dtpStart.MaxDate = dtpEnd.MaxDate
    dtpStart.minDate = dtpEnd.minDate
    dtpStart.Value = dtpEnd.Value   '默认为当天
    Call FillData(True)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsItem
        .Left = Me.ScaleLeft + 50
        .Width = Me.ScaleWidth - 100
        .Top = picSeach.Top + picSeach.Height + 20
        .Height = Me.ScaleHeight - .Top - picDown.Height - 50
    End With
End Sub
 
Private Sub picSeach_Resize()
    Err = 0: On Error Resume Next
    With picSeach
        fraSearch.Left = .ScaleLeft + 50
        fraSearch.Top = .ScaleTop + 50
        fraSearch.Height = .ScaleHeight - 100
        fraSearch.Width = .ScaleWidth - 100
    End With
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        fraBottomSplit.Left = .ScaleLeft
        fraBottomSplit.Top = .ScaleTop
        fraBottomSplit.Width = .ScaleWidth
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - cmdCancel.Width / 2
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub
Private Function FillData(Optional blnDefault As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据
    '入参:blnDefault-是否加载缺省数据,如果是,则以最后一次备货且有库存的入库单作为本次选择的对象,否则根据界面条件来过滤
    '出参:
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-16 10:40:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, rsTemp As ADODB.Recordset, i As Long
    
    On Error GoTo errHandle
    If blnDefault Then
            strWhere = "And A.No =( " & _
            "                                Select Max(No) As No From 药品收发记录 A1, 药品库存 B1  " & _
            "                                Where a1.审核日期 Between Sysdate-7 And Sysdate " & _
            "                                        And  A1.药品id = B1.药品id And A1.库房id = B1.库房id And Nvl(A1.批次, 0) = Nvl(B1.批次, 0)   " & _
            "                                        And Nvl(B1.可用数量, 0) > 0 And A1.库房id = [1] ) "
    Else
        strWhere = ""
        If txtNO.Text <> "" Then strWhere = " And A.NO=[2] "
        If chk入库.Value = 1 Then strWhere = strWhere & " and  (A.审核日期 between [3] and [4] )"
        strWhere = strWhere & " and   A.库房ID=[1] "
        If strWhere = "" Then
            MsgBox "注意:" & vbCrLf & "    查询前必须输入单据号或者入库日期!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    gstrSQL = " " & _
    " Select  a.Id,'' as 选择,A.单据,A.库房ID,A.供药单位ID,a.审核日期,A.药品ID,A.批次, " & _
    "            A.No,A.序号,b.编码||'-'||b.名称 As 卫材名称,b.规格,b.计算单位,A.产地,E.名称 As 供应商,A.批号, " & _
    "           to_char(A.生产日期,'yyyy-mm-dd') as 生产日期, to_char(A.效期,'yyyy-mm-dd') as 效期, " & _
    "           to_char(A.灭菌效期,'yyyy-mm-dd') as 灭菌效期 , " & _
    "           A.实际数量  as 入库数量,LTrim(To_Char(A.零售价,'999999" & gSysPara.Price_Decimal.strFormt_VB & "'))  as 入库零售价,A.零售金额 as 入库零售金额, " & _
    "           to_char(nvl(D.可用数量,0),'9999990.00000') as 本次记帐数量,D.商品条码,D.内部条码, " & _
    "           Decode(B.是否变价,1,'时价',LTrim(To_Char(C1.现价,'999999" & gSysPara.Price_Decimal.strFormt_VB & "'))) as 单价," & _
            IIf(InStr(1, mstrPrivs, "显示库存") > 0, " To_Char(D.可用数量,'9999990.00000')", "Decode(Sign(D.可用数量),1,'有','无')") & " as 库存," & _
    "           D.可用数量" & _
    " From 药品收发记录 A, 收费项目目录 B, 材料特性 C, 收费价目 C1, 药品库存 D,供应商 E " & _
    " Where   a.单据=15  " & strWhere & _
    "              And a.药品ID=C1.收费细目ID And (Sysdate Between C1.执行日期 and Nvl(C1.终止日期,To_Date('3000-01-01','YYYY-MM-DD')))" & _
    "              And a.供药单位ID=e.Id  " & _
    "              And a.药品ID=b.Id And  a.药品ID=C.材料ID " & _
    "              And a.库房ID=D.库房ID And nvl(a.批次,0)=nvl(D.批次,0)    " & _
    "  Order By No,序号"
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng虚拟库房ID, CStr(Trim(txtNO.Text)), _
        CDate(Format(dtpStart.Value, "yyyy-mm-dd")), CDate(Format(dtpEnd.Value, "yyyy-mm-dd")) + 1 - 1 / 24 / 60 / 60)
        
    With vsItem
        .Clear 0: .Cols = 1
        .FixedCols = 1
       Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = UCase(.TextMatrix(0, i))
            .ColData(i) = "0||1"
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "单据" Or .ColKey(i) = "批次" Or Trim(.ColKey(i)) = "可用数量" Then
                .ColHidden(i) = True
                ' ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
                .ColData(i) = "-1||1"
            ElseIf .ColKey(i) = "NO" Or .ColKey(i) = "选择" Or .ColKey(i) = "本次记帐数量" Then
                   .ColData(i) = "1||0"
                   If .ColKey(i) = "选择" Then .ColDataType(i) = flexDTBoolean
                   .ColAlignment(i) = flexAlignCenterCenter
            End If
            If .ColKey(i) Like "*数*" Or .ColKey(i) Like "*价*" Or .ColKey(i) Like "*库存*" Then
                 .ColAlignment(i) = flexAlignRightCenter
            End If
        Next
        '自动列宽
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsItem, Me.Caption, "备货选择列表", False
        If .ColIndex("标志") >= 0 Then .ColWidth(.ColIndex("标志")) = 300
        .Cell(flexcpBackColor, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = &HE7CFBA
        .Cell(flexcpBackColor, 1, .ColIndex("本次记帐数量"), .Rows - 1, .ColIndex("本次记帐数量")) = &HE7CFBA
    End With
    
    FillData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, Me.Name
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "备货选择列表", False, , InStr(1, mstrPrivs, ";附费选项设置;") > 0
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    Dim strNO As String
    If Len(txtNO) < 8 And Len(txtNO) > 0 Then
        strNO = txtNO.Text
        Call MakeNO(68, mlng虚拟库房ID, strNO)
        txtNO.Text = strNO
    End If
End Sub

Private Sub vsItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With vsItem
            Select Case Col
            Case .ColIndex("本次记帐数量")
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "0.00000")
                If Val(.TextMatrix(Row, Col)) <> 0 And GetVsGridBoolColVal(vsItem, Row, .ColIndex("选择")) = False Then
                    vsItem.TextMatrix(Row, .ColIndex("选择")) = 1
                ElseIf Val(.TextMatrix(Row, Col)) = 0 Then
                    vsItem.TextMatrix(Row, .ColIndex("选择")) = 0
                End If
            Case .ColIndex("选择")
                If GetVsGridBoolColVal(vsItem, Row, Col) Then
                    If Val(.TextMatrix(Row, .ColIndex("本次记帐数量"))) = 0 Then
                            .TextMatrix(Row, .ColIndex("本次记帐数量")) = Format(Val(.TextMatrix(Row, .ColIndex("可用数量"))), "0.00000")
                    End If
                End If
            Case Else
            End Select
        End With
End Sub


Private Sub vsItem_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "备货选择列表", False, , InStr(1, mstrPrivs, ";附费选项设置;") > 0
End Sub

Private Sub vsItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "备货选择列表", False, , InStr(1, mstrPrivs, ";附费选项设置;") > 0
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsItem, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsItem, Me.Caption, "备货选择列表", False, , InStr(1, mstrPrivs, ";附费选项设置;") > 0
End Sub
Private Sub picImg_Click()
    Call imgCol_Click
End Sub
Private Function zlBuliedRec() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:构建选中的数据
    '编制:刘兴洪
    '日期:2010-12-16 15:15:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, blnData As Boolean
    '先检查可记帐数量
    With vsItem
        blnData = False
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsItem, i, .ColIndex("选择")) Then
                If Val(.TextMatrix(i, .ColIndex("本次记帐数量"))) <= 0 Then
                    MsgBox "注意:" & "    在第" & i & "行中的本次记帐数量必须大于零,请检查!", vbOKOnly + vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("本次记帐数量")
                    If .RowIsVisible(.Row) = False Or .ColIsVisible(.Col) = False Then
                        Call .ShowCell(.Row, .Col)
                    End If
                    Exit Function
                End If
                If Val(.TextMatrix(i, .ColIndex("本次记帐数量"))) > Val(.TextMatrix(i, .ColIndex("可用数量"))) Then
                    MsgBox "注意:" & "    在第" & i & "行中的本次记帐数量大于了库存数量,请检查!", vbOKOnly + vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("本次记帐数量")
                    If .RowIsVisible(.Row) = False Or .ColIsVisible(.Col) = False Then
                        Call .ShowCell(.Row, .Col)
                    End If
                    Exit Function
                End If
                blnData = True
            End If
        Next
    End With
    If blnData = False Then
        MsgBox "注意:" & "    未选择指定的记帐数据,请检查!", vbOKOnly + vbInformation, gstrSysName
        vsItem.SetFocus
        Exit Function
    End If
    Set mrsSel = New ADODB.Recordset
    With mrsSel
        If .State = 1 Then .Close
        .Fields.Append "虚拟库房ID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "收费项目ID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "材料规格", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "商品条码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "内部条码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "可用库存", adDouble, , adFldIsNullable
        .Fields.Append "数量", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    With vsItem
        For i = 1 To .Rows - 1
            If GetVsGridBoolColVal(vsItem, i, .ColIndex("选择")) Then
                mrsSel.AddNew
                mrsSel!虚拟库房id = mlng虚拟库房ID
                mrsSel!收费项目ID = Val(.TextMatrix(i, .ColIndex("药品ID")))
                mrsSel!批次 = Val(.TextMatrix(i, .ColIndex("批次")))
                mrsSel!材料名称 = Trim(.TextMatrix(i, .ColIndex("卫材名称")))
                mrsSel!材料规格 = Trim(.TextMatrix(i, .ColIndex("规格")))
                mrsSel!内部条码 = Trim(.TextMatrix(i, .ColIndex("内部条码")))
                mrsSel!商品条码 = Trim(.TextMatrix(i, .ColIndex("商品条码")))
                mrsSel!数量 = Trim(.TextMatrix(i, .ColIndex("本次记帐数量")))
                mrsSel!可用库存 = Val(.TextMatrix(i, .ColIndex("可用数量")))
                mrsSel.Update
            End If
        Next
    End With
    zlBuliedRec = True
End Function
Private Sub vsItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem
        Select Case Col
        Case .ColIndex("本次记帐数量"), .ColIndex("选择")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsItem
        Select Case Col
        Case .ColIndex("标志")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsItem_EnterCell()
    '暂未设置
    With vsItem
    End With
End Sub

Private Sub vsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsItem
        If Val(.TextMatrix(.Row, .ColIndex("药品ID"))) = 0 And .Col >= .ColIndex("本次记帐数量") And .Row = .Rows - 1 Then
            gobjCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
End Sub

Private Sub vsItem_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsItem
        Select Case Col
        Case .ColIndex("本次记帐数量")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsItem
        Select Case .Col
            Case .ColIndex("本次记帐数量")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
            
        End Select
    End With
End Sub
Private Sub vsItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    '数据验证
    With vsItem
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
            Case .ColIndex("本次记帐数量")
                If zlDblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                strKey = Format(Val(strKey), "0.00000")
                If Val(strKey) > Val(.TextMatrix(Row, .ColIndex("可用数量"))) Then
                    MsgBox "注意:" & vbCrLf & "    输入的数量大于了库存数量,请检查!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                    Cancel = True: Exit Sub
                End If
                .EditText = strKey
                .TextMatrix(Row, .Col) = strKey
        End Select
    End With
End Sub
 
Private Sub MakeNO(ByVal intBillID As Integer, ByVal lng科室id As Long, ByRef strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据编号规则自动产生号码
    '入参:
    '出参:strNo-返回单据号
    '返回:
    '编制:刘兴洪
    '日期:2010-12-17 14:28:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intYear As Integer, strYear As String
    Dim intMonth As Integer, strMonth As String
    Dim str编号 As String
    Dim rsTemp As New ADODB.Recordset
    
    strNO = UCase(LTrim(strNO))
    intYear = Format(gobjDatabase.Currentdate, "YYYY") - 1990
    strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(gobjDatabase.Currentdate())
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
    If rsTemp.State = 1 Then rsTemp.Close
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord("Select 编号规则 From 号码控制表 Where 项目序号=[1]", "获取单据规则", intBillID)
    
    
    Dim bln年度 As Boolean
    Dim rsTmp As New ADODB.Recordset
    If Nvl(rsTemp!编号规则, 0) = 2 And lng科室id <> 0 Then
        gstrSQL = "Select 工作性质, 部门id, 服务对象 from 部门性质说明 where 工作性质 in ( '卫材库','制剂室','虚拟库房') and 部门ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(gstrSQL, "获取部门性质", lng科室id)
        If rsTmp.EOF Then
            bln年度 = True
        Else
            bln年度 = False
        End If
    Else
        bln年度 = False
    End If
    If Nvl(rsTemp!编号规则, 0) = 0 Or bln年度 Then
        If Len(strNO) < 8 Then strNO = strYear & String(7 - Len(strNO), "0") & strNO
    ElseIf rsTemp!编号规则 = 2 Then
        If rsTemp.State = 1 Then rsTemp.Close
        Set rsTemp = gobjDatabase.OpenSQLRecord("Select 编号 From  科室号码表 Where 项目序号=[1] and nvl(科室ID,0)=[2]", "获取科室编号", intBillID, lng科室id)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Sub
        End If
        If Nvl(rsTemp!编号) = "" Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Sub
        End If
        str编号 = Nvl(rsTemp!编号)
        
        '小于四位，按本月产生号码
        '五位或六位，则认为是指定月份的号码
        '七位，则认为是产生本年指定科室、月份的号码
        '大于等于八位，不处理
        If Len(strNO) <= 4 Then
            strNO = strYear & str编号 & strMonth & String(4 - Len(strNO), "0") & strNO
        ElseIf Len(strNO) <= 6 Then
            strNO = String(6 - Len(strNO), "0") & strNO
            strNO = strYear & str编号 & strNO
        ElseIf Len(strNO) = 7 Then
            strNO = strYear & strNO
        End If
    Else
        MsgBox "不支持这种编号规则！", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub
