VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardInFul 
   BorderStyle     =   0  'None
   Caption         =   "充值记录"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4485
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      _cx             =   14314
      _cy             =   7911
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSquareCardInFul.frx":0000
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
      ExplorerBar     =   7
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmSquareCardInFul.frx":018F
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmSquareCardInFul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mlng消费卡ID As Long, mlng接口编号 As Long
Private mstr卡类型 As String, mstr卡号 As String
'一些公共事件
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '弹出菜单操作
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '弹出菜单操作

'设置相关属性
Public Property Get zlIsHaveData() As Boolean
    zlIsHaveData = mblnHaveData
End Property

Public Property Get zl允许回退() As Boolean
    With vsGrid
        If .Row < 1 Then zl允许回退 = False: Exit Property
        zl允许回退 = Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1
    End With
End Property

Public Property Get zlGet充值ID() As Long
    With vsGrid
        If .Row < 1 Then zlGet充值ID = 0: Exit Property
        zlGet充值ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property

Public Function zlReLoadData(ByVal lng接口编号 As Long, ByVal lng消费卡ID As Long, _
    ByVal str卡类型 As String, ByVal str卡号 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '入参:mcllFilter-过滤条件(目前无)
    '出参:
    '返回:加载成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2009-11-20 16:00:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng消费卡ID = lng消费卡ID: mlng接口编号 = lng接口编号
    mstr卡类型 = str卡类型: mstr卡号 = str卡号
    Err = 0: On Error GoTo Errhand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-11-20 16:05:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("记录状态")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColHidden(.ColIndex("ID")) = True
        .ColHidden(.ColIndex("记录状态")) = True
        
        .ColData(.ColIndex("序号")) = "1|1"
        .ColData(.ColIndex("充值金额")) = "1|0"
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long
    Dim blnHistory As Boolean, strStartDate As String
    Dim strSQL As String
    
    mblnHaveData = False
    Err = 0: On Error GoTo Errhand:
    strSQL = _
        "Select a.Id, a.消费卡id, a.序号, a.记录状态, a.应收金额 As 充值金额, a.扣率 As 充值折扣, a.实收金额 As 缴款金额," & vbNewLine & _
        "       To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 充值时间, a. 操作员姓名, a.缴款人姓名 As 缴款人, 备注" & vbNewLine & _
        "From 病人卡结算记录 A" & vbNewLine & _
        "Where a.记录性质 = 2 And a.消费卡id = [1]" & vbNewLine & _
        "Order By 序号 Desc, 充值时间 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng消费卡ID)
    
    With Me.vsGrid
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("序号")) = Nvl(rsTemp!序号)
            .Cell(flexcpData, lngRow, .ColIndex("序号")) = Nvl(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("记录状态")) = Nvl(rsTemp!记录状态)
            .TextMatrix(lngRow, .ColIndex("充值金额")) = Format(Val(Nvl(rsTemp!充值金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("充值折扣")) = Format(Val(Nvl(rsTemp!充值折扣)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("缴款金额")) = Format(Val(Nvl(rsTemp!缴款金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("充值时间")) = Nvl(rsTemp!充值时间)
            .TextMatrix(lngRow, .ColIndex("充值说明")) = Nvl(rsTemp!备注)
            .TextMatrix(lngRow, .ColIndex("操作员姓名")) = Nvl(rsTemp!操作员姓名)
            .TextMatrix(lngRow, .ColIndex("缴款人")) = Nvl(rsTemp!缴款人)

            Select Case Val(Nvl(rsTemp!记录状态))
            Case 2
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
            Case 3
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000D
            Case Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000012
            End Select
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        mblnHaveData = rsTemp.RecordCount > 0
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "充值列表", True
        .ColWidth(.ColIndex("标志")) = 285
        If .Row <= 0 Then .Row = 1
    End With
    Call vsGrid_AfterRowColChange(-1, 0, 1, 1)
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs

    Call InitVsGrid
    Call vsGrid_GotFocus
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsGrid
        .Left = ScaleLeft: .Top = ScaleTop
        .Width = ScaleWidth: .Height = ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "充值列表", True
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "充值列表", True
End Sub

Private Sub picImg_Click()
    Call imgCol_Click
End Sub
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2009-11-20 16:36:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid
    
    Err = 0: On Error GoTo errH:
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "消费卡充值情况"
    
    objRow.Add "卡类型：" & mstr卡类型
    objRow.Add "卡号：" & mstr卡号
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("标志") Then .ColWidth(intCol) = 0
        Next
    End With
    Set objPrint.Body = vsGrid
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
    
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
End Sub

Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub

Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
    zl_VsGridLostFocus vsGrid, gSysColor.lngGridColorLost
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(vsGrid)
End Sub
