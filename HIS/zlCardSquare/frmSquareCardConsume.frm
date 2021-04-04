VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardConsume 
   BorderStyle     =   0  'None
   Caption         =   "使用消费管理"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   3930
      _cx             =   6932
      _cy             =   3969
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSquareCardConsume.frx":0000
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
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   45
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmSquareCardConsume.frx":0066
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmSquareCardConsume"
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
Public Event zlDblClick(ByVal lng结算ID As Long, ByVal vsGrid As VSFlexGrid)   '弹出菜单操作'设置相关属性

Public Property Get zlIsHaveData() As Boolean
    zlIsHaveData = mblnHaveData
End Property

Public Property Get zlGet消费ID() As Long
    Dim lng卡结算ID As Long
    With vsGrid
        If .Row < 1 Then zlGet消费ID = 0: Exit Property
        lng卡结算ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        zlGet消费ID = lng卡结算ID
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
    Dim i As Long
    Dim strHead As String, varHead As Variant
    
    strHead = "标志,1,285|ID,1,0|序号,4,800|消费卡id,1,0|结算方式,1,1200|结算金额,7,1000|卡号,4,1200|" & _
              "交易流水号,1,1500|交易时间,4,1850|备注,1,1850|结算标志,4,1200|实际票号,4,1200"
    varHead = Split(strHead, "|")
    With vsGrid
        .Cols = UBound(varHead) + 1
        For i = 0 To UBound(varHead)
            .TextMatrix(0, i) = Split(varHead(i), ",")(0)
            .ColKey(i) = Split(varHead(i), ",")(0)
            If .TextMatrix(0, i) = "标志" Then .TextMatrix(0, i) = ""
            .ColAlignment(i) = Split(varHead(i), ",")(1)
            .ColWidth(i) = Split(varHead(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("卡号")) = "1|1"
        .ColData(.ColIndex("消费卡id")) = "-1|1"
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    mblnHaveData = False
    Err = 0: On Error GoTo Errhand:
    strSQL = _
        "Select A.ID, A.序号, A.消费卡id, A.结算方式, Trim(to_char(-1*A.应收金额," & gOraFmtString.FM_金额 & ")) as 结算金额," & vbNewLine & _
        "       A.卡号, A.交易流水号, A.交易时间, A.备注, " & vbNewLine & _
        "       Decode(A.结算标志,1,'√','') as 结算标志,max( C.实际票号) 实际票号 " & vbNewLine & _
        "From 病人卡结算记录 A, 病人预交记录 C " & vbNewLine & _
        "Where A.结算id = C.ID And A.接口编号 = [2] And A.消费卡id = [1] And a.记录性质 = 4" & vbNewLine & _
        "Group By  A.ID, A.序号, A.消费卡id, A.结算方式, A.应收金额, A.卡号, A.交易流水号, A.交易时间, A.备注,A.结算标志" & vbNewLine & _
        "Order by A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng消费卡ID, mlng接口编号)
    
    With Me.vsGrid
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("序号")) = Nvl(rsTemp!序号)
            .TextMatrix(lngRow, .ColIndex("消费卡id")) = Nvl(rsTemp!消费卡ID)
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .TextMatrix(lngRow, .ColIndex("结算金额")) = Format(Val(Nvl(rsTemp!结算金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("卡号")) = Nvl(rsTemp!卡号)
            .TextMatrix(lngRow, .ColIndex("交易流水号")) = Nvl(rsTemp!交易流水号)
            .TextMatrix(lngRow, .ColIndex("交易时间")) = Format(Nvl(rsTemp!交易时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(lngRow, .ColIndex("备注")) = Nvl(rsTemp!备注)
            .TextMatrix(lngRow, .ColIndex("结算标志")) = Nvl(rsTemp!结算标志)
            .TextMatrix(lngRow, .ColIndex("实际票号")) = Nvl(rsTemp!实际票号)

            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "消费列表", True
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    mblnHaveData = rsTemp.RecordCount > 0
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsGrid.Redraw = flexRDBuffered
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "消费列表", True
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "消费列表", True
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
    
    On Error GoTo errH:
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "消费卡消费情况"
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "消费列表", True
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "消费列表", True
End Sub

 Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub

Private Sub vsGrid_DblClick()
    Dim lng卡结算ID As Long
    With vsGrid
        If .MouseRow < 1 Then Exit Sub
        lng卡结算ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lng卡结算ID = 0 Then Exit Sub
        RaiseEvent zlDblClick(lng卡结算ID, vsGrid)
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

 
