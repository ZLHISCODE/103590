VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiCureCardChangeMgr 
   BorderStyle     =   0  'None
   Caption         =   "变动信息"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPatiCureCardChangeMgr.frx":0000
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
      ExplorerBar     =   2
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
            Picture         =   "frmPatiCureCardChangeMgr.frx":0066
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmPatiCureCardChangeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mstrCardNo As String, mlngCardTypeID As Long
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '弹出菜单操作
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '弹出菜单操作

Public Function zlReLoadData(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '返回:加载成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-06-28 15:30:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrCardNo = strCardNo: mlngCardTypeID = lngCardTypeID
    Err = 0: On Error GoTo Errhand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long, strSQL As String
    mblnHaveData = False
    Err = 0: On Error GoTo Errhand:
    '问题号:55447
    '69745:刘尔旋,2014-02-13,读取数据库时间时没有进行格式转换导致时间错误的问题
    strSQL = "" & _
    "   Select A.卡号, A.变动类别 As 变动类别id, " & _
    "          Decode(A.变动类别, 1, '发卡',11,'绑定卡', 2, '换卡', 3, '补卡',13,'补卡停用', 4, '退卡',14,'取消绑定', 5, '密码调整', 6, '挂失', 16, '取消挂失', '无') As 变动类别, " & _
    "          to_char(A.变动时间,'yyyy-mm-dd hh24:mi:ss') as 变动时间, A.变动原因, A.挂失方式, A.操作员姓名, to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间, A.变动id, B.姓名, B.性别, B.年龄 " & _
    "   From 病人医疗卡变动 A, 病人信息 B " & _
    "   Where A.病人id = B.病人id And A.卡类别id = [1] And 卡号 = [2] " & _
    "   Order By 登记时间  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, mstrCardNo)
    With Me.vsGrid
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "变动类别" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*时间" Then
                .ColAlignment(i) = flexAlignCenterCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) = "卡号" Then
                .ColData(i) = "1|1"
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "医疗卡变动列表", True
        .ColWidth(.ColIndex("标志")) = 285
        .ColData(.ColIndex("标志")) = "-1|1"
        
        .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "医疗卡变动列表", True
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "医疗卡变动列表", True
End Sub

Private Sub picImg_Click()
    Call imgCol_Click
End Sub
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2011-06-28 15:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, vsGrid As VSFlexGrid
    Err = 0: On Error GoTo errH:
    gstrSQL = "Select   A.卡号, to_char(A.发卡时间,'yyyy-mm-dd hh24:mi:ss') From 病人医疗卡信息 A where 卡类别ID=[1] and 卡号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCardTypeID, mstrCardNo)
    If rsTemp.EOF = True Then Exit Sub '无卡信息，退出
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "医疗卡变动情况"
    
    objRow.Add "卡号：" & Nvl(rsTemp!卡号)
    objRow.Add "发卡时间：" & Nvl(rsTemp!发卡时间)
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

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "医疗卡变动列表", True
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Then Exit Sub
    
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "医疗卡变动列表", True
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
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(vsGrid)
End Sub
 






