VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcManage 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "变动过程日常管理"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmProcManage.frx":0000
   ScaleHeight     =   4995
   ScaleWidth      =   8265
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000010&
      Height          =   270
      Left            =   5040
      TabIndex        =   5
      Text            =   "输过程名称或修改人后按回车进行定位"
      ToolTipText     =   "请直接按回车键进行过滤"
      Top             =   1815
      Width           =   3135
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      ItemData        =   "frmProcManage.frx":803A
      Left            =   2280
      List            =   "frmProcManage.frx":804A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   2010
   End
   Begin VB.CommandButton cmdCollect 
      Caption         =   "检查变动过程(&C)"
      Height          =   350
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Left            =   -120
      TabIndex        =   2
      Top             =   1680
      Width           =   8280
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修改变动过程(&E)"
      Height          =   350
      Left            =   6480
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "查看变动痕迹(&D)"
      Height          =   350
      Left            =   4680
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfProc 
      Height          =   1815
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   7935
      _cx             =   13996
      _cy             =   3201
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
      ForeColorFixed  =   -2147483636
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   -1  'True
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
   Begin VB.Label lblLocation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "定位"
      Height          =   180
      Left            =   4560
      TabIndex        =   10
      Top             =   1860
      Width           =   360
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统"
      Height          =   180
      Left            =   1800
      TabIndex        =   9
      Top             =   1860
      Width           =   360
   End
   Begin VB.Label lblcaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "变动过程日常管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFEBD7&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmProcManage.frx":8080
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   840
      TabIndex        =   7
      Top             =   720
      Width           =   9360
   End
   Begin VB.Label lblProc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户变动过程清单"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1860
      Width           =   1440
   End
   Begin VB.Image Img 
      Height          =   615
      Left            =   120
      Picture         =   "frmProcManage.frx":8139
      Stretch         =   -1  'True
      Top             =   600
      Width           =   580
   End
End
Attribute VB_Name = "frmProcManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum txtColor
    黑色 = &H80000012
    灰色 = &H80000010
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用

End Function


Private Sub cboSystem_Click()
    Dim i As Long, lngSrow As Long
    
    With vsfProc
        For i = 1 To .Rows - 1
            If cboSystem.List(cboSystem.ListIndex) = "全部" Or cboSystem.List(cboSystem.ListIndex) = "" Then
                lngSrow = 1
                .RowHidden(i) = False
            Else
                If .TextMatrix(i, .ColIndex("系统")) <> Trim(cboSystem.List(cboSystem.ListIndex)) Then
                    .RowHidden(i) = True
                Else
                    If lngSrow = 0 Then lngSrow = i
                    .RowHidden(i) = False
                End If
            End If
        Next
        
        If .Rows > 1 Then
            .Select lngSrow, 0
        Else
            .Select 0, 0
        End If
    End With
End Sub

Private Sub cmdCollect_Click()
    If frmProcCollectCur.ShowMe Then
        Call LoadProc
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim strProc As String
    Dim lngRow As Long, lngTopRow As Long
    
    With vsfProc
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        strProc = LoadBaseProcs(.TextMatrix(.Row, .ColIndex("过程名称")))
        If frmProcEditCommon.ShowMe(.RowData(.Row), .TextMatrix(.Row, .ColIndex("过程名称")), strProc, _
                                         .TextMatrix(.Row, .ColIndex("系统")), .TextMatrix(.Row, .ColIndex("修改说明")), .TextMatrix(.Row, .ColIndex("修改人"))) Then
            '保存后需要重新刷新
            lngRow = .Row: lngTopRow = .TopRow
            Call LoadProc
            .Row = lngRow: .TopRow = lngTopRow
         End If
    End With
End Sub

Private Sub CmdView_Click()
    Dim arrIds() As String, lngIdx As Long
    Dim i As Long
    
    With vsfProc
        If .Row = 0 Then
            MsgBox "没有选中过程", , "提示"
            Exit Sub
        End If
        
        '因为要连续操作,所以把需要调整的过程ID都传到子窗体
        lngIdx = .Row - 1
        ReDim arrIds(.Rows - 2)
        
        For i = 1 To .Rows - 1
            arrIds(i - 1) = .RowData(i) & ":" & .TextMatrix(i, .ColIndex("过程名称"))
        Next
        
    End With
    frmDiffCommon.ShowMe arrIds, lngIdx
End Sub

Private Sub Form_Activate()
    Call LoadProc
End Sub

Private Sub Form_Load()
    Dim strCol As String
    
    strCol = ",350,1;系统,2000,1;过程名称,2000,1;最新标准版本,2000,1;修改人,2000,1;修改时间,2000,1;上次修改人员,2000,1;上次修改时间,2000,1;修改说明,2000,1"
    Call InitTable(vsfProc, strCol)
    vsfProc.FixedCols = 1
    vsfProc.Rows = 1
    vsfProc.Cell(flexcpForeColor, 0, 0, 0, vsfProc.Cols - 1) = &H80000008

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    fraSplit.Width = Me.ScaleWidth + 50
    cmdEdit.Top = Me.ScaleHeight - cmdEdit.Height - 120
    cmdEdit.Left = Me.ScaleWidth - cmdEdit.Width - 120
    
    cmdView.Top = cmdEdit.Top
    cmdView.Left = cmdEdit.Left - cmdView.Width - 80
    
    vsfProc.Width = Me.ScaleWidth - vsfProc.Left - 120
    vsfProc.Height = cmdEdit.Top - vsfProc.Top - 120
    
    txtLocation.Left = vsfProc.Width + vsfProc.Left - txtLocation.Width
    lblLocation.Left = txtLocation.Left - lblLocation.Width - 80
    
    cboSystem.Left = lblLocation.Left - cboSystem.Width - 200
    lblSystem.Left = cboSystem.Left - lblSystem.Width - 80
End Sub


Private Sub LoadProc()
    '加载数据库中保存的变动过程
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Long, strTmp As String
    
    On Error GoTo errH
    ShowFlash "正在加载变动过程..."
    strSQL = "Select a.Id, a.系统编号, a.名称, a.类型, a.状态, a.所有者, a.修改人员, To_Char(a.修改时间, 'yyyy-mm-dd hh24:mi') 修改时间, a.上次修改人员," & vbNewLine & _
                    "       To_Char(a.上次修改时间, 'yyyy-mm-dd hh24:mi') 上次修改时间, a.升级前版本, a.说明, c.名称 系统" & vbNewLine & _
                    "From (Select Distinct a.Id, a.系统编号, a.名称, a.类型, a.状态, a.所有者, a.修改人员, a.修改时间, a.上次修改人员, a.上次修改时间, a.升级前版本, a.说明" & vbNewLine & _
                    "       From Zlprocedure A, Zlproceduretext B" & vbNewLine & _
                    "       Where 类型 = 1 And a.Id = b.过程id And (b.性质 = 1 Or b.性质 = 3)) A, zlSystems C" & vbNewLine & _
                    "Where a.系统编号 = c.编号" & vbNewLine & _
                    "Order By a.系统编号, a.名称"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "加载变动过程")
    
    '加载变动过程
    With vsfProc
        If rsTmp.RecordCount = 0 Then Exit Sub

        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTmp.RecordCount + 1
        .MergeCells = flexMergeRestrictRows
        .MergeCol(.ColIndex("系统")) = True
        
        i = 1
        cboSystem.Clear
        cboSystem.AddItem "全部"
        Do While Not rsTmp.EOF
            '下拉列表数据
            If rsTmp!系统 <> strTmp Then
                strTmp = rsTmp!系统
                cboSystem.AddItem rsTmp!系统
                cboSystem.ItemData(cboSystem.NewIndex) = rsTmp!系统编号
            End If
            
            '表格数据
            .TextMatrix(i, 0) = i
            .TextMatrix(i, .ColIndex("系统")) = rsTmp!系统 & ""
            .TextMatrix(i, .ColIndex("过程名称")) = rsTmp!名称 & ""
            .TextMatrix(i, .ColIndex("修改人")) = rsTmp!修改人员 & ""
            .TextMatrix(i, .ColIndex("修改时间")) = rsTmp!修改时间 & ""
            .TextMatrix(i, .ColIndex("修改说明")) = rsTmp!说明 & ""
            .TextMatrix(i, .ColIndex("最新标准版本")) = rsTmp!升级前版本 & ""
            .TextMatrix(i, .ColIndex("上次修改人员")) = rsTmp!上次修改人员 & ""
            .TextMatrix(i, .ColIndex("上次修改时间")) = rsTmp!上次修改时间 & ""
            .RowData(i) = rsTmp!id & ""
            i = i + 1
            rsTmp.MoveNext
        Loop
        cboSystem.ListIndex = 0
        .AutoResize = True
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
        
    If .Rows > 1 Then
        .Select 1, 0
    End If
    End With
    
    ShowFlash ""
    
    Exit Sub
errH:
    ShowFlash ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, , gstrSysName
End Sub


Private Sub vsfProc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    With vsfProc
        If .Redraw = flexRDNone Then Exit Sub
        If .Rows = 1 Then Exit Sub
        
        .Cell(flexcpForeColor, OldRow, 0) = Color.深灰色
        .Cell(flexcpFontBold, OldRow, 0) = False
        .Cell(flexcpFontBold, NewRow, 0) = True
        .Cell(flexcpForeColor, NewRow, 0) = Color.兰色
    End With

End Sub


Private Sub txtLocation_GotFocus()
    If txtLocation.Text = "输过程名称或修改人后按回车进行定位" Then
        txtLocation.Text = ""
        txtLocation.ForeColor = 黑色
    End If
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    If txtLocation.Text = "" Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    
    GetRowPos vsfProc, txtLocation.Text, "过程名称,修改人,上次修改人员"
End Sub

Private Sub txtLocation_LostFocus()
    If txtLocation.Text = "" Then
        txtLocation.Text = "输过程名称或修改人后按回车进行定位"
        txtLocation.ForeColor = 灰色
    End If
End Sub
