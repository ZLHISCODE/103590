VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIdxInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "外键索引缺失查找和补建"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12600
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8130
      Left            =   5505
      ScaleHeight     =   8130
      ScaleWidth      =   7095
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdExecute 
         Caption         =   "执行(&E)"
         Height          =   350
         Left            =   120
         TabIndex        =   14
         Top             =   7680
         Width           =   1095
      End
      Begin VB.OptionButton optFKey 
         Caption         =   "调整外键"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtToolInfo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmIdxInfo.frx":0000
         Top             =   120
         Width           =   6615
      End
      Begin VB.Frame fraIdx 
         BorderStyle     =   0  'None
         Caption         =   "索引选项"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   360
         TabIndex        =   7
         Top             =   4080
         Width           =   6135
         Begin VB.TextBox txtParaNum 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2160
            TabIndex        =   10
            Text            =   "0"
            ToolTipText     =   "执行索引创建后会自动取消索引的并行属性"
            Top             =   637
            Width           =   975
         End
         Begin VB.CheckBox chkParallel 
            Caption         =   "并行执行"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "执行索引创建后会自动取消索引的并行属性"
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkOnln 
            Caption         =   "在线模式"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lblParaNum 
            AutoSize        =   -1  'True
            Caption         =   "并行度"
            Height          =   180
            Left            =   1560
            TabIndex        =   11
            Top             =   697
            Width           =   540
         End
      End
      Begin VB.Frame fraFKey 
         Enabled         =   0   'False
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   5520
         Width           =   6135
         Begin VB.OptionButton optDisable 
            Caption         =   "禁用约束"
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optDel 
            Caption         =   "删除约束"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.OptionButton optIdx 
         Caption         =   "补建索引"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtSql 
         Height          =   1065
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   6510
         Width           =   6615
      End
      Begin VB.Label lblSql 
         AutoSize        =   -1  'True
         Caption         =   "SQL命令"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   6240
         Width           =   630
      End
      Begin VB.Label lblTip 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   7730
         Width           =   5565
      End
   End
   Begin VB.PictureBox pctLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkZlhis 
         Caption         =   "只检查业务表"
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         ToolTipText     =   "涉及外键的子表和父表均为业务表"
         Top             =   7728
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtRange 
         Height          =   270
         Left            =   1080
         TabIndex        =   21
         Text            =   "100000"
         Top             =   7720
         Width           =   735
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   4320
         TabIndex        =   19
         Top             =   7680
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5295
         _cx             =   9340
         _cy             =   12726
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
         GridColor       =   32768
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   75
         ColWidthMax     =   0
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
         SubtotalPosition=   0
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
      Begin VB.Label lblRange2 
         AutoSize        =   -1  'True
         Caption         =   "行的表"
         Height          =   180
         Left            =   1800
         TabIndex        =   22
         Top             =   7770
         Width           =   540
      End
      Begin VB.Label lblRange1 
         AutoSize        =   -1  'True
         Caption         =   "只检查大于"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   7770
         Width           =   900
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         Caption         =   "外键缺失索引情况"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmIdxInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Public Sub ShowMe()
    Me.Show
End Sub

    
Private Sub chkOnln_Click()
    Call GetSql
End Sub

Private Function Checkindex(ByVal strIndexName As String) As Boolean
'功能：检查指定的索引是否存在
    Dim rstmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From DBA_Indexes Where Index_Name = [1] and Owner = [2]"
    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strIndexName, ".")(1), Split(strIndexName, ".")(0))
    
    Checkindex = rstmp.RecordCount > 0
    Exit Function
errH:
    Call ErrCenter(strSql)
End Function


Private Sub cmdExecute_Click()
    Dim intOldRow As Integer
    Dim strIdx As String, strFkey As String, strQuery As String
    
    On Error GoTo errH
    '弹出信息
    If optIdx Then
        strQuery = "你确定要创建索引“" & optIdx.Tag & "”吗？" & vbCrLf & vbCrLf & _
                    IIf(InStr(vsfGrid.TextMatrix(vsfGrid.Row, vsfGrid.ColIndex("外键字段")), ",") > 0, "如果有其他多余的相似索引（例如：字段顺序不同），请检查后删除，避免影响该表的写入性能。" & vbCrLf, "") & _
                     IIf(chkOnln.Value = 0, "由于没有选择在线创建，创建期间不能对该表进行任何写入操作", "大表上创建索引比较耗时") & ",建议在业务空闲期间进行。"
    Else
        strQuery = "你确定要" & IIf(optDel, "删除", "禁用") & "约束“" & optFKey.Tag & "”吗？" & _
                        IIf(vsfGrid.Cell(flexcpText, vsfGrid.Row, vsfGrid.ColIndex("级联操作")) = "", "", vbCrLf & vbCrLf & "该约束有级联操作，请确保已阅读并理解相关说明。")
    End If
    If MsgBox(strQuery, vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then: Exit Sub
    
    gstrSQL = TrimEx(UCase(txtSql.Text))
    If Not (gstrSQL Like "CREATE INDEX*" Or gstrSQL Like "ALTER TABLE * CONSTRAINT*") Then
        MsgBox "只允许执行索引补建或外键删除、停用的命令！"
        Exit Sub
    End If
    
    Call SetCmdEnable(False)
    intOldRow = vsfGrid.Row
    strIdx = optIdx.Tag
    strFkey = optFKey.Tag
    
    If gstrSQL Like "CREATE INDEX*" Then
        If Checkindex(strIdx) Then
            gcnOracle.Execute "Drop Index " & strIdx
        End If
    End If
    gcnOracle.Execute gstrSQL
    
    With vsfGrid
        .RemoveItem intOldRow
        If intOldRow >= .Rows - .FixedRows Then '保证选中行的位置不变
            .Select .Rows - .FixedRows, 1
        Else
            .Select intOldRow, 1
        End If
        .TopRow = .Row
        If .Row = intOldRow Then
            Call GetSql
        End If
    End With
    
    If optFKey.Value Then
        If optDel.Value Then
            lblTip.Caption = "删除外键 " & strFkey & " 成功 。"
        Else
            lblTip.Caption = "停用外键 " & strFkey & " 成功 。"
        End If
    Else
        lblTip.Caption = "补建索引 " & strIdx & " 成功 。"
    End If
    
    Call SetCmdEnable(True)
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    lblTip.Caption = Err.Description

    Call SetCmdEnable(True)
End Sub

Private Sub cmdRefresh_Click()
    Dim intOldRow As Integer
    
    intOldRow = vsfGrid.Row
    Call SetCmdEnable(False)
    lblTip.Caption = "正在刷新数据，请稍等。"
    lblTip.Refresh
    
    With vsfGrid
        .Rows = .FixedRows
        Call LoadGrid
        If intOldRow > .Rows - .FixedRows Then
            .Select .Rows - .FixedRows, 1
        Else
            .Select intOldRow, 1
        End If
        .TopRow = .Row
    End With
    Call SetCmdEnable(True)
    lblTip.Caption = ""
    
End Sub

Private Sub Form_load()
    Dim strCol As String
    
    '初始化表格，加载数据，设置控件可用性
    strCol = "所有者,1485,1;子表,1485,1;外键名称,2145,1;外键字段,1555,1;父表,1705,1;级联操作,1300,1"
    Call InitTable(vsfGrid, strCol)
    With vsfGrid
        .Editable = flexEDNone
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .AutoSizeMode = flexAutoSizeColWidth
    End With
    
    Call LoadGrid
    
    With vsfGrid
        If .Rows > 2 Then
            vsfGrid.Select 1, 1
        End If
    End With
    
    optIdx.Value = True
    chkOnln.Value = 1
    optDel.Value = True

    txtParaNum.ToolTipText = "当前CPU共" & gintCpuCount & "个，建议并行度 " & gintCpuAdvise & "最大并行度 " & gintCpuMax
    txtParaNum.Text = gintCpuAdvise
    
    lblRange1.Visible = Not gblnIsZlhis
    txtRange.Visible = Not gblnIsZlhis
    lblRange2.Visible = Not gblnIsZlhis
    chkZlhis.Visible = gblnIsZlhis
End Sub

Private Sub Form_Resize()

    '调整左侧位置大小
    pctLeft.Height = Me.ScaleHeight
    pctLeft.Width = Me.ScaleWidth - pctRight.Width - 25
    
End Sub

Private Sub optFKey_Click()
    Call SetEnableFra
    Call GetSql
End Sub

Private Sub optIdx_Click()
    '补建索引
    Call SetEnableFra
    Call GetSql
End Sub

Private Sub optDel_Click()
    Call GetSql
End Sub

Private Sub optDisable_Click()
    Call GetSql
End Sub

Private Sub chkParallel_Click()
    txtParaNum.Enabled = fraIdx.Enabled And chkParallel.Value
    Call GetSql
End Sub

Private Sub SetEnableFra()
'功能：修改选项卡和复选框的可用性
    fraIdx.Enabled = optIdx.Value And Not fraIdx.Enabled
    fraFKey.Enabled = optFKey.Value And Not fraFKey.Enabled
     
    chkOnln.Enabled = fraIdx.Enabled
    chkParallel.Enabled = fraIdx.Enabled
    lblParaNum.Enabled = fraIdx.Enabled
    txtParaNum.Enabled = fraIdx.Enabled And chkParallel.Value
    
    optDel.Enabled = fraFKey.Enabled
    optDisable.Enabled = fraFKey.Enabled
    
End Sub

Private Sub LoadGrid()
'功能： 初始化表格，加载表格数据。
    Dim rsData As ADODB.Recordset, i As Integer
    Dim strTblRange As String
    
    On Error GoTo errH
    
    If gblnIsZlhis Then
        '是否有Zltables这张表
        If gblnHasZltables Then
            strTblRange = " c.Table_Name Not In (Select 表名 From zlBaseCode) " & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And b.Table_Name In (Select 表名 From zlTables Where 分类 In ('B1','B2','B3','C1','C2','C3') )", "") & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And c.Table_Name In (Select 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3'))", "") & vbNewLine
        Else
            strTblRange = " c.Table_Name Not In (Select 表名 From zlBaseCode) " & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And b.Table_Name In (Select 表名 From zlBakTables  " & IIf(gblnHasBigtables, "Union All Select 表名 From Zlbigtables )", ")"), "") & vbNewLine & _
                                        IIf(chkZlhis.Value = 1, "And c.Table_Name In (Select 表名 From zlBakTables  " & IIf(gblnHasBigtables, "Union All Select 表名 From Zlbigtables )", ")"), "") & vbNewLine
        End If
    Else
        strTblRange = "b.Table_Name in (Select Table_Name From Dba_Tables Where Num_Rows > " & Val(txtRange.Text) & ")" & vbNewLine
    End If
    
    '查询出存在外键下子表缺失索引的数据
    'Child_Table-子表   Foreign_Key-外键  Columns-子表引用列  Main_Table-主表 Delete_Rule-外键删除规则
    gstrSQL = "Select Main_Table, Child_Table, Foreign_Key, Columns, Delete_Rule, Owner" & vbNewLine & _
                        "From (Select c.Table_Name As Main_Table, b.Table_Name As Child_Table, b.Constraint_Name As Foreign_Key," & vbNewLine & _
                        "              f_List2str(Cast(Collect(a.Column_Name Order By Position) As t_Strlist), ',', 1) As Columns, b.Delete_Rule," & vbNewLine & _
                        "              b.Owner" & vbNewLine & _
                        "       From Dba_Cons_Columns A, Dba_Constraints B, Dba_Constraints C" & vbNewLine & _
                        "       Where a.Constraint_Name = b.Constraint_Name And b.Status = 'ENABLED' And b.Constraint_Type = 'R' And" & vbNewLine & _
                        "             b.r_Constraint_Name <> '部门表_PK' And b.r_Constraint_Name = c.Constraint_Name And b.r_owner=c.owner  And A.Owner = B.Owner And" & vbNewLine & _
                        strTblRange & _
                        "       Group By c.Table_Name, b.Table_Name, b.Delete_Rule, b.Constraint_Name, b.r_Constraint_Name, b.Owner) A " & vbNewLine & _
                        "Where Not Exists" & vbNewLine & _
                        " (Select 1" & vbNewLine & _
                        "       From (Select Table_Name, Index_Name, Index_Owner," & vbNewLine & _
                        "                     f_List2str(Cast(Collect(Column_Name Order By Column_Position) As t_Strlist), ',', 1) As Columns" & vbNewLine & _
                        "              From Dba_Ind_Columns" & vbNewLine & _
                        "              Group By Table_Name, Index_Name, Index_Owner) B" & vbNewLine & _
                        "       Where a.Owner = b.Index_Owner And b.Table_Name = a.Child_Table And Instr(b.Columns, a.Columns) =1)" & vbNewLine
                        
    Set rsData = OpenSQLRecord(gstrSQL, Me.Caption)
    If rsData.RecordCount = 0 Then
        Call ClearVsf(vsfGrid, "当前环境没有外键子表缺失索引。")
        Exit Sub
    End If
    
    With vsfGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rsData.RecordCount + .FixedRows
        
        i = 1
        While Not rsData.EOF
            .TextMatrix(i, .ColIndex("所有者")) = rsData!Owner
            .TextMatrix(i, .ColIndex("子表")) = rsData!Child_Table
            .TextMatrix(i, .ColIndex("外键名称")) = rsData!Foreign_Key
            .TextMatrix(i, .ColIndex("外键字段")) = rsData!Columns
            .TextMatrix(i, .ColIndex("父表")) = rsData!Main_Table
            
            If rsData!Delete_Rule = "CASCADE" Then
                .TextMatrix(i, .ColIndex("级联操作")) = "更新"
            ElseIf rsData!Delete_Rule = "SET NULL" Then
                .TextMatrix(i, .ColIndex("级联操作")) = "清除"
            Else
                .TextMatrix(i, .ColIndex("级联操作")) = ""
            End If
            
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = BackAlterNate_颜色
            Else
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = Back_颜色
            End If
            
            i = i + 1
            rsData.MoveNext
        Wend
        .AutoSize .ColIndex("子表"), .ColIndex("级联操作"), False
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    Call SetCmdEnable(True)
End Sub

Private Sub pctLeft_Resize()
    On Error Resume Next
    vsfGrid.Width = pctLeft.Width - vsfGrid.Left
    vsfGrid.Height = pctLeft.Height - cmdRefresh.Height - lblPrompt.Height - 280
    
    cmdRefresh.Top = vsfGrid.Top + vsfGrid.Height + 60
    cmdRefresh.Left = vsfGrid.Left + vsfGrid.Width - cmdRefresh.Width
    
    lblRange1.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - lblRange1.Height / 2
    lblRange2.Top = lblRange1.Top
    txtRange.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - txtRange.Height / 2
    
    lblRange1.Left = vsfGrid.Left
    txtRange.Left = lblRange1.Left + lblRange1.Width + 45
    lblRange2.Left = txtRange.Left + txtRange.Width + 45
    
    chkZlhis.Top = cmdRefresh.Top + cmdRefresh.Height / 2 - chkZlhis.Height / 2 + 15
    chkZlhis.Left = cmdRefresh.Left - chkZlhis.Width
End Sub

Private Sub pctRight_Resize()
    
    On Error Resume Next
    txtSql.Height = vsfGrid.Top + vsfGrid.Height - txtSql.Top
    cmdExecute.Top = txtSql.Height + txtSql.Top + 60
    
    lblTip.Top = cmdExecute.Top + cmdRefresh.Height - lblTip.Height
    lblTip.Width = pctRight.Width - (cmdExecute.Left + cmdExecute.Width + 60)
End Sub

Private Sub txtParaNum_Change()
    Call GetSql
End Sub

Private Sub txtParaNum_KeyPress(KeyAscii As Integer)
    Call OnlyIntCK(KeyAscii)
    If Val(txtParaNum.Text & Chr(KeyAscii)) > gintCpuCount And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRange_KeyPress(KeyAscii As Integer)
    Call OnlyIntCK(KeyAscii)
End Sub

Private Sub txtSql_KeyPress(KeyAscii As Integer)
    Call OnlyStrChnCK(KeyAscii, " ", "_", Chr(3), Chr(22))
End Sub

Private Function GetTbSpace(ByVal strTbName As String, ByVal strOwner As String) As String
'功能：根据表名获取索引所在的表空间
'参数： strTbName - 表名
    Dim rstmp As ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "select a.table_name , a.tablespace_name ,count (b.index_name) index_Nums  ,b.tablespace_name Index_tbs" & vbNewLine & _
                    "from dba_tables  a , dba_indexes b" & vbNewLine & _
                    "where a.table_name = b.table_name(+)  and a.table_name = [1] and a.owner = [2]" & vbNewLine & _
                    "and a.temporary = 'N' and a.partitioned = 'NO'   and a.owner = b.owner" & vbNewLine & _
                    "group by  a.table_name,  a.tablespace_name ,b.tablespace_name" & vbNewLine & _
                    "order by index_Nums desc"

    Set rstmp = OpenSQLRecord(gstrSQL, Me.Caption, strTbName, strOwner)
    
    If rstmp.RecordCount = 0 Then Exit Function
    GetTbSpace = IIf(rstmp!index_Nums = 0, rstmp!Tablespace_Name, rstmp!Index_tbs)
    
    Exit Function
errH:
    Call ErrCenter(gstrSQL)

End Function

Private Sub GetSql()
'功能 ：根据界面所选，修改SQL指令
'控件选项保存在对应TAG中，实现语句如下
'create index  [索引]  on [表名]([列]) tablespace [表空间]  nologging [并行度] [在线];
'alter table [表名] [操作] constraint [外键]
    Dim strTbName As String, strCols As String, srtFkey As String
    Dim strTbSpace As String, strOwner As String
    
    With vsfGrid
      
        If .Rows < 2 Or .Row = 0 Then txtSql.Text = "": Exit Sub '防止索引越界
        If .TextMatrix(1, 0) = "当前环境没有外键子表缺失索引。" Then
            txtSql.Text = ""
            cmdExecute.Enabled = False
            Exit Sub
        Else
            cmdExecute.Enabled = True
        End If
        strTbName = .TextMatrix(.Row, .ColIndex("子表"))
        strCols = .TextMatrix(.Row, .ColIndex("外键字段"))
        srtFkey = .TextMatrix(.Row, .ColIndex("外键名称"))
        strOwner = .TextMatrix(.Row, .ColIndex("所有者"))
        strTbSpace = GetTbSpace(strTbName, strOwner)
    End With
    
    If optFKey Then  '删除外键
        optFKey.Tag = srtFkey
        optDel.Tag = IIf(optDel.Value, " Drop", " Disable")
        txtSql.Text = "Alter Table " & strOwner & "." & strTbName & optDel.Tag & " Constraint " & srtFkey
    Else    '建立索引
        If InStr(1, strCols, ",") > 0 Then
            optIdx.Tag = strOwner & "." & strTbName & "_IX_" & Mid(strCols, 1, InStr(1, strCols, ",") - 1)
        Else
            optIdx.Tag = strOwner & "." & strTbName & "_IX_" & strCols
        End If
        chkParallel.Tag = IIf(chkParallel.Value = 1, " Parallel " & txtParaNum.Text & " ", " ")
        chkOnln.Tag = IIf(chkOnln.Value = 1, "online ", " ")
        txtSql.Text = "Create Index " & optIdx.Tag & " On " & strOwner & "." & strTbName & "(" & strCols & ") Tablespace " & strTbSpace & " nologging" & _
                             chkParallel.Tag & chkOnln.Tag
    End If
    txtSql.Tag = txtSql.Text
        
End Sub


Private Sub vsfGrid_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow = 0 Then: Exit Sub   '防止刷新时误报
    If txtSql.Tag <> txtSql.Text Then
        If MsgBox("当前SQL指令已经发生更改，切换后所做修改将丢失。" & vbCrLf & vbCrLf & "是否切换？", _
            vbYesNo + vbQuestion + vbDefaultButton1, "确认") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub vsfGrid_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Call GetSql
End Sub

Private Sub SetCmdEnable(ByVal blnEnable As Boolean)
'功能： 设置按钮可用性和光标样式
    cmdRefresh.Enabled = blnEnable
    cmdExecute.Enabled = cmdRefresh.Enabled
    If blnEnable Then
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbArrowHourglass
    End If
End Sub

