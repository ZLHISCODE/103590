VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmHistSqlPlan 
   Caption         =   "SQL执行计划"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHistSqlPlan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeSyntaxEdit.SyntaxEdit txtSql 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _Version        =   983043
      _ExtentX        =   5318
      _ExtentY        =   3836
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox pctHorLine 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   135
      ScaleWidth      =   9015
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   9015
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
      Height          =   1095
      Index           =   1
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "颜色加重行标识当前语句引起性能问题的原因。"
      Top             =   3240
      Width           =   1815
      _cx             =   3201
      _cy             =   1931
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
      GridColor       =   -2147483643
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
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
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
   Begin MSComctlLib.TabStrip tabPlan 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSql 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL文本"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmHistSqlPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const conCol = "Operation,2000,1;Name,500,1;ID,500,1;Cardinality,500,1;Bytes,500,1;Cost,500,1;Time,500,1;Object_Owner,500,1;Object_Type,500,1"


Private Sub Form_Load()
    With txtSql
        '设置控件的显示颜色方案为：SQL
        .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
        .SyntaxScheme = GetSqlColor
    End With
    
    InitTable vsfPlan(1), conCol
End Sub

Private Sub Form_Resize()
    Dim objVsf As VSFlexGrid
    
    On Error Resume Next
    txtSql.Width = Me.ScaleWidth - txtSql.Left * 2
    pctHorLine.Top = txtSql.Top + txtSql.Height
    pctHorLine.Width = txtSql.Width
    pctHorLine.Left = txtSql.Left
    tabPlan.Width = txtSql.Width
    
    For Each objVsf In vsfPlan
        objVsf.Width = txtSql.Width
        objVsf.Height = Me.ScaleHeight - objVsf.Top - 120
    Next
End Sub


Private Sub pctHorLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objVsf As VSFlexGrid
    Dim intY As Integer, intOldHeight As Integer
    
    On Error Resume Next
    If Button <> 1 Then Exit Sub
    '防止拖动过度，导致界面异常
    If pctHorLine.Top + y < 360 Then Exit Sub
    If pctHorLine.Top + y > 10095 Then Exit Sub
    
    intOldHeight = txtSql.Height
    pctHorLine.Top = pctHorLine.Top + y
    txtSql.Height = Abs(pctHorLine.Top - txtSql.Top)
    tabPlan.Top = pctHorLine.Top + 240
    
    For Each objVsf In vsfPlan
        objVsf.Top = tabPlan.Top + tabPlan.Height
        objVsf.Height = objVsf.Height - (txtSql.Height - intOldHeight)
    Next
    
    Me.Refresh
    
End Sub


Public Sub ShowMe(ByVal strSqlId As String)
    LoadSqlText strSqlId
    LoadSqlPlan strSqlId
    Me.Show 1
End Sub

Private Sub LoadSqlText(ByVal strSqlId As String)
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select sql_text   From dba_hist_sqltext  Where Sql_Id = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "获取SQL文本", strSqlId)
    
    If rsTmp.RecordCount <> 0 Then
        txtSql.Text = rsTmp!sql_text
    End If
    
    Exit Sub
errH:
    MsgBox "获取SQL文本发生错误." & vbNewLine & err.Description
End Sub

Private Sub LoadSqlPlan(ByVal strSqlId As String)
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim intPlanNum  As Integer
    
    On Error GoTo errH
    '获取执行计划子游标的条数
    strSql = "Select a.*, Rownum - 1 Child_Number" & vbNewLine & _
                "From (Select Distinct Plan_Hash_Value From Dba_Hist_Sqlstat Where Sql_Id = [1] Order By Plan_Hash_Value) A"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "获取执行计划个数", strSqlId)
    
    If rsTmp.RecordCount = 0 Then Exit Sub  ' 没有执行计划就退出
    tabPlan.Tabs().Clear
    Do While Not rsTmp.EOF
        '添加TAB
        If rsTmp.RecordCount = 1 Then
            tabPlan.Tabs().Add intPlanNum + 1, , "执行计划"
        Else
            tabPlan.Tabs().Add intPlanNum + 1, , "执行计划" & intPlanNum + 1
        End If

        '添加VSFGRID,Index为1的VSFGRID不用重复加载
        If intPlanNum > 0 Then
            Load vsfPlan(intPlanNum + 1)
            Call InitTable(vsfPlan(intPlanNum + 1), conCol)
        End If
        
        GetSqlPlanByChild vsfPlan(intPlanNum + 1), strSqlId, intPlanNum
        intPlanNum = intPlanNum + 1
        If intPlanNum = 9 Or intPlanNum = rsTmp.RecordCount Then Exit Do  '控制最多显示10个子计划
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    MsgBox "获取执行计划游标数发生错误." & vbNewLine & err.Description
End Sub


Private Sub GetSqlPlanByChild(vsfPlan As VSFlexGrid, strSqlId As String, intChild As Integer)
    '根据SQLID ChildNumber绘制执行计划图
    Dim strSql As String, rsPlan As New ADODB.Recordset
    Dim intRowNum As Integer
    
    On Error GoTo errH
    
    strSql = "Select *" & vbNewLine & _
            "From (Select /*+ no_merge */" & vbNewLine & _
            "        LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name, ID, Cardinality, Bytes," & vbNewLine & _
            "        Cost, Time, Object_Owner, Object_Type" & vbNewLine & _
            "       From (Select *" & vbNewLine & _
            "              From Dba_Hist_Sql_Plan" & vbNewLine & _
            "              Where Sql_Id = [1] And" & vbNewLine & _
            "                    Plan_Hash_Value = (Select Plan_Hash_Value" & vbNewLine & _
            "                                       From (Select Plan_Hash_Value, Rownum - 1 Child_Number" & vbNewLine & _
            "                                              From (Select Distinct Plan_Hash_Value" & vbNewLine & _
            "                                                     From Dba_Hist_Sqlstat" & vbNewLine & _
            "                                                     Where Sql_Id = [1] " & vbNewLine & _
            "                                                     Order By Plan_Hash_Value) A)" & vbNewLine & _
            "                                       Where Child_Number = [2])) A" & vbNewLine & _
            "       Start With a.Id = 0" & vbNewLine & _
            "       Connect By Prior a.Id = a.Parent_Id" & vbNewLine & _
            "       Order By ID, Position)"
    Set rsPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, "获取执行计划", strSqlId, intChild)
    
    With vsfPlan
        .Redraw = flexRDNone
        .FixedCols = 0
        .Editable = flexEDNone
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarSimpleLeaf
        .SubtotalPosition = flexSTAbove
        .AllowUserResizing = flexResizeColumns
        .Rows = .FixedRows
        .Rows = .FixedRows + rsPlan.RecordCount
        intRowNum = 1
        
        If rsPlan.RecordCount = 0 Then
            .Rows = 1
        End If
        Do While Not rsPlan.EOF
            
            .TextMatrix(intRowNum, .ColIndex("Operation")) = "" & LTrim(rsPlan!Operation)
            If InStr(1, "TABLE ACCESS FULL/INDEX FULL SCAN/INDEX SKIP SCAN/ ", LTrim(rsPlan!Operation)) > 0 Then
                 .Cell(flexcpBackColor, intRowNum, 0, intRowNum, .Cols - 1) = &HB3DEF5
            End If
            .TextMatrix(intRowNum, .ColIndex("Name")) = "" & rsPlan!name
            .TextMatrix(intRowNum, .ColIndex("ID")) = "" & rsPlan!Id
            .TextMatrix(intRowNum, .ColIndex("Cardinality")) = "" & rsPlan!Cardinality
            .TextMatrix(intRowNum, .ColIndex("Bytes")) = "" & rsPlan!Bytes
            .TextMatrix(intRowNum, .ColIndex("Cost")) = "" & rsPlan!Cost
            .TextMatrix(intRowNum, .ColIndex("Time")) = "" & rsPlan!Time
            .TextMatrix(intRowNum, .ColIndex("Object_Owner")) = "" & rsPlan!Object_Owner
            .TextMatrix(intRowNum, .ColIndex("Object_Type")) = "" & rsPlan!Object_Type
            
            .RowOutlineLevel(intRowNum) = Len(rsPlan!Operation) - Len(LTrim(rsPlan!Operation)) '以空格个数控制树形结构的等级
            .IsSubtotal(intRowNum) = True
            intRowNum = intRowNum + 1
            rsPlan.MoveNext
        Loop
        .AutoResize = True
        .AutoSize .ColIndex("Operation"), .ColIndex("Object_Owner"), False
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    MsgBox "获取执行计划发生错误." & vbNewLine & err.Description
    
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub tabPlan_Click()
    Dim intPlanNum As Integer
    
    '显示当前选中计划
    If Val(tabPlan.SelectedItem.Index) = Val(tabPlan.Tag) Or tabPlan.Tag = "" Then Exit Sub
    
    vsfPlan(tabPlan.SelectedItem.Index).Visible = True
    vsfPlan(tabPlan.Tag).Visible = False
    tabPlan.Tag = tabPlan.SelectedItem.Index
End Sub


