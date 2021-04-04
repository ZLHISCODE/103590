VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReused 
   BorderStyle     =   0  'None
   ClientHeight    =   7860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   17970
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGO 
      Caption         =   "定位到LOB"
      Height          =   350
      Left            =   16560
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkFree 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "只显示空块"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13560
      TabIndex        =   17
      Top             =   65
      Width           =   1275
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   17970
      TabIndex        =   6
      Top             =   7245
      Width           =   17970
      Begin VB.CommandButton cmdMore 
         Caption         =   "更多(&4)"
         Height          =   350
         Left            =   15618
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox chkOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         Caption         =   "在线重整索引"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   16
         Top             =   180
         Value           =   1  'Checked
         Width           =   1400
      End
      Begin VB.TextBox txtParallel 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   10245
         TabIndex        =   15
         Text            =   "12"
         ToolTipText     =   "设置了并行度后，重整(Move)操作时容易导致新分配的空间位于数据文件的尾部，从而导致无法收缩数据文件。"
         Top             =   160
         Width           =   375
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&S)"
         Height          =   350
         Left            =   2745
         TabIndex        =   13
         Top             =   117
         Width           =   1095
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   1500
         TabIndex        =   12
         Top             =   117
         Width           =   1200
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "回收(&2)"
         Height          =   350
         Left            =   13246
         TabIndex        =   7
         ToolTipText     =   "一般用于大量删除数据后降低高水标记以收回空间"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "重整(&1)"
         Height          =   350
         Left            =   12060
         TabIndex        =   8
         ToolTipText     =   "用于移动块的物理位置以便收缩文件"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "收缩(&3)"
         Height          =   350
         Left            =   14432
         TabIndex        =   9
         ToolTipText     =   "收缩当前表空间中当前数据文件的大小"
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblParallel 
         BackStyle       =   0  'Transparent
         Caption         =   "重整并行度"
         Height          =   255
         Left            =   9315
         TabIndex        =   14
         Top             =   210
         Width           =   930
      End
      Begin VB.Label lblFind 
         BackColor       =   &H00EFF0E0&
         Caption         =   "表或索引名称(&F)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label lblOptPrompt 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         ForeColor       =   &H00400000&
         Height          =   180
         Left            =   3945
         TabIndex        =   10
         Top             =   195
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfExtents 
      Height          =   6375
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   14175
      _cx             =   25003
      _cy             =   11245
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
      ForeColorSel    =   12582912
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsfTbs 
      Height          =   6735
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   3435
      _cx             =   6059
      _cy             =   11880
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   380
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.ComboBox cboFiles 
      Height          =   300
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   65
      Width           =   8895
   End
   Begin VB.Label lblPrompt 
      Caption         =   "当前选中Extent的信息"
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   12855
   End
   Begin VB.Label lblFiles 
      Caption         =   "表空间文件"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblTableSpaces 
      Caption         =   "表空间列表"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuResize 
      Caption         =   "收缩选项"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuResizeAll 
         Caption         =   "收缩全部数据文件"
      End
      Begin VB.Menu mnuResizeTemp 
         Caption         =   "收缩临时数据文件"
      End
      Begin VB.Menu mnuResizeUndo 
         Caption         =   "收缩Undo表空间"
      End
      Begin VB.Menu mnuAddFile 
         Caption         =   "添加数据文件"
      End
   End
End
Attribute VB_Name = "frmReused"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONCOLS As Long = 50
Private Const CONBLOCKS As Long = 8
Private mrsExtents As ADODB.Recordset
Private mrsLobs As ADODB.Recordset
Private mcolCells As Collection
Private mlngRowPre As Long, mlngColPre As Long

Private Enum opt
    P1回收 = 1
    P2重整
    P3收缩
End Enum

Public Sub ShowMe()
    Me.Show
End Sub

Private Sub cboFiles_Click()
    
    '由于循环中使用了doevents，所以需禁用任何可操作的功能
    Call SetCommandEnable(0)
    
    On Error GoTo errH
    
    Call LoadExtents(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")), Val(cboFiles.ItemData(cboFiles.ListIndex)))
    cboFiles.ToolTipText = cboFiles.List(cboFiles.ListIndex)
    Call SetCommandEnable(1)
    
    If Me.Visible And Me.Enabled Then
        vsfExtents.SetFocus
    End If
    vsfExtents.Select vsfExtents.Rows - 1, vsfExtents.Cols - 1
    vsfExtents.TopRow = vsfExtents.Row
    
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub SetCommandEnable(bytEnable As Byte)
'功能：设置命令按钮的可用性
    cmdShrink.Enabled = bytEnable = 1
    cmdMove.Enabled = cmdShrink.Enabled
    cmdResize.Enabled = cmdShrink.Enabled
    cmdMore.Enabled = cmdShrink.Enabled
    chkFree.Enabled = cmdShrink.Enabled
    chkOnline.Enabled = cmdShrink.Enabled
    cmdFind.Enabled = cmdShrink.Enabled
    txtFind.Enabled = cmdShrink.Enabled
    If txtParallel.Locked = False Then txtParallel.Enabled = cmdShrink.Enabled
    
    If cmdGO.Visible Then cmdGO.Enabled = cmdShrink.Enabled
    
    vsfTbs.Enabled = cmdShrink.Enabled
    cboFiles.Enabled = cmdShrink.Enabled
End Sub

Private Sub chkFree_Click()
    If cboFiles.ListIndex >= 0 Then Call cboFiles_Click
End Sub

Private Function ResizeTBS(ByVal strTBS As String, Optional ByVal lngFile As Long) As Boolean
'功能：收缩表空间
'参数：strTBS-表空间名称
'      blnPrompt-数据文件号,不传入时，在不提示的情况下收缩当前表空间的所有数据文件至最小尺寸
    Dim strSql As String, dblMax As Double, dblFileSize As Double, dblLimit As Double, dblBlockSize As Double
    Dim i As Long, blnTry As Boolean
    Dim rsTmp As ADODB.Recordset
           
    On Error GoTo errH
    
    dblBlockSize = Val(vsfTbs.RowData(vsfTbs.Row))
    If dblBlockSize = 0 Then dblBlockSize = 8192
        
    If lngFile <> 0 Then
        dblLimit = CDbl(1024) * 1024 * 2
        
        strSql = "Select a.File_Id, a.Last_Block, b.Bytes" & vbNewLine & _
            "From (Select a.File_Id, Max(a.Block_Id + a.Blocks - 1) Last_Block" & vbNewLine & _
            "       From Dba_Extents A" & vbNewLine & _
            "       Where a.Tablespace_Name = [1] And File_Id = [2]" & vbNewLine & _
            "       Group By a.File_Id) A, Dba_Data_Files B" & vbNewLine & _
            "Where a.File_Id = b.File_Id"
        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS, lngFile)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "在Dba_Extents中没有找到当前表空间及数据文件的记录", vbInformation, "错误"
            Exit Function
        End If
    
        dblMax = rsTmp!Last_Block * dblBlockSize
        dblFileSize = rsTmp!Bytes
        If dblFileSize - dblMax < dblLimit Then '小于2M，不收缩
            If MsgBox("可收缩的空间(" & Round((dblFileSize - dblMax) / 1024) & "KB)小于2M,是否确实要收缩该文件？", vbYesNo + vbDefaultButton2, "提醒") = vbNo Then
                Exit Function
            End If
            dblMax = Round(dblMax / 1024 / 1024) + 1 '取整加1，单位M
        Else
            dblMax = Round(dblMax / 1024 / 1024) + 1 '取整加1，单位M
            If MsgBox("你确定要将当前文件收缩到" & dblMax & "M吗?", vbQuestion + vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
                Exit Function
            End If
        End If
        
        If dblMax >= Round(rsTmp!Bytes / 1024 / 1024) Then
            MsgBox "数据文件已达到最大尺寸，无法更改！", vbInformation
        Else
            Err.Clear
            On Error Resume Next
retry1:     strSql = "Alter Database Datafile " & lngFile & " Resize " & CStr(dblMax) & "M"
            gcnOracle.Execute strSql
            
            If Err.Number <> 0 Then
                If MsgBox("收缩数据文件失败，可能是删除对象后未清空回收站引起的，是否清空后重试？", vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                    Err.Clear
                    strSql = "purge tablespace " & strTBS
                    gcnOracle.Execute strSql
                    GoTo retry1
                Else
                    GoTo errH
                End If
            End If
            ResizeTBS = True
        End If
        
    Else
        dblLimit = CDbl(1024) * 1024 * 10
        
        '可收缩空间小于10M，不收缩，避免在循环中频繁执行收缩
        strSql = "Select a.File_Id, a.Last_Block * " & dblBlockSize & " as MaxBytes, b.Bytes" & vbNewLine & _
                "From (Select a.File_Id, Max(a.Block_Id + a.Blocks - 1) Last_Block" & vbNewLine & _
                "       From Dba_Extents A" & vbNewLine & _
                "       Where a.Tablespace_Name = [1]" & vbNewLine & _
                "       Group By a.File_Id) A, Dba_Data_Files B" & vbNewLine & _
                "Where a.File_Id = b.File_Id And (b.Bytes - a.Last_Block * " & dblBlockSize & ") > " & dblLimit
    
        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS)
        
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            dblMax = Round(rsTmp!MaxBytes / 1024 / 1024) + 1 '取整加1，单位M
            If dblMax < Round(rsTmp!Bytes / 1024 / 1024) Then
                lblOptPrompt.Caption = "收缩" & rsTmp!File_Id & "号数据文件至" & CStr(dblMax) & "M"
                                
                blnTry = False
retry2:         strSql = "Alter Database Datafile " & rsTmp!File_Id & " Resize " & CStr(dblMax) & "M"
                gcnOracle.Execute strSql
                If Err.Number <> 0 And blnTry = False Then
                    Err.Clear
                    strSql = "purge tablespace " & strTBS
                    gcnOracle.Execute strSql
                    blnTry = True
                    GoTo retry2
                Else
                    Err.Clear   '重试一次后跳过
                End If
                
                ResizeTBS = True
            End If
            
            rsTmp.MoveNext
        Next
    End If
    
    Exit Function
errH:
    Call ErrCenter(strSql)
    Call SetCommandEnable(1)
End Function


Private Sub cmdMore_Click()
    Me.PopupMenu mnuResize
End Sub

Private Sub cmdResize_Click()
'功能：执行表空间收缩
    If cboFiles.ListIndex < 0 Then
        MsgBox "请选择一个数据文件！", vbInformation, "提醒"
        If cboFiles.Enabled Then cboFiles.SetFocus
    Else
        Call SetCommandEnable(0)
        
        If ResizeTBS(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")), Val(cboFiles.ItemData(cboFiles.ListIndex))) Then
            lblOptPrompt.Caption = "已完成文件收缩，正在刷新。"
            lblOptPrompt.Refresh
            
            Call RefreshData
            
            lblOptPrompt.Caption = "已完成操作。"
        End If
        
        Call SetCommandEnable(1)
    End If
End Sub

Private Sub RefreshData()
'功能：刷新当前表空间的当前数据文件的段的数据信息

    Dim i As Long, strTBS As String
    Dim lngFile As Long
    
    strTBS = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称"))
    lngFile = cboFiles.ListIndex
    Call LoadTablespaces
    
    vsfTbs.Redraw = flexRDNone
    i = vsfTbs.FindRow(strTBS, , vsfTbs.ColIndex("名称"))
    If i <> -1 Then vsfTbs.Row = i: vsfTbs.TopRow = i
    vsfTbs.Redraw = flexRDDirect
    
    Call LoadFiles(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")))
    If lngFile <= cboFiles.ListCount Then
        cboFiles.ListIndex = lngFile
    Else
        cboFiles.ListIndex = 0
    End If
End Sub

Private Function CheckUnSuportObject(strSegment As String, strOpt As String) As Boolean
'功能：检查指定的表是否存在Move或Shrink不支持的对象
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 1" & vbNewLine & _
            "From All_Tab_Columns" & vbNewLine & _
            "Where Table_Name = [2] And Owner = [1] And Data_Type In ('LONG','LONG RAW','UNDEFINED')"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rsTmp.RecordCount > 0 Then
        lblOptPrompt.Caption = strSegment & "含有LONG,LONG RAW类型字段，不能进行" & strOpt & "操作."
    Else
        CheckUnSuportObject = True
    End If
End Function

Private Function CheckIOT(strSegment As String) As Boolean
'功能：检查指定的索引是否为索引组织表的索引（不支持重建）
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Indexes Where Owner = [1] And Index_Name = [2] And Index_Type = 'IOT - TOP'"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckIOT = rsTmp.RecordCount > 0
End Function

Private Function GetIOTName(strSegment As String) As String
'功能：根据索引组织表的索引名返回索引组织表名(含所有者前缀)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select Table_Owner||'.'||Table_Name as Tab_Name From All_Indexes Where Owner = [1] And Index_Name = [2] And Index_Type = 'IOT - TOP'"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rsTmp.RecordCount > 0 Then
        GetIOTName = rsTmp!Tab_Name
    End If
End Function


Private Function CheckLOBIndex(strSegment As String) As Boolean
'功能：检查指定的索引是否为LOB的索引（不支持重建）
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Lobs Where Owner = [1] And Index_Name = [2]"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckLOBIndex = rsTmp.RecordCount > 0
End Function

Private Function GetLOBNameByIndex(strSegment As String) As String
'功能：检查指定的索引是否为LOB的索引（不支持重建）
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select Segment_Name From All_Lobs Where Owner = [1] And Index_Name = [2]"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rsTmp.RecordCount > 0 Then
        GetLOBNameByIndex = Split(strSegment, ".")(0) & "." & rsTmp!Segment_Name
    End If
End Function

Private Sub ReBuildIndex(ByVal strOwner As String, ByVal strTable As String, ByVal strParallel As String)
'功能：重建某张表上失效的索引
'参数：strOwner=所有者,strTable=表名
'      strParallel=" Parallel X",并行度
    Dim rsTmp As ADODB.Recordset, rsIndex As ADODB.Recordset
    Dim strSql As String
    
    lblOptPrompt.Caption = "正在重建[" & strOwner & "." & strTable & "]上失效的索引"
    On Error GoTo errH
    
    '重建失效的索引
    strSql = "Select Index_Name From DBA_Indexes Where Status='UNUSABLE' And Owner = [1] And Table_Name = [2]"
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strOwner, strTable)
    
    Do While Not rsTmp.EOF
        '如果是分区索引，则要单独处理
        strSql = "Select Partition_Name From Dba_Ind_Partitions Where Index_Owner = [1] And Index_Name = [2]"
        Set rsIndex = OpenSQLRecord(strSql, Me.Caption, strOwner, rsTmp!Index_Name)
        If rsIndex.RecordCount > 0 Then
            Do While Not rsIndex.EOF
                strSql = "Alter Index " & strOwner & "." & rsTmp!Index_Name & " Rebuild Partition " & rsIndex!Partition_Name & " Nologging" & strParallel
                gcnOracle.Execute strSql
                rsIndex.MoveNext
            Loop
        Else
            strSql = "Alter Index " & strOwner & "." & rsTmp!Index_Name & " Rebuild Nologging" & strParallel
            gcnOracle.Execute strSql
        End If
        
        If strParallel <> "" Then
            strSql = "Alter Index " & strOwner & "." & rsTmp!Index_Name & " NOParallel"
            gcnOracle.Execute strSql
        End If
        
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Sub cmdMove_Click()
'功能：执行表或索引的空闲空间重整(Move)
    Dim strSql As String, strType As String, strPartition As String, strSegment As String, strSegmentPre As String, strSegmentAll As String
    Dim strTBSTemp As String, strTbsOriginal As String, strTbsLob As String, strColumn As String, strParallel As String, strTableName As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTbs As ADODB.Recordset
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, c As Long, r As Long
    Dim arrTmp As Variant, blnRemove As Boolean
    Dim strPrompt As String, strOnline As String, strObjName As String
    
    '不收回空间的对象，先移到临时存储的表空间，最后才移回来
    Dim strRemoveIndex As String, strRemoveLob As String, strRemoveTable As String
    Dim strRemovePARTable As String, strRemovePARIndex As String, strRemovePARLOB As String
    Dim datBegin As Date, strTime As String
    
    If CheckExtent(P2重整) = False Then Exit Sub
    On Error GoTo errH
    strType = Trim(lblPrompt.Tag)
    If strType <> "" And InStr(",TABLE,TABLE PARTITION,INDEX,INDEX PARTITION,LOBSEGMENT,LOBINDEX,LOB PARTITION,", "," & strType & ",") = 0 Then '对LOBINDEX对象，则重整其LOBSEGMENT
        Call MsgBox("仅支持对表或索引进行空闲空间收回，不支持的数据类型：" & strType, vbInformation, Me.Caption)
        Exit Sub
    End If
    Call SetCommandEnable(0)
    
reInput:    strTBSTemp = Trim(InputBox("    为了将可能位于数据文件末尾的当前段移到前面，需要将此段先移到一个临时存放的表空间，收缩数据文件之后再移回来。" & vbCrLf & _
                    "    如果选择“取消”按钮，则直接在当前表空间进行重整操作(重整后，当前段可能仍然位于数据文件的末尾)。", "临时存放的表空间", "SYSAUX"))
    If strTBSTemp <> "" Then
        strTBSTemp = UCase(strTBSTemp)
        
        strSql = "Select 1 From DBA_TABLESPACES Where TABLESPACE_NAME = [1]"
        Set rsTbs = OpenSQLRecord(strSql, "表空间检查", strTBSTemp)
        If rsTbs.RecordCount = 0 Then
            MsgBox "输入的表空间不存在，请重新输入", vbExclamation, "提示"
            GoTo reInput
        End If
    End If
    strTbsOriginal = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称"))
    If strTBSTemp = strTbsOriginal Then strTBSTemp = ""
    
    If txtParallel.Text <> "0" Then strParallel = " Parallel " & txtParallel.Text
    If chkOnline.Value = 1 Then strOnline = "Online"
    
    Me.Refresh  '避免表格上的残影
    datBegin = GetCurrentdate
    
    
    '处理一次选择多行多列的情况
    With vsfExtents
        .GetSelection r1, c1, r2, c2
                
        For r = r2 To r1 Step -1
            For c = c2 To c1 Step -1
                strSegment = mcolCells("_" & r & "_" & c)     '含所有者
                If strSegment <> strSegmentPre Then
                    If InStr(strSegmentAll & ",", "," & strSegment & ",") = 0 Then
                    
                        mrsExtents.Filter = "Row=" & r & " And Col=" & c
                        If mrsExtents.RecordCount > 0 Then
                            DoEvents
                            strType = mrsExtents!Segment_Type
                            '1.普通表
                            If strType = "TABLE" Then
                                'mdsys用户下存在类似GridFile1044_TAB这种含有小写字母的表段，但在dba_tables中却查不到记录
                                If CheckUnSuportTable(Split(strSegment, ".")(0), Split(strSegment, ".")(1)) Then
                                
                                    If CheckUnSuportObject(strSegment, "重整(Move)") Then
                                        lblOptPrompt.Caption = "正在对[" & strSegment & "]进行重整"
                                        lblOptPrompt.Refresh
                                        If strTBSTemp = "" Then
                                            strSql = "Alter Table " & strSegment & " Move Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
                                        Else
                                            strSql = "Alter Table " & strSegment & " Move TableSpace " & strTBSTemp & " Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            strRemoveTable = strRemoveTable & "," & strSegment & "||" & strTbsOriginal
                                        End If
                                        
                                        If strParallel <> "" Then
                                            strSql = "Alter Table " & strSegment & " NOParallel"
                                            gcnOracle.Execute strSql
                                        End If
                                    Else
                                        If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "含有Long或Long Raw字段的表:" & strSegment
                                    End If
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "IOT索引的溢出表或含有自定义字段的表:" & strSegment
                                End If
                            '2.分区表(不含LOB分区表)
                            ElseIf strType = "TABLE PARTITION" Then
                                If CheckUnSuportObject(strSegment, "重整(Move)") Then
                                    
                                    strSql = "Select Partition_Name From Dba_Tab_Partitions Where Table_Owner = [1] And Table_Name = [2]"
                                    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                    Do While Not rsTmp.EOF
                                        strPartition = rsTmp!Partition_Name
                                        
                                        lblOptPrompt.Caption = "正在对[" & strSegment & "(" & strPartition & ")]进行重整"
                                        lblOptPrompt.Refresh
                                        
                                        '未加级联更新索引update indexes，在后面调用ReBuildIndex来恢复，因为可能移两次
                                        If strTBSTemp = "" Then
                                            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            
                                            Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
                                        Else
                                            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " TableSpace " & strTBSTemp & " Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            strRemovePARTable = strRemovePARTable & "," & strSegment & "||" & strPartition & "||" & strTbsOriginal
                                        End If
                                        rsTmp.MoveNext
                                    Loop
                                    
                                    If strParallel <> "" Then
                                        strSql = "Alter Table " & strSegment & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "含有Long或Long Raw字段的分区表:" & strSegment
                                End If
                                
                            '3.LOB段（不含LOB分区索引和LOB分区表）
                            ElseIf strType = "LOBSEGMENT" Or strType = "LOBINDEX" Then
                                If strType = "LOBINDEX" Then
                                    strSql = "Select Owner ||'.'|| Segment_Name as Segment_Name From Dba_Lobs Where Owner = [1] And Index_Name = [2]"
                                    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                    If rsTmp.RecordCount > 0 Then
                                        If InStr(strSegmentAll & ",", "," & rsTmp!Segment_Name & ",") = 0 Then
                                            strSegment = rsTmp!Segment_Name
                                        Else
                                            GoTo NextCell '如果LOB已重整过，则跳过
                                        End If
                                    End If
                                End If
                            
                                mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"   '为了取表名和列名
                                If mrsLobs.RecordCount > 0 Then
                                    'mdsys用户下存在类似GridFile1044_TAB这种含有小写字母的表段，但在dba_tables中却查不到记录
                                    If CheckUnSuportTable(mrsLobs!Owner, mrsLobs!Table_name) Then
                                        strTableName = mrsLobs!Owner & "." & mrsLobs!Table_name
                                        strTbsLob = mrsLobs!Tablespace_Name
                                        strColumn = mrsLobs!Column_Name
                                        
                                        lblOptPrompt.Caption = "正在对[" & strTableName & "(" & strColumn & ")]进行重整"
                                        lblOptPrompt.Refresh
                                        If strTBSTemp = "" Then
                                            strSql = "ALTER TABLE " & strTableName & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsLob & ") Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            
                                        Else
                                            strSql = "ALTER TABLE " & strTableName & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTBSTemp & ") Nologging" & strParallel
                                            gcnOracle.Execute strSql
                                            strRemoveLob = strRemoveLob & "," & strTableName & "||" & strColumn & "||" & strTbsLob
                                        End If
                                                                        
                                        'LOB并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
                                    Else
                                        If InStr(strPrompt, ":" & mrsLobs!Table_name) = 0 Then strPrompt = strPrompt & vbCrLf & "未支持的表:" & mrsLobs!Table_name
                                    End If
                                Else
                                    lblOptPrompt.Caption = "在视图Dba_Lobs中未找到LOB对象" & strSegment & "。"
                                End If
                                
                            '4.普通索引
                            ElseIf strType = "INDEX" Then
                                If CheckIOT(strSegment) = False Then    'IOT索引只能通过move原表重建
                                    lblOptPrompt.Caption = "正在对[" & strSegment & "]进行重建"
                                    lblOptPrompt.Refresh
                                    If strTBSTemp = "" Then
                                        strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                    Else
                                        strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " TableSpace " & strTBSTemp & " Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                        strRemoveIndex = strRemoveIndex & "," & strSegment & "||" & strTbsOriginal
                                    End If
                                                                
                                    If strParallel <> "" Then
                                        strSql = "Alter Index " & strSegment & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                    
                                Else 'IOT索引组织表
                                    strObjName = GetIOTName(strSegment)
                                    
                                    lblOptPrompt.Caption = "正在对[" & strObjName & "]进行重整"
                                    lblOptPrompt.Refresh
                                    
                                    If strTBSTemp = "" Then
                                        strSql = "Alter Table " & strObjName & " Move Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                    Else
                                        strSql = "Alter Table " & strObjName & " Move TableSpace " & strTBSTemp & " Nologging" & strParallel
                                        gcnOracle.Execute strSql
                                        strRemoveTable = strRemoveTable & "," & strObjName & "||" & strTbsOriginal
                                    End If
                                    
                                    If strParallel <> "" Then
                                        strSql = "Alter Table " & strObjName & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                End If
                                
                            '5.分区索引
                            ElseIf strType = "INDEX PARTITION" Then
                                If CheckLOBIndex(strSegment) Then
                                    'LOB分区索引跟LOB分区表一起Move
                                    
                                ElseIf CheckIOT(strSegment) = False Then
                                    
                                    strSql = "Select Partition_Name From Dba_Ind_Partitions Where Index_Owner = [1] And Index_Name = [2]"
                                    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                    Do While Not rsTmp.EOF
                                        strPartition = rsTmp!Partition_Name
                                    
                                        lblOptPrompt.Caption = "正在对[" & strSegment & "(" & strPartition & ")]进行重建"
                                        lblOptPrompt.Refresh
                                        If strTBSTemp = "" Then
                                            strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " Nologging" & strParallel & " " & strOnline
                                            gcnOracle.Execute strSql
                                        Else
                                            strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " TableSpace " & strTBSTemp & " Nologging" & strParallel & " " & strOnline
                                            gcnOracle.Execute strSql
                                            strRemovePARIndex = strRemovePARIndex & "," & strSegment & "||" & strPartition & "||" & strTbsOriginal
                                        End If
                                        rsTmp.MoveNext
                                    Loop
                                                                
                                    If strParallel <> "" Then
                                        strSql = "Alter Index " & strSegment & " NOParallel"
                                        gcnOracle.Execute strSql
                                    End If
                                    
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "索引组织表（IOT）的分区索引:" & strSegment
                                End If
                                
                            '6.LOB分区表
                            ElseIf strType = "LOB PARTITION" Then
                                If CheckUnSuportObject(strSegment, "重整(Move)") Then
                                                                        
                                    mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"   '为了取表空间名
                                    If mrsLobs.RecordCount > 0 Then
                                        strTableName = mrsLobs!Owner & "." & mrsLobs!Table_name
                                        strTbsLob = mrsLobs!Tablespace_Name
                                        strColumn = mrsLobs!Column_Name
                                        
                                        strSql = "Select Partition_Name From Dba_Lob_Partitions Where Table_Owner = [1] And Lob_Name = [2]"
                                        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                                        Do While Not rsTmp.EOF
                                            strPartition = rsTmp!Partition_Name
                                            
                                            lblOptPrompt.Caption = "正在对[" & strTableName & "(" & strPartition & ")]进行重整"
                                            lblOptPrompt.Refresh
                                            If strTBSTemp = "" Then
                                                strSql = "Alter Table " & strTableName & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsLob & ") Nologging" & strParallel
                                                gcnOracle.Execute strSql
                                            Else
                                                strSql = "Alter Table " & strTableName & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTBSTemp & ") Nologging" & strParallel
                                                gcnOracle.Execute strSql
                                                strRemovePARLOB = strRemovePARLOB & "," & strTableName & "||" & strPartition & "||" & strColumn & "||" & strTbsLob
                                            End If
                                            rsTmp.MoveNext
                                        Loop
                                        
                                        'LOB分区并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
                                        If strTBSTemp = "" Then Call ReBuildIndex(Split(strTableName, ".")(0), Split(strTableName, ".")(1), strParallel)
                                        
                                    Else
                                        lblOptPrompt.Caption = "在视图Dba_Lobs中未找到LOB对象" & strSegment & "。"
                                    End If
                                Else
                                    If InStr(strPrompt, ":" & strSegment) = 0 Then strPrompt = strPrompt & vbCrLf & "含有Long或Long Raw字段的分区表:" & strSegment
                                End If
                            
                            ElseIf strType <> " " Then
                                lblOptPrompt.Caption = strSegment & ",不支持的对象类型：" & strType
                            End If
                        End If
NextCell:               strSegmentAll = strSegmentAll & "," & strSegment
                    End If
                    strSegmentPre = strSegment
                End If
            Next
            lblOptPrompt.Caption = "已处理完第" & r & "行"
            lblOptPrompt.Refresh
        Next
    End With
    
    If strTBSTemp <> "" Then
        If strRemoveTable & strRemovePARTable & strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
            Call ResizeTBS(strTBSTemp)
        End If
    End If
    
    
    '对没有收回空间的，移到临时存储的表空间的对象，移回原表空间。
reMove: blnRemove = True
    '1.表
   If strRemoveTable <> "" Then
        arrTmp = Split(Mid(strRemoveTable, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strTbsOriginal = Split(arrTmp(r), "||")(1)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
                Call ResizeTBS(strTbsOriginal)
            End If
            
            lblOptPrompt.Caption = "正在将[" & strSegment & "]移回原表空间"
            lblOptPrompt.Refresh
            strSql = "Alter Table " & strSegment & " Move TableSpace " & strTbsOriginal & " Nologging" & strParallel
            gcnOracle.Execute strSql
                        
            If strParallel <> "" Then
                strSql = "Alter Table " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
                        
            Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
        Next
        
        If strRemovePARTable & strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
    '2.分区表
   If strRemovePARTable <> "" Then
        arrTmp = Split(Mid(strRemovePARTable, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strPartition = Split(arrTmp(r), "||")(1)
            strTbsOriginal = Split(arrTmp(r), "||")(2)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
                Call ResizeTBS(strTbsOriginal)
            End If
            
            lblOptPrompt.Caption = "正在将[" & strSegment & "(" & strPartition & ")]移回原表空间"
            lblOptPrompt.Refresh
            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " TableSpace " & strTbsOriginal & " Nologging" & strParallel
            gcnOracle.Execute strSql
            
            
            If strParallel <> "" Then
                strSql = "Alter Table " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
                
            End If
                 
            '移回最后一个分区后重建表上失效的索引
            If r = UBound(arrTmp) Then
                Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
            End If
        Next
        
        If strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
    '3.LOB段
    If strRemoveLob <> "" Then
        arrTmp = Split(Mid(strRemoveLob, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strColumn = Split(arrTmp(r), "||")(1)
            strTbsLob = Split(arrTmp(r), "||")(2)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "正在收缩" & strTbsLob & "的空间"
                Call ResizeTBS(strTbsLob)
            End If
                        
            lblOptPrompt.Caption = "正在将[" & strSegment & "(" & strColumn & ")]移回原表空间"
            lblOptPrompt.Refresh
            strSql = "ALTER TABLE " & strSegment & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsLob & ") Nologging" & strParallel
            gcnOracle.Execute strSql
            
            
            'LOB并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
        Next
        
        If strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "正在收缩" & strTbsLob & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsLob)
        End If
    End If
    
    '4.索引
    If strRemoveIndex <> "" Then
        arrTmp = Split(Mid(strRemoveIndex, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strTbsOriginal = Split(arrTmp(r), "||")(1)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
                Call ResizeTBS(strTbsOriginal)
            End If
                        
            lblOptPrompt.Caption = "正在将[" & strSegment & "]移回原表空间"
            lblOptPrompt.Refresh
            strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " TableSpace " & strTbsOriginal & " Nologging" & strParallel
            gcnOracle.Execute strSql
            
                               
            If strParallel <> "" Then
                strSql = "Alter Index " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
        Next
        
        If strRemovePARIndex & strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
    '5.分区索引
    If strRemovePARIndex <> "" Then
        arrTmp = Split(Mid(strRemovePARIndex, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strPartition = Split(arrTmp(r), "||")(1)
            strTbsOriginal = Split(arrTmp(r), "||")(2)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
                Call ResizeTBS(strTbsOriginal)
            End If
                        
            lblOptPrompt.Caption = "正在将[" & strSegment & "(" & strPartition & ")]移回原表空间"
            lblOptPrompt.Refresh
            strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " TableSpace " & strTbsOriginal & " Nologging" & strParallel & " " & strOnline
            gcnOracle.Execute strSql
            
                               
            If strParallel <> "" Then
                strSql = "Alter Index " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
                
            End If
        Next
        
        If strRemovePARLOB = "" Then
            lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTBSTemp)
            
            lblOptPrompt.Caption = "正在收缩" & strTbsOriginal & "的空间"
            lblOptPrompt.Refresh
            Call ResizeTBS(strTbsOriginal)
        End If
    End If
    
     '6.LOB分区
    If strRemovePARLOB <> "" Then
        arrTmp = Split(Mid(strRemovePARLOB, 2), ",")
        For r = 0 To UBound(arrTmp)
            strSegment = Split(arrTmp(r), "||")(0)
            strPartition = Split(arrTmp(r), "||")(1)
            strColumn = Split(arrTmp(r), "||")(2)
            strTbsLob = Split(arrTmp(r), "||")(3)
            
            DoEvents
            If r = 0 Then
                lblOptPrompt.Caption = "正在收缩" & strTbsLob & "的空间"
                Call ResizeTBS(strTbsLob)
            End If
                        
            lblOptPrompt.Caption = "正在将[" & strSegment & "(" & strPartition & ")]移回原表空间"
            lblOptPrompt.Refresh
            strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsLob & ") Nologging" & strParallel
            gcnOracle.Execute strSql
            
            '移回最后一个分区后重建表上失效的索引
            If r = UBound(arrTmp) Then
                Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
            End If
            
            'LOB分区并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
        Next
                    
        lblOptPrompt.Caption = "正在收缩" & strTBSTemp & "的空间"
        lblOptPrompt.Refresh
        Call ResizeTBS(strTBSTemp)
        
        lblOptPrompt.Caption = "正在收缩" & strTbsLob & "的空间"
        lblOptPrompt.Refresh
        Call ResizeTBS(strTbsLob)
    End If
    
    
    '刷新数据
    Call RefreshData
    
    If strSegment <> "" Then
        mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "'"
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        End If
    End If
    
    strTime = GetTimeString(datBegin, GetCurrentdate)
    
    If strPrompt <> "" Then
        strPrompt = Mid(strPrompt, 2, 1500)
        MsgBox "重整执行完成，本次共耗时：" & strTime & "。" & vbCrLf & _
            "未能支持以下对象的重整：" & vbCrLf & strPrompt, vbInformation, gstrSysName
    Else
        MsgBox "重整执行完成，本次共耗时：" & strTime & "。", vbInformation, gstrSysName
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    Call ErrCenter(strSql)

    If 0 = 1 Then
        Resume
    End If
    If blnRemove = False Then GoTo reMove
    
    If txtParallel.Text <> "0" Then
        Call SetNOParallel(gcnOracle, 0)
        Call SetNOParallel(gcnOracle, 1)
    End If
    
    Call SetCommandEnable(1)
End Sub

Private Function CheckUnSuportTable(ByVal strOwner As String, ByVal strTable As String)
'功能：检查表是否存在('mdsys用户下存在类似GridFile1044_TAB这种含有小写字母的表段，但在dba_tables中却查不到记录)
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    'iot_name不为空的，是IOT索引的溢出表
    'mdsys有一张表SDO_3DTXFMS_TABLE，存在SDO_NUMBER_ARRAY数据类型，导致不能Move
    'Data_Type_Owner为Public的是XMLTYPE
    strSql = "Select 1" & vbNewLine & _
            "From Dba_Tables A" & vbNewLine & _
            "Where Owner = [1] And Table_Name = [2] And Iot_Name Is Null And Not Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From Dba_Tab_Cols B" & vbNewLine & _
            "       Where a.Owner = b.Owner And a.Table_Name = b.Table_Name And Nvl(b.Data_Type_Owner,'PUBLIC') <> 'PUBLIC' And b.Data_Type<> 'XMLTYPE')"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strOwner, strTable)

    CheckUnSuportTable = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    Call ErrCenter(strSql)
End Function

Private Sub cmdShrink_Click()
'功能：执行表或索引的空闲空间收回(Shrink Space)
    Dim strSql As String, strType As String, strSegment As String, strSegmentPre As String, strSegmentAll As String
    Dim rsTmp As ADODB.Recordset
    Dim blnRow_Movement As Boolean
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, r As Long, c As Long
    Dim strSegment_Type As String, strObjName As String
    
    If CheckExtent(P1回收) = False Then Exit Sub
        
    On Error GoTo errH
    strType = Trim(lblPrompt.Tag)
    If strType <> "" And InStr(",TABLE,INDEX,LOBSEGMENT,", "," & strType & ",") = 0 Then
        Call MsgBox("仅支持对表或索引进行空闲空间收回，不支持的数据类型：" & strType, vbInformation, Me.Caption)
        Exit Sub
    End If
    
    Call SetCommandEnable(0)
    vsfExtents.GetSelection r1, c1, r2, c2
    For r = r2 To r1 Step -1
        For c = c2 To c1 Step -1
            strSegment = mcolCells("_" & r & "_" & c)     '含所有者
            strSegment_Type = CStr(vsfExtents.Cell(flexcpData, r, c))
            
            If strSegment & "|" & strSegment_Type <> strSegmentPre Then
                If InStr(strSegmentAll & ",", "," & strSegment & "|" & strSegment_Type & ",") = 0 Then
                    mrsExtents.Filter = "Row=" & r & " And Col=" & c
                    If mrsExtents.RecordCount > 0 Then
                        DoEvents
            
                        strType = mrsExtents!Segment_Type
                        If strType = "TABLE" Then
                            If CheckUnSuportObject(strSegment, "收回(Shrink Space)") Then
                                strSql = "Select Row_Movement From All_Tables Where Table_Name = [1] And Owner = [2]"
                                Set rsTmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(1), Split(strSegment, ".")(0))
                                If rsTmp.RecordCount = 0 Then
                                    Call MsgBox("从视图All_Tables中未找到指定的对象" & strSegment, vbInformation, Me.Caption)
                                    Call SetCommandEnable(1)
                                    Exit Sub
                                End If
                                If rsTmp!Row_Movement = "DISABLED" Then 'enable row movement语句会造成引用表XXX的对象(如存储过程、包、视图等)变为无效
                                    strSql = "Alter Table " & strSegment & " Enable Row Movement"
                                    gcnOracle.Execute strSql
                                    blnRow_Movement = True
                                End If
                                       
                                lblOptPrompt.Caption = "正在对[" & strSegment & "]进行空间收回"
                                                                
                                strSql = "Alter Table " & strSegment & " Shrink Space"
                                gcnOracle.Execute strSql
                                
                                If blnRow_Movement Then
                                    strSql = "Alter Table " & strSegment & " Disable Row Movement"
                                    gcnOracle.Execute strSql
                                End If
                            End If
                            
                        ElseIf strType = "LOBSEGMENT" Then
                            mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"
                            If mrsLobs.RecordCount > 0 Then
                                lblOptPrompt.Caption = "正在对[" & mrsLobs!Table_name & "." & mrsLobs!Column_Name & "]进行空间回收"
                                                                    
                                strSql = "ALTER TABLE " & mrsLobs!Owner & "." & mrsLobs!Table_name & " MODIFY LOB (" & mrsLobs!Column_Name & ") (SHRINK SPACE)"
                                gcnOracle.Execute strSql
                            Else
                                lblOptPrompt.Caption = "在视图Dba_Lobs中未找到LOB对象" & strSegment & "。"
                            End If
                            
                        ElseIf strType = "INDEX" Then
                            If Not CheckIOT(strSegment) Then
                                lblOptPrompt.Caption = "正在对[" & strSegment & "]进行空间收回"
                                                            
                                strSql = "Alter Index " & strSegment & " Shrink Space"
                                gcnOracle.Execute strSql
                            Else
                                strObjName = GetIOTName(strSegment)
                                strSql = "Alter Table " & strObjName & " Shrink Space"
                                gcnOracle.Execute strSql
                            End If
                        ElseIf strType <> " " Then
                            lblOptPrompt.Caption = strSegment & ",不支持的对象类型：" & strType
                        End If
                    End If
                    strSegmentAll = strSegmentAll & "," & strSegment & "|" & strSegment_Type
                End If
                strSegmentPre = strSegment & "|" & strSegment_Type
            End If
        Next
        lblOptPrompt.Caption = "已处理完第" & r & "行"
        lblOptPrompt.Refresh
    Next
    
    '未改变数据文件大小，不用刷新表空间及数据文件列表
    Call LoadExtents(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")), Val(cboFiles.ItemData(cboFiles.ListIndex)))
    
    If strSegment <> "" Then
        mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "'"
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        End If
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    If MsgBox(Err.Description & vbCrLf & "最近一次执行的SQL：" & vbCrLf & strSql & vbCrLf & "可能是因为当前在线业务的影响，是否重试？", vbRetryCancel, "错误") = vbRetry Then
        Resume
    End If
    
    If blnRow_Movement Then
        strSql = "Alter Table " & strSegment & " Disable Row Movement"
        gcnOracle.Execute strSql
    End If
    lblOptPrompt.Caption = ""
    Call SetCommandEnable(1)
End Sub

Private Function CheckExtent(ByVal bytOpt As opt) As Boolean
    Dim strSegment As String, strPrompt As String
    Dim r1&, c1&, r2&, c2&, r&, c&
    
    On Error Resume Next
    If vsfExtents.Row = -1 Or vsfExtents.Col = -1 Then
        MsgBox "请先选中一个单元格再执行本操作", vbInformation, Me.Caption
        Exit Function
    End If
    If mcolCells Is Nothing Then
        MsgBox "请先刷新数据并加载一个存储了数据的单元格再执行本操作", vbInformation, Me.Caption
        Exit Function
    End If
    
    With vsfExtents
        .GetSelection r1, c1, r2, c2
        If r1 = r2 And c1 = c2 Then '仅选中一个单元格时才检查
            strSegment = mcolCells("_" & .Row & "_" & .Col)
            If strSegment = "" Or strSegment = "sys.free" Or cboFiles.ListIndex = -1 Then
                MsgBox "请先选中一个存储了数据的单元格再执行本操作", vbInformation, Me.Caption
                Exit Function
            End If
            mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"
            If mrsLobs.RecordCount > 0 Then
                strSegment = strSegment & "(" & mrsLobs!Table_name & "." & mrsLobs!Column_Name & ")"
            End If
        Else
            strSegment = mcolCells("_" & .Row & "_" & .Col) & "等"
        End If
    End With
        
    If bytOpt = P1回收 Then
        strSegment = "回收(Shrink)一般用于删除大量数据后降低高水标记，以便进行文件收缩操作。操作过程不影响业务运行，建议你在业务空闲期间执行，你确定要对" & vbCrLf & vbTab & strSegment & vbCrLf & "进行回收操作吗？"
        
    ElseIf bytOpt = P2重整 Then
        strSegment = "重整(Move Or Rebuild)一般用于移动块的物理位置，操作过程会锁表，并且需要与该对象等量的空闲空间，可能中断业务运行，建议你在业务空闲期间执行，请慎重。" & vbCrLf & _
                "Move表之后，相关索引会失效，本操作将会自动重建，可能耗时较长，你确定要对" & vbCrLf & vbTab & strSegment & vbCrLf & "进行重整操作吗？"
    End If
    If MsgBox(strSegment, vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
        Exit Function
    End If
        
    CheckExtent = True
End Function


Private Sub Form_Load()
    Dim strCol As String, i As Long
    
    strCol = "行,300,1;状态;名称,1250,1;大小,500,1"
    Call InitTable(vsfTbs, strCol)
    vsfTbs.FixedCols = 1
    
    strCol = ""
    For i = 0 To CONCOLS
        If strCol = "" Then
            strCol = i & ",550,1"
        Else
            strCol = strCol & ";" & i & ",280,4"
        End If
    Next

    Call InitTable(vsfExtents, strCol)
    vsfExtents.FixedCols = 1
    vsfExtents.Rows = vsfExtents.FixedRows
    vsfExtents.TextMatrix(0, 0) = "行\列"
    
    
    Call LoadTablespaces
    
    vsfTbs.Editable = flexEDNone
    vsfExtents.Editable = flexEDNone
    
    Call LoadParallel
    
    'Me.Caption = Me.Caption & "(服务器：" & gstrServer & ")"
End Sub

Private Sub LoadParallel()
'功能：读取并显示并行度
    
    On Error GoTo errH
    If gintCpuCount = 0 Then
        txtParallel.Text = "0"
        txtParallel.Locked = True
        txtParallel.Enabled = False
        lblParallel.ToolTipText = "未能读取到数据库参数cpu_count"
    Else
        txtParallel.Tag = gintCpuCount
        If gintCpuCount < 3 Then
            txtParallel.Text = "0"
            txtParallel.Enabled = False
            lblParallel.ToolTipText = "服务器Cpu数量不足3个，不能进行并行执行"
        ElseIf gintCpuCount < 13 Then
            txtParallel.Text = gintCpuCount \ 2 '一半取整
        Else
            txtParallel.Text = "12"  '即使cpu足够，但仍可能受限于磁盘性能，并行度并非越大越好
        End If
    End If

    Exit Sub
errH:
    Call ErrCenter
End Sub




Private Sub mnuAddFile_Click()
    '添加数据文件
    Dim strTblSpace As String, strQuery As String
    Dim strFileName As String, strFilePth As String
    Dim strSql As String
    
    On Error GoTo errH
    With vsfTbs
        If .Row = -1 Or .Row = 0 Then
            MsgBox "请先选择一个表空间再执行操作。"
            Exit Sub
        End If
        
        strTblSpace = .TextMatrix(.Row, .ColIndex("名称"))
    End With
    
    If strTblSpace = "" Then
        MsgBox "获取表空间名称失败，请重新操作。"
        Exit Sub
    Else
        Call SetCommandEnable(0)
        strFileName = GetDataFile(strTblSpace, strFilePth)
        strQuery = Trim(InputBox("为表空间" & strTblSpace & "添加数据文件" & vbCrLf & vbCrLf & _
                                                            "默认添加的数据文件大小为100M，如果有特殊需要，请手动执行以下指令：" & vbCrLf & vbCrLf & _
                                                            "ALTER TABLESPACE " & strTblSpace & " ADD DATAFILE " & vbCrLf & "'" & strFilePth & strFileName & "' SIZE 100M AUTOEXTEND ON" _
                                                        , "表空间添加数据文件", strFileName))
        
        If strQuery = "" Then
            Call SetCommandEnable(1)
            Exit Sub
        Else
            lblOptPrompt.Caption = "正在为表空间" & strTblSpace & "添加数据文件" & strFilePth & strFileName & "......"
            strSql = "ALTER TABLESPACE " & strTblSpace & " ADD DATAFILE '" & strFilePth & strQuery & "' SIZE 100M AUTOEXTEND ON"
            gcnOracle.Execute strSql
        End If
        lblOptPrompt.Caption = "表空间" & strTblSpace & "添加数据文件" & strFilePth & strQuery & "成功。"
        Call SetCommandEnable(1)
    End If
    
    Exit Sub
errH:
    Call SetCommandEnable(1)
    lblOptPrompt.Caption = "表空间" & strTblSpace & "添加数据文件" & strFilePth & strQuery & "失败。"
    If InStr(Err.Description, "ORA-01537") > 0 Then
        MsgBox "当前表空间已经存在名为" & strQuery & "的数据文件，请重新输入。"
        Exit Sub
    End If
    
    ErrCenter
End Sub

Private Sub mnuResizeAll_Click()
    Call ResizeAll
End Sub

Private Sub mnuResizeTemp_Click()
'收缩临时表空间
    Call ResizeTemp
End Sub

Private Sub mnuResizeUndo_Click()
'收缩Undo表空间
    Call frmResizeUndo.ShowMe(frmReused)
End Sub


Private Sub txtParallel_GotFocus()
    txtParallel.SelStart = 0
    txtParallel.SelLength = Len(txtParallel.Text)
End Sub

Private Sub txtParallel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtParallel_Validate(Cancel As Boolean)
    If Val(txtParallel.Tag) <> 0 Then
        If Val(txtParallel.Text) > Val(txtParallel.Tag) Then
            MsgBox "并行度不能超过cpu个数" & txtParallel.Tag, vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub LoadTablespaces()
    Dim rsTmp As ADODB.Recordset, strSql  As String
    Dim i As Long, lngStart As Long
    
    strSql = "Select a.Status, a.Tablespace_Name, a.Block_Size, Round(Sum(b.Bytes) / 1024 / 1024, 2) Tsize , Max(Decode(b.autoextensible,'YES',0,1)) as autoextensible" & vbNewLine & _
            "From Dba_Tablespaces A, Dba_Data_Files B" & vbNewLine & _
            "Where a.Contents = 'PERMANENT' And a.Tablespace_Name = b.Tablespace_Name And b.Online_status in('ONLINE','SYSTEM')" & vbNewLine & _
            "Group By a.Tablespace_Name, a.Status, a.Block_Size" & vbNewLine & _
            "Order By 4 Desc"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption)
    
    With vsfTbs
        .Redraw = flexRDNone
        lngStart = .FixedRows
        .Rows = lngStart
        .Rows = lngStart + rsTmp.RecordCount
        For i = lngStart To rsTmp.RecordCount
            If rsTmp!autoextensible = 1 Then
                .Cell(flexcpBackColor, i, .ColIndex("行"), i, .ColIndex("大小")) = OFF_颜色
                .Cell(flexcpData, i, .ColIndex("大小")) = "NO"
            End If
            .TextMatrix(i, .ColIndex("行")) = i
            .TextMatrix(i, .ColIndex("状态")) = rsTmp!Status
            .TextMatrix(i, .ColIndex("名称")) = rsTmp!Tablespace_Name
            
            If Val("" & rsTmp!Tsize) > 1024 Then
                .TextMatrix(i, .ColIndex("大小")) = Round(rsTmp!Tsize / 1024, 2) & "G"
            Else
                .TextMatrix(i, .ColIndex("大小")) = rsTmp!Tsize & "M"
            End If
            
            .RowData(i) = Val(rsTmp!Block_Size)
            rsTmp.MoveNext
        Next
        .Redraw = flexRDDirect
    End With

    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    cmdGO.Left = Abs(Me.ScaleWidth - cmdGO.Width - 60)
    lblPrompt.Width = Abs(Me.ScaleWidth - cmdGO.Width)
    
    vsfExtents.Width = Abs(Me.ScaleWidth - vsfExtents.Left - 60)
    vsfExtents.ColWidth(-1) = (vsfExtents.Width - 550 - 120) / 51   '120为滚动条的宽度
    vsfExtents.ColWidth(vsfExtents.FixedRows - 1) = 550
    vsfExtents.RowHeight(-1) = vsfExtents.Width / 51
    
    vsfTbs.Height = Abs(Me.ScaleHeight - vsfTbs.Top - 60 - picBottom.Height)
    vsfExtents.Height = Abs(vsfTbs.Height - lblPrompt.Height)
    
    cmdMore.Left = Abs(Me.ScaleWidth - cmdMore.Width - 60)
    cmdResize.Left = Abs(cmdMore.Left - cmdResize.Width - 60)
    cmdShrink.Left = Abs(cmdResize.Left - cmdShrink.Width - 60)
    cmdMove.Left = Abs(cmdShrink.Left - cmdMove.Width - 60)
    
    chkOnline.Left = Abs(cmdMove.Left - chkOnline.Width - 60)
    txtParallel.Left = Abs(chkOnline.Left - txtParallel.Width - 60)
    lblParallel.Left = Abs(txtParallel.Left - lblParallel.Width)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsExtents = Nothing
    Set mcolCells = Nothing
    Set mrsLobs = Nothing
End Sub

Private Sub LoadLobs(ByVal strTBS As String)
'功能：读取当前表空间的Lob段信息
    Dim strSql As String
 
    strSql = "Select Table_Name, TableSpace_Name, Column_Name, Owner, Segment_Name, Index_Name From Dba_Lobs Where Tablespace_Name = [1]"
    On Error GoTo errH
    Set mrsLobs = OpenSQLRecord(strSql, Me.Caption, strTBS)

    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Public Sub LoadExtents(ByVal strTBS As String, ByVal lngFile As Long)
'功能：加载Extents到单元格
    Dim rsTmp As ADODB.Recordset, strSql  As String, strSegment As String, strPreSegment As String, strFullSegment As String
    Dim i As Long, j As Long, n As Long, lngStart As Long, lngRows As Long
    Dim lngCells As Long, lngFixedCols As Long
    Dim blnFree As Boolean, blnSameCell As Boolean, strFirst As String
    
    lblOptPrompt.Caption = "正在读取数据块信息......"
    lblOptPrompt.Refresh
    
    If chkFree.Value = 1 Then
        strSql = "Select File_Id,Block_Id as Extent_ID, Block_Id as First_Block, Block_Id + Blocks - 1 as Last_Block,Blocks, 'free' as Segment_Name, 'sys.free' as Full_Segment_Name, ' ' as Segment_Type,' ' as Owner" & vbNewLine & _
            "From Dba_Free_Space A" & vbNewLine & _
            "Where Tablespace_Name = [1] And a.File_Id = [2]" & vbNewLine & _
            "Order By First_Block"
    Else
        strSql = "Select a.File_Id,a.Extent_ID, a.Block_Id First_Block, a.Block_Id + a.Blocks - 1 Last_Block,a.Blocks, a.Segment_Name, a.Owner || '.' || a.Segment_Name as Full_Segment_Name, b.Segment_Type, a.Owner" & vbNewLine & _
            "From Dba_Extents A, Dba_Segments B" & vbNewLine & _
            "Where a.Tablespace_Name = [1] And a.File_Id = [2] And a.Segment_Name = b.Segment_Name And a.Owner = b.Owner" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select File_Id,0, Block_Id, Block_Id + Blocks - 1,Blocks, 'free', 'sys.free' as Full_Segment_Name, ' ',' '" & vbNewLine & _
            "From Dba_Free_Space A" & vbNewLine & _
            "Where Tablespace_Name = [1] And a.File_Id = [2]" & vbNewLine & _
            "Order By First_Block"
    End If
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS, lngFile)
            
    Call InitmrsExtents
    
    If rsTmp.RecordCount = 0 Then
        lblOptPrompt.Caption = ""
        lblOptPrompt.Refresh
        vsfExtents.Rows = vsfExtents.FixedRows
        Exit Sub
    End If
    
    lblOptPrompt.Caption = "正在加载数据块信息......"
    lblOptPrompt.Refresh
    lngFixedCols = vsfExtents.FixedCols
    lngStart = vsfExtents.FixedRows
    
    
    '先计算出行数,用于显示进度
    j = lngFixedCols
    lngRows = lngStart
    Do While Not rsTmp.EOF
        lngCells = rsTmp!blocks \ CONBLOCKS '取整
        If rsTmp!blocks <> lngCells * CONBLOCKS Then lngCells = lngCells + 1
        
        For n = 1 To lngCells
            j = j + 1
            If j > CONCOLS Then '换行
                lngRows = lngRows + 1
                j = lngFixedCols
            End If
        Next
        rsTmp.MoveNext
    Loop
    If rsTmp.RecordCount <> 0 Then
        rsTmp.MoveFirst
    End If
    
    vsfExtents.Redraw = flexRDNone  '避免触发事件vsfExtents_AfterRowColChange
    vsfExtents.Rows = lngStart
    
    vsfExtents.Redraw = flexRDDirect
    vsfExtents.ToolTipText = ""
    vsfExtents.Refresh
    lblPrompt.Caption = ""
    vsfExtents.Redraw = flexRDNone
    vsfExtents.Rows = lngStart + lngRows

    vsfExtents.Redraw = flexRDDirect
    
        
    With vsfExtents
        .Redraw = flexRDNone
                
        i = lngStart
        j = .FixedCols
        If i > 0 Then .TextMatrix(1, 0) = 1
        
        Do While Not rsTmp.EOF
            strSegment = rsTmp!Segment_Name
            blnFree = (strSegment = "free")
            strFullSegment = rsTmp!Full_Segment_Name
                                    
            strFirst = Mid$(strSegment, 1, 1)
            If strPreSegment <> strSegment & "|" & rsTmp!Segment_Type Then
                blnSameCell = Mid$(strPreSegment, 1, 1) = strFirst
            Else
                blnSameCell = False
            End If
            
            lngCells = rsTmp!blocks \ CONBLOCKS '取整
            If rsTmp!blocks <> lngCells * CONBLOCKS Then lngCells = lngCells + 1
           
            For n = 1 To lngCells
                If blnFree Then
                    .Cell(flexcpBackColor, i, j) = &HCCEDC7 '空闲空间
                    If n = 1 Then .TextMatrix(i, j) = "B"
                Else
                    .TextMatrix(i, j) = strFirst
                    .Cell(flexcpData, i, j) = CStr(rsTmp!Segment_Type)
                End If
                mcolCells.Add strFullSegment, "_" & i & "_" & j
               
                '第一个字相同，但对象不同，用加粗来区别
                If blnSameCell Then .Cell(flexcpFontItalic, i, j) = True
                
                mrsExtents.AddNew Array("Row", "Col", "Segment_Name", "Extent_ID", "First_Block", "Blocks", "Last_Block", "Segment_Type", "Owner"), _
                            Array(i, j, strSegment, rsTmp!Extent_ID, rsTmp!First_Block, rsTmp!blocks, rsTmp!Last_Block, rsTmp!Segment_Type, rsTmp!Owner)
                               
                j = j + 1
                If j > CONCOLS Then '换行
                    j = lngFixedCols
                    
                    i = i + 1
                   .TextMatrix(i, 0) = i   '行号
                   
                   If i Mod 100 = 0 Then
                     DoEvents
                     lblOptPrompt.Caption = "正在加载信息(" & i & "/" & lngRows & ")"
                   End If
                End If
           Next
           strPreSegment = strSegment & "|" & rsTmp!Segment_Type
           rsTmp.MoveNext
        Loop
        
        '剩余的空单元格加上空值以避免从集合取值时出错
        For n = j To CONCOLS
            mcolCells.Add "", "_" & i & "_" & n
        Next

        .Redraw = flexRDDirect
    End With
    lblOptPrompt.Caption = ""
    Exit Sub
errH:
    Call ErrCenter(strSql)
    If 0 = 1 Then
        Resume
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
    End If
End Sub

Private Sub cmdFind_Click()
    If Not mrsExtents Is Nothing Then
        If InStr(txtFind.Text, "*") > 0 Then
            mrsExtents.Filter = "Segment_Name Like '" & UCase(Trim(txtFind.Text)) & "'"
        Else
            mrsExtents.Filter = "Segment_Name='" & UCase(Trim(txtFind.Text)) & "'"
        End If
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        Else
            lblOptPrompt.Caption = "没有找到匹配的表或索引。"
            txtFind.SetFocus
            txtFind_GotFocus
        End If
    Else
        lblOptPrompt.Caption = "没有找到匹配的表或索引。"
        txtFind.SetFocus
        txtFind_GotFocus
    End If
End Sub


Private Sub cmdGO_Click()
'功能：根据LOB索引或分区索引 定位到LOB对象
    Dim strObjName As String, strSegment As String, strSegment_Type As String
    Dim i As Long, j As Long
    
    If vsfExtents Is Nothing Then
        Exit Sub
    End If
    With vsfExtents
        If .Row < 0 Or .Col < 0 Then Exit Sub
        
        strSegment_Type = .Cell(flexcpData, .Row, .Col)
        strSegment = mcolCells("_" & .Row & "_" & .Col)
        
        If strSegment = "" Then Exit Sub
        
        If strSegment_Type = "LOBINDEX" Or strSegment_Type = "INDEX PARTITION" Then
            strObjName = GetLOBNameByIndex(strSegment)
        End If
        
        If strObjName <> "" Then
            For i = .FixedRows To .Rows - 1
                For j = .FixedCols To .Cols - 1
                    If strObjName = mcolCells("_" & i & "_" & j) Then
                        .Select i, j
                        .TopRow = i
                        .SetFocus
                        strObjName = ""
                        Exit Sub
                    End If
                Next
            Next
            If strObjName <> "" Then Call MsgBox("未找到" & strObjName, vbInformation)
        End If
    End With
End Sub

Private Sub vsfExtents_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfExtents
        Dim strSegment As String, i As Long, lngBlockSize As Long, strSegment_Type As String
        
        If Me.Visible = False Or .Redraw = flexRDNone Or mcolCells Is Nothing Or vsfTbs.Enabled = False Then Exit Sub
        
        .Redraw = flexRDNone
        '先去掉之前选中的段的背景色
        If OldRow > 0 And OldCol > 0 Then
            strSegment = mcolCells("_" & OldRow & "_" & OldCol)
            If strSegment <> "" Then
                strSegment_Type = .Cell(flexcpData, OldRow, OldCol)
                mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "' And Segment_Type='" & strSegment_Type & "'"
                For i = 1 To mrsExtents.RecordCount
                    If mrsExtents!Segment_Name = "free" Then
                        .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &HCCEDC7 '空闲空间
                    Else
                        .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &H80000005 '白色
                    End If
                    .Cell(flexcpForeColor, mrsExtents!Row, mrsExtents!Col) = vbBlack
                    mrsExtents.MoveNext
                Next
            End If
        End If
                
        .Redraw = flexRDDirect
        
        
        '再设置当前选中段的背景色
        .Redraw = flexRDNone
        cmdGO.Visible = False
        
        strSegment = mcolCells("_" & NewRow & "_" & NewCol)
        lblPrompt.Tag = ""
        If strSegment <> "" Then
            strSegment_Type = .Cell(flexcpData, NewRow, NewCol)
            
            If strSegment_Type = "LOBINDEX" Then
                cmdGO.Visible = True
                cmdGO.Caption = "定位到LOB"
            ElseIf strSegment_Type = "INDEX PARTITION" Then
                If CheckLOBIndex(strSegment) Then
                    cmdGO.Visible = True
                    cmdGO.Caption = "定位到LOB"
                End If
            End If
            
            mrsExtents.Filter = "Row=" & NewRow & " And Col=" & NewCol
            If mrsExtents.RecordCount > 0 Then
                If mrsExtents!Segment_Type = "LOBSEGMENT" Then
                    mrsLobs.Filter = "Owner='" & mrsExtents!Owner & "' And Segment_Name='" & mrsExtents!Segment_Name & "'"
                    If mrsLobs.RecordCount > 0 Then .ToolTipText = mrsLobs!Table_name & "." & mrsLobs!Column_Name
                
                ElseIf mrsExtents!Segment_Type = "LOBINDEX" Then
                    mrsLobs.Filter = "Owner='" & mrsExtents!Owner & "' And Index_Name='" & mrsExtents!Segment_Name & "'"
                    If mrsLobs.RecordCount > 0 Then .ToolTipText = mrsLobs!Table_name & "." & mrsLobs!Column_Name & "(Index)"
                Else
                    .ToolTipText = strSegment & "(一个单元格包含" & CONBLOCKS & "个块)"
                End If
                
                lngBlockSize = Val(vsfTbs.RowData(vsfTbs.Row))
                If lngBlockSize = 0 Then lngBlockSize = 8192
                
                If strSegment = "sys.free" Then
                    lblPrompt.Caption = "已格式化的空闲空间，" & mrsExtents!blocks & "块：从" & Round(mrsExtents!First_Block * 8192 / 1024 / 1024, 2) & _
                                        "M到" & Round(mrsExtents!Last_Block * lngBlockSize / 1024 / 1024, 2) & "M"
                Else
                    lblPrompt.Caption = mrsExtents!Segment_Type & "：" & strSegment & "，Extent_ID：" & mrsExtents!Extent_ID & "(" & mrsExtents!blocks & "块，从" & _
                                        Round(mrsExtents!First_Block * lngBlockSize / 1024 / 1024, 2) & "M到" & Round(mrsExtents!Last_Block * lngBlockSize / 1024 / 1024, 2) & "M)"
                    lblPrompt.Tag = mrsExtents!Segment_Type
                End If
            Else
                lblPrompt.Caption = "未选中数据块。"
            End If
            
            mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "' And Segment_Type='" & strSegment_Type & "'"
            For i = 1 To mrsExtents.RecordCount
                .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &H8000000D     '蓝色
                .Cell(flexcpForeColor, mrsExtents!Row, mrsExtents!Col) = &H80000005
                mrsExtents.MoveNext
            Next
        Else
            lblPrompt.Caption = "未选中数据块。"
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfExtents_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    If Me.Visible = False Or vsfExtents.Redraw = flexRDNone Or vsfTbs.Enabled = False Then Exit Sub
    
    lngRow = vsfExtents.MouseRow
    lngCol = vsfExtents.MouseCol
    If lngRow > 0 And lngCol > 0 And Not mcolCells Is Nothing Then
       If (lngRow <> mlngRowPre Or lngCol <> mlngColPre) And lngRow <> vsfExtents.Row And lngCol <> vsfExtents.Col Then
           vsfExtents.ToolTipText = mcolCells("_" & lngRow & "_" & lngCol) & "(一个单元格包含" & CONBLOCKS & "个块)"
           mlngRowPre = lngRow
           mlngColPre = lngCol
       End If
    Else
        vsfExtents.ToolTipText = ""
    End If
End Sub

Private Sub vsfTbs_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If Me.Visible And NewRowSel <> OldRowSel And vsfTbs.Redraw <> flexRDNone Then
        vsfTbs.Refresh
        Call LoadFiles(vsfTbs.TextMatrix(NewRowSel, vsfTbs.ColIndex("名称")))

        If cboFiles.ListCount < 2 Then
            cboFiles.ListIndex = 0
        Else
            vsfExtents.Redraw = flexRDNone '避免触发事件vsfExtents_AfterRowColChange
            vsfExtents.Rows = vsfExtents.FixedRows
            vsfExtents.Redraw = flexRDDirect
            vsfExtents.ToolTipText = ""
            vsfExtents.Refresh
        End If
        
        Call LoadLobs(vsfTbs.TextMatrix(NewRowSel, vsfTbs.ColIndex("名称")))
        
        If vsfTbs.Cell(flexcpData, NewRowSel, vsfTbs.ColIndex("大小")) = "NO" Then
            lblOptPrompt.Caption = "所选表空间中存在自增长属性为NO的数据文件。"
        End If
    End If
End Sub

Private Sub LoadFiles(strTBS As String)
    Dim rsTmp As ADODB.Recordset, strSql  As String
    Dim i As Long, lngStart As Long
    
    strSql = "Select a.File_Name, a.File_Id, Round(a.Bytes / 1024 / 1024) As Fsize, Round(Nvl(Sum(b.Bytes),0) / 1024 / 1024) As Free_Size , a.autoextensible " & vbNewLine & _
            "From Dba_Data_Files A, Dba_Free_Space B" & vbNewLine & _
            "Where a.Tablespace_Name = [1] And a.File_Id = b.File_Id(+) And a.Tablespace_Name = b.Tablespace_Name(+) And a.Online_status in('ONLINE','SYSTEM')" & vbNewLine & _
            "Group By a.File_Name, a.File_Id, a.Bytes,a.autoextensible" & vbNewLine & _
            "Order By a.File_Id"

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTBS)
    
    cboFiles.Clear
    cboFiles.Tag = ""
    For i = 1 To rsTmp.RecordCount
        cboFiles.AddItem rsTmp!FILE_NAME & "(占用" & rsTmp!fsize & "M,空闲" & rsTmp!Free_Size & "M" & IIf(rsTmp!autoextensible & "" <> "YES", ",不自动扩展", "") & ")"
        cboFiles.ItemData(cboFiles.NewIndex) = Val(rsTmp!File_Id)
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub


Private Sub InitmrsExtents()
    
    Set mcolCells = New Collection
    
    Set mrsExtents = New ADODB.Recordset
    mrsExtents.Fields.Append "Row", adBigInt
    mrsExtents.Fields.Append "Col", adBigInt
    mrsExtents.Fields.Append "Owner", adVarChar, 20
    mrsExtents.Fields.Append "Segment_Name", adVarChar, 100
    mrsExtents.Fields.Append "Segment_Type", adVarChar, 20
    
    mrsExtents.Fields.Append "Extent_ID", adBigInt
    mrsExtents.Fields.Append "Blocks", adBigInt
    mrsExtents.Fields.Append "First_Block", adBigInt
    mrsExtents.Fields.Append "Last_Block", adBigInt
    
    mrsExtents.CursorLocation = adUseClient
    mrsExtents.LockType = adLockOptimistic
    mrsExtents.CursorType = adOpenStatic
    mrsExtents.Open
End Sub


Private Sub ResizeAll()
'功能：收缩所有数据文件
    Dim strErr As String
    Dim rsTmp As ADODB.Recordset, rsSize As ADODB.Recordset
    Dim lngBlockSize As Long, lngSumSize As Long
    
    If MsgBox("你确定要收缩所有数据文件吗？" & vbCrLf & vbCrLf & "建议你在业务空闲期间执行，请慎重！", vbYesNo + vbQuestion + vbDefaultButton2, "确认收缩") = vbNo Then
        lblOptPrompt.Caption = "操作被取消。"
        Call SetCommandEnable(1)
        Exit Sub
    End If
    
    Call SetCommandEnable(0)
    '获取Block_size大小
    gstrSQL = "select value from v$parameter where name = 'db_block_size'"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    lngBlockSize = Val("" & rsTmp!Value)
    
    '记录执行操作语句
    lblOptPrompt.Caption = "正在查询待收缩的数据文件。"
    gstrSQL = "Select File_Name,'alter database datafile ''' || Trim(File_Name) || ''' resize ' || Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024+10) || 'm' Cmd" & vbNewLine & _
            "From Dba_Data_Files A, (Select File_Id, Max(Block_Id + Blocks ) Hwm From Dba_Extents Group By File_Id) B" & vbNewLine & _
            "Where a.File_Id = b.File_Id(+) And Exists(Select 1 From Dba_Tablespaces C Where a.Tablespace_Name = c.Tablespace_Name And c.Status = 'ONLINE' And Contents != 'UNDO')" & vbNewLine & _
            "      And Ceil(Blocks * " & lngBlockSize & " / 1024 / 1024) - Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) > 10"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount = 0 Then
        Call MsgBox("没有要收缩数据文件！", vbInformation, "收缩数据文件")
        lblOptPrompt.Caption = ""
        Call SetCommandEnable(1)
        Exit Sub
    Else
    
        If MsgBox("共有" & rsTmp.RecordCount & "个待收缩的数据文件，你确定要收缩这些数据文件吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    lblOptPrompt.Caption = "开始进行收缩操作。"
    
    '执行操作
    '1.记录收缩前的大小
    gstrSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files"
    Set rsSize = OpenSQLRecord(gstrSQL, Me.Caption)
    lngSumSize = rsSize!Mb_Size
    On Error Resume Next
    strErr = ""
    While Not rsTmp.EOF
        lblOptPrompt.Caption = "正在收缩：" & rsTmp!FILE_NAME
        lblOptPrompt.Refresh
        gstrSQL = rsTmp!cmd
        gcnOracle.Execute gstrSQL
        
        If Err.Number <> 0 Then
            strErr = strErr & vbCrLf & rsTmp!cmd & "，错误：" & Err.Description
            Err.Clear
        End If
        
        rsTmp.MoveNext
    Wend
    
    '2.记录收缩后的总大小
    gstrSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files"
    Set rsSize = OpenSQLRecord(gstrSQL, Me.Caption)
    lngSumSize = lngSumSize - rsSize!Mb_Size

    lblOptPrompt.Caption = ""
        
    If strErr <> "" Then
        MsgBox "错误信息：" & strErr, vbExclamation
    Else
        lblOptPrompt.Caption = "操作完成，共收缩了" & lngSumSize & "M的空间。"
    End If
    
    Call RefreshData
    
    Call SetCommandEnable(1)
End Sub

Private Sub ResizeTemp()
    Dim strError As String, strVersion As String, strTbsInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim strSize As String, lngMax As Long
    
    strVersion = getVersion
    If strVersion = "" Then
        Exit Sub
    End If
    
    Call SetCommandEnable(0)

    On Error GoTo errH
    gstrSQL = "Select Tablespace_Name, File_Name, Trunc(Bytes / 1024 / 1024) Siz" & vbNewLine & _
            "From Dba_Temp_Files" & vbNewLine & _
            "Where Bytes / 1024 / 1024 > 10" & vbNewLine & _
            "Order By Tablespace_Name, File_Name"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rsTmp.RecordCount <> 0 Then
        While Not rsTmp.EOF
            strTbsInfo = strTbsInfo & rsTmp!FILE_NAME & "," & rsTmp!Siz & "M" & vbCrLf
            If rsTmp!Siz > lngMax Then lngMax = rsTmp!Siz
            rsTmp.MoveNext
        Wend
        strTbsInfo = "当前临时表空间：" & vbCrLf & vbCrLf & strTbsInfo
        
        '获取重置后的大小
input_line:
        strSize = Trim(InputBox(strTbsInfo & vbCrLf & vbCrLf & "请输入收缩后的数据文件大小(单位M)，小于等于指定值的不收缩，建议你在业务空闲期间执行。", "收缩临时表空间"))
        If strSize = "" Then
            Call SetCommandEnable(1)
            Exit Sub
        Else
            strError = ""
            If Not IsNumeric(strSize) Then
                strError = "请重新输入数字"
            ElseIf Val(strSize) <= 0 Then
                strError = "请重新输入大于零的数字"
            ElseIf Val(strSize) >= lngMax Then
                strError = "请重新输入小于" & lngMax & "的数字。"
            ElseIf InStr(strSize, ".") > 0 Then
                strError = "请重新输入不含小数的数字"
            End If
            
            If strError <> "" Then
                MsgBox strError, vbInformation, gstrSysName
                GoTo input_line
            End If
        End If
        
        On Error Resume Next
        strError = ""
        strTbsInfo = ""
        lblOptPrompt.Caption = ""
        rsTmp.MoveFirst
        rsTmp.Filter = "Siz>" & strSize
        While Not rsTmp.EOF
            lblOptPrompt.Caption = "正在收缩临时表空间 " & rsTmp!Tablespace_Name & "。"
            lblOptPrompt.Refresh
            If strVersion = 11 Then
                '一个表空间有多个数据文件，11GR1是按表空间来收缩的
                '也可以按数据文件逐个收缩: alter tablespace temp shrink tempfile '/u01/app/oracle/oradata/anqing/temp01.dbf' keep 300M;
                If rsTmp!Tablespace_Name <> strTbsInfo Then
                    strTbsInfo = rsTmp!Tablespace_Name
                    gstrSQL = "alter tablespace " & strTbsInfo & "  shrink space keep " & Val(strSize) & "M"
                    gcnOracle.Execute gstrSQL
                End If
            Else
                gstrSQL = "alter database tempfile '" & rsTmp!FILE_NAME & "'  resize " & Val(strSize) & "M"
                gcnOracle.Execute gstrSQL
            End If
            
            If Err <> 0 Then
                strError = strError & vbCrLf & rsTmp!FILE_NAME & vbCrLf & Err.Description
                Err.Clear
            End If
            rsTmp.MoveNext
        Wend
        
        If strError <> "" Then
            MsgBox "收缩表空间出错 " & vbCrLf & strError & vbCrLf & "请重新指定保留文件的大小，或者重启系统后执行收缩。", vbInformation, gstrSysName
        Else
            lblOptPrompt.Caption = "临时表空间收缩完毕！"
        End If
    Else
        MsgBox "当前没有大于10M的临时数据文件，不需要收缩。"
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    Call SetCommandEnable(1)
End Sub


Public Sub SetNOParallel(ByVal cnThis As ADODB.Connection, ByVal bytType As Byte)
'功能：并行执行后会自动为表名索引加上并行度属性，如果不取消，会影响相关SQL的执行计划(全表扫描+并行查询，巨慢)
'参数：bytType：0=索引，1=表

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    If bytType = 0 Then
        strSql = "Select Owner || '.' || Index_Name As Index_Name From DBA_Indexes Where Degree Not In ('1', '0')"
    Else
        strSql = "Select Owner || '.' || Table_Name As Table_Name From DBA_Tables Where Degree !=('         1')"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSql, cnThis, adOpenKeyset, adLockReadOnly
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    '如果有索引组织表，会报错：ORA-25176: 不允许对主键使用存储说明
    On Error Resume Next
    
    While Not rsTmp.EOF
        If bytType = 0 Then
            strSql = "alter index " & rsTmp!Index_Name & " noparallel"
        Else
            strSql = "alter table " & rsTmp!Table_name & " noparallel"
        End If
        cmdTmp.CommandText = strSql
        
        cmdTmp.Execute
        
        rsTmp.MoveNext
    Wend
End Sub


Private Function GetDataFile(ByVal strTblSpace As String, ByRef strLocation As String) As String
    '功能：传入表空间名称，获取新增的数据文件
    '规则：默认新增文件名称为当前数据文件后+1
    '返回值： 新的数据文件名称，strLocation-被修改为数据文件所在路径
    
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strFileName As String, strFilePath As String
    Dim i As Integer, strTmp As String
    
    On Error GoTo errH
    strSql = "Select FILE_NAME From Dba_Data_Files Where tablespace_name = [1] Order By File_Name Desc"
    Set rsTmp = OpenSQLRecord(strSql, "GetDataFile", strTblSpace)
    
    If rsTmp.RecordCount = 0 Then Exit Function

    strFilePath = rsTmp!FILE_NAME
    
    '判断服务器是否为WINDOWS环境,WINDOW为 \,Linux为 /
    strFileName = Mid(strFilePath, InStrRev(strFilePath, IIf(InStr(strFilePath, "\") > 0, "\", "/")) + 1)
    strFilePath = Left(strFilePath, InStrRev(strFilePath, IIf(InStr(strFilePath, "\") > 0, "\", "/")))

    If InStr(strFileName, ".DBF") > 0 Then
        strFileName = Left(strFileName, InStrRev(strFileName, ".DBF") - 1)
        '判断文件名的尾部是否为数字，不是数字就添加字段为01，否则+1
        For i = Len(strFileName) To 1 Step -1
            If InStr("0123456789", Mid(strFileName, i, 1)) > 0 Then
                strTmp = Mid(strFileName, i, 1) & strTmp
            Else
                Exit For
            End If
        Next
        If strTmp <> "" Then
            strFileName = Left(strFileName, InStr(1, strFileName, strTmp) - 1) & Format(Val(strTmp) + 1, "00") & ".DBF"
        Else
            strFileName = strFileName & Format(Val(strTmp) + 1, "00") & ".DBF"
        End If
    End If

    strLocation = strFilePath
    GetDataFile = strFileName
    Exit Function
errH:
    GetDataFile = ""
    ErrCenter
End Function
