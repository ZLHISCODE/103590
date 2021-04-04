VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmUnitRegPlan 
   BorderStyle     =   0  'None
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTime 
      BorderStyle     =   0  'None
      Height          =   3465
      Left            =   5280
      ScaleHeight     =   3465
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   2160
      Width           =   7065
      Begin MSComctlLib.TabStrip tbWeekTime 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   60
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   503
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsThis 
         Height          =   4905
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   11970
         _cx             =   21114
         _cy             =   8652
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
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmUnitRegPlan.frx":0000
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
   Begin XtremeSuiteControls.TabControl tbUunits 
      Height          =   6855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      _Version        =   589884
      _ExtentX        =   8070
      _ExtentY        =   12091
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmUnitRegPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsUnitRegPlan As ADODB.Recordset
Private mlng计划Id As Long
Private mstr合作单位 As String
Private mbytMode    As Byte ' 1序号控制 0没有序号控制
Private mblnNotClick As Boolean
Private mrsUnits As ADODB.Recordset
Private mblnNotChange As Boolean
Private mblnUnitReg As Boolean '是否设置了合作单位预约号
Private mrsUnitReg  As ADODB.Recordset

Private Sub Form_Load()
  '  Call InitPage
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With tbUunits
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
    End With
End Sub

 

Public Sub ShowUnitReg(ByVal lng计划ID As Long)
   Dim ObjItem As TabControlItem
   mlng计划Id = lng计划ID
   If Not mrsUnitRegPlan Is Nothing Then Set mrsUnitRegPlan = Nothing
   If mlng计划Id = 0 Then
    tbUunits.RemoveAll
    '问题号:51156
    'Set ObjItem = tbUunits.InsertItem(1, "", Me.picTime.hWnd, 0)
    'ObjItem.Tag = 1
     tbWeekTime.Tabs.Clear
    'tbUunits.Item(0).Selected = True
     tbUunits.PaintManager.Position = xtpTabPositionBottom
     With tbUunits
         .PaintManager.Position = xtpTabPositionBottom
         '.PaintManager.OneNoteColors = True
         '.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        ' .PaintManager.Layout = xtpTabLayoutCompressed
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
        '.PaintManager.Position = xtpTabPositionTop
     End With
     vsThis.Clear
     vsThis.Top = tbWeekTime.Top = tbWeekTime.Height + 50
     vsThis.Left = 0
     vsThis.Visible = False
    ' vsThis.Rows = 1
     Exit Sub
     
   End If
   If InitData() = False Then Exit Sub
   InitPage
   If mrsUnits Is Nothing Then Exit Sub
   If mrsUnits.RecordCount = 0 Then Exit Sub
   vsThis.Visible = True
End Sub

'Public Sub LoadUnitReg(ByVal lng计划Id As Long, ByVal str合作单位 As String)
'    mlng计划Id = lng计划Id: mstr合作单位 = str合作单位
'    If InitData() = False Then Exit Sub
'
'   ' Call LoadUnitRegPlan(lng计划Id, IIf(mbytMode = 1, True, False), str合作单位)
'
'End Sub

Private Function InitPage() As Boolean
 
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Dim blnExit As Boolean
    Dim str合作单位 As String
    Dim j As Long
    Err = 0: On Error GoTo Errhand:
    tbUunits.RemoveAll
    If mrsUnits Is Nothing Then
        blnExit = True
    Else
        mrsUnits.Filter = 0
        If mrsUnits.RecordCount = 0 Then blnExit = True
    End If
    If blnExit Then
        '问题号:51156
         'Set ObjItem = tbUunits.InsertItem(1, "", Me.picTime.hWnd, 0)
        'ObjItem.Tag = 1
        tbWeekTime.Tabs.Clear
        'tbUunits.Item(0).Selected = True
        vsThis.Clear
        vsThis.Top = tbWeekTime.Top
        tbWeekTime.Visible = False
       ' vsThis.Rows = 1
        Exit Function
    End If
    mrsUnits.MoveFirst
    j = 1
    Do While Not mrsUnits.EOF
        Set ObjItem = tbUunits.InsertItem(j, mrsUnits!合作单位, Me.picTime.hWnd, 0)
        ObjItem.Tag = j
        j = j + 1
        mrsUnits.MoveNext
    Loop
    mblnNotChange = True
    With tbUunits
         .PaintManager.Position = xtpTabPositionBottom
         '问题号:51156
        ' .PaintManager.OneNoteColors = True
         .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        ' .PaintManager.Layout = xtpTabLayoutCompressed
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
        '.PaintManager.Position = xtpTabPositionTop
    End With
    mblnNotChange = False
    If tbUunits.ItemCount > 0 Then
      Call tbUunits_SelectedChanged(tbUunits.Item(0))
    End If
 Exit Function
 
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function InitData() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim BytMode As Byte
    strSQL = "Select 1 As 属性, Nvl(序号控制, 0) As 数据 ,'' as 单位 From 挂号安排计划 Where ID = [1] "
     
    
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)
    If rsTmp.EOF Then Exit Function
    rsTmp.Filter = "属性=1"
    If rsTmp.RecordCount = 0 Then Exit Function
    mbytMode = Val(Nvl(rsTmp!数据))
    strSQL = "    Select   合作单位  From 合作单位计划控制 Where 计划Id = [1]  Group By 合作单位  "
    Set mrsUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng计划Id)
    Set rsTmp = Nothing
    InitData = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
 

Private Sub picTime_Resize()
On Error Resume Next
 With tbWeekTime
    .Left = picTime.ScaleLeft
    .Top = picTime.ScaleTop
    .Width = picTime.ScaleWidth
 End With
 With vsThis
    .Left = picTime.ScaleLeft
    .Top = tbWeekTime.Top + tbWeekTime.Height
    .Width = tbWeekTime.Width
    .Height = picTime.ScaleHeight - .Top
 End With
End Sub

 
 

 

Private Sub tbUunits_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "" Or mblnNotChange Then Exit Sub
    mstr合作单位 = Item.Caption
    Call LoadUnitRegPlan(mlng计划Id, IIf(mbytMode = 1, True, False), mstr合作单位)
End Sub

Private Sub tbWeekTime_Click()
    If mblnNotClick = True Then Exit Sub
    Call LoadUnitRegPlan(mlng计划Id, IIf(mbytMode = 1, True, False), mstr合作单位)
End Sub


 
Private Sub LoadUnitRegPlan(ByVal lng计划ID As Long, ByVal bln序号控制 As Boolean, ByVal str合作单位 As String, _
   Optional blnReload As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:
    '入参:
    '编制:
    '日期:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str星期 As String
    Dim i As Long, r As Integer, strTime As String, strKey As String
    Static lngPre计划Id  As Long
    Static strPre合作单位 As String
    On Error GoTo errHandle
    '加载该挂号项目的的停用时间信息
    If mrsUnitRegPlan Is Nothing Then
        lngPre计划Id = -1
    ElseIf mrsUnitRegPlan.State <> 1 Then
         lngPre计划Id = -1
    End If
    If lngPre计划Id <> lng计划ID Or strPre合作单位 <> str合作单位 Or blnReload Then
        lngPre计划Id = lng计划ID
        strPre合作单位 = str合作单位
        strSQL = "" & _
        "   Select decode(限制项目,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,限制项目 As 星期,序号, " & _
        "       数量 as 限制数量 , 计划Id, 合作单位 " & _
        "   From 合作单位计划控制" & _
        "   Where  计划Id=[1] And 合作单位=[2] And   数量>0 " & _
        "   Order by 排序,序号"
        Set mrsUnitRegPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng计划ID, str合作单位)
        tbWeekTime.Tabs.Clear
    End If
    If tbWeekTime.Tabs.Count = 0 Then
        With mrsUnitRegPlan
            strTime = ""
            mrsUnitRegPlan.Filter = 0
            If mrsUnitRegPlan.RecordCount > 0 Then mrsUnitRegPlan.MoveFirst
            Do While Not .EOF
                If strTime <> Nvl(mrsUnitRegPlan!星期) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsUnitRegPlan!星期), Nvl(mrsUnitRegPlan!星期)
                    strTime = Nvl(mrsUnitRegPlan!星期)
                End If
                .MoveNext
            Loop
            mblnNotClick = True
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            mblnNotClick = False
            mrsUnitRegPlan.Filter = 0
            If mrsUnitRegPlan.RecordCount <> 0 Then mrsUnitRegPlan.MoveFirst
        End With
    End If
    str星期 = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsUnitRegPlan.Filter = "星期='" & str星期 & "'"
    
    With vsThis
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 800: .RowHeightMin = 800
        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0


        .Cols = 8: .FixedCols = 0
        r = 0: i = -1
        Do While Not mrsUnitRegPlan.EOF
           i = i + 1
           If i > .Cols - 1 Then r = r + 1: i = 0
           If Not bln序号控制 Then
                strTime = "预约" & Val(Nvl(mrsUnitRegPlan!限制数量)) & "人" & vbCrLf & vbCrLf
                If Val(Nvl(mrsUnitRegPlan!序号)) <> 0 Then
                    strTime = strTime & Val(Nvl(mrsUnitRegPlan!序号))
                End If
           Else
                If Val(Nvl(mrsUnitRegPlan!序号)) = 0 Then
                    strTime = "预约" & Val(Nvl(mrsUnitRegPlan!限制数量)) & "人" & vbCrLf & vbCrLf
                Else
                    strTime = Val(Nvl(mrsUnitRegPlan!序号))
                End If
           End If
           If r > .Rows - 1 Then .Rows = .Rows + 1
           .TextMatrix(r, i) = strTime
           mrsUnitRegPlan.MoveNext
        Loop
        For i = 0 To .Cols - 1
           .ColAlignment(i) = flexAlignCenterCenter
           .ColWidth(i) = 1200
        Next
        .Redraw = flexRDBuffered
    End With
        Exit Sub


'        Do While Not mrsUnitRegPlan.EOF
'            If str时点 <> Nvl(mrsUnitRegPlan!时点) Then
'                r = r + 1
'                str时点 = Nvl(mrsUnitRegPlan!时点)
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, 0) = str时点
'                i = 0
'            End If
'            i = i + 1
'            strTime = mrsUnitRegPlan!序号 & vbCrLf & vbCrLf
'            strTime = strTime & mrsUnitRegPlan!时间范围
'            If i > .Cols - 1 Then .Cols = .Cols + 1
'            If r > .Rows - 1 Then .Rows = .Rows + 1
'            .TextMatrix(r, i) = strTime
'            If Val(Nvl(mrsUnitRegPlan!是否预约)) = 1 Then
'                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
'                .Cell(flexcpFontBold, r, i, r, i) = True
'            End If
'            mrsUnitRegPlan.MoveNext
'        Loop
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        .Redraw = flexRDBuffered
'    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



