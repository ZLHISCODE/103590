VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDocPrintPatiList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择病人"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   ControlBox      =   0   'False
   Icon            =   "frmDocPrintPatiList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkPrinted 
      Caption         =   "未打印"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4740
      Width           =   855
   End
   Begin VB.CheckBox chkChoose 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   1
      Top             =   4680
      Width           =   1125
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5760
      TabIndex        =   0
      Top             =   4680
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      _cx             =   14631
      _cy             =   7858
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
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
   End
End
Attribute VB_Name = "frmDocPrintPatiList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String
Private mrtReportTye As ReportType

Public Function Showfrm(ByVal vsList As VSFlexGrid, frmParent As Form, ByVal blnCanPrint As Boolean, _
    ByVal rtReportTye As ReportType, ByVal lngDeptID As Long) As String
'参数：vsList病人列表，blnCanPrint 平诊报告需要审核才能打印
    chkPrinted.value = 0
    mstrReturn = ""
    mrtReportTye = rtReportTye
    
    Call InitReleationList
    Call LoadListDate(vsList, blnCanPrint, lngDeptID)
    
    Me.Show 1, frmParent
    Showfrm = mstrReturn
End Function

Private Sub InitReleationList()
'初始化关联列表
    With vsfList
        .Cols = IIf(mrtReportTye = 报告文档编辑器, 12, 11)
        .Rows = 1
        
        .TextMatrix(0, 0) = "医嘱ID"
        .TextMatrix(0, 1) = "姓名"
        .TextMatrix(0, 2) = "来源"
        .TextMatrix(0, 3) = "检查号"
        .TextMatrix(0, 4) = "性别"
        .TextMatrix(0, 5) = "年龄"
        .TextMatrix(0, 6) = "内容"
        .TextMatrix(0, 7) = "部位"
        .TextMatrix(0, 8) = "打印状态"
        .TextMatrix(0, 9) = "PACS报告"
        .TextMatrix(0, 10) = "执行科室"
        If mrtReportTye = 报告文档编辑器 Then .TextMatrix(0, 11) = "报告ID"
                    
        .FixedCols = 0
        .FixedRows = 1
        
        .GridLines = flexGridFlat
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered
        .OutlineCol = 1
        .OutlineBar = flexOutlineBarCompleteLeaf
        .Ellipsis = flexEllipsisEnd
        
        .AllowSelection = False
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
        
        .ColDataType(0) = flexDTBoolean
        .ColHidden(0) = True
        If mrtReportTye = 报告文档编辑器 Then .ColHidden(11) = True
    End With
End Sub

Private Sub LoadListDate(ByVal vsList As VSFlexGrid, ByVal blnCanPrint As Boolean, ByVal lngDeptID As Long)
    Dim i As Integer
    Dim iCount As Integer
    Dim lngOldDeptID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If mrtReportTye = 报告文档编辑器 Then
        vsfList.Redraw = flexRDNone
        
        For i = 1 To vsList.Rows - 1
            With vsList
                strSQL = "Select RawToHex(ID) As ID, 报告打印,报告状态 From 影像报告记录 Where 医嘱ID = [1] And 报告状态 In (2,3,4)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "", .TextMatrix(i, GetColNum(vsList, "医嘱ID")))
                
                '如果有“执行科室ID”，则需要重新根据科室ID读取平诊报告需审核的参数
                If GetColNum(vsList, "执行科室ID") <> 0 Then
                    If lngOldDeptID <> .TextMatrix(i, GetColNum(vsList, "执行科室ID")) Then   '科室ID改变了，重新读取平诊报告打印的参数
                        lngOldDeptID = .TextMatrix(i, GetColNum(vsList, "执行科室ID"))
                        blnCanPrint = GetDeptPara(lngOldDeptID, "平诊需审核才能打报告") = "1"           '平诊需要审核才能打印 =true
                    End If
                Else
                    lngOldDeptID = lngDeptID
                End If
                    
                If rsTmp.RecordCount > 0 Then
                    vsfList.AddItem ""
                    vsfList.RowData(vsfList.Rows - 1) = 0
                    
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 0) = .TextMatrix(i, GetColNum(vsList, "医嘱ID"))
                    vsfList.Cell(flexcpChecked, vsfList.Rows - 1, 1) = 2
                    
                    vsfList.TextMatrix(vsfList.Rows - 1, 1) = .TextMatrix(i, GetColNum(vsList, "姓名"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 2) = .TextMatrix(i, GetColNum(vsList, "来源"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 3) = .TextMatrix(i, GetColNum(vsList, "检查号"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 4) = .TextMatrix(i, GetColNum(vsList, "性别"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 5) = .TextMatrix(i, GetColNum(vsList, "年龄"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 6) = .TextMatrix(i, GetColNum(vsList, "医嘱内容"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 7) = .TextMatrix(i, GetColNum(vsList, "部位方法"))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 8) = IIf(Nvl(rsTmp!报告打印) = 0, "", Nvl(rsTmp!报告打印))
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 9) = 2
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 10) = lngOldDeptID
                    vsfList.Cell(flexcpText, vsfList.Rows - 1, 11) = Nvl(rsTmp!ID)
                    
                    vsfList.IsSubtotal(vsfList.Rows - 1) = True
                    vsfList.RowOutlineLevel(vsfList.Rows - 1) = 1
                    vsfList.RowData(vsfList.Rows - 1) = 1
                    
                    If rsTmp.RecordCount > 0 Then
                        While Not rsTmp.EOF
                            iCount = iCount + 1
                            vsfList.AddItem ""
                            vsfList.RowData(vsfList.Rows - 1) = 1
        
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 0) = .TextMatrix(i, GetColNum(vsList, "医嘱ID"))
                            vsfList.Cell(flexcpChecked, vsfList.Rows - 1, 1) = 2
                            
                            vsfList.TextMatrix(vsfList.Rows - 1, 1) = .TextMatrix(i, GetColNum(vsList, "姓名"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 2) = .TextMatrix(i, GetColNum(vsList, "来源"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 3) = .TextMatrix(i, GetColNum(vsList, "检查号"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 4) = .TextMatrix(i, GetColNum(vsList, "性别"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 5) = .TextMatrix(i, GetColNum(vsList, "年龄"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 6) = .TextMatrix(i, GetColNum(vsList, "医嘱内容"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 7) = .TextMatrix(i, GetColNum(vsList, "部位方法"))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 8) = IIf(Nvl(rsTmp!报告打印) = 0, "", Nvl(rsTmp!报告打印))
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 9) = 2
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 10) = lngOldDeptID
                            vsfList.Cell(flexcpText, vsfList.Rows - 1, 11) = Nvl(rsTmp!ID)
            
                            vsfList.IsSubtotal(vsfList.Rows - 1) = True
                            vsfList.RowOutlineLevel(vsfList.Rows - 1) = 2
        
                            rsTmp.MoveNext
                        Wend
                    End If
            
                    vsfList.Outline 1
                    
                    If vsfList.Rows > 0 Then
                        vsfList.Row = 0
                        vsfList.RowSel = 0
                    End If
            
                    vsfList.Redraw = flexRDBuffered
                End If
            End With
        Next
    Else
        For i = 1 To vsList.Rows - 1
            With vsList
                If .TextMatrix(i, GetColNum(vsList, "检查过程")) = "已报告" _
                    Or .TextMatrix(i, GetColNum(vsList, "检查过程")) = "已审核" _
                    Or .TextMatrix(i, GetColNum(vsList, "检查过程")) = "已完成" Then
                
                    '如果有“执行科室ID”，则需要重新根据科室ID读取平诊报告需审核的参数
                    If GetColNum(vsList, "执行科室ID") <> 0 Then
                        If lngOldDeptID <> .TextMatrix(i, GetColNum(vsList, "执行科室ID")) Then   '科室ID改变了，重新读取平诊报告打印的参数
                            lngOldDeptID = .TextMatrix(i, GetColNum(vsList, "执行科室ID"))
                            blnCanPrint = GetDeptPara(lngOldDeptID, "平诊需审核才能打报告") = "1"           '平诊需要审核才能打印 =true
                        End If
                    Else
                        lngOldDeptID = lngDeptID
                    End If
                    If IIf(blnCanPrint, IIf(.Cell(flexcpData, i, GetColNum(vsList, "紧急")) = 1, .TextMatrix(i, GetColNum(vsList, "报告人")) <> "", .TextMatrix(i, GetColNum(vsList, "复核人")) <> ""), True) Then
                        iCount = iCount + 1
                        vsfList.Rows = vsfList.Rows + 1
                        vsfList.Cell(flexcpText, iCount, 0) = .TextMatrix(i, GetColNum(vsList, "医嘱ID"))
                        vsfList.Cell(flexcpChecked, iCount, 1) = 2
                        vsfList.TextMatrix(iCount, 1) = .TextMatrix(i, GetColNum(vsList, "姓名"))
                        vsfList.Cell(flexcpText, iCount, 2) = .TextMatrix(i, GetColNum(vsList, "来源"))
                        vsfList.Cell(flexcpText, iCount, 3) = .TextMatrix(i, GetColNum(vsList, "检查号"))
                        vsfList.Cell(flexcpText, iCount, 4) = .TextMatrix(i, GetColNum(vsList, "性别"))
                        vsfList.Cell(flexcpText, iCount, 5) = .TextMatrix(i, GetColNum(vsList, "年龄"))
                        vsfList.Cell(flexcpText, iCount, 6) = .TextMatrix(i, GetColNum(vsList, "医嘱内容"))
                        vsfList.Cell(flexcpText, iCount, 7) = .TextMatrix(i, GetColNum(vsList, "部位方法"))
                        vsfList.Cell(flexcpText, iCount, 8) = Nvl(.TextMatrix(i, GetColNum(vsList, "报告打印")), "")
                        vsfList.Cell(flexcpText, iCount, 9) = IIf(mrtReportTye = 电子病历编辑器, 0, 1)
                        vsfList.Cell(flexcpText, iCount, 10) = lngOldDeptID
                    End If
                End If
            End With
        Next
    End If
    
    '自动列宽
    vsfList.AutoSize 0, vsfList.Cols - 1
    '内容靠左
    If vsfList.Rows > 1 Then vsfList.Cell(flexcpAlignment, 1, 1, vsfList.Rows - 1, vsfList.Cols - 1) = flexAlignLeftCenter
    
    Me.Caption = "选择需要打印的医嘱，医嘱总数为：" & iCount
End Sub

Private Sub chkChoose_Click()
    Dim i As Integer
    
    If chkChoose.value = 1 Then
        chkChoose.Caption = "全清(&D)"
        For i = 1 To vsfList.Rows - 1
            vsfList.Cell(flexcpChecked, i, 1) = 1
        Next
    Else
        chkChoose.Caption = "全选(&A)"
        For i = 1 To vsfList.Rows - 1
            vsfList.Cell(flexcpChecked, i, 1) = 2
        Next
    End If
End Sub

Private Sub chkPrinted_Click()
    Dim i As Integer
    
    For i = 1 To vsfList.Rows - 1
        If vsfList.TextMatrix(i, 8) = "" Then vsfList.Cell(flexcpChecked, i, 1) = IIf(chkPrinted.value = 1, 1, 2)
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    '组织返回值，返回值由"医嘱ID（报告文档编辑器:报告ID）-编辑器类型-执行科室ID|医嘱ID（报告文档编辑器:报告ID）-编辑器类型-执行科室ID|..."组成
    Dim i As Long
    
    If mrtReportTye = 报告文档编辑器 Then
        For i = 1 To vsfList.Rows - 1
            If vsfList.Cell(flexcpChecked, i, 1) = 1 And vsfList.RowOutlineLevel(i) = 2 Then
                mstrReturn = mstrReturn & "|" & vsfList.Cell(flexcpText, i, 11) _
                             & "-" & vsfList.Cell(flexcpText, i, 9) & "-" & vsfList.Cell(flexcpText, i, 10)
            End If
        Next
    Else
        For i = 1 To vsfList.Rows - 1
            If vsfList.Cell(flexcpChecked, i, 1) = 1 Then
                mstrReturn = mstrReturn & "|" & vsfList.Cell(flexcpText, i, 0) _
                             & "-" & vsfList.Cell(flexcpText, i, 9) & "-" & vsfList.Cell(flexcpText, i, 10)
            End If
        Next
    End If
    
    mstrReturn = Mid(mstrReturn, 2)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdOK_Click
    End If
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Dim lngCurLevel As Long
    Dim i As Long
    
    If Col <> 1 Then Exit Sub
    
    lngCurLevel = vsfList.RowOutlineLevel(Row)

    For i = Row + 1 To vsfList.Rows - 1
        If vsfList.RowOutlineLevel(i) <= lngCurLevel Then Exit For
        
        vsfList.Cell(flexcpChecked, i, 1) = vsfList.Cell(flexcpChecked, Row, Col)
    Next i
    
    i = Row - 1
    While i >= 1
        If vsfList.RowOutlineLevel(i) < lngCurLevel Then
            If vsfList.Cell(flexcpChecked, Row, 1) = 2 Then
                vsfList.Cell(flexcpChecked, i, 1) = 2
                lngCurLevel = vsfList.RowOutlineLevel(i)
            End If
        End If
        
        i = i - 1
    Wend
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsfList_DblClick()
On Error GoTo errHandle
    If vsfList.Rows <= 0 Or mrtReportTye <> 报告文档编辑器 Then Exit Sub
    
    If vsfList.IsCollapsed(vsfList.Row) = flexOutlineCollapsed Then
        vsfList.IsCollapsed(vsfList.Row) = flexOutlineExpanded
    Else
        vsfList.IsCollapsed(vsfList.Row) = flexOutlineCollapsed
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
