VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmScheme_FilterCfg 
   BorderStyle     =   0  'None
   Caption         =   "查找过滤设置"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFilter 
      Height          =   3255
      Left            =   1080
      ScaleHeight     =   3195
      ScaleWidth      =   9075
      TabIndex        =   9
      Top             =   3720
      Width           =   9135
      Begin VB.CommandButton cmdNextFilter 
         Caption         =   "下移"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   15
         Top             =   2040
         Width           =   1100
      End
      Begin VB.CommandButton cmdLastFilter 
         Caption         =   "上 移"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   14
         Top             =   1560
         Width           =   1100
      End
      Begin VB.Frame fraFilterSet 
         Height          =   30
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   7215
      End
      Begin VB.CommandButton cmdNewFilter 
         Caption         =   "新增快速过滤项"
         Enabled         =   0   'False
         Height          =   465
         Left            =   7800
         TabIndex        =   12
         Top             =   480
         Width           =   1100
      End
      Begin VB.CommandButton cmdDeleteFilter 
         Caption         =   "删 除"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   11
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterReset 
         Caption         =   "重 置"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   10
         Top             =   2520
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFilter 
         Height          =   2655
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   7335
         _cx             =   12938
         _cy             =   4683
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   350
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
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         Caption         =   "快速过滤配置"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.PictureBox picCondition 
      Height          =   3495
      Left            =   1080
      ScaleHeight     =   3435
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmdReset 
         Caption         =   "重 置"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   6
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CommandButton cmdNextCondition 
         Caption         =   "下 移"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   5
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdLastCondition 
         Caption         =   "上 移"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   4
         Top             =   1800
         Width           =   1100
      End
      Begin VB.Frame fraInputSet 
         Height          =   30
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton cmdNewCondition 
         Caption         =   "新增自定义条件"
         Enabled         =   0   'False
         Height          =   465
         Left            =   7800
         TabIndex        =   2
         Top             =   480
         Width           =   1100
      End
      Begin VB.CommandButton cmdDeleteCondition 
         Caption         =   "删 除"
         Enabled         =   0   'False
         Height          =   350
         Left            =   7800
         TabIndex        =   1
         Top             =   1200
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfConditonCfg 
         Height          =   2655
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   7335
         _cx             =   12938
         _cy             =   4683
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   350
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
      Begin VB.Label lblInput 
         AutoSize        =   -1  'True
         Caption         =   "条件录入配置"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1080
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmScheme_FilterCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mobjFilterCfg As clsScSerachCfg
Public mblnIsEdit As Boolean    '是否已编辑

Private mblnState As Boolean    '是否正在编辑状态
Private mblnNewState As Boolean
Private mobjCustomQueryForm As New frmSetDataFrom
Private mstrFilterItem As String
Private mstrQuerySql As String
Private mobjSqlScheme As New clsSqlScheme
Private Const M_STR_INPUTCOL = "录入项目|控件类型|扩展属性|默认值|数据来源|"
Private Const M_STR_FILTERCOL = "过滤项目|选择方式|扩展属性|数据来源|自定义过滤脚本|"
'Private Const M_STR_FILTERCOL = "过滤项目|选择方式|扩展属性|数据来源|自定义过滤脚本|"

Private Const M_STR_INSTYLE = "0-弹窗录入|1-快捷录入|2-快捷+弹窗"
Private Const M_STR_CONSTYLE = "0-文本框|1-日期框|2-时间框|3-长日期框|4-下拉框|5-多选框|6-年龄框|7-互斥框"
Private Const M_STR_CHKSTYLE = "单选|多选"

Private Enum ConColTitlte
    it录入项目 = 0
    it控件类型 = 1
    it扩展属性 = 2
    it默认值 = 3
    it数据来源 = 4
    itIsNew = 5
End Enum

Private Enum FilColTitlte
    ft过滤项目 = 0
    ft选择方式 = 1
    ft扩展属性 = 2
    ft数据来源 = 3
    ft自定义过滤脚本 = 4
    ftIsNew = 5
End Enum

Private Sub cmdDeleteCondition_Click()
    On Error GoTo errHandle
    
    If vsfConditonCfg.Rows < 2 Or IsSelectionRow(vsfConditonCfg) = False Then Exit Sub
    
    If Val(vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.itIsNew)) = 1 Then
        vsfConditonCfg.RemoveItem vsfConditonCfg.Row
        mblnIsEdit = True
        If vsfConditonCfg.Rows < 2 Then
            cmdDeleteCondition.Enabled = False
        End If
    Else
        MsgBox "查询所需条件，不允许进行删除", vbInformation, Me.Caption
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdDeleteFilter_Click()
    On Error GoTo errHandle
    
    If vsfFilter.Rows < 2 Or IsSelectionRow(vsfFilter) = False Then Exit Sub
     
    vsfFilter.RemoveItem vsfFilter.Row
    mblnIsEdit = True
    If vsfFilter.Rows < 2 Then
        cmdDeleteFilter.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdFilterReset_Click()
    On Error GoTo errHandle
    
    Call ShowFilterSet(mobjSqlScheme, 2)
    If vsfFilter.Rows > 1 Then
        cmdDeleteFilter.Enabled = True
    Else
        cmdDeleteFilter.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdLastCondition_Click()
    On Error GoTo errHandle
    
    Call MoveUp(vsfConditonCfg)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdLastFilter_Click()
    On Error GoTo errHandle
    
    Call MoveUp(vsfFilter)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNewCondition_Click()
    On Error GoTo errHandle
    
    If vsfConditonCfg.Rows = 1 Then
        cmdDeleteCondition.Enabled = True
    End If
    
    mblnNewState = True
    Call NewRow(vsfConditonCfg)
    mblnIsEdit = True
    vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.itIsNew) = 1
    Call ConCfgDataDefalut(vsfConditonCfg.Row)
    vsfConditonCfg.EditCell
    mblnNewState = False
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNewFilter_Click()
    On Error GoTo errHandle
    
    If vsfFilter.Rows = 1 Then
        cmdDeleteFilter.Enabled = True
    End If
    Call NewRow(vsfFilter)
    mblnIsEdit = True
    vsfFilter.TextMatrix(vsfFilter.Row, FilColTitlte.ftIsNew) = 1
    Call FiltetDataDefalut(vsfFilter.Row)
    vsfFilter.EditCell
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNextCondition_Click()
    On Error GoTo errHandle
    
    Call MoveDown(vsfConditonCfg)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNextFilter_Click()
    On Error GoTo errHandle
    
    Call MoveDown(vsfFilter)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdReset_Click()
'重置
    On Error GoTo errHandle
    
    Call ShowFilterSet(mobjSqlScheme, 1)
    Call RefreshFilterSet(mstrQuerySql, mobjSqlScheme, True)
    If vsfConditonCfg.Rows > 1 Then
        cmdDeleteCondition.Enabled = True
    Else
        cmdDeleteCondition.Enabled = False
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    mblnNewState = False
    Call InitDockPannel
    Call GridInit(M_STR_INPUTCOL, vsfConditonCfg)
    Call GridInit(M_STR_FILTERCOL, vsfFilter)
    Call GridShow
    Call RefreshWindowState(False)
    Call SetFontSize(gbytFontSize)
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub


Private Sub GridShow()
    With vsfConditonCfg
        .ColHidden(ConColTitlte.itIsNew) = True
        .ColComboList(ConColTitlte.it控件类型) = M_STR_CONSTYLE
        .ColComboList(ConColTitlte.it默认值) = "..."
        .ColComboList(ConColTitlte.it数据来源) = "..."
    End With
    With vsfFilter
        .ColHidden(FilColTitlte.ftIsNew) = True
        .ColComboList(FilColTitlte.ft选择方式) = M_STR_CHKSTYLE
        .ColComboList(FilColTitlte.ft数据来源) = "..."
        .ColComboList(FilColTitlte.ft自定义过滤脚本) = "..."
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjCustomQueryForm Is Nothing Then Unload mobjCustomQueryForm
    
    Set mobjFilterCfg = Nothing
    Set mobjCustomQueryForm = Nothing
    Set mobjSqlScheme = Nothing
End Sub

Private Sub picCondition_Resize()
    On Error Resume Next
    
    '录入配置部分
    lblInput.Move picCondition.ScaleLeft + 100, picCondition.ScaleTop + 100
    fraInputSet.Move lblInput.Left + lblInput.Width, lblInput.Top + lblInput.Height / 2, picCondition.ScaleWidth - lblInput.Left
    vsfConditonCfg.Move picCondition.ScaleLeft + 100, fraInputSet.Top + 200, picCondition.ScaleWidth - 300 - cmdNewCondition.Width, picCondition.ScaleHeight - vsfConditonCfg.Top
    cmdNewCondition.Move vsfConditonCfg.Left + vsfConditonCfg.Width + 100, vsfConditonCfg.Top
    cmdDeleteCondition.Move cmdNewCondition.Left, cmdNewCondition.Top + cmdNewCondition.Height + 100
    cmdLastCondition.Move cmdNewCondition.Left, cmdDeleteCondition.Top + cmdDeleteCondition.Height + 100
    cmdNextCondition.Move cmdNewCondition.Left, cmdLastCondition.Top + cmdLastCondition.Height + 100
    cmdReset.Move cmdNewCondition.Left, cmdNextCondition.Top + cmdNextCondition.Height + 100
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    
    '过滤配置部分
    lblFilter.Move picFilter.ScaleLeft + 100, picFilter.Top + 100
    fraFilterSet.Move lblFilter.Left + lblFilter.Width, lblFilter.Top + lblFilter.Height / 2, picFilter.ScaleWidth - fraFilterSet.Left
    vsfFilter.Move vsfConditonCfg.Left, fraFilterSet.Top + 200, vsfConditonCfg.Width, picFilter.ScaleHeight - 750
    cmdNewFilter.Move cmdNewCondition.Left, vsfFilter.Top
    cmdDeleteFilter.Move cmdNewCondition.Left, cmdNewFilter.Top + cmdNewFilter.Height + 100
    cmdLastFilter.Move cmdNewCondition.Left, cmdDeleteFilter.Top + cmdDeleteFilter.Height + 100
    cmdNextFilter.Move cmdNewCondition.Left, cmdLastFilter.Top + cmdLastFilter.Height + 100
    cmdFilterReset.Move cmdNewCondition.Left, cmdNextFilter.Top + cmdNextFilter.Height + 100
End Sub

Private Sub vsfConditonCfg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    Dim strPara As String
    Dim i As Long
    
     If Col = ConColTitlte.it默认值 Or Col = ConColTitlte.it数据来源 Then
        For i = 1 To Row - 1
            If vsfConditonCfg.RowHidden(i) = False _
                And vsfConditonCfg.TextMatrix(i, ConColTitlte.it控件类型) <> "8-可选框" _
                And vsfConditonCfg.TextMatrix(i, ConColTitlte.it控件类型) <> "9-条件选择框" _
                And Len(Trim(vsfConditonCfg.TextMatrix(i, ConColTitlte.it录入项目))) > 0 Then
                strPara = strPara & "|" & vsfConditonCfg.TextMatrix(i, ConColTitlte.it录入项目)
            End If
        Next
        strPara = Mid(strPara, 2)
        strValue = vsfConditonCfg.TextMatrix(Row, Col)
        strValue = mobjCustomQueryForm.ShowSqlFromWindow(strValue, strPara, IIf(Col = ConColTitlte.it默认值, 1, 2), mblnState, gbytFontSize, Me)
        vsfConditonCfg.TextMatrix(Row, Col) = strValue
    End If
End Sub

Private Sub vsfConditonCfg_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    If mblnState Then
        mblnIsEdit = True
    End If
    
    '带*号的参数表示可对查询的条件方式进行选择
    If Col = ConColTitlte.it控件类型 Then
        If InStr(vsfConditonCfg.TextMatrix(Row, ConColTitlte.it录入项目), "*") = 1 Then
            If vsfConditonCfg.TextMatrix(Row, Col) <> GetConDataChange("ControlType", TControlType.ctCombobox) Then
                vsfConditonCfg.TextMatrix(Row, Col) = GetConDataChange("ControlType", TControlType.ctCombobox)
            End If
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub
 

Private Sub vsfFilter_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    Dim strPara As String
    Dim i As Long
    
    If Col = FilColTitlte.ft数据来源 Then
        For i = 1 To Row - 1
            If vsfFilter.RowHidden(i) = False And Len(Trim(vsfFilter.TextMatrix(i, FilColTitlte.ft过滤项目))) > 0 Then
                strPara = strPara & "|" & vsfFilter.TextMatrix(i, FilColTitlte.ft过滤项目)
            End If
        Next
        strPara = Mid(strPara, 2)
        strValue = vsfFilter.TextMatrix(Row, Col)
        strValue = mobjCustomQueryForm.ShowSqlFromWindow(strValue, strPara, 2, mblnState, gbytFontSize, Me)
        vsfFilter.TextMatrix(Row, Col) = strValue
    End If
    
    If Col = FilColTitlte.ft自定义过滤脚本 Then
        strValue = vsfFilter.TextMatrix(Row, Col)
        strValue = mobjCustomQueryForm.ShowSqlFromWindow(strValue, "", 4, mblnState, gbytFontSize, Me)
        vsfFilter.TextMatrix(Row, Col) = strValue
    End If
End Sub

Private Sub vsfFilter_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    If mblnState Then
        mblnIsEdit = True
    End If
    
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsfFilter_RowColChange()
    Dim strFilterItem As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If vsfFilter.Row < 1 Then Exit Sub
    vsfFilter.Editable = flexEDKbdMouse
    If mblnState Then
        strFilterItem = mstrFilterItem
        If vsfFilter.Col = 0 And vsfFilter.Row > 0 Then
            strFilterItem = InitFilterItem(mstrQuerySql)
            For i = 1 To vsfFilter.Row - 1
                If Val(vsfFilter.TextMatrix(i, FilColTitlte.ftIsNew)) = 1 And Len(Trim(vsfFilter.TextMatrix(i, FilColTitlte.ft过滤项目))) > 0 Then
                    strFilterItem = strFilterItem & "|" & vsfFilter.TextMatrix(i, FilColTitlte.ft过滤项目)
                End If
            Next
            vsfFilter.ColComboList(FilColTitlte.ft过滤项目) = strFilterItem
        End If
    Else
        If Not (vsfFilter.Col = FilColTitlte.ft数据来源 Or vsfFilter.Col = FilColTitlte.ft自定义过滤脚本) Then
            vsfFilter.Editable = flexEDNone
        End If
    End If

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub


Private Sub vsfConditonCfg_RowColChange()
    On Error GoTo errHandle
    
    If vsfConditonCfg.Row < 1 Then Exit Sub
    vsfConditonCfg.Editable = flexEDKbdMouse
    If mblnState Then
        If vsfConditonCfg.Col = ConColTitlte.it录入项目 _
            And Not (Val(vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.itIsNew)) = 1 Or mblnNewState) _
                Or vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.it控件类型) = "8-可选框" _
                Or vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.it控件类型) = "9-条件选择框" Then
            
                If (vsfConditonCfg.Col = ConColTitlte.it控件类型) Or (vsfConditonCfg.Col = ConColTitlte.it录入项目) Then
                    vsfConditonCfg.Editable = flexEDNone
                End If
        End If
    Else
        If Not (vsfConditonCfg.Col = ConColTitlte.it默认值 Or vsfConditonCfg.Col = ConColTitlte.it数据来源) Then
            vsfConditonCfg.Editable = flexEDNone
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub



Private Sub ConCfgDataDefalut(lngRow As Long)
'vsfConditonCfg设置默认值
    With vsfConditonCfg
        .TextMatrix(lngRow, ConColTitlte.it控件类型) = "0-文本框"
        .TextMatrix(lngRow, ConColTitlte.it扩展属性) = ""
        .TextMatrix(lngRow, ConColTitlte.it默认值) = ""
        .TextMatrix(lngRow, ConColTitlte.it数据来源) = ""
    End With
End Sub

Private Sub FiltetDataDefalut(lngRow As Long)
'vsfFilter设置默认值
    With vsfFilter
        .TextMatrix(lngRow, FilColTitlte.ft选择方式) = "单选"
        .TextMatrix(lngRow, FilColTitlte.ft扩展属性) = ""
        .TextMatrix(lngRow, FilColTitlte.ft数据来源) = ""
        .TextMatrix(lngRow, FilColTitlte.ft自定义过滤脚本) = ""
    End With
End Sub

Public Sub SetConditionCfg(objSqlScheme As clsSqlScheme)
    '写入录入配置
    Dim objScSearchCfg As clsScSerachCfg
    Dim objScFilterCfg As clsScFilterCfg
    Dim i As Long
    
    If vsfConditonCfg.Rows < 2 Then Exit Sub
    For i = 1 To vsfConditonCfg.Rows - 1
        Set objScSearchCfg = New clsScSerachCfg
        With vsfConditonCfg
            If Len(.TextMatrix(i, ConColTitlte.it录入项目)) > 0 And .RowHidden(i) = False Then
                objScSearchCfg.Name = .TextMatrix(i, ConColTitlte.it录入项目)
                objScSearchCfg.ControlType = SetConDataChange(i, ConColTitlte.it控件类型)
                objScSearchCfg.Default = .TextMatrix(i, ConColTitlte.it默认值)
                objScSearchCfg.ExtProperty = .TextMatrix(i, ConColTitlte.it扩展属性)
                objScSearchCfg.DataFrom = .TextMatrix(i, ConColTitlte.it数据来源)
                objSqlScheme.AddSerachCfg objScSearchCfg
            End If
        End With
    Next
    
    '快速过滤配置
    For i = 1 To vsfFilter.Rows - 1
        Set objScFilterCfg = New clsScFilterCfg
        With vsfFilter
            If Len(.TextMatrix(i, FilColTitlte.ft过滤项目)) > 0 And .RowHidden(i) = False Then
                objScFilterCfg.Name = .TextMatrix(i, FilColTitlte.ft过滤项目)
                objScFilterCfg.SelectWay = IIf(.TextMatrix(i, FilColTitlte.ft选择方式) = "多选", 1, 0)
                objScFilterCfg.ExtProperty = .TextMatrix(i, FilColTitlte.ft扩展属性)
                objScFilterCfg.DataFrom = .TextMatrix(i, FilColTitlte.ft数据来源)
                objScFilterCfg.CustomScript = .TextMatrix(i, FilColTitlte.ft自定义过滤脚本)
                objSqlScheme.AddFilterCfg objScFilterCfg
            End If
        End With
    Next
End Sub


Private Function SetConDataChange(lngRow As Long, lngCol As Long) As Long
'vsfConditonCfg写入数据转换
    Dim strValue As String
    Dim arrData() As String
    strValue = vsfConditonCfg.TextMatrix(lngRow, lngCol)
    
    If Len(strValue) = 0 Then
        SetConDataChange = 0
        Exit Function
    End If
    
    arrData = Split(strValue, "-")
    SetConDataChange = Val(arrData(0))
End Function

Private Function GetConDataChange(strItem As String, lngNo As Long) As String
'vsfConditonCfg读取数据转换
    Dim arrContent() As String
    Dim arrText() As String
    Dim i As Long
    
    Select Case strItem
        Case "ConColTitlte"
            arrContent = Split(M_STR_INSTYLE, "|")
        Case "ControlType"
            arrContent = Split(M_STR_CONSTYLE, "|")
    End Select
    
    For i = 0 To UBound(arrContent)
        arrText = Split(arrContent(i), "-")
        If lngNo = arrText(0) Then
            GetConDataChange = arrText(0) & "-" & arrText(1)
            Exit Function
        ElseIf lngNo = 8 And strItem = "ControlType" Then
            GetConDataChange = "8-可选框"
        ElseIf lngNo = 9 And strItem = "ControlType" Then
            GetConDataChange = "9-条件选择框"
        End If
    Next
End Function

Public Sub RefreshWindowState(blnState As Boolean)
    mblnState = blnState
    cmdDeleteCondition.Enabled = False
    cmdDeleteFilter.Enabled = False
    cmdLastCondition.Enabled = blnState
    cmdLastFilter.Enabled = blnState
    cmdNewCondition.Enabled = blnState
    cmdNewFilter.Enabled = blnState
    cmdNextCondition.Enabled = blnState
    cmdNextFilter.Enabled = blnState
    cmdReset.Enabled = blnState
    cmdFilterReset.Enabled = blnState
    
    If blnState Then
        vsfConditonCfg.BackColor = &H80000005
        vsfFilter.BackColor = &H80000005
        If vsfConditonCfg.Rows > 1 Then
            cmdDeleteCondition.Enabled = blnState
        End If
        
        If vsfFilter.Rows > 1 Then
            cmdDeleteFilter.Enabled = blnState
        End If
    Else
        vsfConditonCfg.BackColor = &H8000000F
        vsfFilter.BackColor = &H8000000F
    End If
   
End Sub

Private Function InitFilterItem(strSchemeSql As String) As String
'设置可选过滤项目
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim strItem As String
    Dim i As Long

    objSqlParse.init strSchemeSql

    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    mstrFilterItem = ""
    If rsRecord Is Nothing Then Exit Function
    For i = 0 To rsRecord.Fields.Count - 1
        InitFilterItem = InitFilterItem & "|" & rsRecord.Fields(i).Name
    Next
End Function

'Public Sub ClearScheme()
'    vsfConditonCfg.Rows = 1
'    vsfFilter.Rows = 1
'End Sub


Private Function GetQueryItem(strSchemeSql As String) As String
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim i As Long
    
    
    objSqlParse.init strSchemeSql
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    If rsRecord Is Nothing Then Exit Function
    For i = 0 To rsRecord.Fields.Count - 1
        GetQueryItem = GetQueryItem & "|" & rsRecord.Fields(i).Name
    Next
    
    GetQueryItem = GetQueryItem & "|"
End Function

Public Sub ShowFilterSet(objSqlScheme As clsSqlScheme, Optional lngReset As Long)
'界面配置显示
    Dim objScSearchCfg As New clsScSerachCfg
    Dim objScFilterCfg As New clsScFilterCfg
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim strSelecItem As String
    Dim lngRow As Long
    Dim arrQueryPara() As String
    Dim strQueryItem As String
    Dim i As Long
    Dim j As Long
     
    Set mobjSqlScheme = objSqlScheme
    mstrQuerySql = objSqlScheme.Query
    
    objSqlParse.init objSqlScheme.Query
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    
    '显示录入配置
    If lngReset <> 2 Then
        vsfConditonCfg.Rows = 1
        For i = 1 To objSqlScheme.SerachCfgCount
            Set objScSearchCfg = objSqlScheme.SerachCfg(i)
            With vsfConditonCfg
                If InStr(1, UCase(gstrPara & IIf(Len(gstrBasePara) > 0, ",", "") & gstrBasePara), "[" & UCase(objScSearchCfg.Name) & "]") = 0 Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, ConColTitlte.it录入项目) = objScSearchCfg.Name
                    
          
                    .TextMatrix(lngRow, ConColTitlte.it控件类型) = GetConDataChange("ControlType", objScSearchCfg.ControlType)
                  
                    If (.TextMatrix(lngRow, ConColTitlte.it控件类型) = "8-可选框") _
                     Or (.TextMatrix(lngRow, ConColTitlte.it控件类型) = "9-条件选择框") Then
                        vsfConditonCfg.Cell(flexcpBackColor, lngRow, 0, lngRow, vsfConditonCfg.Cols - 1) = &HC0FFFF
                    End If
                    .TextMatrix(lngRow, ConColTitlte.it扩展属性) = objScSearchCfg.ExtProperty
                    .TextMatrix(lngRow, ConColTitlte.it默认值) = objScSearchCfg.Default
                    .TextMatrix(lngRow, ConColTitlte.it数据来源) = objScSearchCfg.DataFrom
                    .TextMatrix(lngRow, ConColTitlte.itIsNew) = IIf(objSqlParse.SqlStruct.HasParName(objScSearchCfg.Name), 0, 1)
                End If
            End With
        Next
    End If
    '快速过滤配置
    If lngReset <> 1 Then
        vsfFilter.Rows = 1
        For i = 1 To objSqlScheme.FilterCfgCount
            Set objScFilterCfg = objSqlScheme.FilterCfg(i)
            With vsfFilter
                .Rows = .Rows + 1
                .TextMatrix(i, FilColTitlte.ft过滤项目) = objScFilterCfg.Name
                .TextMatrix(i, FilColTitlte.ft选择方式) = IIf(objScFilterCfg.SelectWay = swMulti, "多选", "单选")
                .TextMatrix(i, FilColTitlte.ft扩展属性) = objScFilterCfg.ExtProperty
                .TextMatrix(i, FilColTitlte.ft数据来源) = objScFilterCfg.DataFrom
                .TextMatrix(i, FilColTitlte.ft自定义过滤脚本) = objScFilterCfg.CustomScript
                .TextMatrix(i, FilColTitlte.ftIsNew) = IIf(HasSelectItem(objScFilterCfg.Name, rsRecord), 0, 1)
    
            End With
        Next
    End If
End Sub

Public Sub RefreshFilterSet(strQuerySql As String, objSqlScheme As clsSqlScheme, Optional lngReset As Long)
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim strQueryPara As String
    Dim strCurPara As String
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    Dim blnIsCusPara As Boolean
    Dim blnIsHave As Boolean
    Dim blnResetWhere As Boolean
    
    mstrQuerySql = strQuerySql
    Set mobjSqlScheme = objSqlScheme
    objSqlParse.init strQuerySql
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    
    '刷新录入项目设置
    If lngReset <> 2 Then
        For i = 1 To vsfConditonCfg.Rows - 1
            If Val(vsfConditonCfg.TextMatrix(i, ConColTitlte.itIsNew)) <> 1 And (Not vsfConditonCfg.RowHidden(i)) Then
                strQueryPara = strQueryPara & "," & "[" & vsfConditonCfg.TextMatrix(i, ConColTitlte.it录入项目) & "]"
                If Not objSqlParse.SqlStruct.HasParName(vsfConditonCfg.TextMatrix(i, ConColTitlte.it录入项目)) And InStr(1, UCase(gstrPara & IIf(Len(gstrBasePara) > 0, ",", "") & gstrBasePara), "[" & UCase(vsfConditonCfg.TextMatrix(i, ConColTitlte.it录入项目)) & "]") = 0 Then
                    vsfConditonCfg.RowHidden(i) = True
                End If
            End If
        Next
        
        strQueryPara = Mid(strQueryPara, 2)
        For i = 1 To objSqlParse.SqlStruct.ParCount
            blnIsCusPara = False
            blnResetWhere = False
            
            strCurPara = objSqlParse.SqlStruct.AllParameter(i)
            
            If (InStr(strCurPara, "[@") > 0) Then blnIsCusPara = True
            If (InStr(strCurPara, "[*") > 0) Then blnResetWhere = True
            
            If blnIsCusPara Or blnResetWhere Then
                strCurPara = Mid$(strCurPara, 3, InStr(strCurPara, ",") - 3)
                blnIsCusPara = True
            Else
                strCurPara = Mid(strCurPara, 2, Len(strCurPara) - 2)
            End If
            
            If InStr(1, UCase(strQueryPara), "[" & UCase(strCurPara) & "]") = 0 And InStr(1, UCase(gstrPara & IIf(Len(gstrBasePara) > 0, ",", "") & gstrBasePara), "[" & UCase(strCurPara) & "]") = 0 Then
                '是否与自定义重复
                blnIsHave = False
                For j = 1 To vsfConditonCfg.Rows - 1
                    If UCase(Trim(strCurPara)) = UCase(Trim(vsfConditonCfg.TextMatrix(j, ConColTitlte.it录入项目))) And (Not vsfConditonCfg.RowHidden(j)) Then
                        blnIsHave = True
                    End If
                Next
                If Not blnIsHave Then
                    vsfConditonCfg.AddItem strCurPara, vsfConditonCfg.Rows
                    Call ConCfgDataDefalut(vsfConditonCfg.Rows - 1)
                    
                    If blnIsCusPara Then
                        vsfConditonCfg.TextMatrix(vsfConditonCfg.Rows - 1, ConColTitlte.it控件类型) = "8-可选框"
                        vsfConditonCfg.Cell(flexcpBackColor, vsfConditonCfg.Rows - 1, 0, vsfConditonCfg.Rows - 1, vsfConditonCfg.Cols - 1) = &HC0FFFF
                    End If
                    
                    If blnResetWhere Then
                        vsfConditonCfg.TextMatrix(vsfConditonCfg.Rows - 1, ConColTitlte.it控件类型) = "9-条件选择框"
                        vsfConditonCfg.Cell(flexcpBackColor, vsfConditonCfg.Rows - 1, 0, vsfConditonCfg.Rows - 1, vsfConditonCfg.Cols - 1) = &HC0FFFF
                    End If
                End If
            End If
        Next
    End If
    If lngReset <> 1 Then
        '刷新快速过滤设置
        For i = 1 To vsfFilter.Rows - 1
            If Val(vsfFilter.TextMatrix(i, FilColTitlte.ftIsNew)) <> 1 And (Not vsfFilter.RowHidden(i)) Then
                If Not HasSelectItem(vsfFilter.TextMatrix(i, FilColTitlte.ft过滤项目), rsRecord) Then
                    vsfFilter.RowHidden(i) = True
                End If
            End If
        Next
    End If
    
    Set objSqlParse = Nothing
    Set objQuery = Nothing
End Sub

Private Function HasSelectItem(strItem As String, rsRecord As Recordset) As Boolean
    Dim i As Long
    
    HasSelectItem = False
    For i = 0 To rsRecord.Fields.Count - 1
        If UCase(strItem) = UCase(rsRecord.Fields(i).Name) Then
            HasSelectItem = True
            Exit Function
        End If
    Next
End Function

Public Function IsEnabledSave() As Boolean
    Dim blnResult As Boolean
    
    blnResult = CheckRepet(vsfConditonCfg, ConColTitlte.it录入项目)
    If blnResult Then
        MsgBox "条件录入配置中录入项目有重复，请检查", vbInformation, Me.Caption
        IsEnabledSave = False
        Exit Function
    End If
    
    blnResult = CheckRepet(vsfFilter, FilColTitlte.ft过滤项目)
    If blnResult Then
        MsgBox "快速过滤配置中过滤项目有重复，请检查", vbInformation, Me.Caption
        IsEnabledSave = False
        Exit Function
    End If
    
    IsEnabledSave = True
End Function

Public Sub UnloadMe()
    Unload Me
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim lngCmdHeight As Long
    Dim lngCmdWithd As Long
    Dim lngNewCondition As Long
    
    If bytFontSize = 9 Then
        lngCmdHeight = 350
        lngCmdWithd = 1100
        vsfConditonCfg.Width = 7355
        lngNewCondition = 465
        Call SetColWithd(0)
    ElseIf bytFontSize = 12 Then
        lngCmdHeight = 385
        lngCmdWithd = 1300
        vsfConditonCfg.Width = 7155
        lngNewCondition = 545
        Call SetColWithd(1)
    ElseIf bytFontSize = 15 Then
        lngCmdHeight = 420
        lngCmdWithd = 1500
        vsfConditonCfg.Width = 6955
        lngNewCondition = 625
        Call SetColWithd(2)
    End If
    
    lblInput.FontSize = bytFontSize
    lblFilter.FontSize = bytFontSize
    vsfConditonCfg.FontSize = bytFontSize
    vsfFilter.FontSize = bytFontSize
    
    cmdNewCondition.FontSize = bytFontSize
    cmdNewCondition.Height = lngNewCondition
    cmdNewCondition.Width = lngCmdWithd
    cmdDeleteCondition.FontSize = bytFontSize
    cmdDeleteCondition.Height = lngCmdHeight
    cmdDeleteCondition.Width = lngCmdWithd
    cmdLastCondition.FontSize = bytFontSize
    cmdLastCondition.Height = lngCmdHeight
    cmdLastCondition.Width = lngCmdWithd
    cmdNextCondition.FontSize = bytFontSize
    cmdNextCondition.Height = lngCmdHeight
    cmdNextCondition.Width = lngCmdWithd
    cmdReset.FontSize = bytFontSize
    cmdReset.Height = lngCmdHeight
    cmdReset.Width = lngCmdWithd
    
    cmdNewFilter.FontSize = bytFontSize
    cmdNewFilter.Height = lngNewCondition
    cmdNewFilter.Width = lngCmdWithd
    cmdDeleteFilter.FontSize = bytFontSize
    cmdDeleteFilter.Height = lngCmdHeight
    cmdDeleteFilter.Width = lngCmdWithd
    cmdLastFilter.FontSize = bytFontSize
    cmdLastFilter.Height = lngCmdHeight
    cmdLastFilter.Width = lngCmdWithd
    cmdNextFilter.FontSize = bytFontSize
    cmdNextFilter.Height = lngCmdHeight
    cmdNextFilter.Width = lngCmdWithd
    cmdFilterReset.FontSize = bytFontSize
    cmdFilterReset.Height = lngCmdHeight
    cmdFilterReset.Width = lngCmdWithd
    
    Call picCondition_Resize
    Call picFilter_Resize
End Sub

Private Sub SetColWithd(ByVal bytSize As Long)
    Select Case bytSize
        Case 0
            vsfConditonCfg.ColWidth(ConColTitlte.it扩展属性) = 1700
            vsfConditonCfg.ColWidth(ConColTitlte.it控件类型) = 1200
            vsfConditonCfg.ColWidth(ConColTitlte.it录入项目) = 1200
            vsfConditonCfg.ColWidth(ConColTitlte.it默认值) = 2000
            
            vsfFilter.ColWidth(FilColTitlte.ft过滤项目) = 1400
            vsfFilter.ColWidth(FilColTitlte.ft选择方式) = 1400
            vsfFilter.ColWidth(FilColTitlte.ft扩展属性) = 1700
            vsfFilter.ColWidth(FilColTitlte.ft数据来源) = 4000
        Case 1
            vsfConditonCfg.ColWidth(ConColTitlte.it扩展属性) = 2200
            vsfConditonCfg.ColWidth(ConColTitlte.it控件类型) = 1400
            vsfConditonCfg.ColWidth(ConColTitlte.it录入项目) = 1600
            vsfConditonCfg.ColWidth(ConColTitlte.it默认值) = 2000
            
            vsfFilter.ColWidth(FilColTitlte.ft过滤项目) = 1450
            vsfFilter.ColWidth(FilColTitlte.ft选择方式) = 1450
            vsfFilter.ColWidth(FilColTitlte.ft扩展属性) = 2200
            vsfFilter.ColWidth(FilColTitlte.ft数据来源) = 4000
        Case 2
            vsfConditonCfg.ColWidth(ConColTitlte.it扩展属性) = 2700
            vsfConditonCfg.ColWidth(ConColTitlte.it控件类型) = 1600
            vsfConditonCfg.ColWidth(ConColTitlte.it录入项目) = 2000
            vsfConditonCfg.ColWidth(ConColTitlte.it默认值) = 2000
            
            vsfFilter.ColWidth(FilColTitlte.ft过滤项目) = 1700
            vsfFilter.ColWidth(FilColTitlte.ft选择方式) = 1700
            vsfFilter.ColWidth(FilColTitlte.ft扩展属性) = 2700
            vsfFilter.ColWidth(FilColTitlte.ft数据来源) = 4000
    End Select
End Sub

Public Function IsSetted() As Boolean
    IsSetted = vsfConditonCfg.Rows > 1
End Function

'布局绑定
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errHandle
    
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hwnd
        Case 2
            Item.Handle = picFilter.hwnd
    End Select
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

'界面布局
Private Sub InitDockPannel()
    Dim objPane As Pane
    
    On Error GoTo errHandle
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing)
    objPane.Title = "picCondition"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane)
    objPane.Title = "picFilter"
    objPane.Options = PaneNoCaption
    
    Set objPane = Nothing
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Public Function GetSearchNames() As String
On Error GoTo errH
    Dim strTmp As String
    Dim i As Integer
    
    GetSearchNames = ""
    For i = 1 To vsfConditonCfg.Rows - 1
        strTmp = strTmp & ";" & vsfConditonCfg.TextMatrix(i, ConColTitlte.it录入项目) & ";"
    Next
    GetSearchNames = strTmp
    Exit Function
errH:
    GetSearchNames = ""
End Function
