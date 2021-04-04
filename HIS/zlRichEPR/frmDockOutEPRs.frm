VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockOutEPRs 
   BorderStyle     =   0  'None
   Caption         =   "门诊病历记录"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsColumn 
      Height          =   3480
      Left            =   1215
      TabIndex        =   1
      Top             =   1875
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   6138
      Appearance      =   0
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDockOutEPRs.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   315
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frmDockOutEPRs.frx":004E
         ToolTipText     =   "选择需要显示的列(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2655
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2745
      Left            =   225
      TabIndex        =   0
      Top             =   675
      Width           =   7890
      _cx             =   13917
      _cy             =   4842
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      Begin MSComctlLib.ImageList imgThis 
         Left            =   0
         Top             =   1710
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":059C
               Key             =   "书写"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":0B36
               Key             =   "修订"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":10D0
               Key             =   "归档"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockOutEPRs.frx":166A
               Key             =   "转交"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   720
      Top             =   4875
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockOutEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------
'窗体常量
'-----------------------------------------------------
Private Enum mCol
    标志 = 0: ID: 病历种类: 病历名称: 创建人: 创建时间: 保存人: 完成时间: 当前版本: 签名级别: 当前情况: 归档人: 归档日期: 科室ID: 科室名: 处理状态: 打印人: 打印时间: 编辑方式: 申报状态
End Enum

Const conPane_Content = 1
Const conPane_New = 2
Private mstrColWidthConfig As String

'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)       '继承文档对象的“点击诊断参考事件”
Public Event RequestRefresh() '要求主窗体刷新
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者对本程序(1250)的权限串
Private mblnSearch As Boolean   '当前使用者是否具备病历检索(1273)权
Private mlngPatiId As Long      '病人id
Private mlngPageId As Long      '主页id
Private mlngDeptId As Long      '当前操作科室id
Private mblnEdit As Boolean     '是否允许操作
Private mblnMoved As Boolean    '是否转储
Private mlngAdviceID As Long    '医嘱ID
Private mblnOutDoc As Boolean   '是否启用门诊快捷病历

Private WithEvents mfrmNew As frmDockEPRNew
Attribute mfrmNew.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mfrmMonitor As New frmDockEPRMonitor
Attribute mfrmMonitor.VB_VarHelpID = -1
Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR            '表格式病历编辑器
Attribute mObjTabEpr.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Private mbln传染病 As Boolean              '传染病报告卡在病人完成接诊之后也是可以修改的

Private mcbsThis As Object          'CommandBar控件
Private mlngVersion As Long         '选中的文件版本号
Private mblnDisease As Boolean      '是否拥有了1249模块的权限

Private Sub InitColumnSelect()
    On Error Resume Next
    '功能：根据原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vfgThis
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.病历名称, mCol.创建人, mCol.创建时间, mCol.保存人, mCol.完成时间, mCol.当前情况, mCol.科室名
                 vsColumn.Rows = vsColumn.Rows + 1
                 lngRow = vsColumn.Rows - 1
                 vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                 vsColumn.RowData(lngRow) = i
                
                 '固定显示列
                 If InStr(",页面名称,病历名称,", "," & .TextMatrix(0, i) & ",") > 0 Then
                     vsColumn.TextMatrix(lngRow, 0) = 1
                     vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                 End If
            End Select
        Next
    End With
    vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 130
    vsColumn.Row = 1
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Select Case Pane.ID
    Case conPane_Content
        Cancel = True
    Case conPane_New
        Select Case Action
        Case PaneActionClosing, PaneActionClosed: Cancel = False
        Case Else: Cancel = True
        End Select
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Content
        If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
        Item.Handle = mfrmContent.hwnd
    Case conPane_New
        If mfrmNew Is Nothing Then Set mfrmNew = New frmDockEPRNew
        Item.Handle = mfrmNew.hwnd
    End Select
End Sub

Private Sub dkpMan_Resize()
    Dim lngScaleLeft As Long, lngScaleTop  As Long, lngScaleRight  As Long, lngScaleBottom  As Long
    Call Me.dkpMan.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Err = 0: On Error Resume Next
    With Me.vfgThis
        .Left = lngScaleLeft: .Width = lngScaleRight - lngScaleLeft
        .Top = lngScaleTop: .Height = lngScaleBottom - .Top
        .ZOrder 0
    End With
    fraColSel.Move Me.vfgThis.Left + 50, Me.vfgThis.Top + 50
    fraColSel.ZOrder 0
    vsColumn.Move fraColSel.Left, fraColSel.Top + fraColSel.Height
    vsColumn.ZOrder 0
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If vsColumn.Visible Then
        vsColumn.SetFocus '列选择器
    Else
        If Me.vfgThis.Visible Then Me.vfgThis.SetFocus
    End If
    RaiseEvent Activate
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    vsColumn.Visible = False '列选择器
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
Dim i As Integer
    For i = 1 To vfgThis.Rows - 1
        If vfgThis.TextMatrix(i, mCol.ID) = lngRecordId Then
            vfgThis.Cell(flexcpText, i, mCol.打印人) = gstrUserName
            vfgThis.Cell(flexcpText, i, mCol.打印时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm")
            Exit For
        End If
    Next
End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)
    If mblnOutDoc Then RaiseEvent RequestRefresh
End Sub

Private Sub vsColumn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then '关闭列选择器
        If vsColumn.Visible Then
            vsColumn.Visible = False
            vfgThis.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then '打开列选择器
        Call imgColSel_MouseUp(1, 0, 0, 0)
    End If
End Sub

Private Sub vfgThis_KeyDown(KeyCode As Integer, Shift As Integer)
    vsColumn_KeyDown KeyCode, Shift
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsColumn
            If .Visible Then
                .Visible = False
                vfgThis.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vfgThis.ColHidden(.RowData(i)) Or vfgThis.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim lngCol As Long, T As Variant, i As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            T = Split("270;0;0;2200;800;1600;800;1600;0;0;3000;0;0;0;1200;0;800;1600;0", ";")
            vfgThis.ColWidth(lngCol) = T(lngCol)
            vfgThis.ColHidden(lngCol) = False
        Else
            vfgThis.ColWidth(lngCol) = 0
            vfgThis.ColHidden(lngCol) = True
        End If
    End If
    Dim strCols As String
    For i = 0 To 18
        strCols = strCols & IIf(i = 0, "", ";") & vfgThis.ColWidth(i)
    Next
    mstrColWidthConfig = strCols
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    On Error Resume Next
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub
 
Private Sub Form_Load()
    Dim intType As Integer, lngFontSize As Long
    
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "基本") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1250)
    
    mblnOutDoc = Val(zlDatabase.GetPara("显示病历快捷输入", glngSys, 1260, 0, , , intType)) = 1
    
    mstrColWidthConfig = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", _
        "270;0;0;2200;800;1600;800;0;0;0;3000;0;0;0;1200;0;800;1600;0")
    lngFontSize = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", 9)
    vfgThis.FontSize = lngFontSize
    Dim panContent As Pane, panNew As Pane
    mlngPatiId = -1: mlngPageId = -1
    
    Set mfrmContent = New frmDockEPRContent
    Set panContent = dkpMan.CreatePane(conPane_Content, 400, 300, DockBottomOf, Nothing)
    panContent.Title = "病历内容"
    panContent.Options = PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmNew = New frmDockEPRNew
    Set panNew = dkpMan.CreatePane(conPane_New, 200, 400, DockRightOf, Nothing)
    panNew.Title = "新增病历"
    panNew.Options = PaneNoFloatable Or PaneNoHideable
    panNew.Close
    
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Me.dkpMan.Options.ThemedFloatingFrames = True
    mlngVersion = 1  '默认为第1版

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCols As String, i As Long
    If vfgThis.Cols = 19 Then
        For i = 0 To 18
            strCols = strCols & IIf(i = 0, "", ";") & vfgThis.ColWidth(i)
        Next
    Else
        strCols = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", _
            "270;0;0;2200;800;1600;800;0;0;0;3000;0;0;0;1200;0;800;1600;0")
    End If
    mstrColWidthConfig = strCols
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", mstrColWidthConfig
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", vfgThis.FontSize
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmNew Is Nothing Then Unload mfrmNew
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    Set mfrmContent = Nothing
    Set mfrmNew = Nothing
    Set mfrmMonitor = Nothing
    Set mobjDoc = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub mfrmNew_NewClick(ByVal FileId As Long, ByVal babyNum As Long)
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim frmThis As Form, bFinded As Boolean

        
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If Not gobjPlugIn.AddEMRBefore(glngSys, 1250, mlngPatiId, mlngPageId, FileId) Then Exit Sub
        Err.Clear: On Error GoTo 0
    End If
    
    If gstrPrivsEpr = ";;" Then
        MsgBox "您不具备病历编辑相应权限，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If

    If mblnMoved Then
        MsgBox "该病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errHand
    strSQL = "Select 保留 From 病历文件列表 Where  ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, FileId)
    If rs!保留 < 0 Then
        '特殊病历，手术麻醉单
        Exit Sub
    ElseIf rs!保留 = 2 Then '表格式编辑器
        If Not mObjTabEpr Is Nothing Then
            bFinded = mObjTabEpr.Showfrm(FileId, mlngPatiId, mlngPageId, cprPF_门诊, mlngDeptId)
        End If
        If Not bFinded Then
            Set mObjTabEpr = New cTableEPR
            mObjTabEpr.InitOpenEPR Me, cprEM_新增, cprET_单病历编辑, FileId, True, 0, cprPF_门诊, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, , InStr(gstrPrivsEpr, "病历打印") > 0, Val(gstrESign)
        End If
    ElseIf rs!保留 = 4 Then '传染病报告卡编辑器
'        已独立页面
    Else
        For Each frmThis In Forms
            If TypeName(frmThis) = "frmMain" Then
                With frmThis.Document
                    If .EPRFileInfo.ID = FileId And .EPRPatiRecInfo.病人ID = mlngPatiId _
                        And .EPRPatiRecInfo.病人来源 = cprPF_门诊 And .EPRPatiRecInfo.主页ID = mlngPageId _
                        And .EPRPatiRecInfo.科室ID = mlngDeptId And frmThis.ChildMode = False Then
                        frmThis.Show
                        bFinded = True
                    End If
                End With
            End If
        Next
        If bFinded = False Then
            Set mobjDoc = New cEPRDocument
            mobjDoc.InitEPRDoc cprEM_新增, cprET_单病历编辑, FileId, cprPF_门诊, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
            mobjDoc.ShowEPREditor Me
            Me.dkpMan.Panes(conPane_New).Close
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjDoc_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cbrControl As CommandBarControl
    vfgThis.Row = IIf(vfgThis.MouseRow = -1, vfgThis.Rows - 1, vfgThis.MouseRow)
    If Button = vbRightButton And Not mcbsThis Is Nothing Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        
        Set Popup = mcbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"):  cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngRecordId As Long, blnRTFFile As Boolean, byteEdit As Byte
    
    Me.dkpMan.Panes(conPane_New).Close
    Err = 0: On Error Resume Next
    With Me.vfgThis
        If .Cols < mCol.ID + 1 Then Exit Sub
        lngRecordId = Val(.TextMatrix(.Row, mCol.ID))
        byteEdit = Val(.TextMatrix(.Row, mCol.编辑方式))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag = "" And (Val(Me.vfgThis.Tag) <> Me.vfgThis.Row) Then
        Call mfrmContent.zlRefresh(lngRecordId, IIf(mblnEdit = False, "", mstrPrivs), , mblnMoved, blnRTFFile, byteEdit, True)
        If blnRTFFile Then
            If dkpMan.Panes(conPane_Content).Closed = True Then Call dkpMan.Panes(conPane_Content).Select
        ElseIf dkpMan.Panes(conPane_Content).Selected = True Then
            dkpMan.Panes(conPane_Content).Close
        End If
        Me.vfgThis.Tag = Me.vfgThis.Row
    End If
End Sub

'------------------------------------------------------------
'以下为公共方法
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开(&O)…", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…", cbrControl.Index + 1)
        
        '放在导出为XML文件之后
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "列表打印(&T)", cbrControl.Index + 1): cbrControl.BeginGroup = True
    End With
    
    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "病历(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印"): cbrControl.STYLE = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
    End With

    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", cbrMenuBar.Index, False)
        cbrMenuBar.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "病历质量监测(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "病人病历检索(&S)")
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = cbsThis(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        'Set cbrControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印", cbrControl.Index + 1)
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With

    '命令的快键绑定
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '设置不常用命令
    '-----------------------------------------------------
    With cbsThis.Options
    End With
    
    '-----------------------------------------------------
    '当没有书写病历时，根据权限状态，显示增加窗格
    '-----------------------------------------------------
    If Val(Me.vfgThis.TextMatrix(Me.vfgThis.FixedRows, mCol.ID)) = 0 Then
        If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0) Then
            Me.dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(1, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strInfo As String, lFileId As Long, blnCanPrint As Boolean
Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_NewItem Or Control.ID = conMenu_Edit_Archive Or _
                        Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML) Then '已转储病人,修改,删除,新增,归档,打开,导出不允许操作
        MsgBox "该病人的本次就诊数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lFileId = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    bEditor = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.编辑方式))
    blnCanPrint = IIf(Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.完成时间)) = "", InStr(1, gstrPrivsEpr, "未签名打印") > 0, InStr(1, gstrPrivsEpr, "病历打印") > 0) And (Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)
    Select Case Control.ID
    Case conMenu_File_Open
        '病历阅读
        If bEditor = 0 Then
            Dim fViewDoc As New frmEPRView
            If EprPrinted(lFileId) And InStr(mstrPrivs, "取消打印") = 0 Then blnCanPrint = False ''已经打印过且没有取消打印权限,不允许重复打印
            fViewDoc.ShowMe Me, lFileId, , blnCanPrint, , mlngAdviceID
        ElseIf bEditor = 1 Then
            If Not mObjTabEprView Is Nothing Then
                bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_门诊, mlngDeptId)
            End If
            If Not bFinded Then
                mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, True, 0, cprPF_门诊, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, blnCanPrint, Val(gstrESign)
            End If
        ElseIf bEditor = 2 Then
'            传染病已独立页面
        End If
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview
        If EprPrinted(lFileId) And InStr(mstrPrivs, "取消打印") = 0 Then '已经打印过且没有取消打印权限,不允许重复打印
            MsgBox "当前病历已打印，不允许重复打印！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(True)
    Case conMenu_File_Print
        If EprPrinted(lFileId) And InStr(mstrPrivs, "取消打印") = 0 Then '已经打印过且没有取消打印权限,不允许重复打印
            MsgBox "当前病历已打印，不允许重复打印！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(False)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_ExportToXML
        '导出到XML文件
        Dim strF As String
        dlgThis.Filename = "病历_" & Me.vfgThis.TextMatrix(Me.vfgThis.Row, mCol.病历名称) & _
            "(" & Me.vfgThis.TextMatrix(Me.vfgThis.Row, mCol.ID) & "," & mlngVersion & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        On Error GoTo errHand
        strF = dlgThis.Filename
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If bEditor = 1 Then
            '表格式病历
            mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, False, 0, cprPF_门诊, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved
            If mObjTabEprView.zlExportXML(strF) Then
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument '普通住院病历
            DocXML.InitAndOpenEPR lFileId, mlngVersion, , True
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    Case conMenu_File_RowPrint
        Call zlRptPrint(1)
    Case conMenu_Edit_NewItem
        Me.dkpMan.Panes(conPane_New).Select
        Call mfrmNew.zlRefList(1, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
    Case conMenu_Edit_Modify
        If EprPrinted(lFileId) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
        If bEditor = 1 Then
            '表格式病历
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_门诊, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, True, 0, cprPF_门诊, _
                    mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(gstrPrivsEpr, "病历打印") > 0, Val(gstrESign)
            End If
        ElseIf bEditor = 0 Then
            For Each frmThis In Forms
                If frmThis.Name = "frmMain" Then
                    With frmThis.Document
                        On Error Resume Next
                        If .EPRPatiRecInfo.ID = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 1) And .EPRPatiRecInfo.病人ID = mlngPatiId _
                            And .EPRPatiRecInfo.病人来源 = cprPF_门诊 And .EPRPatiRecInfo.主页ID = mlngPageId _
                            And frmThis.ChildMode = False Then
                            frmThis.Show
                            bFinded = True
                        End If
                        If Err.Number <> 0 Then
                            Err.Clear
                            bFinded = True
                        End If
                    End With
                End If
            Next
            If bFinded = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历编辑, lFileId, cprPF_门诊, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
                mobjDoc.ShowEPREditor Me
            End If
        ElseIf bEditor = 2 Then
'            传染病已独立页面
        End If
    Case conMenu_Edit_Delete
        With Me.vfgThis
            strInfo = "真的删除这份“" & .TextMatrix(.Row, mCol.病历名称) & "”吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
            gstrSQL = "Zl_电子病历记录_Delete(" & .TextMatrix(.Row, mCol.ID) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
            
            RaiseEvent RequestRefresh
        End With
    Case conMenu_Edit_Archive
        With Me.vfgThis
            If Trim(.TextMatrix(.Row, mCol.归档人)) = "" Then
                strInfo = "真的将该份“" & .TextMatrix(.Row, mCol.病历名称) & "”归档吗？"
            Else
                strInfo = "真的撤消该份“" & .TextMatrix(.Row, mCol.病历名称) & "”的归档吗？"
            End If
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_电子病历记录_Archive(" & .TextMatrix(.Row, mCol.ID) & "," & IIf(Trim(.TextMatrix(.Row, mCol.归档人)) = "", 0, 1) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
        End With
    Case conMenu_Edit_NoPrint '取消打印标记
        Call PrintCancel(lFileId)
    Case conMenu_Tool_Monitor
        If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
        Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, 1, mlngDeptId, 1, 1)
    Case conMenu_Tool_Search
        Call frmEPRSearchMan.ShowSearchClinic(Me, mlngDeptId)
    Case conMenu_View_Refresh:  Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_SignVerify
        If bEditor = 0 Then
            Call VerifySignature(Me, lFileId, mblnMoved)
        Else '表格式病历，28未处理数字签名情况
            'call
        End If
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnTmp As Boolean
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    With Me.vfgThis
        Select Case Control.ID
        Case conMenu_File_Open, conMenu_File_Excel, conMenu_File_RowPrint
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
        Case conMenu_Edit_NoPrint
            Control.Enabled = InStr(mstrPrivs, "取消打印") > 0 And (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
            If Control.Enabled Then Control.Enabled = Trim(.TextMatrix(.Row, mCol.打印人)) <> ""
            If Control.Enabled Then Control.Enabled = mblnEdit
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0 And InStr(1, gstrPrivsEpr, "病历打印") > 0)
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)
            If Control.Enabled And (Control.ID = conMenu_File_Preview Or Control.ID = conMenu_File_ExportToXML) Then
                Control.Enabled = Val(.TextMatrix(.Row, mCol.编辑方式)) <> 2
            End If
        Case conMenu_Edit_NewItem
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0)
        Case conMenu_Edit_Modify
            If mblnDisease Then
                Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0)
            Else
                If Val(.TextMatrix(.Row, mCol.病历种类)) = 5 Then
                    Control.Enabled = (mbln传染病 And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0)
                Else
                    Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0)
                End If
                If Control.Enabled Then
                    blnTmp = (Val(.TextMatrix(.Row, mCol.处理状态)) <= 0) '已经进入后续处理的病历不能处理
                    If Not blnTmp Then
                        If Val(.TextMatrix(.Row, mCol.申报状态)) = 4 Or Val(.TextMatrix(.Row, mCol.申报状态)) = 5 Then
                            blnTmp = True
                        End If
                    End If
                    Control.Enabled = blnTmp
                End If
            End If
            
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.科室ID)))   '本科病历才可以改
            If Control.Enabled Then
                If Trim(.TextMatrix(.Row, mCol.完成时间)) = "" Then
                    Control.Enabled = (InStr(1, mstrPrivs, "他人病历") > 0 Or Trim(.TextMatrix(.Row, mCol.创建人)) = Trim(gstrUserName))
                ElseIf Trim(.TextMatrix(.Row, mCol.归档人)) = "" And Val(.TextMatrix(.Row, mCol.当前版本)) <= 1 And InStr(1, ",1,2,4,", Val(.TextMatrix(.Row, mCol.签名级别))) > 0 Then
                    Control.Enabled = (InStr(1, mstrPrivs, "他人病历") > 0 Or InStr(1, .TextMatrix(.Row, mCol.保存人), Trim(gstrUserName)) > 0)
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Edit_Delete
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0) And (mblnEdit And mlngPatiId > 0 And (InStr(1, mstrPrivs, "病历书写") > 0 Or InStr(1, mstrPrivs, "强制删除") > 0))
            If Control.Enabled And InStr(1, mstrPrivs, "强制删除") > 0 Then Exit Sub '具备强制删除权限，则不进行后续的判断
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.处理状态)) <= 0)  '已经进入后续处理的病历不能处理
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.科室ID)))   '本科病历才可以删
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.完成时间)) = "")        '未完成病历可以删
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "他人病历") > 0 Or Trim(.TextMatrix(.Row, mCol.创建人)) = Trim(gstrUserName))
        Case conMenu_Edit_Archive
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历归档") > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.处理状态)) <= 0)  '已经进入后续处理的病历不能处理
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.签名级别)) <> 0)         '当前版本已经签名完成才可以归档
            If Trim(.TextMatrix(.Row, mCol.归档人)) = "" Then
                Control.Caption = "归档": Control.Checked = False
            Else
                Control.Caption = "撤档": Control.Checked = True
            End If
        Case conMenu_Tool_Monitor
            Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "质量监测") > 0)
        Case conMenu_Tool_Search: Control.Enabled = mblnSearch
        Case conMenu_Tool_SignVerify
            Control.Enabled = Val(.TextMatrix(.Row, mCol.ID)) <> 0 And Trim(.TextMatrix(.Row, mCol.完成时间)) <> ""
       End Select
    End With
End Sub

Public Sub RefreshList()
    Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    RaiseEvent RequestRefresh
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-小(缺省)，1-大
Dim bytFontSize As Byte

    bytFontSize = Decode(bytSize, 0, 9, 1, 12, bytSize)
    Call mPublic.SetFontSize(Me, bytFontSize)
    Call mPublic.SetFontSize(mfrmNew, bytFontSize)
End Sub
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
                            Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long) As Long
    Dim lngCol As Long, lngRow As Long
    Dim strKind As String
    Dim rsTemp As New ADODB.Recordset
    Dim str传染病病历 As String
    Dim rs传染 As ADODB.Recordset
    Dim str种类 As String
    
    If mlngPatiId = lngPatiID And mlngPageId = lngPageId And blnForce = False Then Exit Function
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '提取是否本部门启用电子签名,科室变更或没取过时提取
        gstrESign = getPassESign(0, lngDeptId)
    End If
    mblnDisease = (GetPrivFunc(glngSys, 1249) <> "")   'true-启用了疾病报告模块;false-不启用疾病报告模块
    
    mlngPatiId = lngPatiID: mlngPageId = lngPageId: mlngAdviceID = lngAdviceID
    mlngDeptId = lngDeptId: mblnEdit = blnEdit: mblnMoved = blnMoved
    
    vsColumn.Visible = False
    Me.vfgThis.Tag = ""
    
    If mblnDisease Then
        str种类 = " r.病历种类 In (1,6) "
    Else
        str种类 = " (r.病历种类 In (1,6) or (r.病历种类=5 And r.编辑方式<>2)) "
    End If

    gstrSQL = "Select r.Id, r.病历种类, r.病历名称, r.创建人 As 创建人, To_Char(r.创建时间, 'yyyy-mm-dd hh24:mi') As 创建时间, r.保存人," & _
            "        To_Char(r.完成时间, 'yy-mm-dd hh24:mi') As 完成时间, r.最后版本 As 当前版本, r.签名级别," & _
            "        Decode(r.最后版本, 1, '', '修订：') || r.保存人 || '在' || To_Char(r.保存时间, 'yyyy-mm-dd hh24:mi') ||" & _
            "        Decode(Nvl(r.签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') As 当前情况, r.归档人, r.归档日期, r.科室id," & _
            "        d.名称 As 科室名, r.处理状态,r.打印人,To_Char(r.打印时间, 'yyyy-mm-dd hh24:mi') As 打印时间,Decode(r.编辑方式,2,Decode(r.病历种类,1,0,r.编辑方式),r.编辑方式) as 编辑方式,null as 申报状态" & _
            " From 电子病历记录 r, 部门表 d" & _
            " Where r.科室id = d.Id And r.病人来源 = 1 And " & str种类 & " And r.病人id = [1] And Nvl(r.主页id, 0) = [2]" & _
            " Order By r.病历种类, r.序号, r.创建时间"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    
    If Not mblnDisease Then
        gstrSQL = "Select a.处理状态,b.id,C.执行部门ID,C.执行状态 From 疾病申报记录 a,电子病历记录 b, 病人挂号记录 c  where a.文件id=b.id and b.病历种类=5" & vbNewLine & _
            "and b.病人id=c.病人id and b.主页id=c.id and  c.id=[1] and a.处理状态 in (4,5)"
        Set rs传染 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPageId)
        
        mbln传染病 = False
        If mblnEdit Then
            mbln传染病 = True
        ElseIf rs传染.RecordCount > 0 Then
            If Val(rs传染!执行部门ID) = lngDeptId And (Val(rs传染!执行状态) = 1 Or Val(rs传染!执行状态) = 2) Then
                mbln传染病 = True
            End If
        End If
        For lngRow = 1 To rs传染.RecordCount
            str传染病病历 = str传染病病历 & "," & rs传染!ID
            rs传染.MoveNext
        Next
    End If
    
    With Me.vfgThis
        .Clear
        Set .DataSource = rsTemp
        
        Dim T As Variant, i As Long
        On Error Resume Next
        T = Split(mstrColWidthConfig, ";")
        If UBound(T) < 18 Then
            mstrColWidthConfig = "270;0;0;2200;800;1600;800;0;0;0;3000;0;0;0;1200;0;800;1600;0;0"
        Else
            For i = 0 To 18
                .ColWidth(i) = T(i)
            Next
        End If
        
        If .FixedRows > 0 Then .ROWHEIGHT(.FixedRows - 1) = .RowHeightMin
        .MergeRow(0) = True
        For lngCol = .FixedCols To .Cols - 1
            .FixedAlignment(lngCol) = flexAlignCenterCenter
        Next
        strKind = ""
        For lngRow = .FixedRows To .Rows - 1
            If strKind <> .TextMatrix(lngRow, mCol.病历种类) Then
                '画分类线条
                If strKind <> "" Then .CellBorderRange lngRow, 0, lngRow, .Cols - 1, RGB(0, 0, 255), 0, 1, 0, 0, 0, 0
                strKind = .TextMatrix(lngRow, mCol.病历种类)
            End If
            If Val(.TextMatrix(lngRow, mCol.处理状态)) > 0 Then
                Set .Cell(flexcpPicture, lngRow, mCol.标志) = imgThis.ListImages("转交").Picture
            ElseIf Trim(.TextMatrix(lngRow, mCol.归档人)) <> "" Then
                Set .Cell(flexcpPicture, lngRow, mCol.标志) = imgThis.ListImages("归档").Picture
            ElseIf Val(.TextMatrix(lngRow, mCol.当前版本)) <= 1 Then
                Set .Cell(flexcpPicture, lngRow, mCol.标志) = imgThis.ListImages("书写").Picture
            Else
                Set .Cell(flexcpPicture, lngRow, mCol.标志) = imgThis.ListImages("修订").Picture
            End If
            If .ROWHEIGHT(lngRow) < .RowHeightMin Then .ROWHEIGHT(lngRow) = .RowHeightMin
            If str传染病病历 <> "" Then
                If InStr(str传染病病历 & ",", "," & Val(.TextMatrix(lngRow, mCol.ID)) & ",") > 0 Then
                    rs传染.Filter = "id=" & Val(.TextMatrix(lngRow, mCol.ID))
                    If Not rs传染.EOF Then
                        .TextMatrix(lngRow, mCol.申报状态) = Val(rs传染!处理状态 & "")
                    End If
                End If
            End If
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        vfgThis.Tag = -1: .Row = 0 '促使vfgthis不选中任何行，不显示任何内容，仅当选中某行时才刷新
        If rsTemp.RecordCount = 1 Then
            .Row = 1
        End If
        Call vfgThis_RowColChange
    End With
    
    Call InitColumnSelect '列选择器
    
    '-----------------------------------------------------
    '当没有书写病历时，根据权限状态，显示增加窗格
    '-----------------------------------------------------
    If Val(Me.vfgThis.TextMatrix(Me.vfgThis.FixedRows, mCol.ID)) = 0 Then
        If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0) Then
            Me.dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(1, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlOpenDefaultEPR(Optional ByVal bytKind As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：自动新增一份缺省的门诊或急诊病历
    '参数：bytKind=1表示初诊病历;=2表示是急诊病历;3=复诊
    '说明：如果当前病人已有病历，则不需要自动增加
    '******************************************************************************************************************
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0) Then
        With vfgThis
            If .Rows = 2 And Val(.TextMatrix(1, mCol.ID)) = 0 Then
                zlOpenDefaultEPR = mfrmNew.zlOpenDefaultEPR(bytKind)
            End If
        End With
    End If
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode=1 打印;2 预览;3 输出到EXCEL
    '       strSubhead，打印的副标题
    '-------------------------------------------------
Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
Dim rsTemp As New ADODB.Recordset
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = "病历书写情况"
    
    '---------------------------------------------
    '获得基本信息
    Dim strSubhead1 As String, strSubhead2 As String
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select r.门诊号, r.姓名, r.性别, r.年龄, r.登记时间, r.No From 病人挂号记录 r Where r.Id =[1] and r.记录性质=1  and r.记录状态=1"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "病人挂号记录", "H病人挂号记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPageId)
    If Not rsTemp.EOF Then
        strSubhead1 = "门诊号:" & rsTemp!门诊号 & "  姓名:" & rsTemp!姓名 & "  性别:" & rsTemp!性别
        strSubhead2 = "日期:" & Format(rsTemp!登记时间, "yyyy-MM-dd") & "(No:" & rsTemp!NO & ")"
    Else
        strSubhead1 = "": strSubhead2 = ""
    End If
    
    Err = 0: On Error GoTo 0
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strSubhead1)
    Call objAppRow.Add(strSubhead2)
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'## 功能：  正式病历预览及打印
'##s
'## 参数：  blnPreview  :是否是预览模式
'################################################################################################################
Private Sub zlEPRPrint(blnPreview As Boolean)
Dim lFileId As Long, strPrintName As String
    
    lFileId = CLng(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    strPrintName = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", "")
    Select Case Val(vfgThis.TextMatrix(vfgThis.Row, mCol.编辑方式))
        Case 0
            Set mfrmPrintPreview = New frmPrintPreview
            mfrmPrintPreview.DoMultiDocPreview Me, cpr门诊病历, , , vfgThis.Cell(flexcpText, vfgThis.Row, mCol.病历种类) _
                            , , lFileId, Not blnPreview, , , mblnMoved, mlngAdviceID, strPrintName, IIf(InStr(mstrPrivs, "取消打印") > 0, 0, 1) '没有"取消打印"权限不允许重复打印，不允许调整打印份数
            Unload mfrmPrintPreview 'ByZT:窗体Load了未显示，没有人为关闭的情况下VB不会自动Unload
            Set mfrmPrintPreview = Nothing
        Case 1
            mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, False, 0, cprPF_门诊, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(gstrPrivsEpr, "病历打印") > 0
            mObjTabEprView.zlPrintDoc Me, blnPreview, strPrintName
        Case 2
'            传染病已独立页面
    End Select
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", strPrintName
End Sub

Private Function EprPrinted(ByVal lngRecordId As Long, Optional strPrintInfo As String) As Boolean
'检查当前病历记录是否已经打印过
Dim rsTemp As ADODB.Recordset
On Error GoTo errHand
    '因要求保留电子病历记录（打印人，打印时间），所以历史数据不转移，记录进行联合查询
    gstrSQL = "Select 打印人, 打印时间 From 电子病历打印 Where 文件id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select 打印人, 打印时间 From 电子病历记录 Where ID = [1] And 打印人 is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.EOF Then Exit Function
    
    Do Until rsTemp.EOF
        strPrintInfo = strPrintInfo & vbCrLf & "打印人：" & Rpad(rsTemp!打印人, 5) & "打印时间：" & Format(rsTemp!打印时间, "yyyy-MM-dd hh:mm")
        rsTemp.MoveNext
    Loop
    strPrintInfo = Mid(strPrintInfo, 3)
    EprPrinted = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function EprIsCommit() As String
'以|分隔方式返回,状态为0 不允许 1 允许，分别控制 新增|删除|撤档

Dim rsTemp As ADODB.Recordset, intNew As Integer, intDel As Integer, intMod As Integer
    gstrSQL = "Select 病案状态 From 病案主页 Where 病人id = [1] And 主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)

    Select Case NVL(rsTemp!病案状态, 0)
        Case 0
            intNew = 1: intDel = 1: intMod = 1
        Case 1 '等待审查
            intNew = 0: intDel = 0: intMod = 0
        Case 2 '拒绝审查
            intNew = 0: intDel = 0: intMod = 1
        Case 3 '正在审查
            intNew = 0: intDel = 0: intMod = 0
        Case 4 '审查反馈
            intNew = 0: intDel = 0: intMod = 1
        Case 5 '审查归档
            intNew = 0: intDel = 0: intMod = 0
        Case 6 '审查整改
            intNew = 0: intDel = 0: intMod = 1
        Case 13 '正在抽查
            intNew = 1: intDel = 1: intMod = 1
        Case 14 '抽查反馈
            intNew = 1: intDel = 1: intMod = 1
        Case 16 '抽查整改
            intNew = 1: intDel = 1: intMod = 1
        Case Else
            intNew = 0: intDel = 0: intMod = 0
    End Select
    EprIsCommit = CStr(intNew) & "|" & CStr(intDel) & "|" & CStr(intMod)
End Function
Private Sub PrintCancel(ByVal lngRecordId As Long)
'取消标记打印
On Error GoTo errHand
    gstrSQL = "Zl_电子病历打印_Cancel(" & lngRecordId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    vfgThis.Cell(flexcpText, vfgThis.Row, mCol.打印人) = ""
    vfgThis.Cell(flexcpText, vfgThis.Row, mCol.打印时间) = ""
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Public Function GetFormOperation() As String
'记录界面选定信息，因为工作站在切换页卡时是释放了对象，换回来时重新初始化刷新的。
    GetFormOperation = ""
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'恢复界面选定信息，工作站在刷新之前调用
End Sub
