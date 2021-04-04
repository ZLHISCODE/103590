VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInEPRs 
   BorderStyle     =   0  'None
   Caption         =   "住院病历记录"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsfFeedback 
      Height          =   1335
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   6255
      _cx             =   11033
      _cy             =   2355
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
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   135
      ScaleHeight     =   3120
      ScaleWidth      =   8145
      TabIndex        =   0
      Top             =   195
      Width           =   8145
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockInEPRs.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3480
         Left            =   735
         TabIndex        =   1
         Top             =   165
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
         FormatString    =   $"frmDockInEPRs.frx":054E
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
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   3105
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   8070
         _cx             =   14235
         _cy             =   5477
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
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
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
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFEBD7&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6855
            Picture         =   "frmDockInEPRs.frx":059C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   4
            Top             =   255
            Width           =   250
         End
         Begin MSComctlLib.ImageList imgThis 
            Left            =   0
            Top             =   1125
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":6DEE
                  Key             =   "书写"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":7388
                  Key             =   "修订"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":7922
                  Key             =   "归档"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":7EBC
                  Key             =   "转交"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInEPRs.frx":8256
                  Key             =   "打印"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   90
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   15
      Top             =   705
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDockInEPRs.frx":EAB8
      Left            =   720
      Top             =   4785
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockInEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------
'窗体常量
'-----------------------------------------------------
Private Enum mCol
    标志 = 0: 病人科室: 页面名称: 病历名称: 创建人: 创建时间: 保存人: 完成时间: 当前版本: 签名级别: 当前情况: 归档人: 归档日期: 科室ID: 科室名: 保留: 处理状态:  新组: ID: 病历种类: 页面编号: 编辑方式: 打印: 婴儿: 申报状态: 反馈记录
End Enum

Const conDefColWidth = "270;0;1200;1600;800;1600;800;1600;500;0;3300;0;0;0;1200;0;0;0;0;0;0;0;0;0;0;0"
Const conPane_List = 1
Const conPane_Content = 2
Const conPane_New = 3
Const mlngModul = 1251
Private mstrColWidthConfig As String
Private mlngfolding As Long
'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)       '继承文档对象的“点击诊断参考事件”
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者对本程序(1250)的权限串
Private mblnSearch As Boolean   '当前使用者是否具备病历检索(1273)权
Private mlngPatiId As Long      '病人id
Private mlngPageId As Long      '主页id
Private mlngDeptId As Long      '当前操作科室id，不一定是当前病人科室
Private mblnEdit As Boolean     '是否允许操作，通常由上级程序根据当前操作科室是否当前病人科室决定。
Private mblnMoved As Boolean    '是否数据已经转储
Private mlngAdviceID As Long    '医嘱ID
Private mintState As Integer    '见clsDockInEPR
Private mblnInsideTools As Boolean '自建工具条
Private mstrPhysicians  As String '病人三级医师名字串
Private mblnAllowDelete As Boolean '是否允许删除
Private mblnShowFinal As Boolean '显示最终版本
Private mlngCurId As Long

Private WithEvents mfrmNew As frmDockEPRNew
Attribute mfrmNew.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmDockInContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mfrmMonitor As New frmDockEPRMonitor
Private mfrmTipInfo As New frmTipInfo

Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR            '表格式病历编辑器
Attribute mObjTabEpr.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mcbsThis As Object          'CommandBar控件
Private mlngVersion As Long         '选中的文件版本号
Private mblnDisease As Boolean      '是否拥有了1249模块的权限


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars Control
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If cbrMain.Count < 1 Then Exit Sub
    zlUpdateCommandBars Control
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Select Case Pane.ID
    Case conPane_New
        Select Case Action
        Case PaneActionClosing, PaneActionClosed: Cancel = False
        Case Else: Cancel = True
        End Select
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_List
            Item.Handle = picList.hwnd
        Case conPane_Content
            Item.Handle = mfrmContent.hwnd
        Case conPane_New
            Item.Handle = mfrmNew.hwnd
    End Select
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
    Call Event_AfterPrinted(lngRecordId)
End Sub
Public Sub Event_AfterPrinted(lngRecordId As Long)
Dim i As Integer
    For i = 1 To vfgThis.Rows - 1
        If vfgThis.TextMatrix(i, mCol.ID) = lngRecordId Then
            vfgThis.Cell(flexcpData, i, mCol.当前情况) = ""
            vfgThis.Cell(flexcpText, i, mCol.打印) = gstrUserName
            Set vfgThis.Cell(flexcpPicture, i, mCol.页面名称) = imgThis.ListImages("打印").Picture
            Exit For
        End If
    Next
End Sub
Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'显示指定病历列表行的历史签名记录
Dim strTipInfo As String, lngRow As Long, strPrint As String
    If picInfo.Visible = False Then Exit Sub
    lngRow = vfgThis.MouseRow
    If lngRow <= 0 Then Exit Sub
    
    strTipInfo = vfgThis.Cell(flexcpData, lngRow, mCol.当前情况)
    
    If strTipInfo = "" Then '如果没有获取过，则立即获取并记录在列表中
        strTipInfo = GetEprSign(vfgThis.TextMatrix(lngRow, mCol.ID))   '提取签名
        Call EprPrinted(vfgThis.TextMatrix(lngRow, mCol.ID), strPrint) '提取打印记录
        strTipInfo = "由 " & Rpad(vfgThis.TextMatrix(lngRow, mCol.创建人), 8) & _
                     "于 " & Rpad(vfgThis.TextMatrix(lngRow, mCol.创建时间), 19) & " 创建" & vbCrLf & strTipInfo
        strTipInfo = strTipInfo & vbCrLf & strPrint
        vfgThis.Cell(flexcpData, lngRow, mCol.当前情况) = strTipInfo
    End If
    
    mfrmTipInfo.ShowTipInfo picInfo.hwnd, strTipInfo, True
End Sub
Private Sub piclist_Resize()
On Error Resume Next
    With vfgThis
        .Top = 0: .Left = 0
        .Width = picList.ScaleWidth: .Height = picList.ScaleHeight
    End With

    fraColSel.Move Me.vfgThis.Left + 50, Me.vfgThis.Top + 50
    fraColSel.ZOrder 0
    vsColumn.Move fraColSel.Left, fraColSel.Top + fraColSel.Height
    vsColumn.ZOrder 0
Err.Clear
End Sub

Private Sub vfgThis_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
If picInfo.Visible Then
    picInfo.Move vfgThis.Cell(flexcpLeft, NewTopRow, mCol.当前情况) + vfgThis.Cell(flexcpWidth, NewTopRow, mCol.当前情况) - picInfo.Width - 30
End If
End Sub

Private Sub vfgThis_Click()
    Dim lngMouseRow As Long, lngMouseCol As Long, lngWidth As Long, i As Long
    With vfgThis
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                If DisplayContent(Val(.TextMatrix(lngMouseRow, mCol.ID))) Then
                    With vsfFeedback
                        .Left = vfgThis.Left + vfgThis.Width - .Width
                        .Top = vfgThis.Top + 300 * (lngMouseRow + 1)
                        .ZOrder
                        .Visible = True
                        .SetFocus
                    End With
                End If
            Else
                vsfFeedback.Visible = False
            End If
        End If
    End With
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

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vfgThis.MouseRow = -1 And Me.Tag = "" Then
        vfgThis.Row = vfgThis.Rows - 1
    End If
End Sub

Private Sub vfgThis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, lngRow As Long
    lngCol = vfgThis.MouseCol: lngRow = vfgThis.MouseRow
    If lngRow <= 0 Then picInfo.Visible = False: Exit Sub
    
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl.Name <> "vfgThis" Then
            vfgThis.SetFocus
        Else
            vfgThis.SetFocus
        End If
    Else
        vfgThis.SetFocus
    End If
    
    If Val(vfgThis.TextMatrix(lngRow, mCol.ID)) <> 0 Then
        If Val(picInfo.Tag) = lngRow And picInfo.Visible Then Exit Sub
        picInfo.Tag = lngRow
        picInfo.Move vfgThis.Cell(flexcpLeft, lngRow, mCol.当前情况) + vfgThis.Cell(flexcpWidth, lngRow, mCol.当前情况) - picInfo.Width - 30, vfgThis.Cell(flexcpTop, lngRow, mCol.当前情况) + 15
        If vfgThis.RowSel = lngRow Then
            picInfo.BackColor = vfgThis.BackColorSel
        Else
            picInfo.BackColor = &H80000005
        End If
        picInfo.Visible = True
    Else
        picInfo.Visible = False
    End If
    If lngRow >= 0 And lngRow < vfgThis.Rows And lngCol >= 0 And lngCol < vfgThis.Cols Then
        If vfgThis.Cell(flexcpFontUnderline, lngRow, lngCol) = True Then
            vfgThis.MousePointer = 54
        Else
            vfgThis.MousePointer = 0
            If vsfFeedback.Visible Then vsfFeedback.Visible = False
        End If
    Else
        If vsfFeedback.Visible Then vsfFeedback.Visible = False
    End If
End Sub

Private Sub vfgThis_SelChange()
    If picInfo.Visible Then
        picInfo.BackColor = vfgThis.BackColorSel
    End If
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim lngCol As Long, T As Variant, i As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            T = Split(conDefColWidth, ";")
            vfgThis.ColWidth(lngCol) = T(lngCol)
            vfgThis.ColHidden(lngCol) = False
        Else
            vfgThis.ColWidth(lngCol) = 0
            vfgThis.ColHidden(lngCol) = True
        End If
    End If
    Dim strCols As String
    For i = 0 To 19
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

Private Sub vsColumn_LostFocus()
    On Error Resume Next
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Dim strCols As String, i As Long
    If vfgThis.Cols = UBound(Split(conDefColWidth, ";")) + 1 Then
        For i = 0 To vfgThis.Cols - 1
            If i = mCol.反馈记录 Then vfgThis.ColWidth(i) = 0
            strCols = strCols & IIf(i = 0, "", ";") & vfgThis.ColWidth(i)
        Next
    Else
        strCols = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", conDefColWidth)
    End If
    mstrColWidthConfig = strCols
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", mstrColWidthConfig
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", vfgThis.FontSize
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "ShowHistory", IIf(mblnShowFinal, "True", "False")
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmNew Is Nothing Then Unload mfrmNew
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    If Not mfrmPrintPreview Is Nothing Then Unload mfrmPrintPreview
    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
    Set mfrmContent = Nothing
    Set mfrmNew = Nothing
    Set mfrmMonitor = Nothing
    Set mobjDoc = Nothing
    Set mfrmPrintPreview = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mfrmTipInfo = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub Form_Load()
Dim panList As Pane, panContent As Pane, panNew As Pane, lngFontSize As Long
    mlngPatiId = -1: mlngPageId = -1
    mblnShowFinal = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "ShowHistory", "True")
    mstrColWidthConfig = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "CWidthConfig", conDefColWidth)
    lngFontSize = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgThis.Name, "FontSize", 9)
    vfgThis.FontSize = lngFontSize
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "基本") > 0)
    mlngfolding = zlDatabase.GetPara("共享病历折叠起始行数", glngSys, mlngModul, "6")
    mstrPrivs = GetPrivFunc(glngSys, 1251)
    
    Set panList = dkpMan.CreatePane(conPane_List, 200, 300, DockTopOf, Nothing)
    panList.Title = "病历列表"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmContent = New frmDockInContent
    Set panContent = dkpMan.CreatePane(conPane_Content, 200, 300, DockBottomOf, Nothing)
    panContent.Title = "病历内容"
    panContent.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set mfrmNew = New frmDockEPRNew
    Set panNew = dkpMan.CreatePane(conPane_New, 100, 400, DockRightOf, Nothing)
    panNew.Title = "新增病历"
    panNew.Options = PaneNoFloatable Or PaneNoHideable
    Set mObjTabEprView = New cTableEPR
    Call mObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDBUser)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With Me.cbrMain
'        .VisualTheme = xtpThemeOfficeXP
        .VisualTheme = xtpThemeOffice2003
        Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Visible = False
        .EnableCustomization False
        With .Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '放在VisualTheme后有效
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
        End With
    End With
    
    Me.dkpMan.SetCommandBars Me.cbrMain
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    mlngVersion = 1  '默认为第1版
End Sub

Private Sub mfrmNew_NewClick(ByVal FileId As Long, ByVal babyNum As Long)
    Dim rs As New ADODB.Recordset, rt As RECT
    Dim strFileName As String, blnResult As Boolean
    Dim frmThis As Form, bFinded As Boolean
    
    If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
    
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If Not gobjPlugIn.AddEMRBefore(glngSys, mlngModul, mlngPatiId, mlngPageId, FileId) Then Exit Sub
        Err.Clear: On Error GoTo 0
    End If
    
    If UserNewEMR Then
        MsgBox "本科室已经开始使用新病历系统，请使用新病历系统书写病历。", vbInformation, gstrSysName
        Exit Sub
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

    If Split(EprIsCommit, "|")(0) = 0 Then
        MsgBox "该病人病案已提交审查，不能新增病历，请取消审查后再试！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TimeLimitOut Then Exit Sub
        
    gstrSQL = "Select 保留,名称 From 病历文件列表 Where  ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
    If rs!保留 < 0 Then
        '特殊病历，手术麻醉单
        Exit Sub
    ElseIf rs!保留 = 2 Then '表格式编辑器
        If Not mObjTabEpr Is Nothing Then
            bFinded = mObjTabEpr.Showfrm(FileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
        End If
        If Not bFinded Then
            Set mObjTabEpr = New cTableEPR
            mObjTabEpr.InitOpenEPR Me, cprEM_新增, cprET_单病历编辑, FileId, True, 0, cprPF_住院, mlngPatiId, mlngPageId, babyNum, mlngDeptId, mlngAdviceID, mstrPrivs, , InStr(gstrPrivsEpr, "病历打印") > 0, Val(gstrESign)
        End If
    ElseIf rs!保留 = 4 Then '传染病报告卡编辑器
'        传染病已独立页面
    Else                    '电子病历RichEpr
        If InStr(rs!名称, "请会诊记录") > 0 Or InStr(rs!名称, "会诊邀请") > 0 Or InStr(rs!名称, "申请会诊记录") > 0 Then '需要对医嘱进行核选
            gstrSQL = "Select a.Id, b.名称 医嘱, c.名称 执行科室, a.开嘱医生, To_Char(a.开始执行时间, 'yyyy-mm-dd hh24:mi:ss') 开始时间," & vbNewLine & _
                        "       To_Char(a.开嘱时间, 'yyyy-mm-dd hh24:mi:ss') 开嘱时间" & vbNewLine & _
                        "From 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
                        "Where a.病人id = [1] And a.主页id = [2] and a.相关ID IS NULL And a.诊疗项目id = b.Id And b.类别 = 'Z' And b.操作类型 = '7' And a.执行科室id = c.Id(+) And" & vbNewLine & _
                        "      Not Exists (Select 1 From 病人医嘱报告 C Where c.医嘱id = a.Id)"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "会诊对应医嘱", mlngPatiId, mlngPageId)
            If rs.RecordCount > 1 Then
                Set rs = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "会诊对应医嘱", False, 1, "会诊对应医嘱有助于多个医嘱的会诊记录成对出现", False, False, False, 0, 0, 0, blnResult, True, True, mlngPatiId, mlngPageId)
                If blnResult = True Then '取消选择
                    MsgBox "会诊医嘱书写会诊记录，需要指定具体医嘱！", vbExclamation, gstrSysName: Exit Sub
                ElseIf rs.State = 1 Then
                    mlngAdviceID = rs!ID
                End If
            ElseIf rs.RecordCount = 1 Then
                mlngAdviceID = rs!ID
            Else '无数据，未开会诊医嘱，或已开会诊医嘱已经书写 请会诊记录 会诊邀请 申请会诊记录;对于医院无要求下达会诊医嘱才允许书写会诊记录予以通过
                'MsgBox "尚未新开会诊医嘱，或已经书写会诊医嘱相关病历，请检查！", vbExclamation, gstrSysName: Exit Sub
            End If
        ElseIf InStr(rs!名称, "会诊") > 0 And InStr(rs!名称, "请") = 0 Then '需要对 请会诊记录、会诊邀请、申请会诊记录 核选
            gstrSQL = "Select b.医嘱id As ID, f.名称 医嘱执行科室, a.病历名称, a.创建人, To_Char(a.创建时间, 'yyyy-mm-dd hh24:mi:ss') 创建时间" & vbNewLine & _
                        "From 电子病历记录 A, 病人医嘱报告 B, 病人医嘱记录 E, 部门表 F" & vbNewLine & _
                        "Where a.病人id = [1] And a.主页id = [2] And a.病历种类 = 2 And a.Id = b.病历id And" & vbNewLine & _
                        "      (Instr(a.病历名称, '请会诊记录') > 0 Or Instr(a.病历名称, '会诊邀请') > 0 Or Instr(a.病历名称, '申请会诊记录') > 0) And b.医嘱id = e.Id And" & vbNewLine & _
                        "      e.相关id Is Null And e.执行科室id = f.Id(+) And Not Exists" & vbNewLine & _
                        " (Select 1" & vbNewLine & _
                        "       From 电子病历记录 C, 病人医嘱报告 D" & vbNewLine & _
                        "       Where d.医嘱id = b.医嘱id And d.病历id = c.Id And c.病历种类 = 2 And Instr(c.病历名称, '会诊') > 0 And Instr(c.病历名称, '请') = 0)"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "会诊记录对应医嘱", mlngPatiId, mlngPageId)
            If rs.RecordCount > 1 Then
                Set rs = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "会诊记录对应邀请", False, 1, "会诊对应邀请有助于多个会诊记录时成对出现", False, False, False, 0, 0, 0, blnResult, True, True, mlngPatiId, mlngPageId)
                If blnResult = True Then '取消选择
                    MsgBox "会诊医嘱书写会诊记录，需要指定会诊邀请！", vbExclamation, gstrSysName: Exit Sub
                ElseIf rs.State = 1 Then
                    mlngAdviceID = rs!ID
                End If
            ElseIf rs.RecordCount = 1 Then
                mlngAdviceID = rs!ID
            Else '无数据 可能是改造前书写的请会诊记录无医嘱ID 或 请会诊记录已经书写会诊记录
                
            End If
        Else
            If zlDatabase.GetPara("共享病历必须先书写被共享病历", glngSys, mlngModul, "1") = 1 Then
                '判断共享文档是否已经书写过
                gstrSQL = "Select ID From 病历文件列表 Where 编号 <> NVL(页面,编号) And ID =[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
                If rs.EOF = False Then '是共享文档
                    gstrSQL = "Select M.ID,M.名称" & vbNewLine & _
                                "       From 病历文件列表 L, 病历文件列表 M" & vbNewLine & _
                                "       Where M.种类 = L.种类 And M.编号 = L.页面 And L.ID =[1]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
                    If rs.EOF Then MsgBox "该病历的共享病历定义失效，请联系系统管理员。", vbInformation, gstrSysName: Exit Sub
                    strFileName = rs!名称
                    gstrSQL = "Select ID" & vbNewLine & _
                                "From 电子病历记录" & vbNewLine & _
                                "Where 病人id = [1] And 主页id =[2] And 文件id+0 =[3]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, Val(rs!ID))
                    If rs.EOF Then
                        MsgBox "该病历的共享病历 [" & strFileName & "] 尚未书写，请检查。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        For Each frmThis In Forms
            If frmThis.Name = "frmMain" Then
                With frmThis.Document
                    If .EPRFileInfo.ID = FileId And .EPRPatiRecInfo.病人ID = mlngPatiId _
                        And .EPRPatiRecInfo.病人来源 = cprPF_住院 And .EPRPatiRecInfo.主页ID = mlngPageId _
                        And .EPRPatiRecInfo.科室ID = mlngDeptId And frmThis.ChildMode = False Then
                        frmThis.Show
                        bFinded = True
                        Exit For
                    End If
                End With
            End If
        Next
        
        If bFinded = False Then
            Set mobjDoc = New cEPRDocument
            mobjDoc.InitEPRDoc cprEM_新增, cprET_单病历编辑, FileId, cprPF_住院, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
            mobjDoc.EPRPatiRecInfo.婴儿 = babyNum
            mobjDoc.ShowEPREditor Me
        End If
    End If
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjDoc_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub

Private Sub vfgThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton And Not mcbsThis Is Nothing Then
        Dim Popup As CommandBar
        Dim ControlBar As CommandBarControl
        
        Set Popup = mcbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅(&U)"): ControlBar.BeginGroup = True
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Archive, "归档(&I)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_Sort, "病历排序(&S)"): ControlBar.BeginGroup = True
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印(&P)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set ControlBar = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "会诊医嘱调整关联(&C)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim byteEdit As Byte
    Dim ControlBar As Object
    On Error GoTo errHand
    Me.dkpMan.Panes(conPane_New).Close
    With Me.vfgThis
        If .Rows <= 1 Then Exit Sub
        If .Cols < mCol.ID + 1 Then Exit Sub
        mlngCurId = Val(.TextMatrix(.Row, mCol.ID))
        byteEdit = Val(.TextMatrix(.Row, mCol.编辑方式))
    End With
    If Not mcbsThis Is Nothing Then
        Set ControlBar = mcbsThis.FindControl(, conMenu_Edit_Delete, , True)
        zlUpdateCommandBars ControlBar
        If Not mcbsThis.FindControl(, conMenu_Edit_Delete, , True) Is Nothing Then
            mblnAllowDelete = mcbsThis.FindControl(, conMenu_Edit_Delete, , True).Enabled
        End If
    End If
    If Me.Tag = "" And (Val(Me.vfgThis.Tag) <> mlngCurId) Then
        Me.Tag = "Refresh" '不能刷得太快，否则报“拒绝权限”
        Call mfrmContent.zlRefresh(mlngCurId, IIf(mblnEdit = False, "", mstrPrivs), mblnMoved, mblnShowFinal, byteEdit, mblnAllowDelete)
        Me.Tag = ""
        Me.vfgThis.Tag = mlngCurId
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zlDefCommandBars(ByVal cbsThis As Object, ByVal blnInsideTools As Boolean)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

    mblnInsideTools = blnInsideTools
    Set mcbsThis = cbsThis
    Set mcbsThis.Icons = zlCommFun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportAll, "导出病人RTF文件(&A)…", cbrControl.Index + 1): cbrControl.ToolTipText = "导出该病人所有全文式病历为RTF"
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "病历排序(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
        Set cbrControl = .Add(xtpControlButton, ID_PATISIGNVerify, "患者签名验证(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "医嘱关联")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowHistory, "显示最终版本(&L)", cbrControl.Index + 1)
        cbrControl.BeginGroup = True
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "病历范文包(&C)")
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    cbrMain.DeleteAll
    If mblnInsideTools Then
        Set cbrToolBar = cbrMain.Add("工具栏", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ContextMenuPresent = False
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印"): cbrControl.STYLE = xtpButtonIconAndCaption
        End With
    Else
        Set cbrToolBar = cbsThis(2)
        For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
            If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
                Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
            End If
        Next
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增", cbrControl.Index + 1): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开", 1)
            .Item(cbrControl.Index + 1).BeginGroup = True
        End With
    End If
    
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
        .AddHiddenCommand conMenu_Edit_Archive
        .AddHiddenCommand conMenu_Edit_Untread
    End With
    
    '-----------------------------------------------------
    '根据权限状态，显示增加窗格
    '-----------------------------------------------------
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0) Then
        If zlDatabase.GetPara("自动显示新增面板", glngSys, mlngModul, "1") = 1 Then
            Me.dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(2, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strInfo As String
Dim rs As New ADODB.Recordset
Dim strSQL As String, lFileId As Long, blnCanPrint As Boolean
Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_Audit Or Control.ID = conMenu_Edit_Archive Or _
                        Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML Or conMenu_Edit_Compend) Then '已转储病人,修改,删除,审核,归档,打开不允许操作
        MsgBox "该病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If

    lFileId = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    bEditor = Val(vfgThis.TextMatrix(vfgThis.Row, mCol.编辑方式))
    Select Case Control.ID
    Case conMenu_File_Open
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        blnCanPrint = InStr(1, gstrPrivsEpr, "病历打印") > 0
        If blnCanPrint Then blnCanPrint = (Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.完成时间)) <> "" Or InStr(1, gstrPrivsEpr, "未签名打印") > 0)
        If blnCanPrint Then blnCanPrint = (Trim(vfgThis.TextMatrix(vfgThis.Row, mCol.归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)
        If blnCanPrint Then blnCanPrint = IIf(EprPrinted(lFileId), InStr(mstrPrivs, "取消打印") > 0, True) '已经打印过且没有取消打印权限,不允许重复打印
        If blnCanPrint Then blnCanPrint = (vfgThis.TextMatrix(vfgThis.Row, mCol.创建人) = gstrUserName Or InStr(1, mstrPrivs, "病历审阅") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0) '本人书写，有病历审阅权限,病人三级医师
        If bEditor = 0 Then
            Dim fViewDoc As New frmEPRView '查看该病历
            fViewDoc.ShowMe Me, lFileId, , blnCanPrint, , mlngAdviceID
        ElseIf bEditor = 1 Then
            If Not mObjTabEprView Is Nothing Then
                bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
            End If
            If Not bFinded Then
                mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历审核, lFileId, True, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, blnCanPrint, Val(gstrESign)
            End If
        ElseIf bEditor = 2 Then
'            传染病已独立页面
        End If
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)) And InStr(mstrPrivs, "取消打印") = 0 Then '已经打印过且没有取消打印权限,不允许重复打印
            MsgBox "当前病历已打印，不允许重复打印！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(True)
    Case conMenu_File_Print
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)) And InStr(mstrPrivs, "取消打印") = 0 Then '已经打印过且没有取消打印权限,不允许重复打印
            MsgBox "当前病历已打印，不允许重复打印！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(False)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_ExportAll: Call ExportAll
    Case conMenu_File_ExportToXML
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        '导出到XML文件
        Dim strF As String
        dlgThis.Filename = "病历_" & vfgThis.TextMatrix(vfgThis.Row, mCol.病历名称) & "(" & vfgThis.TextMatrix(vfgThis.Row, mCol.ID) & "," & mlngVersion & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        strF = dlgThis.Filename
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        With Me.vfgThis
            If Val(.TextMatrix(.Row, mCol.病历种类)) = 2 And Val(.TextMatrix(.Row, mCol.保留)) < 0 Then
                '特殊的住院病历
            ElseIf bEditor = 1 Then
                '表格式病历
                mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历审核, lFileId, False, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved
                If mObjTabEprView.zlExportXML(strF) Then
                    MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
                End If
            Else
                '普通住院病历
                Dim DocXML As New cEPRDocument
                DocXML.InitAndOpenEPR lFileId, mlngVersion, , True
                If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                    DoEvents
                    MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
                End If
            End If
        End With
    Case conMenu_File_RowPrint
        Call zlRptPrint(1)
    Case conMenu_Edit_NewItem
        Me.dkpMan.Panes(conPane_New).Select
        Call mfrmNew.zlRefList(2, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
    Case conMenu_Edit_Modify
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If TimeLimitOut Then Exit Sub '超过补录时限，不允许修改，新增，审核
        '单病历编辑模式
        With Me.vfgThis
            If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
            If Val(.TextMatrix(.Row, mCol.病历种类)) = 2 And Val(.TextMatrix(.Row, mCol.保留)) < 0 Then
                '特殊的住院病历即“麻醉记录”
            ElseIf bEditor = 1 Then
                '表格式病历
                If Not mObjTabEpr Is Nothing Then
                    bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
                End If
                If bFinded = False Then
                    Set mObjTabEpr = New cTableEPR
                    mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, True, 0, cprPF_住院, _
                        mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(1, gstrPrivsEpr, "病历打印") > 0, Val(gstrESign)
                End If
            ElseIf bEditor = 0 Then
                'RichEPR病历
                For Each frmThis In Forms
                    If frmThis.Name = "frmMain" Then
                        On Error Resume Next
                        If frmThis.Document.EPRPatiRecInfo.ID = .TextMatrix(.Row, mCol.ID) And frmThis.Document.EPRPatiRecInfo.病人ID = mlngPatiId _
                            And frmThis.Document.EPRPatiRecInfo.病人来源 = cprPF_住院 And frmThis.Document.EPRPatiRecInfo.主页ID = mlngPageId _
                            And frmThis.ChildMode = False Then
                            frmThis.Show
                            bFinded = True
                        End If
                        If Err.Number <> 0 Then
                            Err.Clear
                            bFinded = True
                        End If
                    End If
                Next
                If bFinded = False Then
                    Set mobjDoc = New cEPRDocument
                    mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历编辑, .TextMatrix(.Row, mCol.ID), cprPF_住院, mlngPatiId, CStr(mlngPageId), 0, mlngDeptId, mlngAdviceID
                    mobjDoc.ShowEPREditor Me, InStr(1, gstrPrivsEpr, "病历打印") > 0
                End If
            ElseIf bEditor = 2 Then
'                传染病已独立显示
            End If
        End With
    Case conMenu_Edit_Delete
        If Split(EprIsCommit, "|")(1) = 0 Then
            MsgBox "该病人病案已提交审查，不能删除，请取消审查后再试！", vbInformation, gstrSysName
            Exit Sub
        End If
    
        With Me.vfgThis
            If EprPrinted(.TextMatrix(.Row, mCol.ID)) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
            strInfo = "真的删除这份“" & .TextMatrix(.Row, mCol.病历名称) & "”吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_电子病历记录_Delete(" & .TextMatrix(.Row, mCol.ID) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
        End With
    Case conMenu_Edit_Audit
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If TimeLimitOut Then Exit Sub '超过补录时限，不允许修改，新增，审核
        If EprPrinted(lFileId) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
        
        If bEditor = 1 Then
            '表格式病历
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_单病历审核, lFileId, True, 0, cprPF_住院, _
                    mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, , Val(gstrESign)
            End If
        Else
            '单病历审核模式
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    On Error Resume Next
                    If frmAudit.Document.EPRPatiRecInfo.ID = lFileId _
                        And frmAudit.Document.EPRPatiRecInfo.病人来源 = cprPF_住院 And frmAudit.Document.EPRPatiRecInfo.病人ID = mlngPatiId _
                        And frmAudit.Document.EPRPatiRecInfo.主页ID = mlngPageId And frmAudit.ChildMode = False Then
                        frmAudit.Show
                        bFindedAudit = True
                    End If
                    If Err.Number <> 0 Then
                        Err.Clear
                        bFindedAudit = True
                    End If
                End If
            Next
            If bFindedAudit = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历审核, Me.vfgThis.TextMatrix(Me.vfgThis.Row, mCol.ID), cprPF_住院, mlngPatiId, CStr(mlngPageId), , mlngDeptId, mlngAdviceID
                mobjDoc.ShowEPREditor Me, InStr(1, gstrPrivsEpr, "病历打印") > 0
            End If
        End If
    Case conMenu_Edit_Archive
        Call EprArchive
    Case conMenu_Edit_Sort
        '排序
        Dim frmSort As New frmEPRSort
        If frmSort.ShowMe(Me, mlngPatiId, mlngPageId, vfgThis.TextMatrix(vfgThis.Row, mCol.病历种类), vfgThis.TextMatrix(vfgThis.Row, mCol.页面编号)) = True Then
            '刷新显示
            Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
        End If
    Case conMenu_Edit_NoPrint '取消打印标记
        If Split(EprIsCommit, "|")(0) = 0 Then
            MsgBox "该病人病案已提交审查，不能撤消打印，请取消审查后再试！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call PrintCancel(CLng(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)))
    Case conMenu_Tool_Monitor
        If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
        Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, 2, mlngDeptId, 1, mintState)
    Case conMenu_Tool_Search: Call frmEPRSearchMan.ShowSearchClinic(Me, mlngDeptId)
    Case conMenu_View_Refresh: Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    Case conMenu_Edit_Compend
        Call modelsApply
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_SignVerify
        If bEditor = 0 Then
            Call VerifySignature(Me, lFileId, mblnMoved)
        Else '表格式病历，28未处理数字签名情况
            'call
        End If
    Case ID_PATISIGNVerify
        Call VerifyPatiSign(Me, lFileId, mblnMoved)
    Case conMenu_View_ShowHistory
        mblnShowFinal = Not mblnShowFinal
        vfgThis.Tag = 0: Call RefreshList
    Case conMenu_Edit_ApplyModi
        Err = 0: On Error GoTo errHand
        Dim lngOrderId As Long
            gstrSQL = "Select a.Id, b.名称 医嘱, c.名称 执行科室, a.开嘱医生, To_Char(a.开始执行时间, 'yyyy-mm-dd hh24:mi:ss') 开始时间," & vbNewLine & _
                        "       To_Char(a.开嘱时间, 'yyyy-mm-dd hh24:mi:ss') 开嘱时间" & vbNewLine & _
                        "From 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
                        "Where a.病人id = [1] And a.主页id = [2] and a.相关ID IS NULL And a.诊疗项目id = b.Id And b.类别 = 'Z' And b.操作类型 = '7' And a.执行科室id = c.Id(+)"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "会诊对应医嘱", mlngPatiId, mlngPageId)
            If rs.RecordCount > 1 Then
                Set rs = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "会诊对应医嘱", False, 1, "会诊对应医嘱有助于多个医嘱的会诊记录成对出现", False, False, False, 0, 0, 0, bFinded, True, True, mlngPatiId, mlngPageId)
                If bFinded = True Then '取消选择
                    MsgBox "会诊医嘱书写会诊记录，需要指定具体医嘱！", vbExclamation, gstrSysName: Exit Sub
                ElseIf rs.State = 1 Then
                    lngOrderId = rs!ID
                End If
            ElseIf rs.RecordCount = 1 Then
                lngOrderId = rs!ID
            Else '无数据，未开会诊医嘱，或已开会诊医嘱已经书写 请会诊记录 会诊邀请 申请会诊记录;对于医院无要求下达会诊医嘱才允许书写会诊记录予以通过
                'MsgBox "尚未新开会诊医嘱，或已经书写会诊医嘱相关病历，请检查！", vbExclamation, gstrSysName:
                Exit Sub
            End If
        gstrSQL = "Zl_病历医嘱报告_Modify(" & lFileId & "," & lngOrderId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Err = 0: On Error GoTo 0
        Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Dim lngCount As Long, blnFinished As Boolean, lngMaxVersion As Long, eSignLevel As EPRSignLevelEnum
    Dim blnTmp As Boolean
    
    With Me.vfgThis
        Select Case Control.ID
        Case conMenu_File_Open, conMenu_File_Excel, conMenu_File_RowPrint
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
        Case conMenu_Edit_NoPrint
            Control.Enabled = InStr(mstrPrivs, "取消打印") > 0 And (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
            If Control.Enabled Then Control.Enabled = Trim(.TextMatrix(.Row, mCol.打印)) <> ""
            If Control.Enabled Then Control.Enabled = mblnEdit
        Case conMenu_Edit_NewItem
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0)
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.ID)) <> 0 And InStr(1, gstrPrivsEpr, "病历打印") > 0)
            If Control.Enabled Then Control.Enabled = IIf(Trim(.TextMatrix(.Row, mCol.完成时间)) = "", InStr(1, gstrPrivsEpr, "未签名打印") > 0, True)
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)
            If Control.Enabled Then Control.Enabled = (vfgThis.TextMatrix(vfgThis.Row, mCol.创建人) = gstrUserName Or InStr(1, mstrPrivs, "病历审阅") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0)   '本人书写，有病历审阅权限,病人三级医师
            If Control.ID = conMenu_File_Preview Or Control.ID = conMenu_File_ExportToXML Then
                If Control.Enabled Then Control.Enabled = Val(.TextMatrix(.Row, mCol.编辑方式)) <> 2
            End If
        Case conMenu_File_ExportAll
            Control.Enabled = (Val(.TextMatrix(1, mCol.ID)) <> 0 And InStr(1, gstrPrivsEpr, "病历打印") > 0)
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "病历审阅") > 0 Or InStr(1, mstrPhysicians, ";" & gstrUserName & ";") > 0)   '本人书写，有病历审阅权限,病人三级医师
        Case conMenu_Edit_Modify
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0)
            If Control.Enabled And Not mblnDisease Then
                blnTmp = (Val(.TextMatrix(.Row, mCol.处理状态)) <= 0)  '已经进入后续处理的病历不能处理
                If Not blnTmp Then
                    If Val(.TextMatrix(.Row, mCol.申报状态)) = 4 Or Val(.TextMatrix(.Row, mCol.申报状态)) = 5 Then
                        blnTmp = True
                    End If
                End If
                Control.Enabled = blnTmp
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
        Case conMenu_Edit_Audit
            Control.Enabled = (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历审阅") > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.处理状态)) <= 0)  '已经进入后续处理的病历不能处理
'            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.科室ID)))   '本科病历才可以审核
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.病历种类)) <> 2 Or Val(.TextMatrix(.Row, mCol.保留)) >= 0) '特殊的住院病历即“麻醉记录”不提供审阅
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.完成时间)) <> "")       '完成病历才可以审
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.归档人)) = "")          '未归档病历可以审
            If Control.Enabled Then Control.Enabled = Val(.TextMatrix(.Row, mCol.编辑方式)) <> 2           '传染病报告卡无需支持修订
        Case conMenu_Edit_Archive
            Control.Enabled = (mblnEdit And mlngPatiId > 0)
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.处理状态)) <= 0)  '已经进入后续处理的病历不能处理
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.签名级别)) <> 0)         '当前版本已经签名完成才可以归档
            If Trim(.TextMatrix(.Row, mCol.归档人)) = "" Then
                Control.Caption = "归档": Control.Checked = False
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "病历归档") > 0)
            Else
                Control.Caption = "撤档": Control.Checked = True
                If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "病历撤档") > 0)
            End If
        Case conMenu_Edit_Sort
            '排序（只有多文档共用页面时才可以调整序号）
            Control.Visible = True: Control.Enabled = True
            Control.Visible = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
        Case conMenu_Edit_Compend
            Control.Enabled = InStr(1, mstrPrivs, "病历书写") > 0
        Case conMenu_Tool_Monitor
            Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "质量监测") > 0)
        Case conMenu_Tool_Search: Control.Enabled = mblnSearch
        Case conMenu_Tool_SignVerify
            Control.Enabled = Val(.TextMatrix(.Row, mCol.ID)) <> 0 And Trim(.TextMatrix(.Row, mCol.完成时间)) <> ""
        Case conMenu_View_ShowHistory
            Control.Checked = mblnShowFinal
        Case conMenu_Edit_ApplyModi
            Control.Visible = (Val(.TextMatrix(.Row, mCol.ID)) <> 0)
            Control.Visible = InStr(.TextMatrix(.Row, mCol.病历名称), "会诊") > 0
            Control.Enabled = Trim(.TextMatrix(.Row, mCol.打印)) = ""
        End Select
    End With
End Sub

Public Sub RefreshList()
    Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
End Sub

Private Sub InitColumnSelect()
    On Error Resume Next
    '功能：根据原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vfgThis
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.病历名称, mCol.创建人, mCol.创建时间, mCol.保存人, mCol.完成时间, mCol.当前版本, mCol.当前情况, mCol.科室名, mCol.婴儿
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
Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-小(缺省)，1-大
Dim bytFontSize As Byte

    bytFontSize = Decode(bytSize, 0, 9, 1, 12, bytSize)
    Call mPublic.SetFontSize(Me, bytFontSize)
    Call mPublic.SetFontSize(mfrmNew, bytFontSize)
End Sub

Private Sub Initvfg()
    With vfgThis
        On Error Resume Next
        mfrmContent.Clear
        .Tag = ""
        .Clear
        .Rows = 1
        .Cols = 26
        .TextMatrix(0, mCol.标志) = "标志"
        .TextMatrix(0, mCol.病人科室) = "病人科室"
        .TextMatrix(0, mCol.页面名称) = "病历名称"
        .TextMatrix(0, mCol.病历名称) = "病历名称"
        .TextMatrix(0, mCol.创建人) = "创建人"
        .TextMatrix(0, mCol.创建时间) = "创建时间"
        .TextMatrix(0, mCol.保存人) = "保存人"
        .TextMatrix(0, mCol.完成时间) = "完成时间"
        .TextMatrix(0, mCol.当前版本) = "版本"
        .TextMatrix(0, mCol.签名级别) = "签名级别"
        .TextMatrix(0, mCol.当前情况) = "当前情况"
        .TextMatrix(0, mCol.归档人) = "归档人"
        .TextMatrix(0, mCol.归档日期) = "归档日期"
        .TextMatrix(0, mCol.科室ID) = "科室ID"
        .TextMatrix(0, mCol.科室名) = "科室名"
        .TextMatrix(0, mCol.保留) = "保留"
        .TextMatrix(0, mCol.处理状态) = "处理状态"
        .TextMatrix(0, mCol.新组) = "新组"
        .TextMatrix(0, mCol.ID) = "ID"
        .TextMatrix(0, mCol.病历种类) = "病历种类"
        .TextMatrix(0, mCol.页面编号) = "页面编号"
        .TextMatrix(0, mCol.编辑方式) = "编辑方式"
        .TextMatrix(0, mCol.打印) = "打印"
        .TextMatrix(0, mCol.申报状态) = "申报状态"
        .TextMatrix(0, mCol.婴儿) = "婴儿"
        .TextMatrix(0, mCol.反馈记录) = "反馈记录"
        .MergeCellsFixed = flexMergeFree
        .MergeCol(mCol.页面名称) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        
        Dim T As Variant, i As Long '调整列宽
        T = Split(mstrColWidthConfig, ";")
        If UBound(T) <> .Cols - 1 Then
            mstrColWidthConfig = conDefColWidth
            T = Split(mstrColWidthConfig, ";")
        End If
        For i = 0 To .Cols - 1
            .ColWidth(i) = T(i)
            .ColHidden(i) = (.ColWidth(i) = 0)
        Next
        
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = mCol.页面名称
        .SubtotalPosition = flexSTAbove
    End With
    
    vsfFeedback.Visible = False
End Sub
Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
                            Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long _
                            , Optional ByVal intState As Integer) As Long
    Dim lngCurId As Long    '当前病历记录ID
    Dim lngCurRow As Long   '刷新后选中行号，默认为0，不选中
    Dim rsTemp As New ADODB.Recordset, rsDis As ADODB.Recordset
    Dim lngCol As Long, lngRow As Long, i As Long
    Dim strKind As String, blnGroupTurnDept As Boolean
    Dim strReportIDs As String
    Dim str传染病病历 As String
    Dim rs传染 As ADODB.Recordset
    Dim str种类 As String
    
    If mlngPatiId = lngPatiID And mlngPageId = lngPageId And blnForce = False Then Exit Function
    lngCurId = IIf(mlngPatiId = lngPatiID, Val(vfgThis.TextMatrix(vfgThis.Row, mCol.ID)), 0) '当前病历刷新前选择行ID
    If lngCurId = 0 Then lngCurId = mlngCurId
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '提取是否本部门启用电子签名,科室变更或没取过时提取
        gstrESign = getPassESign(1, lngDeptId)
    End If
    
    mblnDisease = (GetPrivFunc(glngSys, 1249) <> "")   'true-启用了疾病报告模块;false-不启用疾病报告模块
    
    mlngDeptId = lngDeptId
    mblnEdit = blnEdit
    mblnMoved = blnMoved
    mlngPatiId = lngPatiID
    mlngPageId = lngPageId
    mlngAdviceID = lngAdviceID
    mintState = intState
    vsColumn.Visible = False
    mstrPhysicians = GetPhysicians '提取三级医生姓名
    blnGroupTurnDept = (zlDatabase.GetPara("转科后要求书写的共享病历另起一页打印", glngSys, mlngModul, 1) = 1)
    picInfo.Visible = False
    Call Initvfg
    
    If mblnDisease Then
        str种类 = "r.病历种类 In (2, 6)"
    Else
        str种类 = " r.病历种类 In (2, 5, 6) "
    End If
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select r.科室id 病人科室, Decode(b.名称, Null, r.病历名称, b.名称) As 页面, r.病历名称, r.创建人 As 创建人," & vbNewLine & _
                "       To_Char(r.创建时间, 'yyyy-mm-dd hh24:mi') As 创建时间, r.保存人, To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间," & vbNewLine & _
                "       r.最后版本 As 当前版本, r.签名级别," & vbNewLine & _
                "       Decode(r.最后版本, 1, '书写：', '修订：') || r.保存人 || '在' || To_Char(r.保存时间, 'yyyy-mm-dd hh24:mi') ||" & vbNewLine & _
                "        Decode(Nvl(r.签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') As 当前情况, r.归档人, r.归档日期, r.科室id, d.名称 As 科室名, c.保留, r.处理状态," & vbNewLine & _
                "       Decode(c.编号, b.编号, 1, 0) As 新组, r.Id, r.病历种类, b.编号, r.编辑方式, r.打印人 As 打印, r.婴儿, e.医嘱id" & vbNewLine & _
                "From 电子病历记录 R, 部门表 D, 病历文件列表 C, 病历页面格式 B, 病人医嘱报告 E" & vbNewLine & _
                "Where r.文件id + 0 = c.Id And r.病人来源 = 2 And " & str种类 & " And r.科室id = d.Id And r.病人id = [1] And r.主页id = [2] And" & vbNewLine & _
                "      r.Id = e.病历id(+) And c.种类 = b.种类 And c.页面 = b.编号 And Nvl(c.保留, 0) <> 4" & vbNewLine & _
                "Order By r.病历种类, b.编号, e.医嘱id, r.序号, r.创建时间"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    
    If Not mblnDisease Then
        gstrSQL = "Select a.处理状态,b.id From 疾病申报记录 a,电子病历记录 b, 病案主页 c  where a.文件id=b.id and b.病历种类=5" & vbNewLine & _
            "and b.病人id+0=c.病人id and b.主页id+0=c.主页id and a.姓名=c.姓名 and c.病人id=[1] and c.主页id=[2] and a.处理状态 in (4,5)"
        Set rs传染 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
        
        For lngRow = 1 To rs传染.RecordCount
            str传染病病历 = str传染病病历 & "," & rs传染!ID
            rs传染.MoveNext
        Next
    End If
    
    strKind = ""
    With vfgThis
        .ColWidth(mCol.申报状态) = 0
        .ColHidden(mCol.申报状态) = True
        .ColWidth(mCol.反馈记录) = 0
        .ColHidden(mCol.反馈记录) = True
        Do Until rsTemp.EOF
            .Rows = .Rows + 1
            .IsSubtotal(.Rows - 1) = True

            .TextMatrix(rsTemp.AbsolutePosition, mCol.病人科室) = NVL(rsTemp!病人科室)
            .Cell(flexcpData, rsTemp.AbsolutePosition, mCol.页面名称) = NVL(rsTemp!页面)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.页面名称) = NVL(rsTemp!页面)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.病历名称) = NVL(rsTemp!病历名称)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.创建人) = NVL(rsTemp!创建人)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.创建时间) = NVL(rsTemp!创建时间)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.保存人) = NVL(rsTemp!保存人)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.完成时间) = NVL(rsTemp!完成时间)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.当前版本) = NVL(rsTemp!当前版本)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.签名级别) = NVL(rsTemp!签名级别)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.当前情况) = NVL(rsTemp!当前情况)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.归档人) = NVL(rsTemp!归档人)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.归档日期) = NVL(rsTemp!归档日期)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.科室ID) = NVL(rsTemp!科室ID)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.科室名) = NVL(rsTemp!科室名)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.保留) = NVL(rsTemp!保留)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.处理状态) = NVL(rsTemp!处理状态)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.新组) = NVL(rsTemp!新组)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ID) = NVL(rsTemp!ID)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.病历种类) = NVL(rsTemp!病历种类)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.页面编号) = NVL(rsTemp!编号)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.编辑方式) = NVL(rsTemp!编辑方式)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.打印) = NVL(rsTemp!打印)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.婴儿) = NVL(rsTemp!婴儿)
            If str传染病病历 <> "" Then
                If InStr(str传染病病历 & ",", "," & rsTemp!ID & ",") > 0 Then
                    rs传染.Filter = "id=" & rsTemp!ID
                    If Not rs传染.EOF Then
                        .TextMatrix(rsTemp.AbsolutePosition, mCol.申报状态) = Val(rs传染!处理状态 & "")
                        .ColWidth(mCol.反馈记录) = 1200
                        .ColHidden(mCol.反馈记录) = False
                        .TextMatrix(rsTemp.AbsolutePosition, mCol.反馈记录) = "反馈记录"
                        .Cell(flexcpForeColor, rsTemp.AbsolutePosition, mCol.反馈记录, rsTemp.AbsolutePosition, mCol.反馈记录) = &HFF0000     '蓝色
                        .Cell(flexcpFontUnderline, rsTemp.AbsolutePosition, mCol.反馈记录, rsTemp.AbsolutePosition, mCol.反馈记录) = True
                        End If
                End If
            End If

            '页面名称相同分组，新组时换分组，转科时换分组
            If .Cell(flexcpData, rsTemp.AbsolutePosition - 1, mCol.页面名称) = NVL(rsTemp!页面) And NVL(rsTemp!新组, 0) <> 1 _
                And Not (blnGroupTurnDept And .TextMatrix(rsTemp.AbsolutePosition - 1, mCol.病人科室) <> NVL(rsTemp!病人科室)) Then
                .RowOutlineLevel(rsTemp.AbsolutePosition) = 1
                .TextMatrix(rsTemp.AbsolutePosition, mCol.页面名称) = ""
            Else
                .RowOutlineLevel(rsTemp.AbsolutePosition) = 0
            End If
            
            If strKind <> .TextMatrix(rsTemp.AbsolutePosition, mCol.病历种类) Then '画分类线条
                If strKind <> "" Then Call .CellBorderRange(rsTemp.AbsolutePosition, 0, rsTemp.AbsolutePosition, .Cols - 1, RGB(0, 0, 255), 0, 1, 0, 0, 0, 0)
                strKind = .TextMatrix(rsTemp.AbsolutePosition, mCol.病历种类)
            End If

            If Val(.TextMatrix(rsTemp.AbsolutePosition, mCol.处理状态)) > 0 Then '状态图标
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.标志) = imgThis.ListImages("转交").Picture
            ElseIf Trim(.TextMatrix(rsTemp.AbsolutePosition, mCol.归档人)) <> "" Then
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.标志) = imgThis.ListImages("归档").Picture
            ElseIf Val(.TextMatrix(rsTemp.AbsolutePosition, mCol.当前版本)) <= 1 Then
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.标志) = imgThis.ListImages("书写").Picture
            Else
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.标志) = imgThis.ListImages("修订").Picture
            End If
            .MergeRow(rsTemp.AbsolutePosition) = True
            If Trim(.TextMatrix(rsTemp.AbsolutePosition, mCol.打印)) <> "" Then '打印图标
                 If NVL(rsTemp!页面) <> NVL(rsTemp!病历名称) Or .RowOutlineLevel(rsTemp.AbsolutePosition) = 1 Then
                    .Cell(flexcpPictureAlignment, rsTemp.AbsolutePosition, mCol.病历名称) = flexAlignLeftCenter
                    Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.病历名称) = imgThis.ListImages("打印").Picture
                Else
                    .Cell(flexcpPictureAlignment, rsTemp.AbsolutePosition, mCol.页面名称) = flexAlignLeftCenter
                    Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.页面名称) = imgThis.ListImages("打印").Picture
                End If
            Else
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.页面名称) = Nothing
                Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.病历名称) = Nothing
            End If
            
            If .ROWHEIGHT(rsTemp.AbsolutePosition) < .RowHeightMin Then .ROWHEIGHT(rsTemp.AbsolutePosition) = .RowHeightMin
            If lngCurId = Val(.TextMatrix(rsTemp.AbsolutePosition, mCol.ID)) Then lngCurRow = rsTemp.AbsolutePosition '赋值行号
            rsTemp.MoveNext
        Loop
        
        Call Folding '处理折叠
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngCurRow = 0 Then
            vfgThis.Tag = -1: .Row = 0 '促使vfgthis不选中任何行，不显示任何内容，仅当选中某行时才刷新
        Else
           .Row = lngCurRow
        End If
        Call vfgThis_RowColChange
        zlRefresh = .Rows - 1
    End With
    Call InitColumnSelect '列选择器
    
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "病历书写") > 0) Then '根据权限状态，显示增加窗格
        If zlDatabase.GetPara("自动显示新增面板", glngSys, mlngModul, "1") = 1 Then
            dkpMan.Panes(conPane_New).Select
            Call mfrmNew.zlRefList(2, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs, mlngAdviceID)
        End If
    End If
    Exit Function
   
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Folding()
Dim i As Long, l As Long, N As Long
    l = CLng(vfgThis.Height / vfgThis.RowHeightMin)
    
    If vfgThis.Rows > l Then '可容纳行数小于实际行数
        For i = 1 To vfgThis.Rows - 1
            If vfgThis.RowOutlineLevel(i) = 1 Then '树行首行,开始计数，超过6行时
                N = N + 1
            Else
                N = 0
            End If
            
            If N >= mlngfolding Then
                vfgThis.IsCollapsed(i - mlngfolding) = flexOutlineCollapsed: N = 0
            End If
        Next i
    End If
End Sub

Private Sub AutoResizeCol(ByVal intCol As Integer)
    Dim intRow As Integer
    Dim lngMaxWidth As Long
    
    With vfgThis
        For intRow = .FixedRows To .Rows - 1
            If lngMaxWidth < LenB(.TextMatrix(intRow, intCol)) Then
                lngMaxWidth = LenB(.TextMatrix(intRow, intCol))
            End If
        Next

        If lngMaxWidth > 0 Then
            .ColWidth(intCol) = lngMaxWidth * 90 + 120
        End If
    
    End With
End Sub

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
    Dim strSubhead As String
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select a.住院号, a.姓名 From 病人信息 a Where a.病人id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId)
    If Not rsTemp.EOF Then
        strSubhead = "住院号:" & rsTemp!住院号 & "  姓名:" & rsTemp!姓名
    Else
        strSubhead = ""
    End If
    Err = 0: On Error GoTo 0
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strSubhead)
    Call objAppRow.Add("第" & mlngPageId & "次住院")
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
'##
'## 参数：  blnPreview  :是否是预览模式
'################################################################################################################
Private Sub zlEPRPrint(blnPreview As Boolean)
Dim lFileId As Long, strPrintName As String
Dim r As String, blnOrigMode As Boolean  '是否显示原始状态
    
    lFileId = CLng(vfgThis.TextMatrix(vfgThis.Row, mCol.ID))
    strPrintName = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", "")
    Select Case Val(vfgThis.TextMatrix(vfgThis.Row, mCol.编辑方式))
        Case 0
            Set mfrmPrintPreview = New frmPrintPreview
            r = zlCommFun.ShowMsgBox("病历预览/打印", "请选择病历预览/打印的格式？", "!最终格式(&F),原始格式(&O),取消(&C)", Nothing)
            If r = "最终格式" Then
                blnOrigMode = False
            ElseIf r = "原始格式" Then
                blnOrigMode = True
            Else
                Exit Sub
            End If
            mfrmPrintPreview.DoMultiDocPreview Me, cpr住院病历, mlngPatiId, mlngPageId, _
                        vfgThis.Cell(flexcpText, vfgThis.Row, mCol.病历种类), vfgThis.Cell(flexcpText, vfgThis.Row, mCol.页面编号), _
                        lFileId, Not blnPreview, blnOrigMode, , mblnMoved, mlngAdviceID, , IIf(InStr(mstrPrivs, "取消打印") > 0, 0, 1)    '没有"取消打印"权限不允许重复打印，不允许调整打印份数
            Unload mfrmPrintPreview 'ByZT:窗体Load了未显示，没有人为关闭的情况下VB不会自动Unload
            Set mfrmPrintPreview = Nothing
            If Not blnPreview Then RefreshList '直接打印在此刷新
        Case 1
            mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, False, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, mlngAdviceID, mstrPrivs, mblnMoved, InStr(gstrPrivsEpr, "病历打印") > 0
            mObjTabEprView.zlPrintDoc Me, blnPreview, strPrintName
        Case 2
'            传染病已独立页面，不要此显示
    End Select
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", strPrintName
End Sub

Private Sub modelsApply()
    Dim frmModels As New frmEPRModelsMan, strPrivs As String
    If frmModels.Showfrm(Me, mlngPatiId, mlngPageId, mlngDeptId, gstrPrivsEpr) Then RefreshList
End Sub
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
            intNew = 1: intDel = 1: intMod = 1
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
Private Function GetEprSign(ByVal lngFileID As Long)
'提取病历历史签名记录
Dim rsTemp As ADODB.Recordset, strSign As String
    gstrSQL = "Select 开始版 As 版本, Decode(要素表示, 3, '主任医师', 2, '主治医师', '经治医师') || '身份' || Decode(开始版, 1, '签名', '修订') As 操作," & vbNewLine & _
                "       Decode(Nvl(Instr(内容文本, ';'), 0), 0, 内容文本, Substr(内容文本, 1, Instr(内容文本, ';') - 1)) As 人员," & vbNewLine & _
                "       RTrim(Substr(对象属性, Instr(对象属性, ';', 1, 4) + 1)) As 时间" & vbNewLine & _
                "From 电子病历内容" & vbNewLine & _
                "Where 文件id = [1] And 对象类型 = 8 Order By 对象标记"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名记录", lngFileID)
    Do Until rsTemp.EOF
        strSign = strSign & "由 " & Rpad(NVL(rsTemp!人员), 8) & "于 " & Rpad(NVL(rsTemp!时间), 19) & " 以" & NVL(rsTemp!操作) & vbCrLf
        rsTemp.MoveNext
    Loop
    GetEprSign = strSign
End Function
Private Sub PrintCancel(ByVal lngRecordId As Long)
'取消标记打印
On Error GoTo errHand
    gstrSQL = "Zl_电子病历打印_Cancel(" & lngRecordId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    vfgThis.Cell(flexcpData, vfgThis.Row, mCol.当前情况) = ""
    vfgThis.Cell(flexcpText, vfgThis.Row, mCol.打印) = ""
    Set vfgThis.Cell(flexcpPicture, vfgThis.Row, mCol.页面名称) = Nothing
    Set vfgThis.Cell(flexcpPicture, vfgThis.Row, mCol.病历名称) = Nothing
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        strPrintInfo = strPrintInfo & vbCrLf & "打印人：" & Rpad(rsTemp!打印人, 8) & "打印时间：" & Format(rsTemp!打印时间, "yyyy-MM-dd hh:mm")
        rsTemp.MoveNext
    Loop
    strPrintInfo = Mid(strPrintInfo, 3)
    EprPrinted = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function EprWriteMSG() As Boolean
Dim rsTemp As New ADODB.Recordset, strMsg As String
On Error GoTo errHand
    gstrSQL = "Select 文件ID ID,病历编号 || '-' || 病历名称 病历, 到期时间, 必须" & vbNewLine & _
                "From 电子病历时机" & vbNewLine & _
                "Where 病人id = [1] And 主页id = [2] And 科室id =[3] And 病人来源 = 2 And (Nvl(完成记录id, 0) = 0 And 完成时间 Is Null)" & vbNewLine & _
                "Order By 到期时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mlngDeptId)
    
    Do Until rsTemp.EOF
        strMsg = strMsg & "病历<" & Rpad(rsTemp!病历 & ">", 31) & "尚未书写，最晚完成时间:" & Format(rsTemp!到期时间, "yyyy-MM-dd hh:mm") & "  " & _
                        IIf(NVL(rsTemp!必须, 0) = 0, "但不", "并且") & "是必须书写的，请检查！" & vbCrLf
        rsTemp.MoveNext
    Loop
    
    '内容太多，处理后提示才能看见,只显示十行
    If UBound(Split(strMsg, vbCrLf)) > 9 Then
        strMsg = Mid(strMsg, 1, InStr(710, strMsg, vbCrLf))
        strMsg = strMsg & String(32, Asc("-")) & "以下还有多条记录。"
    End If
    
    If MsgBoxD(Me, strMsg & vbCrLf & "选<是>继续，选<否>取消。", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
        EprWriteMSG = False
    Else
        EprWriteMSG = True
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function TimeLimitOut() As Boolean
'功能:检查是否有转科，出院，预出院情况，有则给出事件和补录时限
Dim rsTemp As New ADODB.Recordset, lngTimeLimit As Long, strReturn As String
    If mintState = 3 Or mintState = 4 Then Exit Function
    
    gstrSQL = "Select Decode(终止原因, 1, '出院', 3, '转科', 10, '预出院') 事件, 终止时间,Trunc((Sysdate - 终止时间) * 24, 5) 当前时差" & vbNewLine & _
                "From 病人变动记录" & vbNewLine & _
                "Where ID = (Select Nvl(Max(ID), 0)" & vbNewLine & _
                "            From 病人变动记录" & vbNewLine & _
                "            Where 病人id = [1] And 主页id = [2] And 终止时间 Is Not Null And 终止原因 In (1, 3, 10))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提出变动记录", mlngPatiId, mlngPageId)
    If rsTemp.EOF Then Exit Function
    
    lngTimeLimit = Val(zlDatabase.GetPara("数据补录时限", 100))
    
    If rsTemp!当前时差 > lngTimeLimit Then
        If rsTemp!事件 = "转科" Then
            strReturn = rsTemp!事件 & "|" & lngTimeLimit
            gstrSQL = "Select 出院科室id From 病案主页 Where 病人id = [1] And 主页id = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出院科室", mlngPatiId, mlngPageId)
            If mlngDeptId = rsTemp!出院科室ID Then strReturn = "" '转科后，在转入科室新增病历不受时限限制
        Else
            strReturn = rsTemp!事件 & "|" & lngTimeLimit
        End If
    End If
    
    If strReturn <> "" Then
        MsgBox "该病人已经" & Split(strReturn, "|")(0) & ",并且超过设定的" & Split(strReturn, "|")(1) & "小时补录时限,不允许变动病历。", vbInformation, gstrSysName
        TimeLimitOut = True
    End If
End Function
Private Function ExportAll() As Boolean
'功能：导出该病人所有全文式病历为RTF
'步骤：1 指定目录
'     2 将文件逐个（共享病历导出为一个文件）加入到控件
'     3 刷新内容对象
'     4 去掉关键字
'     5 保存为RTF，名称为姓名(住院号)_病历名称
Dim strFile As String, strName As String, strPath As String, j As Long
Dim rsTemp As New ADODB.Recordset, strPage As String, lngLen As Long, blnExport As Boolean

    On Error GoTo errHand

    '指定目录
    strPath = zl9ComLib.OS.OpenDir(Me.hwnd, "指定导出目录")
    If strPath = "" Then
        MsgBox "取消指定导出目录，导出失败！", vbExclamation, gstrSysName
        ExportAll = False: Exit Function
    End If
    Call zlCommFun.ShowFlash("请稍等，正在导出文件", Me)
    
    gstrSQL = "Select a.住院号, a.姓名 From 病人信息 a Where a.病人id = [1]" '指定导出文件前辍
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId)
    strName = rsTemp!姓名 & "(住院号_" & rsTemp!住院号 & ")"
    
    strPath = gobjFSO.BuildPath(strPath, rsTemp!姓名) '指定目录下的子目录
    If Not gobjFSO.FolderExists(strPath) Then gobjFSO.CreateFolder strPath '不存在则建立子目录

    
    gfrmPublic.edtPublic.ForceEdit = True
    gfrmPublic.edtBuff.ForceEdit = True
    gfrmPublic.edtPublic.Freeze
    gfrmPublic.edtBuff.Freeze
    For j = 1 To vfgThis.Rows - 1
        If vfgThis.TextMatrix(j, mCol.编辑方式) = 0 Then
            '读取RTF并刷新内容对象
            If vfgThis.RowOutlineLevel(j) = 1 Then '如果当前行与上一行的页面名称相同，则追加，否则单独打开
                Call ReadRTF(gfrmPublic.edtBuff, Val(vfgThis.TextMatrix(j, mCol.ID)), True, mblnMoved)
                gfrmPublic.edtBuff.SelectAll
                gfrmPublic.edtBuff.CopyWithFormat
                lngLen = Len(gfrmPublic.edtBuff.Text)
                If gfrmPublic.edtPublic.Range(lngLen - 2, lngLen).Text = vbCrLf Then '在尾部换行
                    gfrmPublic.edtPublic.Range(lngLen - 2, lngLen).Font.Hidden = False
                Else
                    gfrmPublic.edtPublic.Range(lngLen, lngLen).Text = vbCrLf
                    gfrmPublic.edtPublic.Range(lngLen, lngLen + 2).Font.Hidden = False
                End If
                gfrmPublic.edtPublic.PasteWithFormat
            Else
                strPage = vfgThis.TextMatrix(j, mCol.页面名称)
                Call ReadRTF(gfrmPublic.edtPublic, Val(vfgThis.TextMatrix(j, mCol.ID)), True, mblnMoved)
            End If
            
            
            blnExport = False
            If j = vfgThis.Rows - 1 Then
                blnExport = True
            ElseIf vfgThis.RowOutlineLevel(j + 1) = 0 Then
                blnExport = True
            End If
            
            If blnExport Then
                '清除所有关键字
                Dim i As Long
                Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
                i = 0
                bFinded = FindNextAnyKey(gfrmPublic.edtPublic, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Do While bFinded
                    gfrmPublic.edtPublic.Range(lKSS, lKSE) = ""
                    gfrmPublic.edtPublic.Range(lKSS + lKES - lKSE, lKSS + lKES - lKSE + 16) = ""
                    i = lKSS + (lKES - lKSE)
                    bFinded = FindNextAnyKey(gfrmPublic.edtPublic, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Loop
                
                gfrmPublic.edtPublic.SaveDoc (strPath & "\" & strName & "_" & strPage & Format(vfgThis.TextMatrix(j, mCol.创建时间), "yyyymmddHHmmss") & ".rtf")
            End If
        End If
    Next
    gfrmPublic.edtPublic.ForceEdit = False
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtPublic.UnFreeze
    gfrmPublic.edtBuff.UnFreeze
    Unload gfrmPublic
    Call zlCommFun.StopFlash
    MsgBox "成功导出文件到目录 [" & strPath & "]下!", vbInformation, gstrSysName
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub EprArchive()
Dim strState As String, rsTemp As New ADODB.Recordset, strInfo As String

    On Error GoTo errHand
    gstrSQL = "Select Decode(出院日期, Null, Decode(状态, 3, '预出院', '在院'), '出院') As 病人状态" & vbNewLine & _
                "From 病案主页" & vbNewLine & _
                "Where 病人id = [1] And 主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人当前状态", mlngPatiId, mlngPageId)
    strState = rsTemp!病人状态
    
    With vfgThis
        If Trim(.TextMatrix(.Row, mCol.归档人)) = "" Then
            If Not EprWriteMSG Then Exit Sub
            If strState = "在院" Then
                strInfo = "真的将该份“" & .TextMatrix(.Row, mCol.病历名称) & "”归档吗？"
                If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                gstrSQL = "Zl_电子病历记录_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0)"
            Else
                strInfo = "病人已经" & strState & "，要将病人本次住院全部住院病历归档吗？" & vbCrLf _
                        & "  选择“是”，归档病人本次全部病历；" & vbCrLf _
                        & "  选择“否”，仅归档该份“" & .TextMatrix(.Row, mCol.病历名称) & "”。"
                Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
                Case vbYes: gstrSQL = "Zl_电子病历记录_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0,1)"
                Case vbNo: gstrSQL = "Zl_电子病历记录_Archive(" & .TextMatrix(.Row, mCol.ID) & ",0)"
                Case Else: Exit Sub
                End Select
            End If
        Else
    
            If Split(EprIsCommit, "|")(2) = 0 Then
                MsgBox "该病人病案已提交审查，不能撤档，请取消审查后再试！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strInfo = "需要撤销该病人本次住院所有已归档住院病历吗？" & vbCrLf _
                    & "  选择“是”，撤销该病人本次住院所有已归档住院病历；" & vbCrLf _
                    & "  选择“否”，仅撤消该份“" & .TextMatrix(.Row, mCol.病历名称) & "”的归档。"
            Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
            Case vbYes: gstrSQL = "Zl_电子病历记录_Archive(" & .TextMatrix(.Row, mCol.ID) & ",1,1)"
            Case vbNo: gstrSQL = "Zl_电子病历记录_Archive(" & .TextMatrix(.Row, mCol.ID) & ",1)"
            Case Else: Exit Sub
            End Select
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call Me.zlRefresh(mlngPatiId, mlngPageId, mlngDeptId, mblnEdit, True, mblnMoved, mlngAdviceID)
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function GetPhysicians() As String
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    If mlngPatiId = 0 Then Exit Function
    
    gstrSQL = "Select 经治医师, 主治医师, 主任医师" & vbNewLine & _
            "From 病人变动记录" & vbNewLine & _
            "Where 病人id = [1] And 主页id = [2] And (终止时间 Is Null Or 终止原因 = 1)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医师", mlngPatiId, mlngPageId)
    If rsTemp.EOF Then Exit Function
    GetPhysicians = ";" & NVL(rsTemp!经治医师) & ";" & NVL(rsTemp!主治医师) & ";" & NVL(rsTemp!主任医师) & ";"
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function UserNewEMR() As Boolean
Dim rsTemp As New ADODB.Recordset, lngDeptId As Long
    On Error GoTo errHand
    gstrSQL = "Select ID From 电子病历记录 Where 病人ID=[1] and 主页ID=[2] and 病历种类=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否书写过", mlngPatiId, mlngPageId)
    If Not rsTemp.EOF Then Exit Function '写过老病历
    
    gstrSQL = "Select 出院科室ID From 病案主页 Where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询病人当前科室", mlngPatiId, mlngPageId)
    lngDeptId = NVL(rsTemp!出院科室ID, 0)
    If lngDeptId = 0 Then lngDeptId = mlngDeptId
    
    On Error Resume Next
    gstrSQL = "Select 科室ID From 新版病历启用科室 Where 科室ID=[1]" '没建外键，不能用关联查询
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查本科是否启用", lngDeptId)
    If Err.Number <> 0 Then Err.Clear: Exit Function  '没有控制表
    
    If rsTemp.EOF Then Exit Function '有表但本科没启用
    
    UserNewEMR = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function DisplayContent(ByVal lngId As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHand
    strSQL = "Select  登记人, 登记时间,反馈内容,处理人, 处理时间,处理情况说明  From 疾病报告反馈 where 文件ID = [1] order by 登记时间 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
    If rsTemp.RecordCount = 0 Then
        DisplayContent = False
        Exit Function
    End If

    With vsfFeedback
        .Clear
        .Cols = 6
        .Rows = 1
        .ColWidth(0) = .Width / 10
        .ColWidth(1) = .Width / 10 * 2
        .ColWidth(2) = .Width / 10 * 3
        .ColWidth(3) = .Width / 10
        .ColWidth(4) = .Width / 10
        .ColWidth(5) = .Width / 10 * 2
        .TextMatrix(0, 0) = "登记人"
        .TextMatrix(0, 1) = "登记时间"
        .TextMatrix(0, 2) = "反馈内容"
        .TextMatrix(0, 3) = "处理人"
        .TextMatrix(0, 4) = "处理时间"
        .TextMatrix(0, 5) = "处理情况说明"
    End With
    
    Do Until rsTemp.EOF
        With vsfFeedback
            .Rows = .Rows + 1
            .ROWHEIGHT(.Rows - 1) = 350
            .TextMatrix(.Rows - 1, 0) = NVL(rsTemp!登记人)
            If IsDate(rsTemp!登记时间 & "") Then
                .TextMatrix(.Rows - 1, 1) = Format(rsTemp!登记时间, "yy/mm/dd HH:mm")
            Else
                .TextMatrix(.Rows - 1, 1) = NVL(rsTemp!登记时间)
            End If
            .TextMatrix(.Rows - 1, 1) = NVL(rsTemp!登记时间)
            .TextMatrix(.Rows - 1, 2) = NVL(rsTemp!反馈内容)
            .TextMatrix(.Rows - 1, 3) = NVL(rsTemp!处理人)
            If IsDate(rsTemp!登记时间 & "") Then
                .TextMatrix(.Rows - 1, 1) = Format(rsTemp!处理时间, "yy/mm/dd HH:mm")
            Else
                .TextMatrix(.Rows - 1, 1) = NVL(rsTemp!处理时间)
            End If
            .TextMatrix(.Rows - 1, 5) = NVL(rsTemp!处理情况说明)
        End With
        rsTemp.MoveNext
    Loop
    DisplayContent = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfFeedback_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then '关闭反馈结果查看器
        If vsfFeedback.Visible Then
            vsfFeedback.Visible = False
        End If
    End If
End Sub

Public Function GetFormOperation() As String
'记录界面选定信息，因为工作站在切换页卡时是释放了对象，换回来时重新初始化刷新的。
    GetFormOperation = mlngCurId
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'恢复界面选定信息，工作站在刷新之前调用
    mlngCurId = Val(strValue)
End Sub

Private Sub vsfFeedback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    With vsfFeedback
        If .MouseRow >= 0 And .MouseCol >= 0 Then
            Call zlCommFun.ShowTipInfo(.hWnd, .TextMatrix(.MouseRow, .MouseCol), True, True)
        End If
    End With
End Sub
