VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmExecOnePart 
   Caption         =   "检查医嘱，分部位执行"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "frmExecOnePart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10680
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picOrder 
      Height          =   3735
      Left            =   5160
      ScaleHeight     =   3675
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
      Begin VSFlex8Ctl.VSFlexGrid vsfOrder 
         Height          =   3255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4935
         _cx             =   8705
         _cy             =   5741
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
   Begin VB.Frame frmButton 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   9975
      Begin VB.CommandButton cmdExecPart 
         Caption         =   "分部位执行"
         Height          =   400
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "执行部位医嘱"
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancelPart 
         Caption         =   "分部位取消"
         Height          =   400
         Left            =   4380
         TabIndex        =   5
         ToolTipText     =   "取消部位医嘱的执行状态"
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出"
         Default         =   -1  'True
         Height          =   400
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Frame frmInfo 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.Label lblInfo 
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   10095
      End
      Begin VB.Label lblName 
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   9735
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   1560
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmExecOnePart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjDockExpense As zlPublicExpense.clsDockExpense

Private mlngDeptID As Long
Private mlngSendNo As Long
Private mlngOrderID As Long

Private Enum Order_Column
    col_医嘱ID = 0
    col_相关ID = 1
    col_序号 = 2
    col_医嘱内容 = 3
    col_执行状态 = 4
    col_执行状态描述 = 5
End Enum

Private Sub cmdCancelPart_Click()
    Dim strSql As String
    
    On Error GoTo err
    
    If vsfOrder.Rows < 1 Or vsfOrder.RowSel < 1 Then
        Call MsgBoxD(Me, "没有选中的部位医嘱，不能分部位取消执行。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "ZL_影像检查_CANCEL(" & Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_医嘱ID)) & "," & mlngSendNo & ",1," & mlngDeptID & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    '刷新窗口数据
    RefreshOrder
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdExecPart_Click()
    Dim strSql As String

    On Error GoTo err

    If vsfOrder.Rows < 1 Or vsfOrder.RowSel < 1 Then
        Call MsgBoxD(Me, "没有选中的部位医嘱，不能分部位执行。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    strSql = "Zl_影像检查_单独执行(" & Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_医嘱ID)) & "," & mlngSendNo & ",'" & UserInfo.编号 & _
            "','" & UserInfo.姓名 & "'," & mlngDeptID & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    '刷新窗口数据
    RefreshOrder

    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Public Sub zlShowMe(lngOrderID As Long, strName As String, strAge As String, strSex As String, strState As String, _
    frmParent As Object)
    
    On Error GoTo err
    
    '显示基本信息
    lblName = "姓名：" & strName
    lblInfo = "性别：" & strSex & "      年龄：" & strAge & "      检查状态：" & strState
    
    '初始化窗体
    Call InitForm
    
    mlngOrderID = lngOrderID
    
    '刷新窗口，根据医嘱ID，显示医嘱列表
    Call RefreshOrder
    
    Me.Show 1, frmParent
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitOrder()
    On Error GoTo err
    
    With vsfOrder
        .Rows = 1
        .Cols = 6
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
        
        .TextMatrix(0, col_序号) = "序号"
        .TextMatrix(0, col_医嘱内容) = "医嘱内容"
        .TextMatrix(0, col_执行状态描述) = "执行状态"
        
        .ColWidth(col_序号) = 650
        .ColWidth(col_医嘱内容) = 3500
        .ColWidth(col_执行状态描述) = 600
        
        '隐藏医嘱ID列
        .ColHidden(col_医嘱ID) = True
        .ColHidden(col_相关ID) = True
        .ColHidden(col_执行状态) = True
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub LoadOrder(lngOrderID As Long)
'加载医嘱列表

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    strSql = "select b.医嘱id,a.医嘱内容,a.标本部位 ,a.检查方法,b.执行状态,a.相关id,a.执行科室id,b.发送号," & _
            " Decode(b.执行状态, 0, '未执行', 1, '已完成',  2, '已拒绝', '正在执行') 执行状态描述 " & _
            " from 病人医嘱记录 a ,病人医嘱发送 b " & _
            " where a.id = b.医嘱id And (a.Id = [1] or a.相关ID=[1]) order by 医嘱ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "分部位执行", lngOrderID)

    If rsTemp.EOF = True Then Exit Sub

    With vsfOrder
        If 1 + rsTemp.RecordCount <> .Rows Then
            .Rows = 1 + rsTemp.RecordCount
        End If

        mlngDeptID = rsTemp!执行科室ID
        mlngSendNo = rsTemp!发送号

        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, col_医嘱ID) = rsTemp!医嘱ID
            .TextMatrix(i, col_相关ID) = Nvl(rsTemp!相关ID, 0)
            .TextMatrix(i, col_序号) = IIf(Nvl(rsTemp!相关ID, 0) = 0, "主医嘱", "（" & i - 1 & "）")
            .TextMatrix(i, col_医嘱内容) = IIf(Nvl(rsTemp!相关ID, 0) = 0, rsTemp!医嘱内容, rsTemp!医嘱内容 & "：" & Nvl(rsTemp!标本部位) & "（" & Nvl(rsTemp!检查方法) & "）")
            .TextMatrix(i, col_执行状态) = rsTemp!执行状态
            .TextMatrix(i, col_执行状态描述) = rsTemp!执行状态描述
            rsTemp.MoveNext
        Next i
        
        If .RowSel = 0 Then .RowSel = 1
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = frmInfo.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picOrder.hWnd
    ElseIf Item.ID = 3 Then
        Item.Handle = mobjDockExpense.zlGetForm.hWnd
    ElseIf Item.ID = 4 Then
        Item.Handle = frmButton.hWnd
    End If
End Sub

Private Sub dkpMain_Resize()
    cmdCancelPart.Left = frmButton.Width / 2 - cmdCancelPart.Width / 2
    cmdExecPart.Left = cmdCancelPart.Left - 1000 - cmdExecPart.Width
    cmdExit.Left = cmdCancelPart.Left + cmdCancelPart.Width + 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjDockExpense = Nothing
End Sub

Private Sub picOrder_Resize()
    vsfOrder.Left = 0
    vsfOrder.Top = 0
    vsfOrder.Width = picOrder.Width
    vsfOrder.Height = picOrder.Height
End Sub

Private Sub vsfOrder_SelChange()
    Dim lngOrderID As Long
    Dim blnIsPartOrder As Boolean
    
    On Error GoTo err
    
    '先设置按钮的默认值
    cmdExecPart.Enabled = False
    cmdCancelPart.Enabled = False
        
    If vsfOrder.Rows <= 1 Then Exit Sub
    
    lngOrderID = Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_医嘱ID))
    If lngOrderID = 0 Then Exit Sub
    
    '设置按钮可用性，只有部位医嘱才能分部位执行和取消
    blnIsPartOrder = Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_相关ID)) <> 0
    If blnIsPartOrder = True Then
        '根据执行状态，判断哪个按钮可用
        ' 0, '未执行', 1, '已完成',  2, '已拒绝', 3,'正在执行'
        If Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_执行状态)) = 3 Then
            cmdCancelPart.Enabled = True
            cmdExecPart.Enabled = False
        ElseIf Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_执行状态)) = 0 Then
            cmdExecPart.Enabled = True
            cmdCancelPart.Enabled = False
        End If
    End If
    
    '刷新费用窗口
    If Not mobjDockExpense Is Nothing Then
        Call mobjDockExpense.zlRefresh(mlngDeptID, lngOrderID & ":" & mlngSendNo & ":" & IIf(blnIsPartOrder, 1, 0))
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitFee()
    On Error GoTo err
    If mobjDockExpense Is Nothing Then
        Set mobjDockExpense = New zlPublicExpense.clsDockExpense
        Call mobjDockExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub InitForm()
'------------------------------------------------
'功能：初始化窗口
'参数：无
'返回：无
'------------------------------------------------
    Dim Pane1 As Pane
    Dim Pane2 As Pane
    Dim Pane3 As Pane
    Dim pane4 As Pane
    
    On Error GoTo err
    
    '初始化医嘱列表
    Call InitOrder
    
    '初始化费用列表
    Call InitFee
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .TabPaintManager.BoldSelected = True
        .Options.DefaultPaneOptions = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        
        Set Pane2 = .CreatePane(2, 200, 200, DockLeftOf)
        Set Pane3 = .CreatePane(3, 200, 200, DockRightOf)
        Set Pane1 = .CreatePane(1, 400, 80, DockTopOf, Pane2 And Pane3)
        Set pane4 = .CreatePane(4, 400, 60, DockBottomOf, Nothing)
    
        Pane1.MaxTrackSize.Height = 80
        Pane1.MinTrackSize.Height = 80
        
        pane4.MaxTrackSize.Height = 60
        pane4.MinTrackSize.Height = 60
        
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshOrder()
'------------------------------------------------
'功能：刷新窗口中的医嘱和费用数据
'参数：无
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    '根据医嘱列表，显示费用
    Call LoadOrder(mlngOrderID)
    
    '刷新费用窗体
    Call vsfOrder_SelChange
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub
