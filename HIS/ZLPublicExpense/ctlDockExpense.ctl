VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.UserControl ctlDockExpense 
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ScaleHeight     =   7020
   ScaleWidth      =   10560
   Begin VB.PictureBox picExpense 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2424
      Left            =   150
      ScaleHeight     =   2430
      ScaleWidth      =   7620
      TabIndex        =   2
      Top             =   2280
      Width           =   7620
      Begin VSFlex8Ctl.VSFlexGrid vsExpense 
         Height          =   1440
         Left            =   96
         TabIndex        =   3
         Top             =   108
         Width           =   6960
         _cx             =   12277
         _cy             =   2540
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "批号信息:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   2160
         Width           =   1020
      End
   End
   Begin VB.PictureBox picAdvice 
      BorderStyle     =   0  'None
      Height          =   1668
      Left            =   408
      ScaleHeight     =   1665
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   504
      Width           =   7332
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   1272
         Left            =   24
         TabIndex        =   1
         Top             =   84
         Width           =   6960
         _cx             =   12277
         _cy             =   2244
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "ctlDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------
'说明:
'   如果不使用控件(DockingPance控件),如果绑定窗体是模态窗体,则会死机,
'   如果用控件形式,则不会出现死机情况,因此,将此调整为控件形式
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'入参变量
Private mstrAdviceIDAndPayNums As String
Private mstrAdviceIDFull As String  '完整的医嘱相关信息串,医嘱ID和发送号和独立执行标志(医嘱ID1:发送号1:独立执行,医嘱ID2:发送号2:独立执行,...)
Private mstrNos As String, mbyt记录性质 As Byte, mbyt病人来源 As Byte '按单据号查时
Private mbytFun As Byte '0-按医嘱查找;1-按单据号查找
Private mblnMoved As Boolean
Private mlngModule As Long
Private mlng执行科室ID As Long
Private mobjSquareCard As Object
'-----------------------------------------------------------------------------------------
Private mobjPubAdvice As Object  '公共医嘱对象
Private mfrmParent As Object
Private mobjSaveData As Object
Private mstrPrivsAnnexFee As String
Private mbytFocus As Byte

Private Enum mPaneIdx
    Pan_AdviceList = 1  '医嘱列表
    Pan_FeeList = 2     '费用列表
End Enum
Private Type ty_adviceProperty '医嘱信息
    lng医嘱ID As Long
    lng发送号 As Long
    bln独立执行 As Boolean
    lng病人ID  As Long
    lng主页Id   As Long
    str挂号单   As String
    lng病人科室ID   As Long
    lng病人病区ID   As Long
    lng开嘱科室ID   As Long
    int记录性质   As Integer
    int审核标志   As Integer
    int结算模式   As Integer
    int病人来源 As Integer
    int执行状态 As Integer

    lng相关ID  As Long
    str诊疗类别  As String
    strNO As String
    dat发送时间  As Date
    str费别   As String
    lng计价性质  As Long
    bln门诊记帐 As Boolean
    str计费状态 As String
    strFeeTab As String

End Type
Private mTYAdviceProperty As ty_adviceProperty
Private mrsPrice As ADODB.Recordset '医嘱计价关系

'缺省属性值:
Const m_def_COLOR_FOCUS = &HFFCC99
Const m_def_COLOR_LOST = &HFFEBD7
Const m_def_Tittle = "医嘱附费管理"
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'属性变量:
Dim m_COLOR_FOCUS As OLE_COLOR
Dim m_COLOR_LOST As OLE_COLOR
Dim m_Tittle As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
'Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'事件声明:
Event Click()
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick()
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"

'---------------------------------------------------------------------------------------------------------
'相关事件
Event Activate() '自已激活时
Event RequestRefresh() '要求主窗体刷新
Event StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String) '要求更新主窗体状态栏文字
'bytType:1-附费执行,2-附费取消

Event zlPopupMenu(lng医嘱ID As Long, lng发送号 As Long, strNO As String, int记录性质 As Integer, X As Single, Y As Single)
Private mblnNotFirstSel As Boolean '非第一次选择

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域初始设置
    '编制:刘兴洪
    '日期:2014-05-26 10:30:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(Pan_AdviceList, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "医嘱信息"
    panThis.Handle = picAdvice.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = Pan_AdviceList

    Set panThis = dkpMan.CreatePane(Pan_FeeList, 250, 580, DockBottomOf, panThis)
    panThis.Title = "费用信息"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picExpense.hWnd
    panThis.Tag = Pan_FeeList
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    'zlRestoreDockPanceToReg  Me, dkpMan, "区域"
End Sub
Private Sub picAdvice_Resize()
    Err = 0: On Error Resume Next
    With picAdvice
        vsAdvice.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub
Private Sub picExpense_Resize()
    Err = 0: On Error Resume Next
    With picExpense
        If lblInfo.Visible Then
            vsExpense.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight - lblInfo.Height - 120
            lblInfo.Top = .ScaleHeight - lblInfo.Height - 30
        Else
            vsExpense.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
        End If
    End With
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_AdviceList
        Item.Handle = picAdvice.hWnd
    Case Pan_FeeList
        Item.Handle = picExpense.hWnd
    End Select
End Sub

Private Sub UserControl_Resize()
    dkpMan.RecalcLayout
    Call picAdvice_Resize
    Call picExpense_Resize
End Sub

Private Sub vsAdvice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Tittle, "医嘱信息", True
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call RefreshExpenseData
End Sub
Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsAdvice.AutoSizeMode = flexAutoSizeRowHeight
    Call vsAdvice.AutoSize(vsAdvice.ColIndex("医嘱内容"))
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Tittle, "医嘱信息", True
End Sub
Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = COLOR_FOCUS
    mbytFocus = 1
End Sub
Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = COLOR_LOST
End Sub
Private Sub vsExpense_GotFocus()
    mbytFocus = 2
    vsExpense.BackColorSel = COLOR_FOCUS
End Sub
Private Sub vsExpense_LostFocus()
    vsExpense.BackColorSel = COLOR_LOST
End Sub

Private Sub vsExpense_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "费用信息", True
End Sub
Private Sub vsExpense_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     If OldRow = NewRow Then Exit Sub
     
     With vsExpense
        If NewRow <= 0 Or NewRow >= .Rows Then Exit Sub
        If NewCol <= 0 Or NewCol >= .Cols Then Exit Sub
        Call Load批号信息(NewRow)
        .ForeColorSel = .Cell(flexcpForeColor, NewRow, NewCol)
     End With
End Sub

Private Sub Load批号信息(ByVal lngRow As Long)
    Dim strSql As String, rsInfo As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select 批号" & vbNewLine & _
            "From 药品收发记录" & vbNewLine & _
            "Where 单据 = 21 And 费用id In (Select Max(ID) From " & IIf(mTYAdviceProperty.int病人来源 = 1, "门诊费用记录", "住院费用记录") & " Where NO = [1] And 记录性质 = [2] And 序号 = [3])"
    Set rsInfo = gobjDatabase.OpenSQLRecord(strSql, "Load批号信息", vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("单据号")), _
                                            Val(vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("记录性质"))), Val(vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("序号"))))
    If rsInfo.EOF Then
        lblInfo.Visible = False
        vsExpense.Move picExpense.ScaleLeft, picExpense.ScaleTop, picExpense.ScaleWidth, picExpense.ScaleHeight
        Exit Sub
    End If
    lblInfo.Visible = True
    vsExpense.Move picExpense.ScaleLeft, picExpense.ScaleTop, picExpense.ScaleWidth, picExpense.ScaleHeight - lblInfo.Height - 120
    lblInfo.Caption = "批号信息:" & rsInfo!批号 & "(" & vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("数量")) & _
                        vsExpense.TextMatrix(lngRow, vsExpense.ColIndex("计算单位")) & ")"
    lblInfo.Top = picExpense.ScaleHeight - lblInfo.Height - 30
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Sub

Private Sub vsExpense_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "费用信息", True
End Sub
Private Sub UserControl_Initialize()
    mlngModule = p医嘱附费管理  '医嘱附费管理
    Call InitPancel
    Call InitGridHead(True)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
 

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "强制完全重画一个对象。"
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '编制:刘兴洪
    '日期:2014-05-30 14:48:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytFun <> 1 Then '按医嘱ID加载医嘱数据
        If LoadAdviceData(mstrAdviceIDAndPayNums) = False Then Exit Sub
    Else
        '按单据加载医嘱数据
        If LoadFeeListFromNos(mbyt记录性质, mstrNos, mbyt病人来源, mblnMoved) = False Then Exit Sub
    End If
End Sub
Public Property Get Is未计费() As Boolean
    Is未计费 = InStr(mTYAdviceProperty.str计费状态, ",-1,")
End Property
Public Property Get IsHaveExpenseData() As Boolean
    IsHaveExpenseData = Get单据号 <> ""
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Tittle = m_def_Tittle
    Set UserControl.Font = Ambient.Font
    m_COLOR_FOCUS = m_def_COLOR_FOCUS
    m_COLOR_LOST = m_def_COLOR_LOST
End Sub
'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Tittle = PropBag.ReadProperty("Tittle", m_def_Tittle)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 9)
    m_COLOR_FOCUS = PropBag.ReadProperty("COLOR_FOCUS", m_def_COLOR_FOCUS)
    m_COLOR_LOST = PropBag.ReadProperty("COLOR_LOST", m_def_COLOR_LOST)
End Sub
Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
    Set mobjPubAdvice = Nothing '释放医嘱对象的相关资料
    If gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State = 0 Then Exit Sub
    If gobjDatabase Is Nothing Then Exit Sub
    
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save mlngModule, vsAdvice, Tittle, "医嘱信息", True
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "费用信息", True
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Tittle", m_Tittle, m_def_Tittle)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 9)
    Call PropBag.WriteProperty("COLOR_FOCUS", m_COLOR_FOCUS, m_def_COLOR_FOCUS)
    Call PropBag.WriteProperty("COLOR_LOST", m_COLOR_LOST, m_def_COLOR_LOST)
End Sub
Public Function zlRefresh(ByVal frmMain As Object, ByVal lng科室id As Long, ByVal strAdviceIdAndPayNums As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal strNos As String, _
    Optional ByVal byt记录性质 As Byte, Optional ByVal byt病人来源 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据重新刷新
    '入参:lng科室id-科室ID
    '     strAdviceIdAndPayNums-医嘱ID和发送号和独立执行标志(医嘱ID1:发送号1:独立执行,医嘱ID2:发送号2:独立执行,...)
    '     strNos:单据号(多个传入时,用逗号分离)
    '     byt记录性质:医嘱ID传空时,才传入,单据性质(1-收费单;2-记帐单)
    '     byt病人来源-1-门诊;2-住院
    '     blnMoved -该病人的数据是否已转出
    '     bln单独执行-用于检验项目，一并采集的一组项目，是否针对其中的某一个单独执行
    '出参:
    '编制:刘兴洪
    '日期:2014-04-10 11:02:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = p医嘱附费管理  '医嘱附费管理
    mstrNos = strNos: mbyt记录性质 = byt记录性质: mbyt病人来源 = byt病人来源
    mlng执行科室ID = lng科室id
    mblnMoved = blnMoved
    mbytFun = IIf(strAdviceIdAndPayNums = "", 1, 0)
    mstrAdviceIDFull = strAdviceIdAndPayNums
    mstrPrivsAnnexFee = GetInsidePrivs(p医嘱附费管理)
    Set mfrmParent = frmMain
    mstrAdviceIDAndPayNums = GetAdviceIDAndPayNums(strAdviceIdAndPayNums)
    If mblnMoved = False Then
        mblnMoved = gobjDatabase.TableDataMoved("病人医嘱发送", " (医嘱ID,发送号) IN", " (Select C1 As 医嘱id, C2 As 发送号 From Table(f_Num2list2('" & mstrAdviceIDAndPayNums & "')))")
    End If
    Call VisiblePancel  '显示或隐藏医嘱列表
    If mbytFun <> 1 Then '按医嘱ID加载医嘱数据
        If LoadAdviceData(mstrAdviceIDAndPayNums) = False Then Exit Function
        If mblnNotFirstSel = False Then Call SetDefalutFocus(True)
    Else
        '按单据加载医嘱数据
        If LoadFeeListFromNos(mbyt记录性质, mstrNos, byt病人来源, blnMoved) = False Then Exit Function
        If mblnNotFirstSel = False Then Call SetDefalutFocus(False)
    End If
    mblnNotFirstSel = True
    zlRefresh = True
End Function
Private Function Get独立执行状态(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取独立执行状态
    '入参:lng医嘱ID-医嘱ID
    '     lng发送号-发送号
    '出参:
    '返回:获取独立执行状态
    '编制:刘兴洪
    '日期:2014-05-27 11:12:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, varTemp As Variant
    On Error GoTo errHandle
    varData = Split(mstrAdviceIDFull, ",")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ":::", ":")
        If lng医嘱ID = Val(varTemp(0)) And lng发送号 = Val(varTemp(1)) Then
            Get独立执行状态 = Val(varTemp(2)): Exit For
        End If
    Next
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetAdviceIDAndPayNums(ByVal strAdviceIDFull As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱ID和发送号的字符串
    '入参:strAdviceIDFull(医嘱ID和发送号和独立执行标志(医嘱ID1:发送号1:独立执行,医嘱ID2:发送号2:独立执行,...))
    '返回:返回以医嘱ID和发送号为格式的串(医嘱ID:发送号,医嘱ID:发送号....)
    '编制:刘兴洪
    '日期:2014-05-26 15:37:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim strAdvice As String
    varData = Split(strAdviceIDFull, ",")
    strAdvice = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & "::", ":")
        strAdvice = strAdvice & "," & varTemp(0) & ":" & varTemp(1)
    Next
    If strAdvice <> "" Then strAdvice = Mid(strAdvice, 2)
    GetAdviceIDAndPayNums = strAdvice
End Function
Private Sub VisiblePancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示或隐藏医嘱列表
    '编制:刘兴洪
    '日期:2014-05-26 16:17:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    If dkpMan Is Nothing Then Exit Sub
    Set panThis = dkpMan.FindPane(Pan_AdviceList)
    If panThis Is Nothing Then Exit Sub
    
    If mbytFun = 1 Then
        panThis.Close
    Else
        panThis.Closed = False
    End If
    dkpMan.RecalcLayout
End Sub

Private Sub vsExpense_DblClick()
    '双击查看
    If Get单据号 = "" Then Exit Sub
    If vsExpense.IsSubtotal(vsExpense.Row) = True Then Exit Sub
    
    Call frmTechnicExpense.EditCard(mfrmParent, mstrPrivsAnnexFee, 1, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
         IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), Val(vsExpense.TextMatrix(vsExpense.Row, vsExpense.ColIndex("记录性质"))), mTYAdviceProperty.lng开嘱科室ID, mTYAdviceProperty.lng病人科室ID, 0, "", mTYAdviceProperty.strNO, Get单据号)
End Sub

Private Function LoadAdviceData(ByVal strAdviceIdAndPayNums As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载医保数据
    '入参:strAdviceIDAndPayNums:医嘱ID和发送号字符串，医嘱ID1:发送号1,医嘱ID2:发送号2
    '返回:医嘱数据加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-26 10:51:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim rsTotal As ADODB.Recordset

    If mobjPubAdvice Is Nothing Then
        Call InitGridHead(True): LoadAdviceData = True
        Exit Function
    End If
    If GetAdviceMoney(strAdviceIdAndPayNums, rsTotal) = False Then Set rsTotal = Nothing
     
    'GetExecAdviceRecord:
    '  strIDsAndNos 医嘱ID和发送号字符串，医嘱ID1:发送号1,医嘱ID2:发送号2
    '  rsReturn:返回的记录集,包含的记录集信息有:
    '    医嘱ID,相关ID,发送号,病人ID,主页ID,开始时间,医嘱内容,数次,应收金额,实收金额,医生嘱托,开嘱医生,开嘱时间
    On Error GoTo errHandle
    If mobjPubAdvice.GetExecAdviceRecord(strAdviceIdAndPayNums, rsTemp) = False Then Exit Function
    With vsAdvice
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF

            .TextMatrix(i, .ColIndex("医嘱ID")) = Val(Nvl(rsTemp!医嘱ID))
            .TextMatrix(i, .ColIndex("相关ID")) = Val(Nvl(rsTemp!相关ID))
            .TextMatrix(i, .ColIndex("发送号")) = Val(Nvl(rsTemp!发送号))
            .TextMatrix(i, .ColIndex("病人ID")) = Val(Nvl(rsTemp!病人ID))
            .TextMatrix(i, .ColIndex("独立执行")) = Get独立执行状态(Val(Nvl(rsTemp!医嘱ID)), Val(Nvl(rsTemp!发送号)))
            .TextMatrix(i, .ColIndex("主页ID")) = Val(Nvl(rsTemp!主页ID))
            .TextMatrix(i, .ColIndex("开始时间")) = Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("医嘱内容")) = Nvl(rsTemp!医嘱内容)
            .TextMatrix(i, .ColIndex("数次")) = Nvl(rsTemp!数次)
            If Not rsTotal Is Nothing Then
                If rsTotal.State = 1 Then
                    rsTotal.Filter = "医嘱ID=" & Val(Nvl(rsTemp!医嘱ID)) & " and 发送号=" & Val(Nvl(rsTemp!发送号))
                    If rsTotal.EOF = False Then
                        .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(Nvl(rsTotal!应收金额)), gSysPara.Money_Decimal.strFormt_VB)
                        .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(Nvl(rsTotal!实收金额)), gSysPara.Money_Decimal.strFormt_VB)
                    End If
                End If
            End If
            .TextMatrix(i, .ColIndex("医生嘱托")) = Nvl(rsTemp!医生嘱托)
            .TextMatrix(i, .ColIndex("开嘱医生")) = Nvl(rsTemp!开嘱医生)
            .TextMatrix(i, .ColIndex("开嘱时间")) = Format(rsTemp!开嘱时间, "yyyy-mm-dd HH:MM")
            i = i + 1
            rsTemp.MoveNext
        Loop
        .WordWrap = False
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        If .ColWidth(.ColIndex("医嘱内容")) >= 6000 Then .ColWidth(.ColIndex("医嘱内容")) = 6000
        .Redraw = flexRDBuffered
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsAdvice, Tittle, "医嘱信息", True
        '按医嘱内容,处理行高
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        Call .AutoSize(.ColIndex("医嘱内容"))
        If .Rows > 1 Then
            Call vsAdvice_AfterRowColChange(-1, 0, .Row, .Col)
        End If
        .Redraw = flexRDBuffered
    End With
    
    LoadAdviceData = True
    Exit Function
errHandle:
    vsAdvice.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub InitGridHead(Optional ByRef blnInitHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件的表头等信息
    '入参:blnInitHead-是否初始化列头信息
    '编制:刘兴洪
    '日期:2014-05-26 10:40:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strHeand As String, varData As Variant
    On Error GoTo errHandle

    With vsAdvice
        .Redraw = flexRDNone
        If blnInitHead Then
            strHeand = "" & _
            "医嘱ID,相关ID,独立执行,发送号,病人ID,主页ID, 开始时间,医嘱内容,数次,应收金额,实收金额,医生嘱托,开嘱医生,开嘱时间"
            varData = Split(strHeand, ",")
            .Clear 1
            .Cols = UBound(varData) + 1
            .Rows = 2
           For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
           Next
        End If

        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = Trim(UCase(.TextMatrix(0, i)))
            If i = .ColIndex("医嘱内容") Then .ColWidth(i) = 2500
            'ColData:列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "发送号" Or .ColKey(i) = "独立执行" Then
               .ColHidden(i) = True
               .ColData(i) = "-1||1"
               .ColAlignment(i) = flexAlignLeftCenter
            ElseIf .ColKey(i) Like "*时间" Or .ColKey(i) Like "*日期" _
                Or .ColKey(i) Like "*号" Or .ColKey(i) Like "*单" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColData(i) = "0||0"
            ElseIf .ColKey(i) Like "*数*" Or .ColKey(i) Like "*量" _
                Or .ColKey(i) Like "*额" Then
                .ColAlignment(i) = flexAlignRightCenter
                .ColData(i) = "0||0"
            Else
                .ColAlignment(i) = flexAlignLeftCenter
                .ColData(i) = "0||0"
            End If
            If .ColKey(i) = "医嘱内容" Then
                .ColData(i) = "1||0"
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    With vsExpense
        If blnInitHead Then
            strHeand = "费用类型,记录性质,收费标志,单据类型,单据号,收费细目ID,费别,开单部门,开单人,类别,序号,项目,单价,数量,计算单位,应收金额,实收金额,执行部门,执行情况,执行状态,收费类别,登记时间,操作员姓名"
            varData = Split(strHeand, ",")
            .Clear 1
            .Cols = UBound(varData) + 1
            .Rows = 2
           For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
           Next
        End If
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(UCase(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            .MergeCol(i) = False
            Select Case .ColKey(i)
            Case "单价", "应收金额", "实收金额"
                .ColAlignment(i) = flexAlignRightCenter
            Case "单据号", "单据类型", "费别", "开单部门", "开单人", "记录状态", "执行状态", "收费类别"
                 'If .ColKey(i) <> "单据类型" Then
                 .ColHidden(i) = True
                If .ColKey(i) <> "开单部门" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            Case Else
                If .ColKey(i) Like "*ID" Then
                    .ColHidden(i) = True
                ElseIf .ColKey(i) Like "*时间" Or .ColKey(i) Like "*日期" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            End Select
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserFreezing = flexFreezeNone
        
        .OutlineBar = flexOutlineBarComplete
        .ExplorerBar = flexExSortShow
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsAdvice.Redraw = flexRDBuffered
    vsExpense.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CreatePubAdvice() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建医嘱的公共对象
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-26 10:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubAdvice Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjPubAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
        If Err <> 0 Then
            'Call MsgBox("公共医嘱部件丢失,医嘱信息将显示异常,请与系统管理员联系!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName)
            Err = 0: On Error GoTo 0
            Exit Function
        End If
        Err = 0: On Error GoTo Errhand:
        Call mobjPubAdvice.InitCommon(gcnOracle, glngSys)
    End If
    CreatePubAdvice = True
    Exit Function
Errhand:
    If gobjComlib Is Nothing Then Exit Function
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Private Function LoadFeeListFromNos(ByVal byt记录性质 As Byte, ByVal strNos As String, _
    ByVal byt病人来源 As Byte, ByVal blnMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号和记录性质加载数据
    '入参:byt记录性质:(1-收费;2-记帐)
    '     strNos:单据号,多个用逗号分离
    '     byt病人来源-1-门诊;2-住院
    '     blnMoved-是否转储到历史表空间
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-26 16:23:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String, i As Long, r As Long
    Dim strFeeTab As String
    Dim bln药房单位 As Boolean, str药房单位 As String, str药房包装 As String

    On Error GoTo errHandle
    '药品单位
    bln药房单位 = Val(gobjDatabase.GetPara("药品单位", glngSys, p医嘱附费管理)) <> 0
    If byt病人来源 = 1 Then
        str药房单位 = "门诊单位": str药房包装 = "门诊包装"
    Else
        str药房单位 = "住院单位": str药房包装 = "住院包装"
    End If

    strFeeTab = IIf(byt病人来源 = 1, "门诊费用记录", "住院费用记录")
    strFeeTab = IIf(blnMoved, "H", "") & strFeeTab

    strSql = "" & _
    "   Select mod(A.记录性质,10) as 记录性质,A.记录状态,A.NO as 单据号," & _
    "          A.费别,Nvl(A.价格父号,A.序号) as 序号,A.收费细目ID, " & _
    "          avg(Nvl(A.付数,1)*A.数次) as 数量,sum(A.标准单价) as 标准单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "          Max(decode(A.记录状态,2,0,A.执行状态)) as 执行状态,A.收费类别, " & _
    "          Max(decode(A.记录状态,2,NULL,decode(A.记录性质,11,NULL,to_char(a.登记时间,'yyyy-mm-dd hh24:mi:ss')))) as 登记时间," & _
    "          Max(decode(A.记录状态,2,NULL,decode(A.记录性质,11,NULL,to_char(a.发生时间,'yyyy-mm-dd hh24:mi:ss')))) as 发生时间," & _
    "          Max(decode(A.记录状态,2,NULL,decode(A.记录性质,11,NULL,A.操作员姓名))) as 操作员姓名," & _
    "          A.执行部门ID,A.开单部门ID,A.开单人" & _
    "   From " & strFeeTab & " A,Table(f_str2list([2])) B" & _
    "   Where  mod(A.记录性质,10)=[1]  And  A.NO=B.Column_Value" & _
    "   Group by mod(A.记录性质,10),A.记录状态,A.NO,A.费别,Nvl(A.价格父号,A.序号),A.收费细目ID,A.收费类别," & _
    "           A.执行部门ID,A.开单部门ID,A.开单人"
    
    strSql = "" & _
    "   Select /*+ RULE */ '' as 费用类型,mod(A.记录性质,10) as 记录性质,decode(nvl(max(a.记录状态),0),1,1,0) as 收费标志," & _
    "       Decode( a.记录性质, 1, '收费', 2, '记帐', 3, '记帐', 4, '挂号', '5', '医疗卡', '未知') As 单据类型," & _
    "       A.单据号, " & _
    "       A.收费细目ID,A.费别,M.名称 as 开单部门,A.开单人,C.名称 as 类别,A.序号 ," & _
    "       Nvl(F.名称,B.名称)||Decode(B.规格,NULL,NULL,' '||B.规格) as 项目," & _
    "       sum(A.标准单价" & IIf(bln药房单位, "*Nvl(E." & str药房包装 & ",1)", "") & ") as 单价," & _
    "       Sum(Nvl(A.数量,1)" & IIf(bln药房单位, "/Nvl(E." & str药房包装 & ",1)", "") & ") as 数量," & _
            IIf(bln药房单位, "Decode(E.药品ID,NULL,B.计算单位,E." & str药房单位 & ")", "B.计算单位") & " as 计算单位," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额,D.名称 as 执行部门," & _
    "       Decode(Max(Nvl(A.执行状态,0)),0,'未执行',1,'完全执行',2,'部分执行') as 执行情况," & _
    "       Max(Nvl(A.执行状态,0)) as 执行状态,a.收费类别, " & _
    "       max(A.登记时间) as 登记时间,Max(A.操作员姓名) as 操作员姓名" & _
    " From  (" & strSql & ") A,收费项目目录 B,收费项目类别 C,部门表 D,药品规格 E,收费项目别名 F,部门表 M" & _
    " Where A.收费细目ID=B.ID   And A.收费类别=C.编码 And a.开单部门id = M.Id(+) And A.执行部门ID=D.ID(+)" & _
    "       And B.ID=E.药品ID(+) And A.收费细目ID=F.收费细目ID(+)" & _
    "       And F.码类(+)=1 And F.性质(+)=[3] " & _
    " Group by   a.记录性质,A.单据号,A.费别,M.名称,A.开单人,A.序号,C.名称, A.收费细目ID ,Nvl(F.名称,B.名称),B.规格,B.计算单位,D.名称," & _
    "       a.收费类别,E.药品ID,Nvl(E." & str药房包装 & ",1),E." & str药房单位 & "" & _
    "      " & _
    " Having Nvl(Sum(A.应收金额),0)<>0 Or Nvl(Sum(A.实收金额),0)<>0" & _
    " Order by 单据类型,单据号,序号"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, byt记录性质, strNos, IIf(gSysPara.byt药品名称显示 = 0, 1, 3))
    With vsExpense
        .Redraw = flexRDNone
        .Cols = 1: .FixedCols = 0
        .Rows = 2
        .MergeRow(1) = False
        Set .DataSource = rsTemp
    End With
    Call SetExpenseGridProperty '设置网格属性
    vsExpense.Redraw = flexRDBuffered
    LoadFeeListFromNos = True
    Exit Function
errHandle:
    vsExpense.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetExpenseGridProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置费用网格的相关属性
    '编制:刘兴洪
    '日期:2014-05-26 17:40:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long

    On Error GoTo errHandle
    With vsExpense
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(UCase(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            .MergeCol(i) = False
            Select Case .ColKey(i)
            Case "单价", "应收金额", "实收金额"
                .ColAlignment(i) = flexAlignRightCenter
            Case "序号", "单据号", "单据类型", "费别", "开单部门", "开单人", "记录状态", "执行状态", "收费类别"
                 'If .ColKey(i) <> "单据类型" Then
                 .ColHidden(i) = True
                If .ColKey(i) <> "开单部门" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            Case Else
                If .ColKey(i) Like "*ID" Then
                    .ColHidden(i) = True
                ElseIf .ColKey(i) Like "*时间" Or .ColKey(i) Like "*日期" Then
                    .ColAlignment(i) = flexAlignCenterCenter
                End If
            End Select
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsExpense, Tittle, "费用信息", True
    End With
    '分组处理
    Call ExpenseSplitGroup
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExpenseSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对费用列表信息进行分组显示
    '编制:刘兴洪
    '日期:2014-05-26 16:58:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsExpense
        For i = 0 To .Cols - 1
            If i < .ColIndex("类别") And i <> .ColIndex("单据类型") Then
                If i <> .ColIndex("费用类型") Or mbytFun = 1 Then
                    .ColHidden(i) = True
                End If
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("实收金额"), , &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("应收金额"), , &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove
        If mbytFun = 1 Then
            .Outline .ColIndex("类别")
            .OutlineCol = .ColIndex("类别")
        Else
            .Outline .ColIndex("费用类型")
            .OutlineCol = .ColIndex("费用类型")
        End If
        
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 450
                '将单据号显示在项目名称列中
                If mbytFun = 1 Then
                    .Cell(flexcpText, i, .ColIndex("类别")) = strTemp
                Else
                    .Cell(flexcpText, i, .ColIndex("费用类型")) = strTemp
                End If
                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("单据类型"))
                 strTemp = strTemp & Space(2) & "费别:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("费别"))
                 strTemp = strTemp & Space(2) & "开单部门:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单部门"))
                 strTemp = strTemp & Space(2) & "开单人:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单人"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 If Val(.TextMatrix(i + 1, .ColIndex("记录性质"))) Mod 10 = 1 _
                        And Val(.TextMatrix(i + 1, .ColIndex("收费标志"))) = 1 Then   '已收费蓝色显示
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '深蓝
                        .ForeColorSel = &HC00000
                 Else
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                        .ForeColorSel = vbBlack
                 End If
                 For j = 0 To .Cols - 1
                    If j > .ColIndex("费用类型") And j < .ColIndex("应收金额") Then
                        If mbytFun = 1 Then
                            If j > .ColIndex("类别") Then
                                .Cell(flexcpText, i, j) = strTemp
                                .Cell(flexcpFontBold, i, j) = True
                            End If
                        Else
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = True
                        End If
                       '82582:李南春,2015/2/10,去掉金额中的逗号分隔符
                    ElseIf .ColIndex("实收金额") = j Then
                        .TextMatrix(i, j) = Format(Val(zlFormatNum(.TextMatrix(i, j))), gSysPara.Money_Decimal.strFormt_VB)
                    ElseIf .ColIndex("应收金额") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(zlFormatNum(.TextMatrix(i, j))), gSysPara.Money_Decimal.strFormt_VB)
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("单价")) = Format(Val(zlFormatNum(.TextMatrix(i, .ColIndex("单价")))), gSysPara.Price_Decimal.strFormt_VB)
                .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(zlFormatNum(.TextMatrix(i, .ColIndex("应收金额")))), gSysPara.Money_Decimal.strFormt_VB)
                .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(zlFormatNum(.TextMatrix(i, .ColIndex("实收金额")))), gSysPara.Money_Decimal.strFormt_VB)
            End If
        Next
        If mbytFun = 1 Then
            Call .AutoSize(.ColIndex("类别"))
        Else
            Call .AutoSize(.ColIndex("费用类型"))
        End If
        Call .AutoSize(.ColIndex("单价"))
        For j = 0 To .Cols - 1
            If j > .ColIndex("项目") And j < .ColIndex("应收金额") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

'-------------------------------------------
Private Function SetAdviceProperty(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, _
    ByVal bln独立执行 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置医嘱的相关属性
    '编制:刘兴洪
    '日期:2014-05-27 10:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyAdviceProperty As ty_adviceProperty
    Dim strSql As String, rsTemp As ADODB.Recordset


    On Error GoTo errHandle
    With tyAdviceProperty
        .lng医嘱ID = lng医嘱ID
        .lng发送号 = lng发送号
        .bln独立执行 = bln独立执行
    End With
    mTYAdviceProperty = tyAdviceProperty

    strSql = _
    " Select A.病人ID,A.主页ID,A.挂号单,A.病人科室ID,D.当前病区id,A.开嘱科室ID,A.病人来源,C.结算模式,A.诊疗类别,E.计价性质," & _
    "       Decode(A.诊疗类别,'D',Nvl(A.相关ID,A.ID),A.相关ID) as 相关ID,B.NO,B.记录性质,Nvl(B.门诊记帐,0) as 门诊记帐," & _
    "       B.执行状态,B.发送时间,Nvl(D.费别,C.费别) as 费别,d.审核标志 " & _
    " From 病人信息 C,病案主页 D," & IIf(mblnMoved, "H", "") & "病人医嘱记录 A," & IIf(mblnMoved, "H", "") & "病人医嘱发送 B,诊疗项目目录 E" & _
    " Where A.ID=B.医嘱ID And A.ID=[1] And B.发送号=[2] And A.诊疗项目ID=E.ID" & _
    " And A.病人ID=C.病人ID And A.病人ID=D.病人ID(+) And A.主页ID=D.主页ID(+)"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号)
    If rsTemp.RecordCount = 0 Then SetAdviceProperty = True: Exit Function


    With mTYAdviceProperty
        .lng相关ID = Val(Nvl(rsTemp!相关ID))
        .lng病人ID = Val(Nvl(rsTemp!病人ID))
        .lng主页Id = Val(Nvl(rsTemp!主页ID))
        .str挂号单 = Nvl(rsTemp!挂号单)
        .lng病人科室ID = Val(Nvl(rsTemp!病人科室id))
        .lng病人病区ID = Val(Nvl(rsTemp!当前病区ID))
        .lng开嘱科室ID = Val(Nvl(rsTemp!开嘱科室id))
        .int记录性质 = Val(Nvl(rsTemp!记录性质))
        .int审核标志 = Val(Nvl(rsTemp!审核标志))
        .int结算模式 = Val(Nvl(rsTemp!结算模式))
        .str诊疗类别 = Nvl(rsTemp!诊疗类别)
        .strNO = Nvl(rsTemp!NO)
        .int执行状态 = Val(Nvl(rsTemp!执行状态))
        .dat发送时间 = rsTemp!发送时间
        .str费别 = Nvl(rsTemp!费别)
        .lng计价性质 = Val(Nvl(rsTemp!计价性质))
        .str计费状态 = GetSendFeeState()
        .int病人来源 = Val(Nvl(rsTemp!病人来源))
        .bln门诊记帐 = Val(Nvl(rsTemp!门诊记帐))
         '门诊和住院医生站可发送门诊记帐，存在门诊费用记录中
         '以前的门诊医生站发送为门诊记帐时，rsTemp!门诊记帐的值为空，未修正历史数据
        .strFeeTab = "门诊费用记录"
        If .int病人来源 = 2 Then
            If .bln门诊记帐 Then
                .int病人来源 = 1    '当成门诊病人(这种情况一般是门诊留观病人)
            Else
                .strFeeTab = "住院费用记录"
            End If
        End If
        '检验组合和多部位检查项目的综合执行状态
        If (.str诊疗类别 = "C" Or .str诊疗类别 = "D") And Not .bln独立执行 Then
            strSql = "" & _
            "   Select 执行状态 From 病人医嘱发送 " & _
            "   Where 发送号=[1]  And 医嘱ID IN(Select ID From " & IIf(mblnMoved, "H", "") & "病人医嘱记录 Where (ID=[2] Or 相关ID=[2]) And 诊疗类别 In('C','D'))"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, lng发送号, IIf(.lng相关ID <> 0, .lng相关ID, .lng医嘱ID))
            strSql = ""
            Do While Not rsTemp.EOF
                If InStr(strSql, Nvl(rsTemp!执行状态, 0)) = 0 Then
                    strSql = strSql & Nvl(rsTemp!执行状态, 0)
                End If
                rsTemp.MoveNext
            Loop
            .int执行状态 = IIf(Len(strSql) = 1, Val(strSql), 3)
        End If
    End With
    SetAdviceProperty = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSendFeeState() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定医嘱某次发送之后的计费状态
    '入参:lng医嘱ID-检验组合主项目,手术主项目,第一个检验项目的医嘱ID(即在医技站中显示的项目的)
    '     lng发送号-发送号
    '     bln单独执行-组合项目是否独立执行
    '出参:
    '返回:",-1,0,1,"：其中-1=无需计费,1=已计费,0=未计费,对于门诊单据，2=部分收费,3=全部收费
    '编制:刘兴洪
    '日期:2014-05-27 09:33:50
    '说明:获取指定医嘱某次发送之后的计费状态，主要考虑一些组合医嘱有多种计费的状态
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSql As String

    On Error GoTo errH

    If mTYAdviceProperty.bln独立执行 Then
        strSql = "Select Distinct 计费状态 From 病人医嘱发送 Where 医嘱ID=[1] And 发送号=[2]"
    Else
        strSql = "Select ID From 病人医嘱记录 Where (ID=[3] Or 相关ID=[3]) And 诊疗类别=[4]"
        strSql = "Select Distinct 计费状态 From 病人医嘱发送 Where 医嘱ID IN(" & strSql & ") And 发送号=[2]"
    End If
    If mblnMoved Then
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
        strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, IIf(mTYAdviceProperty.lng相关ID <> 0, mTYAdviceProperty.lng相关ID, mTYAdviceProperty.lng医嘱ID), mTYAdviceProperty.str诊疗类别)
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "," & IIf(Val("" & rsTmp!计费状态) > 1, 1, Val("" & rsTmp!计费状态))
        rsTmp.MoveNext
    Loop
    If strSql <> "" Then GetSendFeeState = strSql & ","
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
Private Function LoadFeeDataFromAdvice(ByVal lng医嘱ID As Long, _
    ByVal lng发送号 As Long, ByVal bln独立执行 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定医嘱的主要费用及附加费用
    '入参:lng医嘱ID-当前医嘱ID
    '     lng发送号-发送号
    '     bln单独执行-组合项目是否独立执行
    '编制:刘兴洪
    '日期:2014-05-27 11:02:09
    '说明：1.包含医嘱本身的主费用及附加费用,主费用可能尚未产生
    '      2.目前单据暂不支持部份退费,所以清单中只需简单显示
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDataMoved As Boolean
    Dim dbl应收 As Double, dbl实收 As Double
    Dim strIF As String
    Dim tyAdviceProperty As ty_adviceProperty
    
    On Error GoTo errHandle

    '加载指定医嘱的明细数据
    Set mrsPrice = Nothing
    If lng医嘱ID = 0 Or lng发送号 = 0 Then
        vsExpense.Rows = 2: vsExpense.Clear 1
        vsExpense.Subtotal flexSTClear
        mTYAdviceProperty = tyAdviceProperty '115514,清除医嘱信息
        LoadFeeDataFromAdvice = True: Exit Function
    End If
        
        
    '加载指定医嘱的相关属性
    Call SetAdviceProperty(lng医嘱ID, lng发送号, bln独立执行)

    blnDataMoved = mblnMoved
    If Not blnDataMoved Then
        blnDataMoved = gobjDatabase.DateMoved(mTYAdviceProperty.dat发送时间)
    End If
    
    If LoadDataFromAdvices(blnDataMoved) = False Then Exit Function
    LoadFeeDataFromAdvice = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetNotFeeSQL(ByVal blnDataMoved As Boolean, ByRef str医嘱IDs As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:追加未计费部分的费用
    '出参:str医嘱IDs-返回涉及的医嘱IDs
    '返回:返回SQL
    '编制:刘兴洪
    '日期:2014-05-27 15:05:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, rsHead As ADODB.Recordset
    Dim int来源 As Integer, lng执行部门ID As Long
    Dim str登记时间 As String

    On Error GoTo errHandle

    str医嘱IDs = ""

    '存在未计费状态,直接读取收费关系显示
    If InStr(mTYAdviceProperty.str计费状态, ",0,") = 0 Then Exit Function

    Call LoadAdvicePrice(mblnMoved)

    If mrsPrice Is Nothing Then Exit Function
    If mrsPrice.State <> 1 Then Exit Function

    int来源 = IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)
    strSql = "" & _
    "   Select A.开嘱科室ID,B.名称 as 开嘱科室,Nvl(停嘱医生,开嘱医生) as 开嘱医生,病人来源,开始执行时间  " & _
    "   From 病人医嘱记录 A,部门表 B  " & _
    "   Where A.开嘱科室ID=B.ID And A.ID=[1]"
    If blnDataMoved Then
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
    End If

    Set rsHead = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng医嘱ID)
    If rsHead.EOF Then Exit Function
    str登记时间 = Format(Nvl(rsHead!开始执行时间), "yyyy-mm-dd HH:MM:SS")
    If str登记时间 = "" Then str登记时间 = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Dim lng项目id As Long, int父号 As Long, i As Long
    strSql = ""
    With mrsPrice
        If mrsPrice.RecordCount <> 0 Then .MoveFirst
        For i = 1 To .RecordCount
            If lng项目id <> !收费细目ID Then int父号 = i

            lng执行部门ID = Get收费执行科室ID(mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, !类别, !收费细目ID, !执行科室, mTYAdviceProperty.lng病人科室ID, Nvl(!开嘱科室id, 0), int来源)

            strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & _
            " Select 1 as 费用类型,'[未计费]' as NO," & mTYAdviceProperty.int记录性质 & " as 记录性质,1 as 记录状态," & _
            "       '" & mTYAdviceProperty.str费别 & "' as 费别," & i & " as 序号," & IIf(int父号 = i, "-NULL", int父号) & " as 价格父号," & _
                    !医嘱ID & " as 医嘱序号,'" & !类别 & "' as 收费类别," & !收费细目ID & " as 收费细目ID," & _
                  rsHead!开嘱科室id & " as 开单科室ID ," & "'" & Nvl(rsHead!开嘱医生) & "'  as 开单人," & lng执行部门ID & " as 执行部门ID," & _
            "       0 as 执行状态," & !收入项目ID & " as 收入项目ID,1 as 付数," & !数量 & " as 数次," & !单价 & " as 标准单价," & _
                     !应收 & " as 应收金额," & !实收 & " as 实收金额, " & _
            "       To_Date('" & str登记时间 & "','YYYY-MM-DD HH24:MI:SS') as 登记时间, " & _
            "       To_Date('" & str登记时间 & "','YYYY-MM-DD HH24:MI:SS') as 发生时间, '" & _
                    IIf(Val(Nvl(rsHead!病人来源)) = 3, Nvl(rsHead!开嘱医生), UserInfo.姓名) & "' as 操作员 , 0 as 已收费" & _
            " From Dual"
            lng项目id = !收费细目ID
             str医嘱IDs = str医嘱IDs & "," & !医嘱ID
            .MoveNext
        Next
        If strSql = "" Then Exit Function
        str医嘱IDs = Mid(str医嘱IDs, 2) '取检验组合中涉及的医嘱ID
    End With
    GetNotFeeSQL = " Union ALL " & strSql
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDataFromAdvices(ByVal blnDataMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:追加医嘱对应的主费用和附加费用
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-27 16:10:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strSQL1 As String, strSQL2 As String
    Dim strFeeTab As String, strIDs As String
    Dim bln药房单位 As String, str药房单位 As String, str药房包装 As String
    Dim strWith As String, int来源 As Integer
    Dim str医嘱IDs As String, rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    strSql = ""
    strFeeTab = mTYAdviceProperty.strFeeTab
    If strFeeTab = "" Then Exit Function
     '药品单位
    int来源 = IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)
    bln药房单位 = Val(gobjDatabase.GetPara("药品单位", glngSys, p医嘱附费管理)) <> 0
    If int来源 = 1 Then
        str药房单位 = "门诊单位": str药房包装 = "门诊包装"
    Else
        str药房单位 = "住院单位": str药房包装 = "住院包装"
    End If

    '包含检查部位，附加手术，检验组合的费用
    If mTYAdviceProperty.bln独立执行 Then
        str医嘱IDs = "Select [1] From Dual"
    Else
        '多部位检查，包含所有部位和方法的
        '1.检查;手术
        '2.检验
        str医嘱IDs = _
        " Select ID From 病人医嘱记录 Where ID=[1] Or (相关ID=[1] And 诊疗类别 IN('F','D'))" & _
        " Union ALL " & _
        " Select ID From 病人医嘱记录 Where 诊疗类别='C' And 相关ID=[2]"
    End If
    If mblnMoved Then
        str医嘱IDs = Replace(str医嘱IDs, "病人医嘱记录", "H病人医嘱记录")
    ElseIf blnDataMoved Then
        str医嘱IDs = str医嘱IDs & " Union ALL " & Replace(str医嘱IDs, "病人医嘱记录", "H病人医嘱记录")
    End If

    '存在已计费状态,应该可以直接读取主费用部份
    '只有一张单据,可能含其它医嘱费用;
    '显示原始单据信息和剩余部份金额
    If InStr(mTYAdviceProperty.str计费状态, ",1,") > 0 Then
        '包含检查部位，附加手术，检验组合的费用
        strSQL1 = _
        " Select 1 as 费用类型,A.记录性质,Decode(B.记录状态,0,0,1) as 已收费," & _
        "       A.NO,B.费别,Sum(B.应收金额) as 应收金额,Sum(B.实收金额) as 实收金额," & _
        "       C.名称 as 开单科室,B.开单人,Max(Decode(Floor(x.记录性质/10), 0, x.登记时间, Null)) As 登记时间, " & _
        "       Max(Decode(Floor(x.记录性质/10), 0, Nvl(x.操作员姓名, x.划价人), Null)) As 操作员 " & _
        " From 病人医嘱发送 A," & strFeeTab & " B,部门表 C," & strFeeTab & " X" & _
        " Where A.医嘱ID IN(" & str医嘱IDs & ") And A.发送号=[3]" & _
        "   And A.NO=B.NO And A.记录性质=Decode(B.记录性质,11,1,B.记录性质)" & _
        "   And A.医嘱ID=B.医嘱序号+0 And B.开单部门ID=C.ID" & _
        "   And B.NO=X.NO And B.记录性质=X.记录性质 And B.序号=X.序号 And X.记录状态 IN(0,1,3)" & _
        " Group by A.记录性质,Decode(B.记录状态,0,0,1),A.NO,B.费别,C.名称,B.开单人"
        If mblnMoved Then
            strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
            strSQL1 = Replace(strSQL1, strFeeTab, "H" & strFeeTab)
        ElseIf blnDataMoved Then
            strSQL2 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
            strSQL2 = Replace(strSQL2, strFeeTab, "H" & strFeeTab)
            strSQL1 = strSQL1 & " Union ALL " & strSQL2
            strSQL1 = _
                " Select A.费用类型,A.记录性质,A.已收费,A.NO,A.费别," & _
                "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
                " A.开单科室,A.开单人,A.登记时间,A.操作员 From (" & strSQL1 & ") A" & _
                " Group by A.费用类型,A.记录性质,A.已收费,A.NO,A.费别,A.开单科室,A.开单人,A.登记时间,A.操作员"
        End If
        strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & strSQL1
    End If

    '附费用部份(显示原始单据信息和剩余部份金额)
    strSQL1 = _
    " Select 2 as 费用类型,A.记录性质,Decode(B.记录状态,0,0,1) as 已收费," & _
    "       A.NO,B.费别,Sum(B.应收金额) as 应收金额,Sum(B.实收金额) as 实收金额," & _
    "       C.名称 as 开单科室,B.开单人,Max(Decode(Floor(x.记录性质/10), 0, x.登记时间, Null)) As 登记时间, " & _
    "       Max(Decode(Floor(x.记录性质/10), 0, Nvl(x.操作员姓名, x.划价人), Null)) As 操作员 " & _
    " From 病人医嘱附费 A," & strFeeTab & " B,部门表 C," & strFeeTab & " X" & _
    " Where A.医嘱ID IN(" & str医嘱IDs & ") And A.发送号=[3]" & _
    "       And A.NO=B.NO And A.记录性质=Decode(B.记录性质,11,1,B.记录性质)" & _
    "       And A.医嘱ID=B.医嘱序号 And B.开单部门ID=C.ID" & _
    "       And B.NO=X.NO And B.记录性质=X.记录性质 And B.序号=X.序号 And X.记录状态 IN(0,1,3)" & _
    " Group by A.记录性质,Decode(B.记录状态,0,0,1),A.NO,B.费别,C.名称,B.开单人"

    If mblnMoved Then
        strSQL1 = Replace(strSQL1, "病人医嘱附费", "H病人医嘱附费")
        strSQL1 = Replace(strSQL1, strFeeTab, "H" & strFeeTab)
    ElseIf blnDataMoved Then
        strSQL2 = Replace(strSQL1, "病人医嘱附费", "H病人医嘱附费")
        strSQL2 = Replace(strSQL2, strFeeTab, "H" & strFeeTab)
        strSQL1 = strSQL1 & " Union ALL " & strSQL2
        strSQL1 = _
        " Select A.费用类型,A.记录性质,A.已收费,A.NO,A.费别," & _
        "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
        "       A.开单科室,A.开单人,A.登记时间,A.操作员 From (" & strSQL1 & ") A" & _
        " Group by A.费用类型,A.记录性质,A.已收费,A.NO,A.费别,A.开单科室,A.开单人,A.登记时间,A.操作员"
    End If
    strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & strSQL1
    strWith = "With 单据列表 as (" & strSql & ") "

    '85232:李南春,2015/5/29,读取病人费用记录时排除已作废的记录
    '以原始记录为准求(审核的登记时间和执行状态)
    strSql = _
        " Select C.费用类型,A.NO,Decode(A.记录性质,11,1,A.记录性质) As 记录性质,A.记录状态,A.费别, " & _
        "        A.序号,A.价格父号,A.医嘱序号,A.收费类别,A.收费细目ID,A.开单部门ID,A.开单人,A.执行部门ID, " & _
        "        Max(a.执行状态) Over(Partition By a.记录性质,a.No,a.序号) As 执行状态, " & _
        "        A.收入项目ID,A.付数,A.数次,A.标准单价,A.应收金额,A.实收金额, " & _
        "        A.发生时间,C.登记时间," & _
        "        C.操作员 as 操作员姓名,C.已收费" & _
        " From " & strFeeTab & " A,单据列表 C" & _
        " Where Decode(A.记录性质,11,1,A.记录性质)= C.记录性质 And A.NO=C.NO "

    If mblnMoved Then
        strSql = Replace(strSql, strFeeTab, "H" & strFeeTab)
    ElseIf blnDataMoved Then
        strSql = strSql & " Union ALL " & Replace(strSql, strFeeTab, "H" & strFeeTab)
    End If
    strSql = strSql & GetNotFeeSQL(blnDataMoved, strIDs)

    strSql = strWith & vbCrLf & strSql
    '已删除或退费销帐,则不显示
    '80752,冉俊明,2014-12-23,显示出已删除或退费销帐记录
    '    " Having Nvl(Sum(A.应收金额),0)<>0 Or Nvl(Sum(A.实收金额),0)<>0"
    strSql = "" & _
    " Select　decode(nvl(A.费用类型,1),1,'主费用','附加费用') as 费用类型, " & _
    "       A.记录性质,A.已收费 as 收费标志,decode(A.记录性质,1,'收费单','记帐单') as 单据类型, " & _
    "       A.NO as 单据号,A.收费细目ID," & _
    "       A.费别,L.名称 as 开单部门,A.开单人, " & _
    "       C.名称 as 类别,Nvl(A.价格父号,A.序号) as 序号," & _
    "       Nvl(F.名称,B.名称)||Decode(B.规格,NULL,NULL,' '||B.规格) as 项目," & _
    "       A.标准单价" & IIf(bln药房单位, "*Nvl(E." & str药房包装 & ",1)", "") & " as 单价," & _
    "       Sum(Nvl(A.付数,1)*A.数次" & IIf(bln药房单位, "/Nvl(E." & str药房包装 & ",1)", "") & ") as 数量," & _
            IIf(bln药房单位, "Decode(E.药品ID,NULL,B.计算单位,E." & str药房单位 & ")", "B.计算单位") & " as 计算单位," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额,D.名称 as 执行部门," & _
    "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行') as 执行情况, " & _
    "       Nvl(A.执行状态,0) as 执行状态,a.收费类别," & _
    "       to_Char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间,A.操作员姓名" & _
    " From (" & strSql & ") A,收费项目目录 B,收费项目类别 C," & _
    "      部门表 D,部门表 L,药品规格 E,收费项目别名 F" & _
    " Where A.收费细目ID=B.ID And A.收费类别=C.编码  And A.开单部门ID=L.ID(+) And A.执行部门ID=D.ID(+)" & _
    "       And B.ID=E.药品ID(+) And A.收费细目ID=F.收费细目ID(+)" & _
    "       And F.码类(+)=1 And F.性质(+)=[4] And A.医嘱序号+0 IN(" & str医嘱IDs & ")" & _
    " Group by A.收费细目ID,A.费用类型,A.记录性质,A.已收费,A.NO,A.费别,L.名称,A.开单人,A.操作员姓名, " & _
    "       to_Char(A.登记时间,'yyyy-mm-dd hh24:mi:ss'),Nvl(A.价格父号,A.序号),C.名称 , " & _
    "       Nvl(F.名称,B.名称),B.规格,B.计算单位,D.名称," & _
    "       A.标准单价,Nvl(A.执行状态,0),a.收费类别,E.药品ID,Nvl(E." & str药房包装 & ",1),E." & str药房单位 & _
    " Order by  费用类型 Desc,记录性质,登记时间 Desc, No, 序号"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng医嘱ID, _
        mTYAdviceProperty.lng相关ID, mTYAdviceProperty.lng发送号, IIf(gSysPara.byt药品名称显示 = 0, 1, 3))

    With vsExpense
        .Redraw = flexRDNone
        .Cols = 1: .FixedCols = 0
        .Rows = 2
        .MergeRow(1) = False
        Set .DataSource = rsTemp
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        If rsTemp.RecordCount = 0 Then
            .Subtotal flexSTClear
            .Rows = 2
            .Clear 1
        Else
            Call SetExpenseGridProperty '设置网格属性
        End If
    End With
    
    
    zl_vsGrid_Para_Save mlngModule, vsExpense, Tittle, "费用信息", True
    
    vsExpense.Redraw = flexRDBuffered
    LoadDataFromAdvices = True
    Exit Function
errHandle:
    vsExpense.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadAdvicePrice(ByVal blnMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定医嘱的计价关系到临时记录集
    '返回:读取计价关系成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-27 11:59:19
    '说明:要计算的项目应该不是叮嘱,院外执行,无需计费
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl数量 As Double, bln附加手术 As Boolean
    Dim strSql As String, strIF As String, strPrice As String
    Dim blnHaveSub As Boolean, lng主收入ID As Long
    Dim cur合计 As Currency, i As Long, j As Long
    Dim str药品价格等级 As String, str卫材价格等级 As String, str普通价格等级 As String
    Dim strWherePriceGrade As String

    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "开嘱科室ID", adBigInt
    mrsPrice.Fields.Append "类别", adVarChar, 10
    mrsPrice.Fields.Append "收费细目ID", adBigInt
    mrsPrice.Fields.Append "计算单位", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "附加手术", adInteger
    mrsPrice.Fields.Append "执行科室", adInteger
    mrsPrice.Fields.Append "收入项目ID", adBigInt
    mrsPrice.Fields.Append "收据费目", adVarChar, 50, adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble
    mrsPrice.Fields.Append "单价", adDouble
    mrsPrice.Fields.Append "应收", adCurrency
    mrsPrice.Fields.Append "实收", adCurrency
    mrsPrice.Fields.Append "从项", adInteger

    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open

    On Error GoTo errH

    '读取要计算主费用的医嘱记录
    '包含检查部位，附加手术，检验组合的费用
    If mTYAdviceProperty.bln独立执行 Then
        strIF = " And B.ID=[1]"
    Else
        'F,D:检查,手术
        'C:检验
        strIF = " And (B.ID=[1] Or (B.相关ID=[1] And B.诊疗类别 IN('F','D')) Or (B.相关ID=[2] And B.诊疗类别='C'))"
    End If

    strSql = _
    " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID,A.执行部门ID," & _
    "        Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量,B.标本部位,B.检查方法,B.执行标记,b.病人ID,b.主页ID" & _
    " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
    " Where Nvl(A.计费状态,0)=0 " & strIF & _
    "   And A.医嘱ID=B.ID And A.发送号=[3]" & _
    "   And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
    " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID," & _
    "       A.执行部门ID,A.发送数次,B.标本部位,B.检查方法,B.执行标记,b.病人ID,b.主页ID" & _
    " Having Nvl(A.发送数次,Sum(Nvl(C.本次数次,0)))<>0" & _
    " Order by 序号"
    If blnMoved Then
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
        strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
        strSql = Replace(strSql, "病人医嘱执行", "H病人医嘱执行")
    End If
    Set rsAdvice = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng相关ID, mTYAdviceProperty.lng发送号)
    
    If rsAdvice.RecordCount > 0 Then
        '医嘱ID只有一个，肯定是同一个病人
        If GetPriceGradeStartType() > 0 Then
            Call GetPriceGrade(gstrNodeNo, Val(Nvl(rsAdvice!病人ID)), Val(Nvl(rsAdvice!主页ID)), "", str药品价格等级, str卫材价格等级, str普通价格等级)
        End If
        If str药品价格等级 <> "" Or str卫材价格等级 <> "" Or str普通价格等级 <> "" Then
            strWherePriceGrade = _
                "      And ((Instr(';5;6;7;', ';' || c.类别 || ';') > 0 And b.价格等级 = [8])" & vbNewLine & _
                "            Or (Instr(';4;', ';' || c.类别 || ';') > 0 And b.价格等级 = [9])" & vbNewLine & _
                "            Or (Instr(';4;5;6;7;', ';' || c.类别 || ';') = 0 And b.价格等级 = [10])" & vbNewLine & _
                "            Or (b.价格等级 Is Null" & vbNewLine & _
                "                And Not Exists (Select 1" & vbNewLine & _
                "                                From 收费价目" & vbNewLine & _
                "                                Where b.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                      And ((Instr(';5;6;7;', ';' || c.类别 || ';') > 0 And 价格等级 = [8])" & vbNewLine & _
                "                                            Or (Instr(';4;', ';' || c.类别 || ';') > 0 And 价格等级 = [9])" & vbNewLine & _
                "                                            Or (Instr(';4;5;6;7;', ';' || c.类别 || ';') = 0 And 价格等级 = [10])))))"
        Else
            strWherePriceGrade = " And b.价格等级 Is Null "
        End If
    End If
    
    For i = 1 To rsAdvice.RecordCount
        dbl数量 = Nvl(rsAdvice!数量, 0)

        '读取对应的收费价目:只读取固定对照,且不是变价的对照
        bln附加手术 = (rsAdvice!诊疗类别 = "F" And Not IsNull(rsAdvice!相关ID))
        '由于没有加部位等条件，所以要用Distinct
        strPrice = "" & _
        "Select * From (" & _
        "   Select Distinct C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,c.适用科室id" & _
        "               ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
        "   From 诊疗收费关系 C Where C.诊疗项目ID=[1]" & _
        "           And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = [6] And C.病人来源 = [7])" & _
        "   ) Where Nvl(适用科室id, 0) = Top"

        strSql = _
        " Select Nvl(A.从属项目,0) as 从项,A.收费项目ID,A.收费数量,B.收入项目ID,D.收据费目," & _
        "       C.类别,C.计算单位,C.执行科室,Decode(C.是否变价,1,B.缺省价格,B.现价) as 单价,C.屏蔽费别," & _
                IIf(bln附加手术, "Nvl(B.附术收费率,100)/100", "1") & " as 附术率" & _
        " From (" & strPrice & ") A,收费价目 B,收费项目目录 C,收入项目 D," & _
        "      (Select [1] as 诊疗项目ID,Decode([2],0,Null,[2]) as 相关ID," & _
        "               Decode([3],'None',Null,[3]) as 标本部位,Decode([4],'None',Null,[4]) as 检查方法,[5] as 执行标记 From Dual " & _
        "       ) X" & _
        " Where A.诊疗项目ID=X.诊疗项目ID" & _
        "       And (   X.相关ID is Null And X.执行标记 IN(1,2) And A.费用性质=1" & _
        "               Or X.标本部位=A.检查部位 And X.检查方法=A.检查方法 And Nvl(A.费用性质,0)=0" & _
        "               Or X.检查方法 is Null And Nvl(A.费用性质,0)=0 And A.检查部位 is Null And A.检查方法 is Null)" & _
        "       And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And B.收入项目ID=D.ID" & _
        "       And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                strWherePriceGrade & vbNewLine & _
        "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        "       And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        "       And Nvl(A.固有对照,0)=1 And Nvl(C.是否变价,0)=0" & _
        " Order By 收费项目ID,从项"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, Val(rsAdvice!诊疗项目ID), Val(Nvl(rsAdvice!相关ID, 0)), _
            CStr(Nvl(rsAdvice!标本部位, "None")), CStr(Nvl(rsAdvice!检查方法, "None")), Val(Nvl(rsAdvice!执行标记, 0)), _
            mlng执行科室ID, mTYAdviceProperty.int病人来源, str药品价格等级, str卫材价格等级, str普通价格等级)

        blnHaveSub = False: lng主收入ID = 0: cur合计 = 0
        If Not rsTemp.EOF And gSysPara.bln从项汇总折扣 Then
            rsTemp.Filter = "从项=1"
            If Not rsTemp.EOF Then blnHaveSub = True
            rsTemp.Filter = "从项=0"
            If Not rsTemp.EOF Then lng主收入ID = rsTemp!收入项目ID
            rsTemp.Filter = 0
        End If

        For j = 1 To rsTemp.RecordCount
            mrsPrice.AddNew
            mrsPrice!医嘱ID = rsAdvice!医嘱ID
            mrsPrice!开嘱科室id = rsAdvice!开嘱科室id
            mrsPrice!类别 = rsTemp!类别
            mrsPrice!收费细目ID = rsTemp!收费项目ID
            mrsPrice!计算单位 = Nvl(rsTemp!计算单位)
            mrsPrice!附加手术 = IIf(bln附加手术, 1, 0)
            mrsPrice!执行科室 = Nvl(rsTemp!执行科室, 0)
            mrsPrice!收入项目ID = rsTemp!收入项目ID
            mrsPrice!收据费目 = rsTemp!收据费目
            mrsPrice!单价 = Format(Nvl(rsTemp!单价, 0), gSysPara.Price_Decimal.strFormt_VB)
            mrsPrice!数量 = Format(Nvl(rsTemp!收费数量, 0) * dbl数量, "0.00000")
            mrsPrice!应收 = Format(mrsPrice!数量 * mrsPrice!单价 * rsTemp!附术率, gSysPara.Money_Decimal.strFormt_VB)
            mrsPrice!从项 = rsTemp!从项
            If gSysPara.bln从项汇总折扣 And blnHaveSub Then
                mrsPrice!实收 = mrsPrice!应收
                cur合计 = cur合计 + mrsPrice!实收
            ElseIf Nvl(rsTemp!屏蔽费别, 0) = 0 Then
                mrsPrice!实收 = Format(ActualMoney(mTYAdviceProperty.str费别, mrsPrice!收入项目ID, mrsPrice!应收, _
                    rsTemp!收费项目ID, Nvl(rsAdvice!执行部门ID, 0), Nvl(rsTemp!收费数量, 0) * dbl数量, 0), gSysPara.Money_Decimal.strFormt_VB)
            Else
                mrsPrice!实收 = mrsPrice!应收
            End If
            mrsPrice.Update
            rsTemp.MoveNext
        Next

        If gSysPara.bln从项汇总折扣 And blnHaveSub And lng主收入ID <> 0 Then
            cur合计 = Format(ActualMoney(mTYAdviceProperty.str费别, lng主收入ID, cur合计), gSysPara.Money_Decimal.strFormt_VB) - cur合计
            mrsPrice.Filter = "从项=0"
            mrsPrice!实收 = Nvl(mrsPrice!实收, 0) + cur合计
            mrsPrice.Update
            mrsPrice.Filter = 0
        End If
        rsAdvice.MoveNext
    Next
    If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
    LoadAdvicePrice = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Set mrsPrice = Nothing
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,医嘱附费管理
Public Property Get Tittle() As String
    Tittle = m_Tittle
End Property

Public Property Let Tittle(ByVal New_Tittle As String)
    m_Tittle = New_Tittle
    PropertyChanged "Tittle"
End Property


'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property


Public Sub zlPrintData(bytStyle As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出费用清单
    '入参:bytStyle=1-打印,2-预览,3-输出到Excel
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-29 17:00:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mTYAdviceProperty.lng病人ID = 0 Then Exit Sub
    
    strSql = "" & _
    "Select Nvl(Nvl(B.姓名, C.姓名), A.姓名) 姓名, Nvl(Nvl(B.性别, C.性别), A.性别) 性别, Nvl(Nvl(B.年龄, C.年龄), A.年龄) 年龄, A.门诊号, B.住院号" & vbNewLine & _
    "From 病人信息 A, 病案主页 B, 病人挂号记录 C" & vbNewLine & _
    "Where A.病人id = B.病人id(+) And B.主页id(+) = [2] And A.病人id = C.病人id(+) And A.门诊号 = C.门诊号(+) And A.病人id = [1]"

    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id)
    If rsTmp.EOF Then Exit Sub
    
    '表头
    objOut.Title.Text = IIf(mbytFun = 1, "医院费用清单", "医院附费清单")
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add "病人：" & Nvl(rsTmp!姓名) & " 性别：" & Nvl(rsTmp!性别) & " 年龄：" & Nvl(rsTmp!年龄)
    If mTYAdviceProperty.lng主页Id <> 0 Then
        objRow.Add "住院号：" & Nvl(rsTmp!住院号)
    Else
        objRow.Add "门诊号：" & Nvl(rsTmp!门诊号)
    End If
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(gobjDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsExpense
    
    '输出
    vsExpense.Redraw = False
    lngRow = vsExpense.Row: lngCol = vsExpense.Col
        
    strWidth = ""
    For i = 0 To vsExpense.Cols - 1
        strWidth = strWidth & "," & vsExpense.ColWidth(i)
        If i <= vsExpense.FixedCols - 1 Or vsExpense.ColHidden(i) Then
            vsExpense.ColWidth(i) = 0
        End If
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsExpense.Cols - 1
        vsExpense.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    vsExpense.Row = lngRow: vsExpense.Col = lngCol
    vsExpense.Redraw = True
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Public Function zlBuildMainExpense(Optional ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:生成主费用
    '返回:生成成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objParent As Object
    If mbytFun = 1 Then Exit Function
    Set objParent = frmMain
    If frmMain Is Nothing Then Set objParent = mfrmParent
    If InStr(mTYAdviceProperty.str计费状态, ",-1,") > 0 Then
        zlBuildMainExpense = FuncFeeMainAppend(objParent)
    Else
        zlBuildMainExpense = FuncFeeMain
    End If
End Function

Private Function FuncFeeMainAppend(ByVal frmMain As Object, Optional strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入主费用
    '出参:strOutNos-保存成功的单据号
    '返回:生成成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:42:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
         IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), mTYAdviceProperty.int记录性质, mTYAdviceProperty.lng开嘱科室ID, mTYAdviceProperty.lng病人科室ID, 0, "", mTYAdviceProperty.strNO, "", "", , , , strOutNos, _
         , , mobjSquareCard) Then
         FuncFeeMainAppend = True
         Call RefreshExpenseData
    End If
End Function

Private Function FuncFeeMain(Optional strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:生成主费用
    '出参:strOutNos-保存成功的单据号
    '返回:生成成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:45:09
    '---------------------------------------------------------------------------------------------------------------------------------------------

 
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim int来源 As Integer, lng医嘱ID As Long
    Dim int价格父号 As Integer, lng项目id As Long, lng执行部门ID As Long
    Dim lng病人病区ID As Long, lng病人科室ID As Long, lng类别ID As Long
    Dim arrSQL As Variant, arrCountSQL As Variant, strSql As String, strDate As String, i As Long, j As Long
    Dim int保险项目否 As Integer, lng保险大类ID As Long, str保险编码 As String, cur统筹金额 As Currency, str费用类型 As String
    Dim lng开嘱科室ID As Long, str开嘱医生 As String, int序号 As Integer, strMsg As String
    Dim int父序号 As Integer, strTmp As String
    Dim blnTrans  As Boolean
    
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mrsPrice Is Nothing Then Exit Function
    If mrsPrice.RecordCount = 0 Then
        MsgBox "该执行项目没有可以计费的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mTYAdviceProperty.int记录性质 = 1 Then
        If BillExistBalance(mTYAdviceProperty.strNO) Then
            MsgBox "单据 " & mTYAdviceProperty.strNO & " 已经收费，不能再生成这张单据的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf mTYAdviceProperty.int记录性质 = 2 Then
        '住院出院病人费用限制
        If mTYAdviceProperty.int病人来源 = 2 Then
            If Not PatiCanBilling(mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mstrPrivsAnnexFee, p医嘱附费管理) Then Exit Function
        End If
    End If
    
    If MsgBox("确实要生成该项目的主费用吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
            
    int来源 = IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)
    
    Screen.MousePointer = 11
    
    '获取病人的信息
    strSql = "Select Nvl(Nvl(B.姓名, C.姓名), A.姓名) 姓名, Nvl(Nvl(B.性别, C.性别), A.性别) 性别, Nvl(Nvl(B.年龄, C.年龄), A.年龄) 年龄," & vbNewLine & _
            "       Nvl(B.费别, A.费别) As 费别, A.门诊号, B.住院号, Nvl(A.当前床号, B.出院病床) As 床号, Nvl(A.当前病区id, B.当前病区id) As 病人病区id," & vbNewLine & _
            "       Nvl(A.当前科室id, B.出院科室id) As 病人科室id, Nvl(B.险类, A.险类) As 险类, D.编码 As 付款码" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人挂号记录 C, 医疗付款方式 D" & vbNewLine & _
            "Where A.病人id = B.病人id(+) And B.主页id(+) = [2] And A.病人id = C.病人id(+) And A.门诊号 = C.门诊号(+) And A.医疗付款方式 = D.名称(+) And A.病人id=[1]"

    On Error GoTo errH
    Set rsPati = gobjDatabase.OpenSQLRecord(strSql, Tittle, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id)
    
    '可能对照费用为药品费用
    If mTYAdviceProperty.int记录性质 = 1 Then
        lng类别ID = ExistIOClass(8) '门诊划价单
    Else
        lng类别ID = ExistIOClass(9) '门诊/住院记帐单
    End If
    
    '可能发送时已自动生成了部份主费用,现在是手工生成剩余部份。
    '1.因为单据号相同,所以要保持序号连续
    '2.如果是生成收费划价单，要保证一张单据中登记时间相同(不然收费无法处理)
    '3.第2点的情况，如果部份主费用已经收费，则不允许再生成主费用
    int序号 = GetBillMax序号(mTYAdviceProperty.strNO, mTYAdviceProperty.int记录性质, strDate, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1))
    If mTYAdviceProperty.int记录性质 = 2 Or strDate = "" Then
        strDate = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strDate = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    arrSQL = Array()
    arrCountSQL = Array()
    With mrsPrice
        .MoveFirst
        For i = 1 To .RecordCount
            '获取对应的医嘱信息
            If lng医嘱ID <> !医嘱ID Then
                strSql = "Select 医嘱期效,病人科室ID,开嘱科室ID,开嘱医生,婴儿,执行频次,计价特性 From 病人医嘱记录 Where ID=[1]"
                Set rsAdvice = gobjDatabase.OpenSQLRecord(strSql, Tittle, Val(!医嘱ID))
                
                '将当前这条计费医嘱标记为已计费
                ReDim Preserve arrCountSQL(UBound(arrCountSQL) + 1)
                arrCountSQL(UBound(arrCountSQL)) = "ZL_病人医嘱发送_计费(" & !医嘱ID & "," & mTYAdviceProperty.lng发送号 & ")"
                
                int父序号 = 0
            End If
            lng医嘱ID = !医嘱ID
            
            '病人病区科室
            lng病人病区ID = Nvl(rsPati!病人病区ID, 0)
            lng病人科室ID = Nvl(rsPati!病人科室id, 0)
            If lng病人科室ID = 0 Then
                lng病人病区ID = Nvl(rsAdvice!病人科室id, 0)
                lng病人科室ID = Nvl(rsAdvice!病人科室id, 0)
            End If
            If lng病人科室ID = 0 Then
                lng病人病区ID = UserInfo.部门ID
                lng病人科室ID = UserInfo.部门ID
            End If
            
            '开单科室及开单人
            lng开嘱科室ID = rsAdvice!开嘱科室id
            str开嘱医生 = rsAdvice!开嘱医生
            
            '每个收费项目的处理
            If lng项目id <> !收费细目ID Then
                int价格父号 = int序号 '获取价格父号
                If !从项 = 0 Then int父序号 = int序号  '读取Price时是分医嘱将主项排在前面的
                lng执行部门ID = Get收费执行科室ID(mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, !类别, !收费细目ID, !执行科室, Nvl(rsAdvice!病人科室id, 0), Nvl(rsAdvice!开嘱科室id, 0), int来源)
                            
                '获取保险项目信息
                If int来源 = 2 And Not IsNull(rsPati!险类) Then
                    strMsg = gclsInsure.GetItemInsure(mTYAdviceProperty.lng病人ID, !收费细目ID, !实收, False, rsPati!险类, "||" & !数量)
                    If strMsg <> "" Then
                        int保险项目否 = Val(Split(strMsg, ";")(0))
                        lng保险大类ID = Val(Split(strMsg, ";")(1))
                        cur统筹金额 = Format(Val(Split(strMsg, ";")(2)), gSysPara.Money_Decimal.strFormt_VB)
                        str保险编码 = CStr(Split(strMsg, ";")(3))
                        If UBound(Split(strMsg, ";")) >= 5 Then
                            If Split(strMsg, ";")(5) <> "" Then
                                str费用类型 = Split(strMsg, ";")(5)
                            End If
                        End If
                    End If
                End If
            End If
            lng项目id = !收费细目ID
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If int来源 = 1 Then
                If mTYAdviceProperty.int记录性质 = 1 Then
                    '生成门诊划价单据
                    arrSQL(UBound(arrSQL)) = lng项目id & ";" & _
                        "zl_门诊划价记录_Insert('" & mTYAdviceProperty.strNO & "'," & int序号 & "," & mTYAdviceProperty.lng病人ID & ",NULL," & _
                        IIf(IsNull(rsPati!门诊号), "NULL", "'" & rsPati!门诊号 & "'") & ",'" & Nvl(rsPati!付款码) & "','" & Nvl(rsPati!姓名) & "'," & _
                        "'" & Nvl(rsPati!性别) & "','" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!费别) & "',NULL," & _
                        lng病人科室ID & "," & lng开嘱科室ID & ",'" & str开嘱医生 & "'," & _
                        IIf(Val(Nvl(!从项)) = 1, ZVal(int父序号), "NULL") & "," & lng项目id & ",'" & !类别 & "','" & !计算单位 & "',NULL,1," & !数量 & "," & _
                        !附加手术 & "," & ZVal(lng执行部门ID) & "," & IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & _
                        !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & !应收 & "," & !实收 & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.姓名 & "',NULL," & _
                        !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ",1,'" & str保险编码 & "','" & str费用类型 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & "," & _
                        "NULL,0,NULL,NULL," & ZVal(lng病人病区ID) & ")"
                Else
                    '生成门诊记帐单据
                    arrSQL(UBound(arrSQL)) = lng项目id & ";" & _
                        "zl_门诊记帐记录_Insert('" & mTYAdviceProperty.strNO & "'," & int序号 & "," & mTYAdviceProperty.lng病人ID & "," & _
                        IIf(IsNull(rsPati!门诊号), "NULL", "'" & rsPati!门诊号 & "'") & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                        "'" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!费别) & "',NULL," & ZVal(rsAdvice!婴儿) & "," & _
                        lng病人科室ID & "," & lng开嘱科室ID & "," & _
                        "'" & str开嘱医生 & "'," & IIf(!从项 = 1, ZVal(int父序号), "NULL") & "," & lng项目id & ",'" & !类别 & "'," & _
                        "'" & !计算单位 & "',1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                        IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                        !应收 & "," & !实收 & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "',NULL,NULL," & !医嘱ID & "," & _
                        "'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ",1,NULL,0,NULL," & ZVal(mTYAdviceProperty.lng主页Id) & "," & ZVal(lng病人病区ID) & ")"
                End If
            Else
                '生成住院记帐单据
                arrSQL(UBound(arrSQL)) = lng项目id & ";" & _
                    "zl_住院记帐记录_Insert('" & mTYAdviceProperty.strNO & "'," & int序号 & "," & mTYAdviceProperty.lng病人ID & "," & ZVal(mTYAdviceProperty.lng主页Id) & "," & _
                    IIf(IsNull(rsPati!住院号), "NULL", "'" & rsPati!住院号 & "'") & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                    "'" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!床号) & "','" & Nvl(rsPati!费别) & "'," & _
                    lng病人病区ID & "," & lng病人科室ID & ",NULL," & ZVal(rsAdvice!婴儿) & "," & _
                    lng开嘱科室ID & ",'" & str开嘱医生 & "'," & IIf(!从项 = 1, ZVal(int父序号), "NULL") & "," & lng项目id & ",'" & !类别 & "'," & _
                    "'" & !计算单位 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                    "1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                    IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                    !应收 & "," & !实收 & "," & cur统筹金额 & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & ZVal(lng类别ID) & ",NULL,NULL,NULL," & _
                    !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                    Nvl(rsAdvice!计价特性, 0) & ",NULL,'" & str费用类型 & "')"
            End If
            
            int序号 = int序号 + 1
            
            .MoveNext
        Next
    End With
    
     '对SQL序列按收费细目ID排序
    For i = 0 To UBound(arrSQL) - 1
        For j = i + 1 To UBound(arrSQL)
            If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                strTmp = CStr(arrSQL(j))
                arrSQL(j) = arrSQL(i)
                arrSQL(i) = strTmp
            End If
        Next
    Next
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    For i = 0 To UBound(arrCountSQL)
        Call gobjDatabase.ExecuteProcedure(CStr(arrCountSQL(i)), Tittle)
    Next
    
    For i = 0 To UBound(arrSQL)
        Call gobjDatabase.ExecuteProcedure(CStr(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)), Tittle)
    Next
    
    '在提交前进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, mTYAdviceProperty.lng病人ID, rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, mTYAdviceProperty.lng病人ID, rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, mTYAdviceProperty.strNO, 2, 1, strMsg, , rsPati!险类) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    strOutNos = mTYAdviceProperty.strNO
    
    '在提交后进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, mTYAdviceProperty.lng病人ID, rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, mTYAdviceProperty.lng病人ID, rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, mTYAdviceProperty.strNO, 2, 1, strMsg, , rsPati!险类) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "单据""" & mTYAdviceProperty.strNO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Screen.MousePointer = 0
    FuncFeeMain = True
    MsgBox "执行项目的主费用生成成功。", vbInformation, gstrSysName
    Call RefreshExpenseData
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function zlFuncFeeNewPrice(ByVal frmMain As Object, Optional strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补充附加收费单据
    '入参:objMain-调用的主窗体
    '出参:strOutNos-成功的收费单据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 14:08:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
         IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), 1, mlng执行科室ID, mTYAdviceProperty.lng病人科室ID, 0, "", "", "", "", , , , strOutNos, , objSaveData, mobjSquareCard) Then
         zlFuncFeeNewPrice = True
         Call Refresh
         Call RefreshExpenseData
    End If

End Function
Public Function zlFuncFeeNewBilling(ByVal frmMain As Object, Optional strOutNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补记帐单据
    '出参:strOutNos-返回成功保成的单据
    '返回:补费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 16:03:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
         IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), 2, mlng执行科室ID, mTYAdviceProperty.lng病人科室ID, 0, "", "", "", "", , , , strOutNos, , , mobjSquareCard) Then
         zlFuncFeeNewBilling = True
         Call Refresh
         Call RefreshExpenseData
    End If
End Function
Public Function zlFuncFeeNewNull(ByVal frmMain As Object, Optional strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补零耗费用
    '出参:strOutNOs-保存成功的单据号
    '返回:补费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 16:02:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
         IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), 2, mlng执行科室ID, mTYAdviceProperty.lng病人科室ID, 0, "", "", "", "", , , True, strOutNos, , objSaveData, mobjSquareCard) Then
         zlFuncFeeNewNull = True
         Call Refresh
        Call RefreshExpenseData
    End If
End Function

Public Function zlFuncFeeModi(Optional frmMain As Object, Optional int病人来源 As Integer, Optional int记录性质 As Integer, Optional strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:改费
    '入参:int病人来源-1-门诊;2-住院
    '     int记录性质-记录性质
    '     strNO-按指定的单据进行改费
    '出参:
    '返回:修改成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 17:48:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln零耗 As Boolean
    Dim strFeeTab As String
    If strNO <> "" Then
      '按单据号进行改费
      If int病人来源 = 1 Then   '门诊
        strFeeTab = "门诊费用记录"
      Else
        strFeeTab = "住院费用记录"
      End If
        If gobjDatabase.NOMoved(strFeeTab, strNO, "记录性质=", int记录性质) Then
            MsgBox "费用单据 " & strNO & " 已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
 
        If int记录性质 = 2 Then
            If Not BillIdentical(strNO, IIf(int病人来源 = 2, 2, 1)) Then
                MsgBox "单据""" & strNO & """中包含部份未审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If BillExistDelete(strNO, int记录性质, IIf(int病人来源 = 2, 2, 1)) Then
            MsgBox "该单据包含已" & IIf(int记录性质 = 1, "退费", "销帐") & "费用,不允许修改！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '如果包含部分执行或全部执行的项目,则不一定可以全部冲销,不允许修改
        If HaveExecute(strNO, int记录性质, False, IIf(int病人来源 = 2, 2, 1)) Then
            MsgBox "该单据中包含完全执行或部分执行的项目,不允许修改！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If int记录性质 = 2 Then
           bln零耗 = BillisZeroLog(strNO, IIf(int病人来源 = 2, 2, 1))
        End If
        
        If zlIs备货材料(strNO, int记录性质) Then
            If frmStuffCharge.zlBillEdit(frmMain, 0, p医嘱附费管理, mstrPrivsAnnexFee, int记录性质, strNO, IIf(int病人来源 = 2, 2, 1), 0, 0, _
                0, 0, , , , , 0, 0, , , , , , objSaveData, mobjSquareCard) = False Then Exit Function
            zlFuncFeeModi = True
            Exit Function
        End If
        
        If Not frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, 0, 0, 0, 0, _
             IIf(int病人来源 = 2, 2, 1), int记录性质, 0, 0, 0, "", "", strNO, "", , , bln零耗, , , objSaveData, mobjSquareCard) Then Exit Function
        zlFuncFeeModi = True
        Exit Function
    End If
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    
    If vsExpense.TextMatrix(vsExpense.Row, 0) = "主费用" And mTYAdviceProperty.lng计价性质 <> 1 Then
        MsgBox "执行项目的主费用不能修改。如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If

    With vsExpense
        '78225:李南春,2014/9/24,获取正确的单据号和记录性质
        strNO = Get单据号
        If strNO = "" Or strNO = "[未计费]" Then Exit Function
        int记录性质 = Get记录性质
        
        If gobjDatabase.DateMoved(mTYAdviceProperty.dat发送时间) Then
            If gobjDatabase.NOMoved(mTYAdviceProperty.strFeeTab, strNO, "记录性质=", int记录性质) Then
                MsgBox "费用单据 " & strNO & " 已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If int记录性质 = 2 Then
            If Not BillIdentical(strNO, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)) Then
                MsgBox "单据""" & strNO & """中包含部份未审核或分多次审核的内容，不允许修改。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If Val(.TextMatrix(.Row, .ColIndex("记录性质"))) Mod 10 = 1 _
                        And Val(.TextMatrix(.Row, .ColIndex("收费标志"))) = 1 Then
            MsgBox "该单据已经收费，不允许修改。", vbInformation, gstrSysName
            Exit Function
        End If
        If BillExistDelete(strNO, int记录性质, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)) Then
            MsgBox "该单据包含已" & IIf(int记录性质 = 1, "退费", "销帐") & "费用,不允许修改！", vbInformation, gstrSysName
            Exit Function
        End If
        '如果包含部分执行或全部执行的项目,则不一定可以全部冲销,不允许修改
        If HaveExecute(strNO, int记录性质, False, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)) Then
            MsgBox "该单据中包含完全执行或部分执行的项目,不允许修改！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If int记录性质 = 2 Then
       bln零耗 = BillisZeroLog(strNO, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1))
    End If
    
    If zlIs备货材料(strNO, int记录性质) Then
        If frmStuffCharge.zlBillEdit(frmMain, 0, p医嘱附费管理, mstrPrivsAnnexFee, int记录性质, strNO, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
            mlng执行科室ID, mTYAdviceProperty.lng病人科室ID, , , , , mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, , , , , , objSaveData, mobjSquareCard) Then
            zlFuncFeeModi = True
            Call Refresh
            Call RefreshExpenseData
        End If
        Exit Function
    End If
    '78225:李南春,2014/9/24,获取正确的单据号和记录性质
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 0, mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
        IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), int记录性质, mlng执行科室ID, mTYAdviceProperty.lng病人科室ID, 0, "", "", strNO, "", , , bln零耗, , , objSaveData, mobjSquareCard) Then
        zlFuncFeeModi = True
        Call Refresh
        Call RefreshExpenseData
    End If
End Function

Public Function zlFuncFeeDel(Optional frmMain As Object, Optional int病人来源 As Integer, _
    Optional int记录性质 As Integer, Optional strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除附费
    '入参:int病人来源-1-门诊;2-住院
    '     int记录性质-记录性质
    '     strNO-按指定的单据进行改费
    '出参:
    '返回:删除成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 18:06:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFeeTab As String
    If strNO <> "" Then
        '按单据号进行改费
        If int病人来源 = 1 Then   '门诊
          strFeeTab = "门诊费用记录"
        Else
          strFeeTab = "住院费用记录"
        End If
      
        '按单据改费
        If gobjDatabase.NOMoved(strFeeTab, strNO, "记录性质=", int记录性质) Then
            MsgBox "费用单据 " & strNO & " 已经转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If int记录性质 = 2 Then
            If Not BillIdentical(strNO, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)) Then
                MsgBox "单据""" & strNO & """中包含部份未审核或分多次审核的内容，不允许删除。", vbInformation, gstrSysName
                Exit Function
            End If
            '住院出院病人费用限制
            If int病人来源 = 2 Then
                If Not PatiCanBilling(mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mstrPrivsAnnexFee, p医嘱附费管理) Then Exit Function
            End If
        End If
        If zlIs备货材料(strNO, int记录性质) Then
            If frmStuffCharge.zlBillEdit(frmMain, 3, p医嘱附费管理, mstrPrivsAnnexFee, int记录性质, strNO, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), _
                , , , , , , , , 0, , , , , , , objSaveData, mobjSquareCard) = False Then Exit Function
            zlFuncFeeDel = True
            Exit Function
        End If
        
        If Not frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 3, 0, 0, 0, 0, _
             IIf(int病人来源 = 2, 2, 1), int记录性质, 0, 0, 0, "", "", strNO, "", , , False, , , objSaveData, mobjSquareCard) Then Exit Function
        zlFuncFeeDel = True
        Exit Function
    End If
    
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsExpense
        '78225:李南春,2014/9/24,获取正确的单据号和记录性质
        strNO = .TextMatrix(.Row, .ColIndex("单据号"))
        If strNO = "" Or strNO = "[未计费]" Then Exit Function
        int记录性质 = Get记录性质
        
       If gobjDatabase.DateMoved(mTYAdviceProperty.dat发送时间) Then
            If gobjDatabase.NOMoved(mTYAdviceProperty.strFeeTab, strNO, "记录性质=", int记录性质) Then
                MsgBox "费用单据 " & strNO & " 已经转出到后备数据库，不允许操作。" & vbCrLf & _
                    "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If int记录性质 = 2 Then
            If Not BillIdentical(strNO, IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1)) Then
                MsgBox "单据""" & strNO & """中包含部份未审核或分多次审核的内容，不允许删除。", vbInformation, gstrSysName
                Exit Function
            End If
            '住院出院病人费用限制
            If mTYAdviceProperty.int病人来源 = 2 Then
                If Not PatiCanBilling(mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mstrPrivsAnnexFee, p医嘱附费管理) Then Exit Function
            End If
        End If
        If Val(.TextMatrix(.Row, .ColIndex("记录性质"))) Mod 10 = 1 _
                        And Val(.TextMatrix(.Row, .ColIndex("收费标志"))) = 1 Then
            MsgBox "该单据已经收费，不允许删除。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If vsExpense.TextMatrix(vsExpense.Row, vsExpense.ColIndex("费用类型")) = "主费用" Then
        If InStr(mstrPrivsAnnexFee, "删除主费用") = 0 Then
            MsgBox "你没有删除主费用的权限，不能删除主费用。", vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox("主费用删除后不能重新产生,你确实要删除项目主费用吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    If zlIs备货材料(strNO, int记录性质) Then
    
        If frmStuffCharge.zlBillEdit(frmMain, 3, p医嘱附费管理, mstrPrivsAnnexFee, int记录性质, strNO, _
            IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), , , , , , , , , IIf(mTYAdviceProperty.bln独立执行, mTYAdviceProperty.lng医嘱ID, 0), _
            , , , , , , objSaveData, mobjSquareCard) = False Then
            Exit Function
        End If
        zlFuncFeeDel = True
        Call Refresh
        Call RefreshExpenseData
        Exit Function
    End If
    
    If frmTechnicExpense.EditCard(frmMain, mstrPrivsAnnexFee, 3, IIf(mTYAdviceProperty.bln独立执行, mTYAdviceProperty.lng医嘱ID, 0), mTYAdviceProperty.lng发送号, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, _
         IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), int记录性质, 0, 0, 0, "", "", strNO, "", , , False, , , objSaveData, mobjSquareCard) Then
         zlFuncFeeDel = True
         Call Refresh
         Call RefreshExpenseData
    End If
    
End Function

Public Function zlFuncExtraFeeExe(ByVal frmMain As Object, ByVal bytType As Byte, ByVal strMainPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费执行与取消执行
    '入参:bytType=0-取消执行,1-执行
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 18:19:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim objParent As Object
    Dim strSql As String
    Dim strNO As String, int记录性质 As Integer, blnIsAbnormal As Boolean, blnTrans As Boolean
    Dim i As Long, blnDo As Boolean, str序号 As String, strDate As String
    Dim arr序号() As String, arrSQL As Variant
    Dim str类别 As String, str类别名 As String, curMoney As Currency
    Dim blnRefresh As Boolean, lngRow As Long
    Dim blnJudge As Boolean, blnTrace As Boolean, blnHaveDrug As Boolean, blnHave卫材 As Boolean, strMsg As String
    
    '1.如果单据全是药品或非自动发料的跟踪在用卫材，则不处理执行，因为这些是发药才表示执行。
    '2.跟踪在用的卫材在执行时根据系统参数决定是否自动发料。
    blnDo = False
    If mTYAdviceProperty.int审核标志 >= 1 And gSysPara.byt病人审核方式 = 1 Then
        MsgBox "该病人的费用正在审核阶段，不允许操作医嘱和费用。", vbInformation, gstrSysName
        Exit Function
    End If
    '78789:李南春,2014/10/22,附费执行与取消执行时，只检查选择的单据
    strNO = Get单据号
    int记录性质 = Get记录性质
    
    If strNO = "" Then
        MsgBox "当前无单据信息，不能" & IIf(bytType = 0, "取消", "") & "执行登记", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    With vsExpense
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpText, i, .ColIndex("单据号")) = strNO Then
                lngRow = lngRow + 1
                If InStr(",西成药,中成药,中草药,", "," & .Cell(flexcpText, i, .ColIndex("类别")) & ",") = 0 Then
                    blnJudge = True
                    If .Cell(flexcpText, i, .ColIndex("类别")) = "卫材" And Not gSysPara.bln卫材执行发料 Then
                        '判断是否卫材跟踪在用
                        strSql = " Select 1  From 材料特性 where 材料ID=[1] And 跟踪在用=1 "
                        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, Val(.Cell(flexcpText, i, .ColIndex("项目"))))
                        If rsTmp.RecordCount = 0 Then
                            blnJudge = True
                        Else
                            blnHave卫材 = True
                            blnJudge = False
                        End If
                    End If
                    If blnJudge Then
                        blnDo = True
                        '非药品单据，不存在部分记录执行的情况。
                        If bytType = 1 Then
                            If .Cell(flexcpText, i, .ColIndex("执行状态")) = 1 Then
                                MsgBox "该单据已经完全执行，无需再次登记执行。", vbQuestion, gstrSysName
                                Exit Function
                            End If
                        Else
                            If .Cell(flexcpText, i, .ColIndex("执行状态")) = 0 Then
                                MsgBox "该单据未执行，无需取消执行。", vbQuestion, gstrSysName
                                Exit Function
                            End If
                        End If
                        str序号 = str序号 & "," & .Cell(flexcpText, i, .ColIndex("序号"))
                    End If
                Else
                    blnHaveDrug = True
                End If
            End If
        Next
        If blnDo = False Then
            If blnHaveDrug And blnHave卫材 Then
                strMsg = "药品通过发药和退药来处理执行或取消执行，" & vbNewLine & "在卫材非自动发料情况下,卫材通过发料来处理执行或取消执行，" _
                        & vbNewLine & "不允许直接登记或取消执行。"
            ElseIf blnHaveDrug Or blnHave卫材 Then
                strMsg = IIf(blnHaveDrug, "药品通过发药和退药来处理执行或取消执行，不允许直接登记或取消执行。", _
                    " 在卫材非自动发料情况下,卫材通过发料来处理执行或取消执行，" & vbNewLine & "不允许直接登记或取消执行。")
            End If
            'strMsg = ""为只含有跟踪在用卫材,且卫材自动发料
            If strMsg <> "" Then
                MsgBox strMsg, vbQuestion, gstrSysName
                Exit Function
            End If
        End If
        str序号 = Mid(str序号, 2)
    End With
    arrSQL = Array()
    
    If bytType = 1 Then
        If int记录性质 = 2 And gSysPara.bln执行后审核 Then
            curMoney = GetUnAuditBill(strNO, int记录性质, str类别, str类别名)
        End If
        
        If mTYAdviceProperty.int病人来源 = 2 Then
            '住院记帐单据
            
            '卫料自动发料时的费用检查
            If gSysPara.bln卫材执行发料 And Not gSysPara.bln执行后审核 Then
                If Not CheckStuffAudit(strNO, int记录性质) Then
                    MsgBox "操作不能继续。" & vbCrLf & vbCrLf & "单据中存在尚未审核的卫材费用，不能在执行之后自动发料。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '执行后自动审核时，对出院病人检查强制记帐权限
            If gSysPara.bln执行后审核 And curMoney <> 0 Then
                If Not PatiCanBilling(mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mstrPrivsAnnexFee, p医嘱附费管理) Then Exit Function
            End If
        Else
            '门诊记帐的，可以执行后自动审核，否则须先交费或审核
            If int记录性质 = 1 Or int记录性质 = 2 And Not gSysPara.bln执行后审核 Then
                If Not CheckFinishCharge(strNO, int记录性质, blnIsAbnormal) Then
                    If gSysPara.bln执行前先结算 And Not mobjSquareCard Is Nothing Then
                        If blnIsAbnormal Then
                            MsgBox "该病人还存在异常费用，请检查。", vbInformation, gstrSysName
                            Exit Function
                        End If
                        '门诊一卡通,项目执行前必须先收费或先记帐审核,不传单据号，根据医嘱ID读取所有未收费单据或未审核的记帐单
                        blnRefresh = mobjSquareCard.zlSquareAffirm(Me, p医嘱附费管理, strMainPrivs, mTYAdviceProperty.lng病人ID, 0, False, int记录性质, strNO, "")
                        If Not blnRefresh Then
                            Exit Function
                        End If
                    Else
                        MsgBox "该病人还存在未" & IIf(int记录性质, "收费", "审核记帐") & "的费用，请检查。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '对记帐费用进行报警
        If int记录性质 = 2 And curMoney <> 0 And gSysPara.bln执行后审核 Then
            If Not FinishBillingWarn(Me, mstrPrivsAnnexFee, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mTYAdviceProperty.lng病人病区ID, curMoney, str类别, str类别名) Then Exit Function
        End If
        
        If MsgBox("你确定要将单据""" & strNO & """登记为已执行吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then Exit Function
        
        strDate = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        arr序号 = Split(str序号, ",")
        '可能排除了药品行
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人费用记录_Execute('" & strNO & "'," & int记录性质 & IIf(UBound(arr序号) + 1 = lngRow, ",Null,", ",'" & str序号 & "',") & mTYAdviceProperty.int病人来源 & ",Null,'" & UserInfo.姓名 & "'," & strDate & ")"
    Else
        If MsgBox("你确定要取消单据""" & strNO & """的执行登记吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then Exit Function
        arr序号 = Split(str序号, ",")
        '可能排除了药品行
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人费用记录_UnExecute('" & strNO & "'," & int记录性质 & IIf(UBound(arr序号) + 1 = lngRow, ",Null,", ",'" & str序号 & "',") & mTYAdviceProperty.int病人来源 & ")"
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Tittle)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    lngRow = vsExpense.Row
    If blnRefresh Then
        Call LoadFeeDataFromAdvice(mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.bln独立执行)
    End If
    vsExpense.Row = lngRow
    '刷新费用明细清单
    Call LoadFeeDataFromAdvice(mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, mTYAdviceProperty.bln独立执行)
    RaiseEvent StatusTextUpdate(IIf(bytType = 0, 2, 1), IIf(bytType = 0, "取消", "") & "执行操作成功。")
    zlFuncExtraFeeExe = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetUnAuditBill(ByVal strNO As String, ByVal int记录性质 As Integer, str类别 As String, str类别名 As String) As Currency
'功能：获取未审核的记帐单据的金额和类别，用于记帐报警
'参数：
'返回：str类别,str类别名=用于报警提示
'说明：
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, curMoney As Currency
    
    str类别 = "": str类别名 = ""
    
    On Error GoTo errH
    strSql = _
        " Select B.编码,B.名称,Sum(A.实收金额) as 金额" & _
        " From 住院费用记录 A,收费项目类别 B" & _
        " Where A.NO = [1] And A.记录性质 = [2] And A.记录状态 = 0 And A.收费类别=B.编码" & _
        " Group by B.编码,B.名称"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strNO, int记录性质)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!金额, 0)
        str类别 = str类别 & rsTmp!编码
        str类别名 = str类别名 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    str类别名 = Mid(str类别名, 2)
    GetUnAuditBill = curMoney
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function CheckStuffAudit(ByVal strNO As String, ByVal int记录性质 As Integer) As Boolean
'功能：判断单据中跟踪在用卫材是否存在未审核的记帐费用
'参数：
'返回：
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = " Select Nvl(Sum(A.实收金额),0) as 金额" & _
        " From 住院费用记录 A,材料特性 B" & _
        " Where NO = [1] And A.记录性质 = [2] And A.记录状态 = 0 And A.收费类别='4' And A.收费细目ID=B.材料ID And B.跟踪在用=1"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strNO, int记录性质)
    If Not rsTmp.EOF Then
        CheckStuffAudit = rsTmp!金额 = 0
    Else
        CheckStuffAudit = True
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function CheckFinishCharge(ByVal strNO As String, ByVal int记录性质 As Integer, ByRef blnIsAbnormal As Boolean) As Boolean
'功能：判断指定的单据是否已收费，以及是否存在异常费用
'参数：blnIsAbnormal=是否存在收费异常的记录

    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 记录状态,执行状态 From 门诊费用记录 Where NO = [1] And 记录性质 = [2] And (记录状态 = 0 or 执行状态 = 9)"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strNO, int记录性质)
    rsTmp.Filter = "记录状态=0"
    If rsTmp.RecordCount > 0 Then
        CheckFinishCharge = False
    Else
        CheckFinishCharge = True
        rsTmp.Filter = "执行状态=9"
        If rsTmp.RecordCount > 0 Then blnIsAbnormal = True
    End If
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function FuncExtraFeeMove(Optional frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费转移
    '入参:frmMain-凋用的主窗体
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 18:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, int记录性质 As Integer
    Dim objParent As Object
    
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsExpense
        If Not Is附加费 Then
            MsgBox "当前选择的单据不是补充的附加费用。", vbInformation, gstrSysName
            Exit Function
        End If
        
        strNO = Get单据号
        int记录性质 = Get记录性质
    End With
    If Not frmMain Is Nothing Then
        Set objParent = frmMain
    Else
        Set objParent = Me
    End If
    
    If frmExtraFeemove.ShowMe(objParent, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mTYAdviceProperty.str挂号单, mTYAdviceProperty.lng相关ID, mTYAdviceProperty.str诊疗类别, mlng执行科室ID, _
        strNO, int记录性质, IIf(mTYAdviceProperty.strFeeTab = "住院费用记录", 2, 1)) Then
        Call Refresh
    End If
End Function
Public Function zlFuncPlugIn(ByVal frmMain As Object, ByVal Control As CommandBarControl) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行外挂
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-29 17:50:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String
    If CreatePlugIn(p医嘱附费管理) = False Then Exit Function
    strNO = Get单据号
    On Error Resume Next
    Call gobjPlugIn.ExecuteFunc(glngSys, p医嘱附费管理, Control.Parameter, _
        mTYAdviceProperty.lng病人ID, IIf(mTYAdviceProperty.str挂号单 = "", mTYAdviceProperty.lng主页Id, mTYAdviceProperty.str挂号单), mTYAdviceProperty.lng医嘱ID, strNO)
    Call zlPlugInErrH(Err, "ExecuteFunc")
    Err.Clear: On Error GoTo 0
End Function
Public Function zlFuncAdviceReCharge(ByVal intType As Integer, Optional frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:费用销帐申请和审核
    '入参:intType=1-申请;2-审核
    '     frmMain-调用的主窗口
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 18:24:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Dim strCommon As String, intAtom As Integer
    Dim objParent As Object, strNO As String, lngAdviceID As Long
    Set objParent = frmMain
    If objParent Is Nothing Then Set objParent = mfrmParent
    
    '调用费用部件功能
    On Error Resume Next
    If gobjInExse Is Nothing Then
        Set gobjInExse = CreateObject("zl9InExse.clsInExse")
        If gobjInExse Is Nothing Then Exit Function
    End If
    Err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    
    If intType = 1 Then
        lngAdviceID = mTYAdviceProperty.lng医嘱ID
        strNO = Get单据号
        
        Select Case mbytFocus
        Case 1
            '按医嘱
            blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 0, 1, mlng执行科室ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng病人ID, , lngAdviceID)
        Case 2
            '按费用
            blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 0, 1, mlng执行科室ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng病人ID, strNO)
        Case Else
            '按病人
            blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 0, 1, mlng执行科室ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng病人ID)
        End Select
    ElseIf intType = 2 Then
        blnOK = gobjInExse.CallReCharge(objParent, gcnOracle, gstrDBUser, glngSys, 1, 1, mlng执行科室ID, mstrPrivsAnnexFee, mTYAdviceProperty.lng病人ID)
    End If
    Call GlobalDeleteAtom(intAtom)
    zlFuncAdviceReCharge = blnOK
    
    If blnOK And frmMain Is Nothing Then RaiseEvent RequestRefresh
End Function
 

Private Sub vsExpense_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strNO As String, int记录性质 As Integer, lng发送号 As Long, lng医嘱ID As Long
    If Button = 1 Then Exit Sub
    
    With vsExpense
        strNO = Get单据号
        If strNO = "[未计费]" Then strNO = ""
        int记录性质 = Val(Get记录性质)
    End With
    RaiseEvent zlPopupMenu(mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, strNO, int记录性质, X, Y)
End Sub

Public Function IsFunValied(ByVal bytType As Byte, ByVal bytPrivsCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判检查某功能是否有效
    '入参: bytType- 1-修改附费;2-删除附费;3-附费转移;4-附费执行;5-附费取消执行;6-销帐申请;7-销帐审核
    '      bytPrivsCheck -检查权限:0-不检查权限;1-检查数据和权限;2-仅检查权限
    '出参:
    '返回:功能有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 17:00:52
    '说明:
    '   1.根据附费列表中的内容,检查某项功能是否有效
    '   2.根据权限检查某项功能是否有效
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValied As Boolean, bln附加费用 As Boolean, bln主费用 As Boolean
    
    On Error GoTo errHandle
    
    bln主费用 = Is主费用: bln附加费用 = Is附加费
    
    Select Case bytType
    Case 1 '修改附费
        blnValied = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) _
                   And (bln附加费用 Or mTYAdviceProperty.lng计价性质 = 1)
        
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";修改费用;") > 0
        End If
    
    Case 2 '删除附费
        blnValied = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) _
            And (bln附加费用 Or bln主费用)
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";删除费用;") > 0
        End If
    Case 3 '附费转移
         blnValied = mTYAdviceProperty.lng医嘱ID <> 0 And bln附加费用
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";补充附加费用;") > 0
        End If
    Case 4 '附费执行
         blnValied = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) And bln附加费用
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied)
        End If
    Case 5 '附费取消执行
        blnValied = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) And bln附加费用
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied)
        End If
    Case 6 '销帐申请
        blnValied = mTYAdviceProperty.int病人来源 = 2 And mlng执行科室ID <> 0
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And Not (InStr(mstrPrivsAnnexFee, ";药品销帐申请;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";卫材销帐申请;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";诊疗销帐申请;") = 0)
        End If
    Case 7 '销帐审核
        blnValied = mTYAdviceProperty.int病人来源 = 2 And mlng执行科室ID <> 0
        If bytPrivsCheck <> 0 Then
            blnValied = IIf(bytPrivsCheck = 2, True, blnValied) And InStr(mstrPrivsAnnexFee, ";销帐审核;") > 0
        End If
    Case Else
        blnValied = True
    End Select
    IsFunValied = blnValied
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlFuncStuffCharge(ByVal frmMain As Object, ByVal int记录性质 As Integer, _
    Optional strOutNos As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:备货卫材记帐
    '入参:int记录性质:1-收费(划价),2-记帐(门/住)
    '出参:strOutNos-保存成功的备货卫材收费或记帐单据单据
    '编制:刘兴洪
    '日期:2010-12-14 13:17:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mTYAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
        Exit Function
    End If
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    If frmStuffCharge.zlBillEdit(frmMain, 0, p医嘱附费管理, mstrPrivsAnnexFee, int记录性质, "", _
        IIf(mTYAdviceProperty.int病人来源 = 2, 2, 1), mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mlng执行科室ID, mTYAdviceProperty.lng病人科室ID, _
        0, "", False, "", mTYAdviceProperty.lng医嘱ID, mTYAdviceProperty.lng发送号, "", , , strOutNos, , objSaveData, mobjSquareCard) = True Then
        zlFuncStuffCharge = True
        Call Refresh
        Call RefreshExpenseData
    End If
End Function

Public Function zlFuncExtraFeeMove(Optional frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费转移
    '入参:frmMain-凋用的主窗体
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 18:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, int记录性质 As Integer
    Dim objParent As Object
    
    If mTYAdviceProperty.lng医嘱ID = 0 Then Exit Function
    If mblnMoved Then
        MsgBox "该病人的本次" & IIf(mTYAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    If Not Is附加费 Then
        MsgBox "当前选择的单据不是补充的附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    strNO = Get单据号
    int记录性质 = Get记录性质
    If Not frmMain Is Nothing Then
        Set objParent = frmMain
    Else
        Set objParent = mfrmParent
    End If
    If frmExtraFeemove.ShowMe(objParent, mTYAdviceProperty.lng病人ID, mTYAdviceProperty.lng主页Id, mTYAdviceProperty.str挂号单, mTYAdviceProperty.lng相关ID, mTYAdviceProperty.str诊疗类别, mlng执行科室ID, _
        strNO, int记录性质, IIf(mTYAdviceProperty.strFeeTab = "住院费用记录", 2, 1)) Then
        zlFuncExtraFeeMove = True
        Call Refresh
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal cbsMain As Object, ByVal Control As CommandBarControl)
    Dim blnEnabled As Boolean
    If cbsMain Is Nothing Then Exit Sub
    
    If vsExpense.Redraw = flexRDNone Then Exit Sub
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '输出费用清单
        Control.Enabled = IsHaveExpenseData
    Case conMenu_Edit_Append
        
        If Is未计费 Then
            Control.Enabled = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3)
            If Control.Parent Is cbsMain(2) Then
                Control.Caption = "补充主费"
            Else
                Control.Caption = "补充主费用(&N)"
            End If
        Else
            blnEnabled = Not mrsPrice Is Nothing
            If blnEnabled Then blnEnabled = mrsPrice.RecordCount <> 0
            Control.Enabled = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) And InStr(mTYAdviceProperty.str计费状态, ",0,") > 0 And blnEnabled
            If Control.Parent Is cbsMain(2) Then
                Control.Caption = "生成主费"
            Else
                Control.Caption = "生成主费用(&N)"
            End If
        End If
    Case conMenu_Edit_ExtraFeeMove  '附费转移
        Control.Enabled = mTYAdviceProperty.lng医嘱ID <> 0 And Is附加费
    Case conMenu_Edit_ExtraFeeExe   '附费执行
        Control.Enabled = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) _
            And Is附加费
    Case conMenu_Edit_ExtraFeeUnExe '附费取消执行
        Control.Enabled = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) _
            And Is附加费
    Case conMenu_Edit_NewItem
        Control.Enabled = mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3)
    '78929:李南春,2014/10/27，点击小计时，改费不可用
    Case conMenu_Edit_Modify
        Control.Visible = InStr(mstrPrivsAnnexFee, ";修改费用;") > 0
        Control.Enabled = Control.Visible And mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) And _
        (Is附加费 Or mTYAdviceProperty.lng计价性质 = 1) And Not IIf(vsExpense.Row < 0, True, vsExpense.IsSubtotal(vsExpense.Row))
    Case conMenu_Edit_Delete
        Control.Visible = InStr(mstrPrivsAnnexFee, ";删除费用;") > 0
        Control.Enabled = Control.Visible And mTYAdviceProperty.lng医嘱ID <> 0 And (mTYAdviceProperty.int执行状态 = 0 Or mTYAdviceProperty.int执行状态 = 3) _
            And (Is附加费 Or Is主费用)
    Case conMenu_Edit_ChargeDelApply, conMenu_Edit_ChargeDelAudit '销帐申请审核
        Control.Enabled = mTYAdviceProperty.int病人来源 = 2 And mlng执行科室ID <> 0
    End Select
End Sub
Public Property Get Is附加费() As Boolean
    If mbytFun = 1 Then Exit Function
    With vsExpense
        If .Rows <= 1 Then Exit Property
        If .Row < 0 Then Exit Property
        If .ColIndex("费用类型") Then Exit Property
        Is附加费 = .TextMatrix(.Row, .ColIndex("费用类型")) = "附加费用"
    End With
End Property
Public Property Get Is主费用() As Boolean
    If mbytFun = 1 Then Exit Property
    With vsExpense
        If .Rows <= 1 Then Exit Property
        If .Row < 0 Then Exit Property
        If .ColIndex("费用类型") Then Exit Property
        Is主费用 = .TextMatrix(.Row, .ColIndex("费用类型")) = "主费用"
    End With
End Property
Private Function Get单据号() As String
    Dim lngRow As Long
    With vsExpense
        If .Rows <= 1 Then Exit Function
        If .Row <= 0 Then Exit Function
        If .ColIndex("单据号") < 0 Then Exit Function
        If .IsSubtotal(.Row) Then
            lngRow = .Row + 1
            If lngRow > .Rows - 1 Then lngRow = .Row
            Get单据号 = .TextMatrix(lngRow, .ColIndex("单据号"))
        Else
            Get单据号 = .TextMatrix(.Row, .ColIndex("单据号"))
        End If
    End With
End Function
Private Function Get记录性质() As Integer
    Dim lngRow As Long, int记录性质 As Integer
    With vsExpense
        int记录性质 = mTYAdviceProperty.int记录性质
        
        If .Rows <= 1 Then Get记录性质 = int记录性质: Exit Function
        If .Row < 0 Then Get记录性质 = int记录性质: Exit Function
        If .ColIndex("记录性质") < 0 Then Get记录性质 = int记录性质: Exit Function
        
        If .IsSubtotal(.Row) Then
            lngRow = .Row + 1
            If lngRow > .Rows - 1 Then lngRow = .Row
            Get记录性质 = Val(.TextMatrix(lngRow, .ColIndex("记录性质")))
        Else
            Get记录性质 = Val(.TextMatrix(.Row, .ColIndex("记录性质")))
        End If
    End With
End Function

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean
    
    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
    Case conMenu_Edit_Append, conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_ExtraFeeMove
        If InStr(mstrPrivsAnnexFee, ";补充附加费用;") = 0 Then blnVisible = False
    Case conMenu_Edit_NewItem * 10# + 1  '补充收费费用
        If mTYAdviceProperty.int病人来源 = 2 Or InStr(mstrPrivsAnnexFee, ";补充收费费用;") = 0 Or mTYAdviceProperty.int结算模式 = 1 Then blnVisible = False
    Case conMenu_Edit_NewItem * 10# + 2     '补充记帐费用
        If mTYAdviceProperty.int病人来源 = 2 Then
            If InStr(mstrPrivsAnnexFee, ";补充住院记帐费用;") = 0 Then blnVisible = False
        Else
            If InStr(mstrPrivsAnnexFee, ";补充门诊记帐费用;") = 0 Then blnVisible = False
        End If
    Case conMenu_Edit_NewItem * 10# + 3 '补充零耗费用
        If InStr(mstrPrivsAnnexFee, ";补充零耗费用;") = 0 Then blnVisible = False
    Case conMenu_Edit_NewItem * 10# + 5  '备货卫材记帐和收费
        If mTYAdviceProperty.int病人来源 = 2 Then
            If InStr(mstrPrivsAnnexFee, ";补充备货卫材费用;") = 0 Or InStr(mstrPrivsAnnexFee, ";补充住院记帐费用;") = 0 Then
                blnVisible = False
            End If
        Else
            If InStr(mstrPrivsAnnexFee, ";补充备货卫材费用;") = 0 Or InStr(mstrPrivsAnnexFee, ";补充门诊记帐费用;") = 0 Then
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_NewItem * 10# + 4     '
        If mTYAdviceProperty.int病人来源 = 2 Then
            blnVisible = False
        Else
            If InStr(mstrPrivsAnnexFee, ";补充备货卫材费用;") = 0 Or _
               InStr(mstrPrivsAnnexFee, ";补充收费费用;") = 0 Then blnVisible = False
        End If
    Case conMenu_Edit_ChargeDelApply '销帐申请
        '刘兴洪 问题: 34873   日期:2010-12-22 13:50:07
        '55380
        If InStr(mstrPrivsAnnexFee, ";药品销帐申请;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";卫材销帐申请;") = 0 _
            And InStr(mstrPrivsAnnexFee, ";诊疗销帐申请;") = 0 Then blnVisible = False
    Case conMenu_Edit_ChargeDelAudit '销帐审核
        If InStr(mstrPrivsAnnexFee, ";销帐审核;") = 0 Then blnVisible = False
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub
Private Sub RefreshExpenseData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新费用数据
    '编制:刘兴洪
    '日期:2014-05-30 14:54:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng医嘱ID As Long, lng发送号 As Long, bln独立执行 As Boolean
    If mbytFun = 1 Then
        Call LoadFeeListFromNos(mbyt记录性质, mstrNos, mbyt病人来源, mblnMoved)
    Else
        With vsAdvice
            If .Row > 0 And .Row <= .Rows - 1 Then
                lng医嘱ID = Val(.TextMatrix(.Row, .ColIndex("医嘱ID")))
                lng发送号 = Val(.TextMatrix(.Row, .ColIndex("发送号")))
                bln独立执行 = Val(.TextMatrix(.Row, .ColIndex("独立执行"))) = 1
            End If
        End With
        Call LoadFeeDataFromAdvice(lng医嘱ID, lng发送号, bln独立执行)
    End If
End Sub

Private Function GetAdviceMoney(ByVal strAdviceIdAndPayNums As String, _
    ByRef rsTemp As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取相关医嘱的应收金额和实收金额
    '入参:strAdviceIdAndPayNums-医嘱ID和发送号(医嘱ID:发送号,...)
    '出参:rsTemp-返回:医嘱ID,发送号,应收金额,实收金额
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-30 15:13:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, blnMoved As Boolean
    Dim strSQL1 As String, strSub As String
    Dim strWith As String, strWith1 As String, strTemp As String
    Dim rsSub As ADODB.Recordset, str医嘱IDs As String
    Dim strSubWith As String, strSubWith1 As String
    
    On Error GoTo errHandle
    '124405,李南春，2018/4/17，读取历史数据时会生成两条相同的with子句
    blnMoved = gobjDatabase.TableDataMoved("病人医嘱发送", " (医嘱ID,发送号) IN", " (Select C1 As 医嘱id, C2 As 发送号 From Table(f_Num2list2('" & strAdviceIdAndPayNums & "')))")
    
    '85249:李南春,2015/5/29,with预交嵌套使用
    strWith = "" & _
    "   Select A.医嘱ID,a.发送号,A.记录性质,A.NO" & vbNewLine & _
    "   From 病人医嘱发送 A" & vbNewLine & _
    "   Where (A.医嘱id,A.发送号) IN (Select /*+cardinality(B,10)*/B.C1, B.C2 From Table(f_Num2list2([1])) B)" & vbNewLine & _
    "   Union ALL " & vbNewLine & _
    "   Select A.医嘱ID,A.发送号,A.记录性质,A.NO " & vbNewLine & _
    "   From 病人医嘱附费 A" & vbNewLine & _
    "   Where (A.医嘱id,A.发送号) IN (Select /*+cardinality(B,10)*/B.C1, B.C2 From Table(f_Num2list2([1])) B) "
    
    If blnMoved Then
        strWith1 = strWith
        strWith1 = Replace(strWith1, "病人医嘱附费", "H病人医嘱附费")
        strWith1 = Replace(strWith1, "病人医嘱发送", "H病人医嘱发送")
        strWith = strWith & " Union ALL " & strWith1
    End If
    strWith = "   With 医嘱单据 as ( " & strWith & " )"
    
    strTemp = "Select A.Id As 医嘱ID, a.相关id As 主ID, a.诊疗类别 From 病人医嘱记录 A Where A.相关ID IN (Select /*+cardinality(B,10)*/B.C1 From Table(f_Num2list2([1])) B)"
    Set rsSub = gobjDatabase.OpenSQLRecord(strTemp, Tittle, strAdviceIdAndPayNums)
    If Not rsSub.EOF Then
        strSubWith = "" & _
        "   Select a.医嘱id, a.发送号, a.记录性质, a.No, m.相关id As 主id" & vbNewLine & _
        "   From 病人医嘱发送 a, 病人医嘱记录 m" & vbNewLine & _
        "   Where a.医嘱id = m.Id And (a.医嘱id, a.发送号, a.No) In" & vbNewLine & _
        "      (Select a.Id As 医嘱id, m.发送号, m.No" & vbNewLine & _
        "                         From 病人医嘱记录 a, 病人医嘱发送 m" & vbNewLine & _
        "                         Where a.相关id = m.医嘱id And" & vbNewLine & _
        "                               (m.医嘱id, m.发送号) In (Select /*+cardinality(b,10) */" & vbNewLine & _
        "                                                    b.C1, b.C2" & vbNewLine & _
        "                                                   From Table(f_Num2list2([1])) b))" & vbNewLine & _
        "   Union ALL " & vbNewLine & _
        "   Select a.医嘱id, a.发送号, a.记录性质, a.No, m.相关id As 主id" & vbNewLine & _
        "   From 病人医嘱附费 a, 病人医嘱记录 m" & vbNewLine & _
        "   Where a.医嘱id = m.Id And" & vbNewLine & _
        "      (a.医嘱id, a.发送号) In (Select a.Id As 医嘱id, m.发送号" & vbNewLine & _
        "                          From 病人医嘱记录 a, 病人医嘱附费 m" & vbNewLine & _
        "                          Where a.相关id = m.医嘱id And" & vbNewLine & _
        "                                (m.医嘱id, m.发送号) In (Select /*+cardinality(b,10) */" & vbNewLine & _
        "                                                     b.C1, b.C2" & vbNewLine & _
        "                                                    From Table(f_Num2list2([1])) b))"

        If blnMoved Then
            strSubWith1 = strSubWith
            strSubWith1 = Replace(strSubWith1, "病人医嘱附费", "H病人医嘱附费")
            strSubWith1 = Replace(strSubWith1, "病人医嘱发送", "H病人医嘱发送")
            strSubWith1 = Replace(strSubWith1, "病人医嘱记录", "H病人医嘱记录")
            strSubWith = strSubWith & " Union ALL " & strSubWith1
        End If
        strSubWith = "   ,医嘱关联单据 as ( " & strSubWith & " )"
        
        strSql = "" & _
        "   Select B.医嘱ID,B.发送号,A.应收金额,A.实收金额" & vbNewLine & _
        "   From 门诊费用记录 A,医嘱单据 B " & vbNewLine & _
        "   Where mod(A.记录性质,10)=B.记录性质  And A.NO=B.NO And A.医嘱序号=B.医嘱ID " & vbNewLine & _
        "   Union All" & _
        "   Select B.主ID As 医嘱ID,B.发送号,A.应收金额,A.实收金额" & vbNewLine & _
        "   From 门诊费用记录 A,医嘱关联单据 B,病人医嘱记录 C " & vbNewLine & _
        "   Where mod(A.记录性质,10)=B.记录性质 And a.医嘱序号=c.id And c.诊疗类别 In ('C','D','F') And A.NO=B.NO And A.医嘱序号=B.医嘱ID "
    Else
        '87435,一张单据多个医嘱时，通过医嘱ID统计费用不正确 'And A.医嘱序号=B.医嘱ID
        strSql = "" & _
        "   Select B.医嘱ID,B.发送号,sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额" & vbNewLine & _
        "   From 门诊费用记录 A,医嘱单据 B " & vbNewLine & _
        "   Where mod(A.记录性质,10)=B.记录性质 And A.NO=B.NO And A.医嘱序号=B.医嘱ID " & vbNewLine & _
        "   Group By B.医嘱ID,B.发送号"
    End If
    
    strSql = strSql & " UNION ALL " & vbNewLine & _
    Replace(Replace(strSql, "门诊费用记录", "住院费用记录"), "mod(A.记录性质,10)", "A.记录性质")
    
    If blnMoved Then
        strSQL1 = strSql
        strSQL1 = Replace(strSQL1, "病人医嘱附费", "H病人医嘱附费")
        strSQL1 = Replace(strSQL1, "病人医嘱发送", "H病人医嘱发送")
        strSQL1 = Replace(strSQL1, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSQL1, "住院费用记录", "H住院费用记录")
        strSql = strSql & " Union ALL " & strSQL1
    End If
    
    strSql = strWith & strSubWith & vbCrLf & strSql
    
    strSql = "" & _
    "   Select 医嘱ID,发送号,sum(应收金额) as 应收金额,Sum(实收金额) as 实收金额" & vbNewLine & _
    "   From (" & strSql & ")  " & vbNewLine & _
    "   Group By 医嘱ID,发送号"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Tittle, strAdviceIdAndPayNums, str医嘱IDs)
    GetAdviceMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "指定给定层的每一行出现的字体大小(以磅为单位)。"
    FontSize = UserControl.FontSize
    Call ReSetFontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    Call ReSetFontSize
    PropertyChanged "FontSize"
End Property
Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置字体大小
    '编制:刘兴洪
    '日期:2012-06-18 16:52:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngFontSize As Single
    sngFontSize = UserControl.FontSize
    
    Err = 0: On Error Resume Next
    picAdvice.FontSize = sngFontSize
    picExpense.FontSize = sngFontSize
    dkpMan.PaintManager.CaptionFont.Size = sngFontSize
    dkpMan.PanelPaintManager.Font.Size = sngFontSize
    Call gobjControl.VSFSetFontSize(vsAdvice, sngFontSize)
    Call gobjControl.VSFSetFontSize(vsExpense, sngFontSize)
    With vsExpense
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = sngFontSize
    End With
    dkpMan.RecalcLayout
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,&HFFCC99
Public Property Get COLOR_FOCUS() As OLE_COLOR
    COLOR_FOCUS = m_COLOR_FOCUS
End Property

Public Property Let COLOR_FOCUS(ByVal New_COLOR_FOCUS As OLE_COLOR)
    m_COLOR_FOCUS = New_COLOR_FOCUS
    PropertyChanged "COLOR_FOCUS"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,&HFFEBD7
Public Property Get COLOR_LOST() As OLE_COLOR
    COLOR_LOST = m_COLOR_LOST
End Property

Public Property Let COLOR_LOST(ByVal New_COLOR_LOST As OLE_COLOR)
    m_COLOR_LOST = New_COLOR_LOST
    PropertyChanged "COLOR_LOST"
End Property
Public Sub SetDefalutFocus(ByVal blnAdviceFocus As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的光标列表
    '入参:blnAdviceFocus-是否指向医嘱列表(true-缺省指向医嘱;false-为明细列表)
    '编制:刘兴洪
    '日期:2014-06-03 11:29:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnAdviceFocus Then
        If vsAdvice.Enabled And vsAdvice.Visible Then vsAdvice.SetFocus
        Call vsAdvice_GotFocus
        Call vsExpense_LostFocus
    Else
        If vsExpense.Enabled And vsExpense.Visible Then vsExpense.SetFocus
        Call vsExpense_GotFocus
        Call vsAdvice_LostFocus
    End If
End Sub
Public Function zlInitCommon(ByRef objSquareCard As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化公共接口
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-08 10:25:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If CreatePubAdvice = False Then Exit Function
    If objSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If mobjSquareCard.zlInitComponents(Me, p医嘱附费管理, glngSys, gstrDBUser, gcnOracle, False) = False Then
            Set mobjSquareCard = Nothing
            MsgBox "医疗卡部件（zl9CardSquare）初始化失败！", vbInformation, gstrSysName
        End If
    Else
        Set mobjSquareCard = objSquareCard
    End If
    zlInitCommon = True
End Function

