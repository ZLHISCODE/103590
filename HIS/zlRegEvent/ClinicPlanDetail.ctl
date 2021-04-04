VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.UserControl ClinicPlanDetail 
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   ScaleHeight     =   8640
   ScaleWidth      =   11475
   Begin VB.PictureBox picUnit 
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   5280
      ScaleHeight     =   3510
      ScaleWidth      =   4890
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4050
      Width           =   4890
      Begin zl9RegEvent.ClinicPlanUnit cpuUnit 
         Height          =   2595
         Left            =   480
         TabIndex        =   14
         Top             =   330
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   4577
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shpUnit 
         BorderColor     =   &H8000000A&
         Height          =   2865
         Left            =   360
         Top             =   210
         Width           =   3885
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   270
      ScaleHeight     =   345
      ScaleWidth      =   11025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11025
      Begin VB.ComboBox cboDespeakType 
         Height          =   300
         Left            =   7545
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   30
         Width           =   1845
      End
      Begin VB.CheckBox chk启用时段 
         Caption         =   "启用时段"
         Height          =   360
         Left            =   5505
         TabIndex        =   6
         Top             =   0
         Width           =   1140
      End
      Begin VB.CheckBox chk序号控制 
         Caption         =   "启用序号控制"
         Height          =   360
         Left            =   4005
         TabIndex        =   5
         Top             =   0
         Width           =   1425
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   4
         Top             =   30
         Width           =   1170
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   615
         TabIndex        =   2
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label lbl预约控制 
         AutoSize        =   -1  'True
         Caption         =   "预约控制"
         Height          =   180
         Left            =   6795
         TabIndex        =   7
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "限约数"
         Height          =   180
         Index           =   1
         Left            =   1980
         TabIndex        =   3
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "限号数"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.PictureBox picRoom 
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   255
      ScaleHeight     =   3510
      ScaleWidth      =   4890
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4065
      Width           =   4890
      Begin zl9RegEvent.ClinicPlanOffice cpoRoom 
         Height          =   2985
         Left            =   270
         TabIndex        =   12
         Top             =   270
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5265
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape shapRoom 
         BorderColor     =   &H8000000A&
         Height          =   3315
         Left            =   180
         Top             =   60
         Width           =   4530
      End
   End
   Begin VB.PictureBox picWorkTime 
      BorderStyle     =   0  'None
      Height          =   3225
      Left            =   570
      ScaleHeight     =   3225
      ScaleWidth      =   5820
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   510
      Width           =   5820
      Begin zl9RegEvent.ClinicPlanWorkTimeNum WorkTimeNum 
         Height          =   2835
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   5001
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IsDataChanged   =   -1  'True
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "ClinicPlanDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'缺省属性值
Const m_def_BackColor = vbButtonFace
Const m_def_BackStyle = 0
'属性变量:
Dim m_IsDataChanged As Boolean
Dim m_EditMode As gRegistPlanEditMode
Private m_BackColor As Long

Private Enum mPan_IDX
    pan_FilterSet = 1
    pan_WorkTimeNum = 2
    pan_room = 3
    pan_CooperateUnit = 4   '合作单位
End Enum
Private mobj出诊记录 As 出诊记录
Private mobj所有门诊诊室 As 分诊诊室集
Private mobj所有合作单位 As 合作单位控制集
Private mblnNotClick As Boolean
Private mblnGetFocus As Boolean '文本框是否获得了焦点，因为改变了文本框的值，直接点击菜单不会触发失去焦点事件
Private mobjCurActiveControl As Object '当前激活控件
Private mblnValiedCanSave As Boolean
'缺省属性值:
Const m_def_IsDataChanged = False
Const m_def_EditMode = 0
'事件声明:
Event DataIsChanged()



Public Function LoadData(ByVal obj出诊记录 As 出诊记录, ByVal obj所有合作单位 As 合作单位控制集, _
    Optional ByVal obj所有门诊诊室 As 分诊诊室集) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载出诊记录
    '入参:
    '   obj出诊记录-出诊记录对象
    '   obj所有合作单位 - 合作单位控制集对象
    '   obj所有门诊诊室 - 分诊诊室集对象
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    m_IsDataChanged = False
    Set mobj出诊记录 = obj出诊记录
    If mobj出诊记录 Is Nothing Then Set mobj出诊记录 = New 出诊记录
    Set mobj所有门诊诊室 = obj所有门诊诊室: Set mobj所有合作单位 = obj所有合作单位
    If Not obj出诊记录 Is Nothing Then m_IsDataChanged = obj出诊记录.是否修改
    
    Call LockWindowUpdate(UserControl.Hwnd)
    blnOK = InitData
    With mobj出诊记录
        blnOK = cpoRoom.LoadData(.安排门诊诊室集, mobj所有门诊诊室, m_IsDataChanged)
        blnOK = WorkTimeNum.LoadData(.号序信息集, .上班时段, , m_IsDataChanged)
        blnOK = cpuUnit.LoadData(.合作单位控制集, .号序信息集, mobj所有合作单位, m_IsDataChanged)

        '未设置限号数时，缺省按照出诊频次计算
        If Val(txtEdit(0).Text) = 0 Then
            If .号序信息集.出诊频次 <> 0 And .是否分时段 Then
                .限号数 = GetMinuteCount(.上班时段.开始时间, .上班时段.结束时间, .上班时段.休息时段) \ .号序信息集.出诊频次
                txtEdit(0).Text = .限号数
                WorkTimeNum.限号数 = .限号数
            End If
        End If
    End With
    Call LockWindowUpdate(0)
    mblnNotClick = False
    LoadData = blnOK
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化面版信息
    '编制:刘兴洪
    '日期:2016-01-13 09:40:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    With cboDespeakType
        .Clear
        .AddItem "0-允许预约": .ItemData(.NewIndex) = 0
        .AddItem "1-禁止预约": .ItemData(.NewIndex) = 1
        .AddItem "2-仅禁止三方机构预约": .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer
     
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    intType = IIf(mobj出诊记录.预约控制 <= 2, mobj出诊记录.预约控制, 0)
    zlControl.CboLocate cboDespeakType, intType, True
    
    txtEdit(1).Text = IIf(mobj出诊记录.限约数 = 0, "", mobj出诊记录.限约数)
    txtEdit(0).Text = IIf(mobj出诊记录.限号数 = 0, "", mobj出诊记录.限号数)
    chk序号控制.Value = IIf(mobj出诊记录.是否序号控制, 1, 0)
    chk启用时段.Value = IIf(mobj出诊记录.是否分时段, 1, 0)
    mblnNotClick = False
    SetPancelVisible
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cboDespeakType_GotFocus()
    Set mobjCurActiveControl = cboDespeakType
End Sub

Private Sub cboDespeakType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk启用时段_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk序号控制_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cpoRoom_DataIsChanged()
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
End Sub

Private Sub cpoRoom_GotFocus()
    Set mobjCurActiveControl = cpoRoom
End Sub

Private Sub cpuUnit_DataIsChanged()
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
End Sub

Private Sub cpuUnit_GotFocus()
    Set mobjCurActiveControl = cpuUnit
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub cboDespeakType_Click()
    On Error GoTo Errhand
    
    If mblnNotClick Then Exit Sub
    
    If Not mobj出诊记录 Is Nothing Then mobj出诊记录.预约控制 = cboDespeakType.ListIndex
    WorkTimeNum.预约控制 = cboDespeakType.ListIndex
    cpuUnit.预约控制 = cboDespeakType.ListIndex
    
    '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
    If cboDespeakType.ListIndex = 1 Then
        txtEdit(1).Text = ""
        WorkTimeNum.限约数 = 0
        If chk序号控制.Value = vbUnchecked Then
            '禁止预约时，如果未启用序号控制，则启用时段也自动取消
            If chk启用时段.Value = vbChecked Then
                chk启用时段.Value = vbUnchecked
            End If
        End If
    Else
        If Val(txtEdit(0).Text) <> 0 Then
            txtEdit(1).Text = txtEdit(0).Text
            WorkTimeNum.限约数 = Val(txtEdit(0).Text)
        End If
    End If
    Me.EditMode = m_EditMode '重设编辑状态
    
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chk启用时段_Click()
    On Error GoTo Errhand
    If mblnNotClick = True Then Exit Sub
    
    Call SetPancelVisible
    WorkTimeNum.启用时段 = IIf(chk启用时段.Value = 1, True, False)
    Call SetFocusControl
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub chk启用时段_GotFocus()
    chk启用时段.BackColor = GCTRL_SELBACK_COLOR
    Set mobjCurActiveControl = chk启用时段
End Sub

Private Sub chk启用时段_LostFocus()
    chk启用时段.BackColor = picFilter.BackColor
End Sub

Private Sub chk序号控制_Click()
    On Error GoTo Errhand
    If mblnNotClick = True Then Exit Sub
    
    WorkTimeNum.启用序号控制 = IIf(chk序号控制.Value = 1, True, False)
    If chk序号控制.Enabled And chk序号控制.Visible Then chk序号控制.SetFocus
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chk序号控制_GotFocus()
    chk序号控制.BackColor = GCTRL_SELBACK_COLOR
    Set mobjCurActiveControl = chk序号控制
End Sub

Private Sub chk序号控制_LostFocus()
    chk序号控制.BackColor = picFilter.BackColor
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case pan_FilterSet
        Item.Handle = picFilter.Hwnd
    Case pan_WorkTimeNum
        Item.Handle = picWorkTime.Hwnd
    Case pan_room
        Item.Handle = picRoom.Hwnd
    Case pan_CooperateUnit
        Item.Handle = PicUnit.Hwnd
    End Select
End Sub

Private Sub picRoom_GotFocus()
    Call SetFocusControl
End Sub

Private Sub SetFocusControl()
    '改变焦点位置
    On Error Resume Next
    If mobjCurActiveControl Is Nothing Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If mobjCurActiveControl.Visible And mobjCurActiveControl.Enabled Then
            mobjCurActiveControl.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub picUnit_GotFocus()
    Call SetFocusControl
End Sub

Private Sub PicUnit_Resize()
    Err = 0: On Error Resume Next
    With PicUnit
        shpUnit.Left = .ScaleLeft
        shpUnit.Top = .ScaleTop
        shpUnit.Width = .ScaleWidth - shapRoom.Left * 2
        shpUnit.Height = .ScaleHeight - shapRoom.Top * 2
        
        cpuUnit.Left = .ScaleLeft + 30
        cpuUnit.Top = .ScaleTop + 30
        cpuUnit.Width = .ScaleWidth - 60
        cpuUnit.Height = .ScaleHeight - 60
    End With
End Sub

Private Sub picRoom_Resize()
    Err = 0: On Error Resume Next
    With picRoom
        shapRoom.Left = .ScaleLeft
        shapRoom.Top = .ScaleTop
        shapRoom.Width = .ScaleWidth - shapRoom.Left * 2
        shapRoom.Height = .ScaleHeight - shapRoom.Top * 2
        
        cpoRoom.Left = .ScaleLeft + 30
        cpoRoom.Top = .ScaleTop + 30
        cpoRoom.Width = .ScaleWidth - 60
        cpoRoom.Height = .ScaleHeight - 60
    End With
End Sub

Private Sub picWorkTime_GotFocus()
    Call SetFocusControl
End Sub

Private Sub picWorkTime_Resize()
    Err = 0: On Error Resume Next
    With picWorkTime
        WorkTimeNum.Left = .ScaleLeft
        WorkTimeNum.Top = .ScaleTop
        WorkTimeNum.Width = .ScaleWidth
        WorkTimeNum.Height = .ScaleHeight
    End With
End Sub
Private Sub SetPancelVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置Dock控件的显示
    '编制:刘兴洪
    '日期:2016-01-13 17:50:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPan As Pane
    
    Err = 0: On Error GoTo Errhand:
    If Not mobj出诊记录 Is Nothing Then
        If m_EditMode = ED_RegistPlan_UpdateUnit Then
            picFilter.Enabled = False

            Set objPan = dkpMain.FindPane(mPan_IDX.pan_WorkTimeNum)
            If Not objPan Is Nothing Then
                If Not objPan.Closed Then objPan.Close
            End If

            Set objPan = dkpMain.FindPane(mPan_IDX.pan_room)
            If Not objPan Is Nothing Then
                If Not objPan.Closed Then objPan.Close
            End If

            Set objPan = dkpMain.FindPane(mPan_IDX.pan_CooperateUnit)
            If Not objPan Is Nothing Then
                objPan.Closed = False
            End If
            Exit Sub
        End If
    End If

    If chk启用时段.Value = 0 Then
        Set objPan = dkpMain.FindPane(mPan_IDX.pan_WorkTimeNum)
        If Not objPan Is Nothing Then
            If Not objPan.Closed Then objPan.Close
        End If
    Else
        Set objPan = dkpMain.FindPane(mPan_IDX.pan_WorkTimeNum)
        If Not objPan Is Nothing Then
            objPan.Closed = False
        End If
    End If

    '无合作单位和预约方式时也不显示合作单位预约挂号控制
    If cboDespeakType.ItemData(cboDespeakType.ListIndex) = 1 Then
        Set objPan = dkpMain.FindPane(mPan_IDX.pan_CooperateUnit)
        If Not objPan Is Nothing Then
            If Not objPan.Closed Then objPan.Close
        End If
    Else
        Set objPan = dkpMain.FindPane(mPan_IDX.pan_CooperateUnit)
        If Not objPan Is Nothing Then
            objPan.Closed = False
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化Docking控件
    '编制:刘兴洪
    '日期:2016-01-08 14:34:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    Dim strReg As String
    Dim panThis As Pane
    
    On Error GoTo Errhand
    sngWidth = picFilter.Width / Screen.TwipsPerPixelX
    sngHeight = picFilter.Height / Screen.TwipsPerPixelY
    
    Set panThis = dkpMain.CreatePane(pan_FilterSet, sngWidth, sngHeight, DockTopOf)
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Title = "": panThis.Tag = pan_FilterSet
    panThis.Handle = picFilter.Hwnd
    panThis.MinTrackSize.Height = sngHeight
    panThis.MaxTrackSize.Height = sngHeight
    
    Set panThis = dkpMain.CreatePane(pan_WorkTimeNum, sngWidth, 300, DockBottomOf, panThis)
    panThis.Title = "上班时间"
    panThis.Tag = pan_WorkTimeNum
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picWorkTime.Hwnd
    
    Set panThis = dkpMain.CreatePane(pan_room, sngWidth * 4 / 7, 300, DockBottomOf, panThis)
    panThis.Title = "分诊诊室"
    panThis.Tag = pan_room
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picRoom.Hwnd
    
    Set panThis = dkpMain.CreatePane(pan_CooperateUnit, sngWidth, 300, DockRightOf, panThis)
    panThis.Title = "合作单位控制"
    panThis.Tag = pan_CooperateUnit
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = PicUnit.Hwnd
     
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    dkpMain.NormalizeSplitters
    'Set dkpMain.PaintManager.CaptionFont = use.Font
    
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtEdit_Change(index As Integer)
    If index = 0 Then
        txtEdit(1).Enabled = m_EditMode = ED_RegistPlan_Edit And Val(txtEdit(0).Text) <> 0 And cboDespeakType.ListIndex <> 1
        SetEnabledBackColor UserControl.Controls
    End If
    
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
End Sub

Private Sub txtEdit_GotFocus(index As Integer)
    zlControl.TxtSelAll txtEdit(index)
    Set mobjCurActiveControl = txtEdit(index)
    mblnGetFocus = True
End Sub

Private Sub txtEdit_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If Len(Trim(txtEdit(index).Text)) >= 9 And txtEdit(index).SelText = "" Then KeyAscii = 0
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEdit_LostFocus(index As Integer)
    On Error GoTo Errhand
    mblnGetFocus = False
    If index = 0 Then
        WorkTimeNum.限号数 = Val(txtEdit(index).Text)
    Else
        WorkTimeNum.限约数 = Val(txtEdit(index).Text)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtEdit_Validate(index As Integer, Cancel As Boolean)
    If index = 0 Then
        If Val(txtEdit(0).Text) > 10000 Then
            MsgBox "限号数太大，请重新输入！", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            mblnValiedCanSave = False
            Cancel = True: Exit Sub
        End If
        If Val(txtEdit(0).Text) = 0 Then
            txtEdit(1).Text = ""
            WorkTimeNum.限约数 = 0
        ElseIf WorkTimeNum.预约控制 <> 1 Then
            txtEdit(1).Text = Val(txtEdit(0).Text)
            WorkTimeNum.限约数 = Val(txtEdit(0).Text)
        End If
    Else
        If Val(txtEdit(1).Text) > Val(txtEdit(0).Text) Then
            MsgBox "限约数应小于等于限号数！", vbInformation, gstrSysName
            txtEdit(1).Text = txtEdit(0).Text
            zlControl.TxtSelAll txtEdit(1)
            mblnValiedCanSave = False
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Call InitFace
    Call InitPanel
    Call SetPancelVisible
End Sub

Public Property Get Get出诊记录() As 出诊记录
    Dim obj出诊记录 As New 出诊记录
    
    On Error GoTo Errhand
    If mblnGetFocus Then
        '保证失去焦点时触发事件
        If UserControl.ActiveControl Is txtEdit(0) Then
            Call txtEdit_Validate(0, False)
            Call txtEdit_LostFocus(0)
        Else
            Call txtEdit_Validate(1, False)
            Call txtEdit_LostFocus(1)
        End If
    End If
    '数据未改变，直接返回原集合的副本
    If m_IsDataChanged = False And mobj出诊记录.是否修改 = False Then
        Set Get出诊记录 = mobj出诊记录.Clone
        Exit Function
    End If
    
    '数据已改变，重新构造集合对象
    Set obj出诊记录 = mobj出诊记录.Clone
    With obj出诊记录
        .是否修改 = True
        .限号数 = Val(txtEdit(0).Text)
        .限约数 = Val(txtEdit(1).Text)
        .是否分时段 = chk启用时段.Value = 1
        .是否序号控制 = chk序号控制.Value = 1
        .预约控制 = cboDespeakType.ItemData(cboDespeakType.ListIndex)
        If .限约数 = 0 And .限号数 <> 0 Then .限约数 = .限号数
        If .预约控制 = 1 Or .限号数 = 0 Then .限约数 = 0
        
        If Not .安排门诊诊室集 Is Nothing Then .安排门诊诊室集.RemoveAll
        If dkpMain(mPan_IDX.pan_room).Closed = False Then Set .安排门诊诊室集 = cpoRoom.Get安排门诊诊室集
        If Not .号序信息集 Is Nothing Then .号序信息集.RemoveAll
        If dkpMain(mPan_IDX.pan_WorkTimeNum).Closed = False Or .是否序号控制 Then Set .号序信息集 = WorkTimeNum.Get号序集
        If Not .合作单位控制集 Is Nothing Then .合作单位控制集.RemoveAll
        If dkpMain(mPan_IDX.pan_CooperateUnit).Closed = False Then Set .合作单位控制集 = cpuUnit.Get合作单位控制信息集
        .分诊方式 = .安排门诊诊室集.分诊方式
        .是否独占 = .合作单位控制集.是否独占
    End With
    Set Get出诊记录 = obj出诊记录
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_EditMode = m_def_EditMode
    m_IsDataChanged = m_def_IsDataChanged
    Set mobjCurActiveControl = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    
    SetBackColor Controls, m_BackColor
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
    
    WorkTimeNum.EditMode = m_EditMode
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    WorkTimeNum.诊疗频次 = PropBag.ReadProperty("诊疗频次", 5)
    cpoRoom.医生姓名 = PropBag.ReadProperty("医生姓名", "")
End Sub

Private Sub UserControl_Terminate()
    Set mobj出诊记录 = Nothing
    Set mobj所有门诊诊室 = Nothing
    Set mobj所有合作单位 = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("诊疗频次", WorkTimeNum.诊疗频次, 5)
    Call PropBag.WriteProperty("医生姓名", cpoRoom.医生姓名, "")
End Sub

Private Sub WorkTimeNum_DataIsChanged()
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
End Sub

Private Sub WorkTimeNum_GotFocus()
    Set mobjCurActiveControl = WorkTimeNum
End Sub

Private Sub WorkTimeNum_TimeIntervalsChanged(ByVal obj号序信息集 As 号序信息集, ByVal blnClearUnit As Boolean)
    On Error GoTo Errhand
    Set mobj出诊记录.合作单位控制集 = cpuUnit.Get合作单位控制信息集
    If mobj出诊记录.合作单位控制集 Is Nothing Then Set mobj出诊记录.合作单位控制集 = New 合作单位控制集
    If blnClearUnit Then
        '改变了分时段或序号控制时，如果预约控制方式是按序号控制预约，则清除已设置信息
        If mobj出诊记录.合作单位控制集.预约控制方式 = 3 Then
            mobj出诊记录.合作单位控制集.RemoveAll
        End If
    End If
    
    cpuUnit.LoadData mobj出诊记录.合作单位控制集, obj号序信息集, mobj所有合作单位
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function IsValied() As Boolean
    '检查数据
    Dim intCount As Integer
    
    Err = 0: On Error GoTo errHandler
    If mblnGetFocus Then
        '保证失去焦点时触发事件
        If UserControl.ActiveControl Is txtEdit(0) Then
            mblnValiedCanSave = True
            Call txtEdit_Validate(0, False)
            If mblnValiedCanSave = False Then Exit Function
            Call txtEdit_LostFocus(0)
        Else
            mblnValiedCanSave = True
            Call txtEdit_Validate(1, False)
            If mblnValiedCanSave = False Then Exit Function
            Call txtEdit_LostFocus(1)
        End If
    End If
    
    '数据未改变不检查
    If m_IsDataChanged = False Then IsValied = True: Exit Function
    If zlCommFun.ActualLen(txtEdit(0).Text) > 9 Then
        MsgBox "限号数不能超过999999999！", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If zlCommFun.ActualLen(txtEdit(1).Text) > 9 Then
        MsgBox "限约数不能超过999999999！", vbInformation, gstrSysName
        If txtEdit(1).Visible And txtEdit(1).Enabled Then txtEdit(1).SetFocus
        zlControl.TxtSelAll txtEdit(1)
        Exit Function
    End If
    
    If chk序号控制.Value = 1 And Val(txtEdit(0)) = 0 Then
        MsgBox "启用序号控制必须设置限号数！", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    
    If chk启用时段.Value = 1 And Val(txtEdit(0)) = 0 Then
        MsgBox "启用时段必须设置限号数！", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If

    If Val(txtEdit(0).Text) <> 0 Then
        If Val(txtEdit(0).Text) <> 0 And Val(txtEdit(0).Text) < Val(txtEdit(1).Text) Then
            MsgBox "限约数不能超过限号数！", vbInformation, gstrSysName
            If txtEdit(1).Visible And txtEdit(1).Enabled Then txtEdit(1).SetFocus
            txtEdit(1).Text = txtEdit(0).Text
            zlControl.TxtSelAll txtEdit(1)
            Exit Function
        End If
    ElseIf Val(txtEdit(1).Text) <> 0 Then
        MsgBox "设置了限约数必须设置限号数！", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    
    '门诊诊室
    If cpoRoom.IsValied() = False Then Exit Function
    '号序
    If dkpMain(mPan_IDX.pan_WorkTimeNum).Closed = False Then
        If WorkTimeNum.IsValied(m_IsDataChanged) = False Then Exit Function
    End If
    '合作单位
    If dkpMain(mPan_IDX.pan_CooperateUnit).Closed = False Then
        If cpuUnit.IsValied(m_IsDataChanged) = False Then Exit Function
    End If
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor Controls, m_BackColor
'    On Error Resume Next
'    dkpMain.PanelPaintManager.OneNoteColors = True
End Property
'注意！不要删除或修改下列被注释的行！
'MemberInfo=26,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    SetEnabled UserControl.Controls, m_EditMode = ED_RegistPlan_Edit
    If Not mobj出诊记录 Is Nothing Then
        txtEdit(1).Enabled = m_EditMode = ED_RegistPlan_Edit And mobj出诊记录.预约控制 <> 1
    Else
        txtEdit(1).Enabled = m_EditMode = ED_RegistPlan_Edit
    End If
    SetEnabledBackColor UserControl.Controls
    SetPancelVisible
    
    WorkTimeNum.EditMode = m_EditMode
    cpuUnit.EditMode = m_EditMode
    cpoRoom.EditMode = m_EditMode
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=WorkTimeNum,WorkTimeNum,-1,诊疗频次
Public Property Get 诊疗频次() As Integer
    诊疗频次 = WorkTimeNum.诊疗频次
End Property

Public Property Let 诊疗频次(ByVal New_诊疗频次 As Integer)
    WorkTimeNum.诊疗频次() = New_诊疗频次
    PropertyChanged "诊疗频次"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,""
Public Property Get 医生姓名() As String
    医生姓名 = cpoRoom.医生姓名
End Property

Public Property Let 医生姓名(ByVal New_医生姓名 As String)
    cpoRoom.医生姓名 = New_医生姓名
    PropertyChanged "医生姓名"
End Property

