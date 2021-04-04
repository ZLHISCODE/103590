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
            Name            =   "����"
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
      Begin VB.CheckBox chk����ʱ�� 
         Caption         =   "����ʱ��"
         Height          =   360
         Left            =   5505
         TabIndex        =   6
         Top             =   0
         Width           =   1140
      End
      Begin VB.CheckBox chk��ſ��� 
         Caption         =   "������ſ���"
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
      Begin VB.Label lblԤԼ���� 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����"
         Height          =   180
         Left            =   6795
         TabIndex        =   7
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��Լ��"
         Height          =   180
         Index           =   1
         Left            =   1980
         TabIndex        =   3
         Top             =   90
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�޺���"
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
            Name            =   "����"
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
            Name            =   "����"
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
'ȱʡ����ֵ
Const m_def_BackColor = vbButtonFace
Const m_def_BackStyle = 0
'���Ա���:
Dim m_IsDataChanged As Boolean
Dim m_EditMode As gRegistPlanEditMode
Private m_BackColor As Long

Private Enum mPan_IDX
    pan_FilterSet = 1
    pan_WorkTimeNum = 2
    pan_room = 3
    pan_CooperateUnit = 4   '������λ
End Enum
Private mobj�����¼ As �����¼
Private mobj������������ As �������Ҽ�
Private mobj���к�����λ As ������λ���Ƽ�
Private mblnNotClick As Boolean
Private mblnGetFocus As Boolean '�ı����Ƿ����˽��㣬��Ϊ�ı����ı����ֵ��ֱ�ӵ���˵����ᴥ��ʧȥ�����¼�
Private mobjCurActiveControl As Object '��ǰ����ؼ�
Private mblnValiedCanSave As Boolean
'ȱʡ����ֵ:
Const m_def_IsDataChanged = False
Const m_def_EditMode = 0
'�¼�����:
Event DataIsChanged()



Public Function LoadData(ByVal obj�����¼ As �����¼, ByVal obj���к�����λ As ������λ���Ƽ�, _
    Optional ByVal obj������������ As �������Ҽ�) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س����¼
    '���:
    '   obj�����¼-�����¼����
    '   obj���к�����λ - ������λ���Ƽ�����
    '   obj������������ - �������Ҽ�����
    '����:
    '����:
    '����:���˺�
    '����:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    m_IsDataChanged = False
    Set mobj�����¼ = obj�����¼
    If mobj�����¼ Is Nothing Then Set mobj�����¼ = New �����¼
    Set mobj������������ = obj������������: Set mobj���к�����λ = obj���к�����λ
    If Not obj�����¼ Is Nothing Then m_IsDataChanged = obj�����¼.�Ƿ��޸�
    
    Call LockWindowUpdate(UserControl.Hwnd)
    blnOK = InitData
    With mobj�����¼
        blnOK = cpoRoom.LoadData(.�����������Ҽ�, mobj������������, m_IsDataChanged)
        blnOK = WorkTimeNum.LoadData(.������Ϣ��, .�ϰ�ʱ��, , m_IsDataChanged)
        blnOK = cpuUnit.LoadData(.������λ���Ƽ�, .������Ϣ��, mobj���к�����λ, m_IsDataChanged)

        'δ�����޺���ʱ��ȱʡ���ճ���Ƶ�μ���
        If Val(txtEdit(0).Text) = 0 Then
            If .������Ϣ��.����Ƶ�� <> 0 And .�Ƿ��ʱ�� Then
                .�޺��� = GetMinuteCount(.�ϰ�ʱ��.��ʼʱ��, .�ϰ�ʱ��.����ʱ��, .�ϰ�ʱ��.��Ϣʱ��) \ .������Ϣ��.����Ƶ��
                txtEdit(0).Text = .�޺���
                WorkTimeNum.�޺��� = .�޺���
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
    '����:��ʼ�������Ϣ
    '����:���˺�
    '����:2016-01-13 09:40:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    With cboDespeakType
        .Clear
        .AddItem "0-����ԤԼ": .ItemData(.NewIndex) = 0
        .AddItem "1-��ֹԤԼ": .ItemData(.NewIndex) = 1
        .AddItem "2-����ֹ��������ԤԼ": .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer
     
    Err = 0: On Error GoTo Errhand:
    mblnNotClick = True
    intType = IIf(mobj�����¼.ԤԼ���� <= 2, mobj�����¼.ԤԼ����, 0)
    zlControl.CboLocate cboDespeakType, intType, True
    
    txtEdit(1).Text = IIf(mobj�����¼.��Լ�� = 0, "", mobj�����¼.��Լ��)
    txtEdit(0).Text = IIf(mobj�����¼.�޺��� = 0, "", mobj�����¼.�޺���)
    chk��ſ���.Value = IIf(mobj�����¼.�Ƿ���ſ���, 1, 0)
    chk����ʱ��.Value = IIf(mobj�����¼.�Ƿ��ʱ��, 1, 0)
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

Private Sub chk����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk��ſ���_KeyPress(KeyAscii As Integer)
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
    
    If Not mobj�����¼ Is Nothing Then mobj�����¼.ԤԼ���� = cboDespeakType.ListIndex
    WorkTimeNum.ԤԼ���� = cboDespeakType.ListIndex
    cpuUnit.ԤԼ���� = cboDespeakType.ListIndex
    
    '0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
    If cboDespeakType.ListIndex = 1 Then
        txtEdit(1).Text = ""
        WorkTimeNum.��Լ�� = 0
        If chk��ſ���.Value = vbUnchecked Then
            '��ֹԤԼʱ�����δ������ſ��ƣ�������ʱ��Ҳ�Զ�ȡ��
            If chk����ʱ��.Value = vbChecked Then
                chk����ʱ��.Value = vbUnchecked
            End If
        End If
    Else
        If Val(txtEdit(0).Text) <> 0 Then
            txtEdit(1).Text = txtEdit(0).Text
            WorkTimeNum.��Լ�� = Val(txtEdit(0).Text)
        End If
    End If
    Me.EditMode = m_EditMode '����༭״̬
    
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chk����ʱ��_Click()
    On Error GoTo Errhand
    If mblnNotClick = True Then Exit Sub
    
    Call SetPancelVisible
    WorkTimeNum.����ʱ�� = IIf(chk����ʱ��.Value = 1, True, False)
    Call SetFocusControl
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub chk����ʱ��_GotFocus()
    chk����ʱ��.BackColor = GCTRL_SELBACK_COLOR
    Set mobjCurActiveControl = chk����ʱ��
End Sub

Private Sub chk����ʱ��_LostFocus()
    chk����ʱ��.BackColor = picFilter.BackColor
End Sub

Private Sub chk��ſ���_Click()
    On Error GoTo Errhand
    If mblnNotClick = True Then Exit Sub
    
    WorkTimeNum.������ſ��� = IIf(chk��ſ���.Value = 1, True, False)
    If chk��ſ���.Enabled And chk��ſ���.Visible Then chk��ſ���.SetFocus
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chk��ſ���_GotFocus()
    chk��ſ���.BackColor = GCTRL_SELBACK_COLOR
    Set mobjCurActiveControl = chk��ſ���
End Sub

Private Sub chk��ſ���_LostFocus()
    chk��ſ���.BackColor = picFilter.BackColor
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
    '�ı佹��λ��
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
    '����:����Dock�ؼ�����ʾ
    '����:���˺�
    '����:2016-01-13 17:50:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPan As Pane
    
    Err = 0: On Error GoTo Errhand:
    If Not mobj�����¼ Is Nothing Then
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

    If chk����ʱ��.Value = 0 Then
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

    '�޺�����λ��ԤԼ��ʽʱҲ����ʾ������λԤԼ�Һſ���
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
    '����:��ʼ��Docking�ؼ�
    '����:���˺�
    '����:2016-01-08 14:34:56
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
    panThis.Title = "�ϰ�ʱ��"
    panThis.Tag = pan_WorkTimeNum
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picWorkTime.Hwnd
    
    Set panThis = dkpMain.CreatePane(pan_room, sngWidth * 4 / 7, 300, DockBottomOf, panThis)
    panThis.Title = "��������"
    panThis.Tag = pan_room
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picRoom.Hwnd
    
    Set panThis = dkpMain.CreatePane(pan_CooperateUnit, sngWidth, 300, DockRightOf, panThis)
    panThis.Title = "������λ����"
    panThis.Tag = pan_CooperateUnit
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = PicUnit.Hwnd
     
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    dkpMain.NormalizeSplitters
    'Set dkpMain.PaintManager.CaptionFont = use.Font
    
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
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
        WorkTimeNum.�޺��� = Val(txtEdit(index).Text)
    Else
        WorkTimeNum.��Լ�� = Val(txtEdit(index).Text)
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
            MsgBox "�޺���̫�����������룡", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(0)
            mblnValiedCanSave = False
            Cancel = True: Exit Sub
        End If
        If Val(txtEdit(0).Text) = 0 Then
            txtEdit(1).Text = ""
            WorkTimeNum.��Լ�� = 0
        ElseIf WorkTimeNum.ԤԼ���� <> 1 Then
            txtEdit(1).Text = Val(txtEdit(0).Text)
            WorkTimeNum.��Լ�� = Val(txtEdit(0).Text)
        End If
    Else
        If Val(txtEdit(1).Text) > Val(txtEdit(0).Text) Then
            MsgBox "��Լ��ӦС�ڵ����޺�����", vbInformation, gstrSysName
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

Public Property Get Get�����¼() As �����¼
    Dim obj�����¼ As New �����¼
    
    On Error GoTo Errhand
    If mblnGetFocus Then
        '��֤ʧȥ����ʱ�����¼�
        If UserControl.ActiveControl Is txtEdit(0) Then
            Call txtEdit_Validate(0, False)
            Call txtEdit_LostFocus(0)
        Else
            Call txtEdit_Validate(1, False)
            Call txtEdit_LostFocus(1)
        End If
    End If
    '����δ�ı䣬ֱ�ӷ���ԭ���ϵĸ���
    If m_IsDataChanged = False And mobj�����¼.�Ƿ��޸� = False Then
        Set Get�����¼ = mobj�����¼.Clone
        Exit Function
    End If
    
    '�����Ѹı䣬���¹��켯�϶���
    Set obj�����¼ = mobj�����¼.Clone
    With obj�����¼
        .�Ƿ��޸� = True
        .�޺��� = Val(txtEdit(0).Text)
        .��Լ�� = Val(txtEdit(1).Text)
        .�Ƿ��ʱ�� = chk����ʱ��.Value = 1
        .�Ƿ���ſ��� = chk��ſ���.Value = 1
        .ԤԼ���� = cboDespeakType.ItemData(cboDespeakType.ListIndex)
        If .��Լ�� = 0 And .�޺��� <> 0 Then .��Լ�� = .�޺���
        If .ԤԼ���� = 1 Or .�޺��� = 0 Then .��Լ�� = 0
        
        If Not .�����������Ҽ� Is Nothing Then .�����������Ҽ�.RemoveAll
        If dkpMain(mPan_IDX.pan_room).Closed = False Then Set .�����������Ҽ� = cpoRoom.Get�����������Ҽ�
        If Not .������Ϣ�� Is Nothing Then .������Ϣ��.RemoveAll
        If dkpMain(mPan_IDX.pan_WorkTimeNum).Closed = False Or .�Ƿ���ſ��� Then Set .������Ϣ�� = WorkTimeNum.Get����
        If Not .������λ���Ƽ� Is Nothing Then .������λ���Ƽ�.RemoveAll
        If dkpMain(mPan_IDX.pan_CooperateUnit).Closed = False Then Set .������λ���Ƽ� = cpuUnit.Get������λ������Ϣ��
        .���﷽ʽ = .�����������Ҽ�.���﷽ʽ
        .�Ƿ��ռ = .������λ���Ƽ�.�Ƿ��ռ
    End With
    Set Get�����¼ = obj�����¼
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
    WorkTimeNum.����Ƶ�� = PropBag.ReadProperty("����Ƶ��", 5)
    cpoRoom.ҽ������ = PropBag.ReadProperty("ҽ������", "")
End Sub

Private Sub UserControl_Terminate()
    Set mobj�����¼ = Nothing
    Set mobj������������ = Nothing
    Set mobj���к�����λ = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("����Ƶ��", WorkTimeNum.����Ƶ��, 5)
    Call PropBag.WriteProperty("ҽ������", cpoRoom.ҽ������, "")
End Sub

Private Sub WorkTimeNum_DataIsChanged()
    m_IsDataChanged = True
    RaiseEvent DataIsChanged
End Sub

Private Sub WorkTimeNum_GotFocus()
    Set mobjCurActiveControl = WorkTimeNum
End Sub

Private Sub WorkTimeNum_TimeIntervalsChanged(ByVal obj������Ϣ�� As ������Ϣ��, ByVal blnClearUnit As Boolean)
    On Error GoTo Errhand
    Set mobj�����¼.������λ���Ƽ� = cpuUnit.Get������λ������Ϣ��
    If mobj�����¼.������λ���Ƽ� Is Nothing Then Set mobj�����¼.������λ���Ƽ� = New ������λ���Ƽ�
    If blnClearUnit Then
        '�ı��˷�ʱ�λ���ſ���ʱ�����ԤԼ���Ʒ�ʽ�ǰ���ſ���ԤԼ���������������Ϣ
        If mobj�����¼.������λ���Ƽ�.ԤԼ���Ʒ�ʽ = 3 Then
            mobj�����¼.������λ���Ƽ�.RemoveAll
        End If
    End If
    
    cpuUnit.LoadData mobj�����¼.������λ���Ƽ�, obj������Ϣ��, mobj���к�����λ
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function IsValied() As Boolean
    '�������
    Dim intCount As Integer
    
    Err = 0: On Error GoTo errHandler
    If mblnGetFocus Then
        '��֤ʧȥ����ʱ�����¼�
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
    
    '����δ�ı䲻���
    If m_IsDataChanged = False Then IsValied = True: Exit Function
    If zlCommFun.ActualLen(txtEdit(0).Text) > 9 Then
        MsgBox "�޺������ܳ���999999999��", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If zlCommFun.ActualLen(txtEdit(1).Text) > 9 Then
        MsgBox "��Լ�����ܳ���999999999��", vbInformation, gstrSysName
        If txtEdit(1).Visible And txtEdit(1).Enabled Then txtEdit(1).SetFocus
        zlControl.TxtSelAll txtEdit(1)
        Exit Function
    End If
    
    If chk��ſ���.Value = 1 And Val(txtEdit(0)) = 0 Then
        MsgBox "������ſ��Ʊ��������޺�����", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    
    If chk����ʱ��.Value = 1 And Val(txtEdit(0)) = 0 Then
        MsgBox "����ʱ�α��������޺�����", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If

    If Val(txtEdit(0).Text) <> 0 Then
        If Val(txtEdit(0).Text) <> 0 And Val(txtEdit(0).Text) < Val(txtEdit(1).Text) Then
            MsgBox "��Լ�����ܳ����޺�����", vbInformation, gstrSysName
            If txtEdit(1).Visible And txtEdit(1).Enabled Then txtEdit(1).SetFocus
            txtEdit(1).Text = txtEdit(0).Text
            zlControl.TxtSelAll txtEdit(1)
            Exit Function
        End If
    ElseIf Val(txtEdit(1).Text) <> 0 Then
        MsgBox "��������Լ�����������޺�����", vbInformation, gstrSysName
        If txtEdit(0).Visible And txtEdit(0).Enabled Then txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    
    '��������
    If cpoRoom.IsValied() = False Then Exit Function
    '����
    If dkpMain(mPan_IDX.pan_WorkTimeNum).Closed = False Then
        If WorkTimeNum.IsValied(m_IsDataChanged) = False Then Exit Function
    End If
    '������λ
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=26,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    SetEnabled UserControl.Controls, m_EditMode = ED_RegistPlan_Edit
    If Not mobj�����¼ Is Nothing Then
        txtEdit(1).Enabled = m_EditMode = ED_RegistPlan_Edit And mobj�����¼.ԤԼ���� <> 1
    Else
        txtEdit(1).Enabled = m_EditMode = ED_RegistPlan_Edit
    End If
    SetEnabledBackColor UserControl.Controls
    SetPancelVisible
    
    WorkTimeNum.EditMode = m_EditMode
    cpuUnit.EditMode = m_EditMode
    cpoRoom.EditMode = m_EditMode
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=WorkTimeNum,WorkTimeNum,-1,����Ƶ��
Public Property Get ����Ƶ��() As Integer
    ����Ƶ�� = WorkTimeNum.����Ƶ��
End Property

Public Property Let ����Ƶ��(ByVal New_����Ƶ�� As Integer)
    WorkTimeNum.����Ƶ��() = New_����Ƶ��
    PropertyChanged "����Ƶ��"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=13,0,0,""
Public Property Get ҽ������() As String
    ҽ������ = cpoRoom.ҽ������
End Property

Public Property Let ҽ������(ByVal New_ҽ������ As String)
    cpoRoom.ҽ������ = New_ҽ������
    PropertyChanged "ҽ������"
End Property

