VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "��ҳ�༭"
   ClientHeight    =   9855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15105
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   15105
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox PicForm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   7920
      ScaleHeight     =   2865
      ScaleWidth      =   6345
      TabIndex        =   5
      Top             =   600
      Width           =   6375
   End
   Begin VB.PictureBox PicDirectory 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   3720
      ScaleHeight     =   3585
      ScaleWidth      =   2865
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   2895
      Begin MSComctlLib.TreeView tvDirectory 
         Height          =   3255
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5741
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox PicErr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   1320
      ScaleHeight     =   3015
      ScaleWidth      =   11265
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   11265
      Begin VSFlex8Ctl.VSFlexGrid vsErr 
         Height          =   2325
         Left            =   2160
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   7680
         _cx             =   13547
         _cy             =   4101
         Appearance      =   3
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         GridColorFixed  =   16777215
         TreeColor       =   16777215
         FloodColor      =   16777215
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
         ColWidthMin     =   2000
         ColWidthMax     =   2000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmMain.frx":6852
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   4
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
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   9495
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   617
            Picture         =   "frmMain.frx":68B0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23733
         EndProperty
      EndProperty
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
   Begin zlSubclass.Subclass subcMain 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image imgWarn 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   3120
      Picture         =   "frmMain.frx":7144
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   645
   End
   Begin VB.Image imgError 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   1320
      Picture         =   "frmMain.frx":992A
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   645
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":C336
      Left            =   1680
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnModifyVisible As Boolean
Private mfrmMedRecEdit As Object    '�༭����

Public Function ShowMe(ByVal blnModal As Boolean) As Boolean
        If gclsPros.FuncType = f������ҳ Then
            Select Case gclsPros.MedPageSandard
                Case ST_��������׼
                    Set mfrmMedRecEdit = New frmPageMedRecEdit
                Case ST_�Ĵ�ʡ��׼
                    Set mfrmMedRecEdit = New frmPageMedRecEdit_SC
                Case ST_����ʡ��׼
                    Set mfrmMedRecEdit = New frmPageMedRecEdit_YN
                Case ST_����ʡ��׼
                    Set mfrmMedRecEdit = New frmPageMedRecEdit_HN
            End Select
        ElseIf gclsPros.FuncType = fҽ����ҳ Then
            Select Case gclsPros.MedPageSandard
                Case ST_��������׼
                    Set mfrmMedRecEdit = New frmInMedRecEdit
                Case ST_�Ĵ�ʡ��׼
                    Set mfrmMedRecEdit = New frmInMedRecEdit_SC
                Case ST_����ʡ��׼
                    Set mfrmMedRecEdit = New frmInMedRecEdit_YN
                Case ST_����ʡ��׼
                    Set mfrmMedRecEdit = New frmInMedRecEdit_HN
            End Select
        End If
    Set gclsPros.CurrentForm = mfrmMedRecEdit
    If blnModal Then
        Me.Show 1, gclsPros.MainForm
    Else
        Me.Show , gclsPros.MainForm
    End If
    ShowMe = True
End Function

Private Sub InitDkpMain()
'��ʼ���沼��
    Dim PaneLeft As Pane
    Dim PaneMain As Pane
    Dim PaneBottom As Pane
    On Error GoTo Errhand
    With Me.dkpMain
        .SetCommandBars cbsMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    dkpMain.DestroyAll
    Set PaneLeft = dkpMain.CreatePane(Pane_����, 25, 250, DockLeftOf, Nothing)
    PaneLeft.Options = PaneNoCloseable + PaneNoFloatable
    PaneLeft.MaxTrackSize.Width = 300
    PaneLeft.Title = "����"
    Set PaneMain = dkpMain.CreatePane(Pane_��ҳ, 250, 250, DockRightOf, Nothing)
    PaneMain.Options = PaneNoCloseable + PaneNoCaption + PaneNoHideable + PaneNoFloatable
    Set PaneBottom = dkpMain.CreatePane(Pane_���, 80, 80, DockBottomOf, PaneMain)
    PaneBottom.Options = PaneNoFloatable + PaneNoHideable
    PaneBottom.Title = "�����Ϣ"
    PaneBottom.Closed = True
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim strTmp As String, arrTmp As Variant
    Dim objControl As CommandBarControl
    Dim i As Long
    
    If CommandBar Is Nothing Then Exit Sub
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
    Case conMenu_Tool_PlugIn
        Call CreatePlugInOK(gclsPros.Module)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFuncNames(gclsPros.SysNo, gclsPros.Module, 6)
            Call zlPlugInErrH(Err, "GetFuncNames")
            Err.Clear: On Error GoTo 0
        End If
        If strTmp <> "" Then
            With CommandBar.Controls
                If .Count = 0 Then
                    strTmp = Replace(strTmp, "Auto:", "")
                    arrTmp = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp)
                        Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = arrTmp(i)
                    Next
                End If
            End With
        End If
    End Select
End Sub



Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionClosed Then
        If Pane.ID = Pane_��� Then
            Call VsErrClick("")
        End If
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = Pane_���� Then
        Item.Handle = PicDirectory.hwnd
    ElseIf Item.ID = Pane_��ҳ Then
        Item.Handle = PicForm.hwnd
    ElseIf Item.ID = Pane_��� Then
        Item.Handle = PicErr.hwnd
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strMsg As String
'    Call gclsPros.CurrentForm.ValidateControls
    Call ShowInfectInfo(False)
    Me.stbThis.Panels(2).Text = ""
    Select Case Control.ID
        Case conMenu_Manage_Preview            'Ԥ��
            Call menuPageOperate(MOP_Ԥ��)
        Case conMenu_Manage_Preview * 10# + 1  'Ԥ��
            Call menuPageOperate(MOP_Ԥ��, 1)
        Case conMenu_Manage_Preview * 10# + 2  'Ԥ��
            Call menuPageOperate(MOP_Ԥ��, 2)
        Case conMenu_Manage_Preview * 10# + 3  'Ԥ��
            Call menuPageOperate(MOP_Ԥ��, 3)
        Case conMenu_Manage_Preview * 10# + 4  'Ԥ��
            Call menuPageOperate(MOP_Ԥ��, 4)
        Case conMenu_Manage_Print             '��ӡ
            Call menuPageOperate(MOP_��ӡ)
        Case conMenu_Manage_Print * 10# + 1   '��ӡ
            Call menuPageOperate(MOP_��ӡ, 1)
        Case conMenu_Manage_Print * 10# + 2   '��ӡ
            Call menuPageOperate(MOP_��ӡ, 2)
        Case conMenu_Manage_Print * 10# + 3   '��ӡ
            Call menuPageOperate(MOP_��ӡ, 3)
        Case conMenu_Manage_Print * 10# + 4   '��ӡ
            Call menuPageOperate(MOP_��ӡ, 4)
        Case conMenu_Manage_Print * 10# + 5   '��ӡ
            Call menuPageOperate(MOP_��ӡ, 5)
        Case conMenu_Manage_Print * 10# + 6   '��ӡ
            Call menuPageOperate(MOP_��ӡ, 6)
        Case conMenu_Manage_Modify  '�޸Ĳ��˻�����Ϣ
            Call ModifyPatiInfo
        Case conMenu_Manage_Audit   '���
            If gclsPros.FuncType = fҽ����ҳ And gclsPros.PatiInfo!�������� = 1 Then
                Call Check����
            Else
                Call CheckMedPageData(True, True)
            End If
            If gColErr.Count > 0 Or gColWarn.Count > 0 Then
                strMsg = "�����ϣ�����" & CStr(gColErr.Count) & "������" & CStr(gColWarn.Count) & "�����棡"
            Else
                strMsg = "�����ϣ�û�з��ִ���"
            End If
            If gColErr.Count > 0 Or gColWarn.Count > 0 Then
                strMsg = "�����ϣ�����" & CStr(gColErr.Count) & "������" & CStr(gColWarn.Count) & "�����棡"
            Else
                strMsg = "�����ϣ�û�з��ִ���"
            End If
            Me.stbThis.Panels(2).Text = strMsg
        Case conMenu_Manage_Save    '����
            Call menuPageOperate(MOP_ȷ��)
        Case conMenu_Manage_Exit    '�˳�
            Unload Me
        Case conMenu_Manage_Up      '��һ��
            Call CmdUPClick
        Case conMenu_Manage_Down    '��һ��
            Call CmdDownClick
        Case conMenu_Manage_Help
            Call CmdHelpClick
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '�����ҹ���ִ��
            Call ExeDiagPlugIn(Control.Parameter)
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim blnChange As Boolean
    
    If Not Me.Visible Then Exit Sub
    Err = 0: On Error Resume Next
    
    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            blnChange = blnChange Or gfrmMecCol(i).gblnchange
        Next
    End If
    
    If gclsPros.FuncType = fҽ����ҳ Then
        Select Case Control.ID
            Case conMenu_Manage_Modify  '�޸Ĳ��˻�����Ϣ
                Control.Visible = mblnModifyVisible
            Case conMenu_Manage_Save
                Control.Visible = True
                Control.Enabled = gclsPros.InfosChange
                If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                    Control.Enabled = blnChange Or gclsPros.InfosChange
                End If
            Case conMenu_Manage_Audit
                Control.Visible = True
                Control.Enabled = gclsPros.InfosChange
                If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                    Control.Enabled = blnChange Or gclsPros.InfosChange
                End If
        End Select
    ElseIf gclsPros.FuncType = f������ҳ Then
        If gclsPros.OpenMode = EM_���� Then
            Select Case Control.ID
                Case conMenu_Manage_Audit
                    Control.Visible = False
                Case conMenu_Manage_Save
                    Control.Visible = False
            End Select
        Else
            Select Case Control.ID
                Case conMenu_Manage_Audit
                    Control.Visible = True
                    Control.Enabled = gclsPros.InfosChange
                    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                        Control.Enabled = blnChange Or gclsPros.InfosChange
                    End If
                Case conMenu_Manage_Save
                    Control.Visible = True
                    Control.Enabled = gclsPros.InfosChange
                    If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                        Control.Enabled = blnChange Or gclsPros.InfosChange
                    End If
            End Select
        End If
                
        If gclsPros.OpenMode = EM_���� Or gclsPros.OpenMode = EM_�༭ Then
            Select Case Control.ID
                Case conMenu_Manage_Down
                    Control.Visible = True
                Case conMenu_Manage_Up
                    Control.Visible = True
            End Select
        Else
            Select Case Control.ID
                Case conMenu_Manage_Down
                    Control.Visible = False
                Case conMenu_Manage_Up
                    Control.Visible = False
            End Select
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Errhand
    subcMain.hwnd = Me.hwnd
    subcMain.Messages(WM_MOUSEWHEEL) = True
    '��ʼ���˵�����ʼ���˵���λ��ֻ�ܹ��������
    Call InitCommandBar
    
    '���Ӵ��嵽PicForm����
    SetParent mfrmMedRecEdit.hwnd, PicForm.hwnd
    
    If Not gclsPros.LoadFinish Then
        mfrmMedRecEdit.picMain.Visible = False
    End If
    
    Call InitDkpMain
    
    mblnModifyVisible = False
    If gclsPros.FuncType = fҽ����ҳ Then
        Me.Caption = "סԺ��ҳ�༭"
        mblnModifyVisible = InStr(GetInsidePrivs(p������Ϣ��������), "������Ϣ����") > 0
    ElseIf gclsPros.FuncType = f������ҳ Then
        If gclsPros.OpenMode = EM_���� Then
            Me.Caption = "������ҳ����"
        Else
            Me.Caption = "������ҳ�༭"
        End If
    End If
    gblnUnload = True
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = 2   '������ҳռ����Ļ�Ͽ���˴������
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If Not gclsPros.LoadFinish Then Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmMedRecEdit Is Nothing Then
        Call FormUnLoad(Cancel)
        If Cancel Then Exit Sub
        Set mfrmMedRecEdit = Nothing
    End If
    Me.subcMain.Messages(WM_MOUSEWHEEL) = False
    Call SaveWinState(Me, App.ProductName)
    gblnUnload = False
End Sub

Private Sub InitCommandBar()
'���ܣ������ڹ��������岿��
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    Dim lngIdx As Long
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    
    '�˵�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "������", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        '�����չ����
        Call CreatePlugInOK(gclsPros.Module)
        If Not gobjPlugIn Is Nothing Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����")
            objPopup.BeginGroup = True
        End If
    End With
    

    '����������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False                   '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False                   '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagHideWrap                '��������Ȳ���ʱҲ������
    With objBar.Controls
        If gclsPros.FuncType = fҽ����ҳ Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Manage_Preview, "Ԥ��", lngIdx + 1): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Select Case gclsPros.MedPageSandard
                    Case ST_��������׼, ST_����ʡ��׼
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Preview * 10# + 1, "����")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Preview * 10# + 2, "����")
                    Case ST_�Ĵ�ʡ��׼, ST_����ʡ��׼
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Preview * 10# + 1, "����")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Preview * 10# + 2, "����")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Preview * 10# + 3, "��ҳ1")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Preview * 10# + 4, "��ҳ2")
                End Select
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
            
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Manage_Print, "��ӡ", lngIdx + 1): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Select Case gclsPros.MedPageSandard
                    Case ST_��������׼, ST_����ʡ��׼
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 1, "����")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 2, "����")
                    Case ST_�Ĵ�ʡ��׼, ST_����ʡ��׼
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 1, "����")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 2, "����")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 3, "��ҳ1")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 4, "��ҳ2")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 5, "����+��ҳ1")
                        Set objControl = .Add(xtpControlButton, conMenu_Manage_Print * 10# + 6, "����+��ҳ2")
                End Select
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Modify, "�޸Ĳ��˻�����Ϣ")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Audit, "���")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Save, "����")
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Exit, "�˳�")
            objControl.Style = xtpButtonIconAndCaption
        ElseIf gclsPros.FuncType = f������ҳ Then
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Up, "��һ��")
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Down, "��һ��")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Audit, "���")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Save, "����")
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Exit, "�˳�")
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Help, "����")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
        End If
    End With
    
     With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_Manage_Save
        .Add FCONTROL, vbKeyM, conMenu_Manage_Modify
        .Add FCONTROL, vbKeyP, conMenu_Manage_Print       '��ӡ
        .Add FCONTROL, vbKeyE, conMenu_Manage_Exit
        .Add 0, vbKeyF2, conMenu_Manage_Audit
        .Add 0, vbKeyF1, conMenu_Manage_Help
    End With
End Sub

Private Sub PicDirectory_Resize()
    On Error Resume Next
    tvDirectory.Move PicDirectory.ScaleLeft, PicDirectory.ScaleTop, PicDirectory.ScaleWidth, PicDirectory.ScaleHeight
End Sub

Private Sub PicErr_Resize()
    On Error Resume Next
    vsErr.Move PicErr.ScaleLeft, PicErr.ScaleTop, PicErr.ScaleWidth, PicErr.ScaleHeight - 375
End Sub

Private Sub PicForm_Resize()
 On Error Resume Next
    Dim vRect As RECT
    Dim x As Long, Y As Long
    GetWindowRect PicForm.hwnd, vRect
    x = vRect.Right - vRect.Left
    Y = vRect.Bottom - vRect.Top
    If x < 0 Then x = 0
    If Y < 0 Then Y = 0
    SetWindowPos mfrmMedRecEdit.hwnd, 0, 0, 0, x, Y, &H40 Or &H20
    ShowWindow mfrmMedRecEdit.hwnd, SW_RESTORE
End Sub

Private Sub subcMain_WndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    Call SubCMainWndProc(msg, wParam, lParam, Result)
End Sub

Private Sub tvDirectory_GotFocus()
    Call ShowInfectInfo(False)
End Sub

Private Sub tvDirectory_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strKEY As String, intIndex As Integer
    strKEY = Node.Key
    If strKEY Like "key-*" Then
        intIndex = CInt(Val(Mid(strKEY, InStr(strKEY, "-") + 1)))
        Call ChangePage(True, intIndex)
    End If
End Sub

Private Sub vsErr_Click()
    Dim strErrID As String
    On Error Resume Next
    If vsErr.MouseRow > vsErr.FixedRows And vsErr.MouseRow < vsErr.Rows Then
        strErrID = NVL(vsErr.Cell(flexcpData, vsErr.Row, ERR_ID))
        Call VsErrClick(strErrID)
    Else
        Call VsErrClick("")
    End If
End Sub
