VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGradeManage 
   Caption         =   "���Ӳ�������"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   12240
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   5280
      ScaleHeight     =   3585
      ScaleWidth      =   4470
      TabIndex        =   3
      Top             =   1620
      Width           =   4470
      Begin XtremeSuiteControls.TabControl tbcTask 
         Height          =   1830
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   9735
      TabIndex        =   2
      Top             =   105
      Width           =   1125
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   450
      ScaleHeight     =   3855
      ScaleWidth      =   3690
      TabIndex        =   0
      Top             =   1320
      Width           =   3690
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2385
         Picture         =   "frmGradeManage.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2520
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   375
         Width           =   2670
         _cx             =   4710
         _cy             =   4445
         Appearance      =   1
         BorderStyle     =   1
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.TextBox txt���� 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   450
         TabIndex        =   8
         Top             =   30
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   30
         TabIndex        =   6
         Top             =   75
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7620
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15743
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmGradeManage.frx":0049
      Left            =   630
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmGradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''���弶��������
''######################################################################################################################
'Private mstrPrivs As String
'Private mblnStartUp As Boolean
'Private mblnAllowClose As Boolean
'Private mstrCondition As String
'Private mstrFindKey As String
'Private mlngTmp As Long
'Private mobjFindKey As CommandBarControl
'Private mclsVsf(0) As clsVsf
'Private mlngModul As Long
'Private mintIndex As Integer
'Private mbytMode As Byte
'Private mfrmChildMedrec As frmChildMedrec
'Private WithEvents mfrmGradeEdit As frmGradeEdit
'
''######################################################################################################################
'
'Public Property Get ģ���() As Long
'    ģ��� = mlngModul
'End Property
'
'Private Function InitCommandBar() As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim objMenu As CommandBarPopup
'    Dim objBar As CommandBar
'    Dim objPopup As CommandBarPopup
'    Dim objControl As CommandBarControl
'    Dim cbrCustom As CommandBarControlCustom
'
'    '------------------------------------------------------------------------------------------------------------------
'    '��ʼ����
'
'    Call CommandBarInit(cbsMain)
'
'    '------------------------------------------------------------------------------------------------------------------
'    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
'
'    cbsMain.ActiveMenuBar.Title = "�˵�"
'    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
'
'    '�ļ�
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
'    objMenu.ID = conMenu_FilePopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "ȫ����ӡ(&L)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
'
'    '�༭
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
'    objMenu.ID = conMenu_EditPopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "��������(&A)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸Ľ��(&M)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Append, "��������(&R)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ�����(&D)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "ͨ�����(&P)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "ȡ�����(&C)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "ȫ��ѡ��(&L)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "ȡ��ѡ��(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "����ѡ��(&B)")
'
'    '�鿴
'    '------------------------------------------------------------------------------------------------------------------
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
'    objMenu.ID = conMenu_ViewPopup
'    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
'    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "��������(&F)", True)
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
'
'
'    '����
'    '------------------------------------------------------------------------------------------------------------------
'    Call CreateHelpMenu(cbsMain)
'
'    '���˵��Ҳ�Ĳ���
'    '------------------------------------------------------------------------------------------------------------------
'    cbsMain.ActiveMenuBar.SetIconSize 16, 16
'    mstrFindKey = GetPara("��λ����", mlngModul, True, "No")
'    If mstrFindKey = "" Then mstrFindKey = "No"
'    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
'    mobjFindKey.IconId = conMenu_View_Find
'    mobjFindKey.Flags = xtpFlagRightAlign
'    mobjFindKey.STYLE = xtpButtonIconAndCaption
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.����", , , "����")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.סԺ��", , , "סԺ��")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.����", , , "����")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.���￨��", , , "���￨��")
'    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, ""): cbrCustom.Handle = txtLocation.Hwnd: cbrCustom.Flags = xtpFlagRightAlign
'    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
'    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
'
'    '����������:������������
'    '------------------------------------------------------------------------------------------------------------------
'    Set objBar = cbsMain.Add("��׼", xtpBarTop)
'    objBar.ContextMenuPresent = False
'    objBar.ShowTextBelowIcons = False
'    objBar.EnableDocking xtpFlagStretched
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "���")
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "����", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
'    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
'
'    '����Ŀ����:���������������Ѵ���
'    '------------------------------------------------------------------------------------------------------------------
'    With cbsMain.KeyBindings
'        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
'        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
'        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
'        .Add FCONTROL, vbKeyV, conMenu_File_Preview
'        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
'        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify          '�޸�
'        .Add FCONTROL, vbKeyR, conMenu_Edit_Append          '����
'        .Add FCONTROL, vbKeyF, conMenu_View_Filter          '����
'        .Add 0, vbKeyDelete, conMenu_Edit_Delete            'ɾ��
'        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
'        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
'        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
'    End With
'
'End Function
'
'Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intRow As Integer
'    Dim rs As New ADODB.Recordset
'    Dim rsSQL As New ADODB.Recordset
'    Dim strTmp As String
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    Call SQLRecord(rsSQL)
'
'    Select Case strCommand
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ʼ�ؼ�"
'
'        Set mclsVsf(0) = New clsVsf
'        With mclsVsf(0)
'            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
'            Call .ClearColumn
'            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
'            Call .AppendColumn("No", 900, flexAlignLeftCenter, flexDTString, "", , True)
'            Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, "", , True)
'            Call .AppendColumn("��¼״̬", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
'            Call .AppendColumn("����ʱ��", 1440, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", , True)
'            Call .AppendColumn("��������", 1440, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", , True)
'            Call .AppendColumn("��������", 1500, flexAlignLeftCenter, flexDTString, "", , True)
'            .AppendRows = True
'        End With
'
'        '��ʼ�˵���������
'        '--------------------------------------------------------------------------------------------------------------
'        Call InitCommandBar
'
'        '����ͣ������
'        '--------------------------------------------------------------------------------------------------------------
'        Dim objPane As Pane
'        Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "������Ϣ": objPane.Options = PaneNoCaption
'        Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing): objPane.Title = "��ϸ����": objPane.Options = PaneNoCaption
'
'
'        dkpMain.SetCommandBars cbsMain
'        Call DockPannelInit(dkpMain)
'
'
'        Call TabControlInit(tbcTask)
'        With tbcTask
'            .PaintManager.BoldSelected = True
'
'            Set mfrmGradeEdit = New frmGradeEdit
'            Set mfrmChildMedrec = New frmChildMedrec
'
'            Call mfrmGradeEdit.InitData(Me, True)
'
'            .InsertItem 0, "��������", mfrmGradeEdit.Hwnd, 0
'            .InsertItem 1, "��ҳ��¼", mfrmChildMedrec.Hwnd, 0
'
'            .Item(0).Selected = True
'
'        End With
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ʼ����"
'
''        '��������������Ŀ�������г�ʼ��
''        Call ParamCreate(mrsCondition)
''        Call ParamAdd(mrsCondition, "��ʼ���ݺ�", "")
''        Call ParamAdd(mrsCondition, "�������ݺ�", "")
''        Call ParamAdd(mrsCondition, "������", "")
''        Call ParamAdd(mrsCondition, "��׼��", "")
''        Call ParamAdd(mrsCondition, "�ܾ���", "")
''
''        Call ParamAdd(mrsCondition, "�µǼǵ���", "1")
''        Call ParamAdd(mrsCondition, "�Ǽǿ�ʼ����", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "�Ǽǽ�������", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "����׼����", "0")
''        Call ParamAdd(mrsCondition, "��׼��ʼ����", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "��׼��������", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "�Ѿܾ�����", "0")
''        Call ParamAdd(mrsCondition, "�ܾ���ʼ����", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
''        Call ParamAdd(mrsCondition, "�ܾ���������", Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
''
''        '��ȡȱʡ�Ľ�������Ǽǲ�ѯʱ�䷶Χ
''        strTmp = GetPara("�Ǽ�ȱʡ��Χ", mlngModul, True, "��  ��")
''        If strTmp = "" Then strTmp = "��  ��"
''        Call ParamWrite(mrsCondition, "�Ǽǿ�ʼ����", GetDateTime(strTmp, 1))
''        Call ParamWrite(mrsCondition, "�Ǽǽ�������", GetDateTime(strTmp, 2))
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "�ؼ�״̬"
'
''        If vsf(0).Enabled <> Not DataChanged Then
''            vsf(0).Enabled = Not DataChanged
''            vsf(0).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
''            tbcTask.Enabled = Not DataChanged
''        End If
''        stbThis.Panels(3).Enabled = DataChanged
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "ˢ��״̬"
'
''        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then
''            strTmp = "��ǰ��û���κε��Ӳ����������뵥��"
''        Else
''            strTmp = "���� " & vsf(0).Rows - 1 & " �����Ӳ����������뵥��"
''        End If
''
''        stbThis.Panels(2).Text = strTmp
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "ˢ������"
'
'        Call ExecuteCommand("��ȡ���˼�¼")
'        Call ExecuteCommand("��ȡ��������")
'        Call ExecuteCommand("��ȡ��ҳ��¼")
'
'        Call ExecuteCommand("ˢ��״̬")
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��������"
'
''        mrsCondition.Filter = ""
''        ExecuteCommand = frmCISBorrowFilter.ShowPara(Me, mrsCondition)
''
''        GoTo endHand
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ȡ���뵥��"
'
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ȡ��������"
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "ǰһ��"
'        With vsf(0)
'            If .Row > 1 Then
'                .Row = .Row - 1
'                .ShowCell .Row, .Col
'            End If
'        End With
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��һ��"
'        With vsf(0)
'            If .Row < .Rows - 1 Then
'                .Row = .Row + 1
'                .ShowCell .Row, .Col
'            End If
'        End With
'    '------------------------------------------------------------------------------------------------------------------
'    Case "��ע���"
'
'        If Val(GetPara("ʹ�ø��Ի����", , , True)) = 1 Then
'            'ʹ�ø��Ի�����
'
'            mstrFindKey = Trim(GetPara("��λ����", mlngModul, True, "No"))
'            mclsVsf(0).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)), ""))
'        End If
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case "дע���"
'        If Val(GetPara("ʹ�ø��Ի����", , , True)) = 1 Then
'            'ʹ�ø��Ի�����
'            Call SetPara("��λ����", mstrFindKey, mlngModul, True)
'        End If
'        Call SetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)), mclsVsf(0).SaveStateToString)
'    End Select
'
'    ExecuteCommand = True
'
'    GoTo endHand
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'
'    '------------------------------------------------------------------------------------------------------------------
'endHand:
'
'
'End Function
'
''######################################################################################################################
'
'Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Dim objControl As CommandBarControl
'    Dim lngLoop As Long
'
'    Select Case Control.ID
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_File_Parameter
'        Call frmCISBorrowPara.ShowEdit(Me, mstrPrivs)
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_NewItem
'
'        Call ExecuteCommand("���ӽ�������")
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Delete                'ɾ����������
'
'        If ExecuteCommand("ɾ����������") Then
'            Call ExecuteCommand("�Ƴ���������")
'        End If
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Audit                '��׼��������
'
'        Call ExecuteCommand("��׼��������")
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case conMenu_Manage_Refuse                '�ܾ���������
'
'        Call ExecuteCommand("�ܾ���������")
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_Edit_Transf_Cancle                  '�ָ�����
'
'        Call ExecuteCommand("�ָ�����")
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Filter '����
'
'        If ExecuteCommand("��������") Then
'            Call ExecuteCommand("ˢ������")
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Forward
'        Call ExecuteCommand("ǰһ��")
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Backward
'        Call ExecuteCommand("��һ��")
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Option
'        mobjFindKey.Execute
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_LocationItem
'
'        mstrFindKey = Control.Parameter
'        mobjFindKey.Caption = mstrFindKey
'        cbsMain.RecalcLayout
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case conMenu_View_Location
'
'        LocationObj txtLocation
'
'    '------------------------------------------------------------------------------------------------------------------
'    Case Else
'
'        If Control.ID > 400 And Control.ID < 500 Then
'            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "ID=" & Val(vsf(0).RowData(vsf(0).Row)))
'        Else
'             '��ҵ���޹صĹ��ܣ������Ĺ���
'            Call CommandBarExecutePublic(Control, Me, vsf(0), "���Ӳ����������뵥")
'        End If
'
'    End Select
'End Sub
'
'Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    If stbThis.Visible Then Bottom = stbThis.Height
'End Sub
'
'Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    On Error GoTo errHand
'
'    With vsf(0)
'        Select Case Control.ID
'        '--------------------------------------------------------------------------------------------------------------
'        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
'            Control.Enabled = (Val(.RowData(.Row)) > 0)
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_File_Parameter, conMenu_View_Filter, conMenu_View_Refresh
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_EditPopup
''            Control.Visible = (tbcTask.Selected.Index = 0)
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_NewItem
''            Control.Visible = IsPrivs(mstrPrivs, "�Ǽ�����") And tbcTask.Selected.Index = 0
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Delete
''            Control.Visible = IsPrivs(mstrPrivs, "�Ǽ�����") And tbcTask.Selected.Index = 0
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Audit
''
''            Control.Visible = IsPrivs(mstrPrivs, "��������") And tbcTask.Selected.Index = 0
''
'
'
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Manage_Refuse
''            Control.Visible = IsPrivs(mstrPrivs, "��������") And tbcTask.Selected.Index = 0
''            With vsf(0)
''                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1
''            End With
''
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Untread
''            Control.Visible = IsPrivs(mstrPrivs, "��������") And tbcTask.Selected.Index = 0
''            With vsf(0)
''                Control.Visible = Control.Visible And (Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) > 1)
''                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) > 1
''                Control.Caption = IIf(Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 2, "������׼(&B)", "���˾ܾ�(&B)")
''            End With
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
''            Control.Visible = IsPrivs(mstrPrivs, "�Ǽ�����") And tbcTask.Selected.Index = 0
''            Control.Enabled = Control.Visible And DataChanged = True
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_Forward
''
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (.Row > 1 And DataChanged = False)
''            Case 1
''                Control.Enabled = (mfrmChildPatientView.VsfBody.Row > 1)
''
''            End Select
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_Backward
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)
''            Case 1
''                Control.Enabled = (mfrmChildPatientView.VsfBody.Row < mfrmChildPatientView.VsfBody.Rows - 1)
''            End Select
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_LocationItem        '
''            Control.Checked = (mstrFindKey = Control.Parameter)
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (DataChanged = False)
''            Case 1
''
''            End Select
''        '--------------------------------------------------------------------------------------------------------------
''        Case conMenu_View_Location, conMenu_View_Column
''            Select Case tbcTask.Selected.Index
''            Case 0
''                Control.Enabled = (DataChanged = False)
''            Case 1
''
''            End Select
'        '--------------------------------------------------------------------------------------------------------------
'        Case Else
'            Call CommandBarUpdatePublic(Control, Me)
'        End Select
'    End With
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'End Sub
'
'Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'
'    Select Case Item.ID
'    Case 1
'        Item.Handle = picPane(1).Hwnd
'    Case 2
'        Item.Handle = picPane(0).Hwnd
'    End Select
'
'End Sub
'
'Private Sub Form_Activate()
'    If mblnStartUp = False Then Exit Sub
'    mblnStartUp = False
'    DoEvents
'
'    If ExecuteCommand("��ʼ����") = False Then GoTo errHand
'
'    Call ExecuteCommand("ˢ������")
'
'    mblnAllowClose = True
'    Exit Sub
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mblnAllowClose = True
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    mblnStartUp = True
'    mblnAllowClose = False
'
'    picPane(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
'
'
'    mstrPrivs = UserInfo.ģ��Ȩ��
'    mlngModul = ParamInfo.ģ���
'
'    Call ExecuteCommand("��ʼ�ؼ�")
'    Call ExecuteCommand("��ע���")
'
'    Call RestoreWinState(Me, App.ProductName)
'    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, gblnShowInTaskBar)
'
'End Sub
'
'Private Sub Form_Resize()
'    On Error Resume Next
'
'    Call SetPaneRange(dkpMain, 1, 100, 100, 300, Me.ScaleHeight)
'    dkpMain.RecalcLayout
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'
'    Cancel = Not mblnAllowClose
'
'    If Cancel = False Then
'
'        If Cancel = False Then
'
'            Call ExecuteCommand("дע���")
'
'            Call SaveWinState(Me, App.ProductName)
'
'            Set mclsVsf(0) = Nothing
'
'            On Error Resume Next
'
'            Unload mfrmGradeEdit
'            Unload mfrmChildMedrec
'
'        End If
'    End If
'
'End Sub
'
''�Զ�����̻���
''######################################################################################################################
'
'Private Sub picPane_Resize(Index As Integer)
'    On Error Resume Next
'
'    Select Case Index
'    Case 0
'        tbcTask.Move 0, 0, picPane(Index).Width, picPane(Index).Height
'    Case 1
'        txt����.Move txt����.Left, txt����.Top, picPane(Index).Width - txt����.Left - 30
'        cmdSelect.Move txt����.Left + txt����.Width - cmdSelect.Width - 30, txt����.Top + 30
'        vsf(0).Move 0, vsf(0).Top, picPane(Index).Width, picPane(Index).Height - vsf(0).Top
'        mclsVsf(0).AppendRows = True
'    End Select
'End Sub
'
'Private Sub tbcTask_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    mintIndex = Item.Index
'End Sub
'
'Private Sub txtLocation_GotFocus()
'    Call zlControl.TxtSelAll(txtLocation)
'End Sub
'
'Private Sub txtLocation_KeyPress(KeyAscii As Integer)
'    Dim lngRow As Long
'    Dim intCol As Integer
'    Dim bytMatch As Byte
'
'    If KeyAscii = vbKeyReturn Then
'        lngRow = -1
'        bytMatch = 2
'
'        intCol = mclsVsf(0).ColIndex(mstrFindKey)
'        If intCol >= 0 Then
'            lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch, vsf(0).Row + 1)
'            If lngRow = -1 Then
'                lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
'            End If
'            If lngRow > 0 And vsf(0).Row <> lngRow Then
'                vsf(0).Row = lngRow
'                vsf(0).ShowCell vsf(0).Row, vsf(0).Col
'            End If
'        End If
'
'        Call LocationObj(txtLocation)
'    End If
'End Sub
'
'Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
'    mclsVsf(Index).AppendRows = True
'End Sub
'
'Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf(Index).AppendRows = True
'End Sub
'
'Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    mclsVsf(Index).AppendRows = True
'End Sub
'
'Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim cbrPopupBar As CommandBar
'
'    If Button = 2 And Index = 0 Then
'        Call SendLMouseButton(vsf(Index).Hwnd, x, y)
'
'        Set cbrPopupBar = CopyMenu(cbsMain, 2)
'        If cbrPopupBar Is Nothing Then Exit Sub
'
'        cbrPopupBar.ShowPopup
'    End If
'
'End Sub
'
'
