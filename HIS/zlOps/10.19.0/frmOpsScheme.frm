VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmOpsScheme 
   Caption         =   "������������"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11295
   Icon            =   "frmOpsScheme.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   6555
      TabIndex        =   5
      ToolTipText     =   "��ݼ���F3"
      Top             =   1095
      Width           =   1320
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   2970
      TabIndex        =   2
      Top             =   1200
      Width           =   2970
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2145
         Index           =   0
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Width           =   2520
         _cx             =   4445
         _cy             =   3784
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   2
      Left            =   4425
      ScaleHeight     =   2865
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   2370
      Width           =   5265
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2025
         Left            =   690
         TabIndex        =   1
         Top             =   240
         Width           =   2700
         _Version        =   589884
         _ExtentX        =   4762
         _ExtentY        =   3572
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6690
      Width           =   11295
      _ExtentX        =   19923
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
            Object.Width           =   14076
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
      Bindings        =   "frmOpsScheme.frx":6852
      Left            =   1170
      Top             =   195
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmOpsScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������
'######################################################################################################################

'��������


'��������
Private mstrPrivs As String
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mclsVsf(0) As New clsVsf
Private mlngTmp As Long
Private mobjFindKey As CommandBarPopup
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean
Private mlngģ��� As Long
Private WithEvents mfrmChildSchemeEdit As frmChildSchemeEdit
Attribute mfrmChildSchemeEdit.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeDrug As frmChildSchemeDrug
Attribute mfrmChildSchemeDrug.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeCharge As frmChildSchemeCharge
Attribute mfrmChildSchemeCharge.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeMaterial As frmChildSchemeMaterial
Attribute mfrmChildSchemeMaterial.VB_VarHelpID = -1
Private WithEvents mfrmChildSchemeOps As frmChildSchemeOps
Attribute mfrmChildSchemeOps.VB_VarHelpID = -1

'######################################################################################################################

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildSchemeEdit.DataChanged = blnData
    mfrmChildSchemeDrug.DataChanged = blnData
    mfrmChildSchemeCharge.DataChanged = blnData
    mfrmChildSchemeMaterial.DataChanged = blnData
    mfrmChildSchemeOps.DataChanged = blnData

    If mfrmChildSchemeEdit.DataChanged Or mfrmChildSchemeDrug.DataChanged Or mfrmChildSchemeCharge.DataChanged Or mfrmChildSchemeMaterial.DataChanged Or mfrmChildSchemeOps.DataChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmChildSchemeEdit Is Nothing) And Not (mfrmChildSchemeDrug Is Nothing) And Not (mfrmChildSchemeCharge Is Nothing) And Not (mfrmChildSchemeMaterial Is Nothing) And Not (mfrmChildSchemeCharge Is Nothing) And Not (mfrmChildSchemeOps Is Nothing) Then
        DataChanged = mfrmChildSchemeEdit.DataChanged Or mfrmChildSchemeDrug.DataChanged Or mfrmChildSchemeCharge.DataChanged Or mfrmChildSchemeMaterial.DataChanged Or mfrmChildSchemeCharge.DataChanged Or mfrmChildSchemeOps.DataChanged
    End If
End Property

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)

    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "���ӷ���(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "��������(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "�������(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������(&R)")
    
    
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Jump, "������ת(&J)")

    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & ParamInfo.��Ʒ����)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.��Ʒ���� & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.��Ʒ���� & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)
    
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = Trim(GetRegister(˽��ģ��, Me.Name, "��λ����", "����"))
    If mstrFindKey = "" Then mstrFindKey = "����"

    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.����", , , "����")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.����", , , "����")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.����", , , "����")
    
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings
        .Add 0, vbKeyF6, conMenu_View_Jump                  '��ת
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save                  '����
        
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save            '����
        
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add 0, vbKeyF4, conMenu_View_Option                'ѡ��λ����
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
        
    End With
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 200, 100, DockLeftOf, Nothing)
    objPane.Title = "���������б�"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 350, 300, DockRightOf, Nothing)
    objPane.Title = "��ϸ����"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 350, 150, DockBottomOf, objPane)
    objPane.Title = "��������"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)

End Sub

Private Function InitTabControl() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With

        Set .Icons = frmPubIcons.imgPublic.Icons


        Set mfrmChildSchemeDrug = New frmChildSchemeDrug
        Call mfrmChildSchemeDrug.InitData(Me, IsPrivs(mstrPrivs, "��ɾ��"))

        Set mfrmChildSchemeMaterial = New frmChildSchemeMaterial
        Call mfrmChildSchemeMaterial.InitData(Me, IsPrivs(mstrPrivs, "��ɾ��"))
        
        Set mfrmChildSchemeOps = New frmChildSchemeOps
        Call mfrmChildSchemeOps.InitData(Me, IsPrivs(mstrPrivs, "��ɾ��"))
        
        Set mfrmChildSchemeCharge = New frmChildSchemeCharge
        Call mfrmChildSchemeCharge.InitData(Me, IsPrivs(mstrPrivs, "��ɾ��"))

        .InsertItem 0, "��ҩ����", mfrmChildSchemeDrug.hWnd, 0
        .InsertItem 1, "���Ϸ���", mfrmChildSchemeMaterial.hWnd, 0
        .InsertItem 2, "���Ʒ���", mfrmChildSchemeCharge.hWnd, 0
        .InsertItem 3, "��������", mfrmChildSchemeOps.hWnd, 0

        .Item(0).Selected = True
        
    End With
    
    InitTabControl = True
    
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        Set mclsVsf(0) = New clsVsf
        With mclsVsf(0)
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("����", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 750, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("˵��", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With

        
        '��ʼ�˵���������
        Call InitCommandBar
        
        '��ʼ����ָ�����
        Call InitDockPannel
        Call InitTabControl
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
        
        If vsf(0).Enabled <> Not DataChanged Then
            vsf(0).Enabled = Not DataChanged
            vsf(0).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
        End If
        stbThis.Panels(3).Enabled = DataChanged

    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
    
        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then
            strTmp = "��ǰ��û�ж�������������"
        Else
            strTmp = "�������� " & vsf(0).Rows - 1 & " ������������"
        End If

        stbThis.Panels(2).Text = strTmp
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
    
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("ˢ�·�������")
        Call ExecuteCommand("ˢ��״̬")
            
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��ָ������"

        strSQL = "SELECT '����' As ͼ��,A.ID,A.����,A.����,A.����,A.˵�� FROM ���������ο� A Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngTmp)
        If rs.BOF = True Then Exit Function
                
        intRow = mclsVsf(0).FindRow(mlngTmp, -1)
        If intRow > 0 Then
            '�Ѽ���
            vsf(0).Row = intRow
        Else
            'δ����
            If Val(vsf(0).RowData(vsf(0).Rows - 1)) > 0 Then vsf(0).Rows = vsf(0).Rows + 1
            vsf(0).Row = vsf(0).Rows - 1
        End If
        
        Call mclsVsf(0).LoadGridRow(vsf(0).Row, rs)
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("ˢ�·�������")
        Call ExecuteCommand("ˢ��״̬")
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
        
        '���ԭ������
        Call mclsVsf(0).ClearGrid

        '��ȡ��������
        strSQL = "SELECT '����' As ͼ��,A.ID,A.����,A.����,A.����,A.˵�� FROM ���������ο� A "
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rs.BOF = False Then Call mclsVsf(0).LoadGrid(rs)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
    
        Call mfrmChildSchemeEdit.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
            
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ�·�������"
        
        Call ExecuteCommand("��ȡ��ҩ����")
        Call ExecuteCommand("��ȡ���Ϸ���")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("��ȡ���÷���")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��ҩ����"
        
        Call mfrmChildSchemeDrug.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
                    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ���Ϸ���"
        
        Call mfrmChildSchemeMaterial.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
        
        Call mfrmChildSchemeOps.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ���÷���"
        
        Call mfrmChildSchemeCharge.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))

    '------------------------------------------------------------------------------------------------------------------
    Case "������������"
        
        mblnNew = True

        If Val(vsf(0).RowData(vsf(0).Rows - 1)) > 0 Then vsf(0).Rows = vsf(0).Rows + 1
        vsf(0).Row = vsf(0).Rows - 1
        vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        
        Call ExecuteCommand("ˢ�¸�������")

        Call mfrmChildSchemeEdit.NewData(0, mlngTmp)

        Call mfrmChildSchemeDrug.NewData(mlngTmp)
        Call mfrmChildSchemeMaterial.NewData(mlngTmp)
        Call mfrmChildSchemeOps.NewData(mlngTmp)
        Call mfrmChildSchemeCharge.NewData(mlngTmp)
        
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ����������"
        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then Exit Function
        
        If MsgBox("���Ƿ����Ҫɾ����" & vsf(0).TextMatrix(vsf(0).Row, mclsVsf(0).ColIndex("����")) & "������������", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
            strSQL = "ZL_���������ο�_DELETE(" & Val(vsf(0).RowData(vsf(0).Row)) & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        End If
        Exit Function

    '------------------------------------------------------------------------------------------------------------------
    Case "�Ƴ���������"
    
        If vsf(0).Rows > 2 Then
            vsf(0).RemoveItem vsf(0).Row
            mclsVsf(0).AppendRows = True
        Else
            Call mclsVsf(0).ClearGrid
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ָ�����"
    
        '1.�ָ���������
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeEdit.DataChanged Then
            If Val(vsf(0).RowData(vsf(0).Row)) = 0 And vsf(0).Rows > 2 Then
                vsf(0).Rows = vsf(0).Rows - 1
                vsf(0).Row = vsf(0).Rows - 1
            End If

            Call ExecuteCommand("��ȡ��������")
            mfrmChildSchemeEdit.DataChanged = False
        End If

        '2.�ָ���ҩ����
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeDrug.DataChanged Then
            Call ExecuteCommand("��ȡ��ҩ����")
            mfrmChildSchemeDrug.DataChanged = False
        End If

        '3.�ָ����Ϸ���
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeMaterial.DataChanged Then
            Call ExecuteCommand("��ȡ���Ϸ���")
            mfrmChildSchemeMaterial.DataChanged = False
        End If

        '4.�ָ���������
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeOps.DataChanged Then
            Call ExecuteCommand("��ȡ��������")
            mfrmChildSchemeOps.DataChanged = False
        End If

        '5.�ָ����÷���
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeCharge.DataChanged Then
            Call ExecuteCommand("��ȡ���÷���")
            mfrmChildSchemeCharge.DataChanged = False
        End If

        mblnNew = False
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"
    
        '1.У����ϸ����
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeEdit.DataChanged Then
            If mfrmChildSchemeEdit.ValidData = False Then Exit Function
        End If

        '2.У����ҩ����
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeDrug.DataChanged Then
            If mfrmChildSchemeDrug.ValidData = False Then Exit Function
        End If

        '3.У����Ϸ���
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeMaterial.DataChanged Then
            If mfrmChildSchemeMaterial.ValidData = False Then Exit Function
        End If

        '4.У����������
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeOps.DataChanged Then
            If mfrmChildSchemeOps.ValidData = False Then Exit Function
        End If

        '5.У����÷���
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeCharge.DataChanged Then
            If mfrmChildSchemeCharge.ValidData = False Then Exit Function
        End If

        ExecuteCommand = True
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))

        '1.������ϸ����
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeEdit.DataChanged Then

            If mfrmChildSchemeEdit.SaveData(rsSQL, mlngTmp) = False Then Exit Function

        End If

        '2.������ҩ����
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeDrug.DataChanged Then
            If mfrmChildSchemeDrug.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        '3.������Ϸ���
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeMaterial.DataChanged Then
            If mfrmChildSchemeMaterial.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        '4.������������
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeOps.DataChanged Then
            If mfrmChildSchemeOps.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        '5.������÷���
        '--------------------------------------------------------------------------------------------------------------
        If mfrmChildSchemeCharge.DataChanged Then
            If mfrmChildSchemeCharge.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If

        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)

        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ǰһ��"
        If vsf(0).Row > 1 Then
            vsf(0).Row = vsf(0).Row - 1
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
            Call ExecuteCommand("��ȡ��������")
            Call ExecuteCommand("ˢ�·�������")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��һ��"
        If vsf(0).Row < vsf(0).Rows - 1 Then
            vsf(0).Row = vsf(0).Row + 1
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
            Call ExecuteCommand("��ȡ��������")
            Call ExecuteCommand("ˢ�·�������")
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ע���"
        
        If Val(GetRegister(˽��ȫ��, "", "ʹ�ø��Ի����", "0")) = 1 Then
            'ʹ�ø��Ի�����
            
'            dkpMain.LoadStateFromString GetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, "")
            
            mstrFindKey = Trim(GetRegister(˽��ģ��, Me.Name, "��λ����", "����"))
            mclsVsf(0).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)), ""))
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "дע���"
        If Val(GetRegister(˽��ȫ��, "", "ʹ�ø��Ի����", "0")) = 1 Then
            'ʹ�ø��Ի�����
            Call SetRegister(˽��ģ��, Me.Name, "��λ����", mstrFindKey)
        End If
        Call SetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)), mclsVsf(0).SaveStateToString)
    End Select

    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem               '������������
        mlngTmp = 0
        Call ExecuteCommand("������������")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_CopyNewItem           '����������������
    
        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))
        Call ExecuteCommand("������������")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                'ɾ����������

        If ExecuteCommand("ɾ����������") Then
            Call ExecuteCommand("�Ƴ���������")
            Call ExecuteCommand("��ȡ��������")
            Call ExecuteCommand("ˢ�·�������")
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save                  '������������
    
        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then
                DataChanged = False
                Call ExecuteCommand("ˢ��ָ������")
                mblnNew = False
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle                  '�ָ���������
    
        Call ExecuteCommand("�ָ�����")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If tbcPage.Selected.Index + 1 <= tbcPage.ItemCount - 1 Then
            tbcPage.Item(tbcPage.Selected.Index + 1).Selected = True
        Else
            tbcPage.Item(0).Selected = True
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        Call ExecuteCommand("ǰһ��")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        Call ExecuteCommand("��һ��")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        
        mobjFindKey.Execute
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
    
        LocationObj txtLocation
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               'ˢ��

        Call ExecuteCommand("ˢ������")
        
    '------------------------------------------------------------------------------------------------------------------
    Case Else

        If Control.ID > 400 And Control.ID < 500 Then
            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
        Else
             '��ҵ���޹صĹ��ܣ������Ĺ���
            Call CommandBarExecutePublic(Control, Me, vsf(0), "���������嵥")
        End If

    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error GoTo errHand

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel
    
        Control.Enabled = (Val(vsf(0).RowData(vsf(0).Row)) > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup
        
        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_CopyNewItem
    
        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = (DataChanged = False And Control.Visible)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                   '�޸�,ɾ��
    
        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = (Val(vsf(0).RowData(vsf(0).Row)) > 0 And DataChanged = False And Control.Visible)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
        
        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = (DataChanged And Control.Visible)
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        Control.Enabled = (vsf(0).Row > 1 And DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        Control.Enabled = (vsf(0).Row < vsf(0).Rows - 1 And DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem        '
        Control.Checked = (mstrFindKey = Control.Parameter)
        Control.Enabled = (DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
         Control.Enabled = (DataChanged = False)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Find, conMenu_View_Refresh
        Control.Enabled = (DataChanged = False And Control.Visible)
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        Call CommandBarUpdatePublic(Control, Me)
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

End Sub

Private Sub cbsSearch_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Execute(Control)
End Sub

Private Sub cbsSearch_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Update(Control)
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Set mfrmChildSchemeEdit = New frmChildSchemeEdit
        Item.Handle = mfrmChildSchemeEdit.hWnd
        Call mfrmChildSchemeEdit.InitData(Me, IsPrivs(mstrPrivs, "��ɾ��"))
    Case 3
        Item.Handle = picPane(2).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents

    If ExecuteCommand("��ʼ����") = False Then GoTo errHand

    Call ExecuteCommand("ˢ������")

    mblnAllowClose = True
    Exit Sub

errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    mblnAllowClose = False
    
    mstrPrivs = UserInfo.ģ��Ȩ��
    mlngģ��� = ParamInfo.ģ���
    
    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��ע���")

    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, ParamInfo.ϵͳ��, ParamInfo.ģ���, UserInfo.ģ��Ȩ��)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Call SetPaneRange(dkpMain, 1, 100, 60, 200, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 15, 150, Me.ScaleWidth, 150)

    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = Not mblnAllowClose
    
    If Cancel = False Then
    
        If DataChanged Then
            Cancel = (MsgBox("�޸ĺ�����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
        End If
        
        If Cancel = False Then
        
            Call ExecuteCommand("дע���")
            
            Call SaveWinState(Me, App.ProductName)
            
            Set mclsVsf(0) = Nothing
            
            Unload mfrmChildSchemeEdit
            Unload mfrmChildSchemeDrug
            Unload mfrmChildSchemeMaterial
            Unload mfrmChildSchemeOps
            Unload mfrmChildSchemeCharge
        End If
    End If

End Sub

Private Sub mfrmChildSchemeCharge_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildSchemeDrug_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildSchemeEdit_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildSchemeMaterial_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmChildSchemeOps_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf(0).AppendRows = True
    Case 2
        tbcPage.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim bytMatch As Byte
    
    If KeyAscii = vbKeyReturn Then
    
        lngRow = -1
        bytMatch = 0
        Select Case mstrFindKey
        Case "����"
            bytMatch = 2
            intCol = mclsVsf(0).ColIndex("����")
        Case "����"
            bytMatch = 2
            intCol = mclsVsf(0).ColIndex("����")
        Case "����"
            bytMatch = 2
            intCol = mclsVsf(0).ColIndex("����")
        Case Else
            Exit Sub
        End Select
        
        lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch, vsf(0).Row + 1)
        If lngRow = -1 Then
            lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
        End If
        If lngRow > 0 And vsf(0).Row <> lngRow Then
            vsf(0).Row = lngRow
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        End If
        
        Call LocationObj(txtLocation)
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Call mclsVsf(Index).AfterMoveColumn(Col, Position)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Index = 0 Then
        If OldRow = NewRow Then Exit Sub
        Call mclsVsf(Index).SelectRow(OldRow, NewRow)
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("ˢ�·�������")
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey)
    vsf(Index).ShowCell vsf(Index).Row, vsf(Index).Col
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf(Index).SaveKey = Val(vsf(Index).RowData(vsf(Index).Row))
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = mclsVsf(Index).ColIndex("ͼ��"))
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                
                Set cbrPopupBar = CopyMenu(cbsMain, 2)
                If cbrPopupBar Is Nothing Then Exit Sub
                cbrPopupBar.ShowPopup
            End If
        End Select
        
    End Select
End Sub


