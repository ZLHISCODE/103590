VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISBorrow 
   Caption         =   "���Ӳ�������"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11595
   Icon            =   "frmCISBorrow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11595
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   1
      Left            =   5745
      ScaleHeight     =   3015
      ScaleWidth      =   4500
      TabIndex        =   3
      Top             =   1005
      Width           =   4500
      Begin VB.Frame fra 
         Height          =   630
         Left            =   0
         TabIndex        =   6
         Top             =   -90
         Width           =   4410
         Begin VB.CheckBox chk 
            Caption         =   "&4.�ѹ黹"
            Height          =   180
            Index           =   3
            Left            =   3270
            TabIndex        =   10
            Top             =   255
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.CheckBox chk 
            Caption         =   "&3.�Ѿܾ�"
            Height          =   180
            Index           =   2
            Left            =   2220
            TabIndex        =   9
            Top             =   255
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox chk 
            Caption         =   "&2.����׼"
            Height          =   180
            Index           =   1
            Left            =   1185
            TabIndex        =   8
            Top             =   255
            Value           =   1  'Checked
            Width           =   1110
         End
         Begin VB.CheckBox chk 
            Caption         =   "&1.������"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   255
            Value           =   1  'Checked
            Width           =   1020
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   9720
      TabIndex        =   2
      Top             =   105
      Width           =   1125
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   0
      Left            =   285
      ScaleHeight     =   3585
      ScaleWidth      =   4470
      TabIndex        =   1
      Top             =   900
      Width           =   4470
      Begin XtremeSuiteControls.TabControl tbcTask 
         Height          =   1830
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
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
            Object.Width           =   16563
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
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
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCISBorrow.frx":6852
      Left            =   615
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCISBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���弶��������
'######################################################################################################################
Private mstrPrivs As String
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mstrCondition As String
Private mstrFindKey As String
Private mlngTmp As Long
Private mobjFindKey As CommandBarControl
Private mclsVsf(0) As clsVsf
Private mlngModul As Long
Private mintIndex As Integer
Private mbytMode As Byte

Private mobjPrintView As CommandBarControl
Private mobjPrintPatient As CommandBarControl
Private mobjPrint As CommandBarControl

Private mrsCondition As New ADODB.Recordset
Private mfrmChildDocumentView As frmChildDocumentView
Private mblnBorrowReason As Boolean
Private mblnBorrowAccount As Boolean

Private WithEvents mfrmCISBorrowEdit As frmCISBorrowEdit
Attribute mfrmCISBorrowEdit.VB_VarHelpID = -1
Private WithEvents mfrmChildPatientView As frmChildPatient
Attribute mfrmChildPatientView.VB_VarHelpID = -1

'######################################################################################################################

Public Property Get ģ���() As Long
    ģ��� = mlngModul
End Property

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmCISBorrowEdit.DataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmCISBorrowEdit Is Nothing) Then
        DataChanged = mfrmCISBorrowEdit.DataChanged
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
    
    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    
    Set mobjPrintPatient = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BatPrint, "��ӡ�������е���(&B)", True)
    Set mobjPrintView = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrintView, "Ԥ���ĵ�(&E)")
    Set mobjPrint = NewCommandBar(objMenu, xtpControlButton, conMenu_File_BillPrint, "��ӡ�ĵ�(&T)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "��������(&M)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '�༭
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "��������(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "��׼����(&Y)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Refuse, "�ܾ���׼(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "������׼(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Send, "�黹����(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "�������(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������(&C)")
       
    '�鿴
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Column, "ѡ������(&H)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Filter, "����(&F)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    
            
    '����
    '------------------------------------------------------------------------------------------------------------------
    Call CreateHelpMenu(cbsMain)
    
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    mstrFindKey = GetPara("��λ����", mlngModul, "No")
    If mstrFindKey = "" Then mstrFindKey = "No"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.No", , , "No")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.������", , , "������")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.����ʱ��", , , "����ʱ��")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&4.��������", , , "��������")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&5.����", , , "����")
'    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&6.סԺ��", , , "סԺ��")

    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, ""): cbrCustom.Handle = txtLocation.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��"): objControl.Flags = xtpFlagRightAlign: objControl.STYLE = xtpButtonIcon
    
    '����������:������������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("��׼", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "��׼")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "�黹")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    
'    Set objControl = NewToolBar(objBar, xtpControlPopup, conMenu_File_Print, "test")
    
    '����Ŀ����:���������������Ѵ���
    '------------------------------------------------------------------------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save           '����
        .Add 0, vbKeyF12, conMenu_File_Parameter            '��������
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyV, conMenu_File_Preview
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyF, conMenu_View_Filter         '����
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With

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
            Call .Initialize(Me.Controls, vsf(0), True, False, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("��¼״̬", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("No", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("��������", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("��������", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��׼��", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��׼ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("��������", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("�ܽ���", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�ܽ�ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
            Call .AppendColumn("�ܽ�����", 900, flexAlignLeftCenter, flexDTString, "", , True)
            
            .SysHidden(.ColIndex("��¼״̬")) = True
            
            .AppendRows = True
        End With
        
        '��ʼ�˵���������
        '--------------------------------------------------------------------------------------------------------------
        Call InitCommandBar
        
        '����ͣ������
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing): objPane.Title = "�༭": objPane.Options = PaneNoCaption
        Set objPane = dkpMain.CreatePane(3, 100, 100, DockRightOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption

        dkpMain.SetCommandBars cbsMain
        Call DockPannelInit(dkpMain)


        Call TabControlInit(tbcTask)
        With tbcTask
            .PaintManager.BoldSelected = True
            
            Set mfrmChildPatientView = New frmChildPatient
            Call mfrmChildPatientView.zlInitData(Me, 4, mstrPrivs)
            
            .InsertItem 0, "�������뵥", picPane(1).hWnd, 1
            If IsPrivs(mstrPrivs, "���Ĳ���") Then
                .InsertItem 1, "�Ķ����Ӳ���", mfrmChildPatientView.hWnd, 2
            End If
            .Item(0).Selected = True
        End With
        
        mlngTmp = Val(GetPara("�ϴ�״̬", ģ���, "0"))
        If mlngTmp >= 0 And mlngTmp <= 1 And tbcTask.ItemCount > mlngTmp Then tbcTask.Item(mlngTmp).Selected = True
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        '��������������Ŀ�������г�ʼ��
        Call ParamCreate(mrsCondition)
        Call ParamAdd(mrsCondition, "��ʼ���ݺ�", "")
        Call ParamAdd(mrsCondition, "�������ݺ�", "")
        Call ParamAdd(mrsCondition, "������", "")
        Call ParamAdd(mrsCondition, "��׼��", "")
        Call ParamAdd(mrsCondition, "�ܾ���", "")
        
        Call ParamAdd(mrsCondition, "�µǼǵ���", "1")
        Call ParamAdd(mrsCondition, "�Ǽǿ�ʼ����", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "�Ǽǽ�������", Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        Call ParamAdd(mrsCondition, "����׼����", "0")
        Call ParamAdd(mrsCondition, "��׼��ʼ����", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "��׼��������", Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        Call ParamAdd(mrsCondition, "�Ѿܾ�����", "0")
        Call ParamAdd(mrsCondition, "�ܾ���ʼ����", Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd"))
        Call ParamAdd(mrsCondition, "�ܾ���������", Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59")
        
        Call ParamAdd(mrsCondition, "סԺ��", "")
        Call ParamAdd(mrsCondition, "��������", "")
        Call ParamAdd(mrsCondition, "��������", "")
        
        '��ȡȱʡ�Ľ�������Ǽǲ�ѯʱ�䷶Χ
        strTmp = GetPara("�Ǽ�ȱʡ��Χ", mlngModul, "��  ��")
        If strTmp = "" Then strTmp = "��  ��"
        Call ParamWrite(mrsCondition, "�Ǽǿ�ʼ����", GetDateTime(strTmp, 1))
        Call ParamWrite(mrsCondition, "�Ǽǽ�������", GetDateTime(strTmp, 2))
        mblnBorrowReason = zlDatabase.GetPara("����¼�����ԭ��", ParamInfo.ϵͳ��, ParamInfo.ģ���, "0", , IsPrivs(mstrPrivs, "��������"))
        mblnBorrowAccount = zlDatabase.GetPara("��������¼�����ԭ��", ParamInfo.ϵͳ��, ParamInfo.ģ���, "0", , IsPrivs(mstrPrivs, "��������"))
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
        
        If vsf(0).Enabled <> Not DataChanged Then
            vsf(0).Enabled = Not DataChanged
            vsf(0).ForeColor = IIf(DataChanged, COLOR.���ɫ, COLOR.��ɫ)
            tbcTask.Enabled = Not DataChanged
        End If
        stbThis.Panels(3).Enabled = DataChanged
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
        
        If mintIndex = 0 Then
            With vsf(0)
                If Val(.RowData(.Row)) = 0 Then
                    strTmp = "��ǰ��û���κε��Ӳ����������뵥��"
                Else
                    strTmp = "���� " & .Rows - 1 & " �����Ӳ����������뵥��"
                End If
            End With
        Else
            With mfrmChildPatientView.VsfBody
                
                If Val(.RowData(.Row)) = 0 Then
                    strTmp = "��ǰ��û�������Բ��ĵĵ��Ӳ�����"
                Else
                    strTmp = "��ǰ���� " & .Rows - 1 & " �������Բ��ĵĵ��Ӳ�����"
                End If
            
            End With
        End If

        stbThis.Panels(2).Text = strTmp
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        
        Call ExecuteCommand("��ȡ���뵥��")
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("��ȡ���Ĳ���")
        Call ExecuteCommand("ˢ��״̬")
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        mrsCondition.Filter = ""
        ExecuteCommand = frmCISBorrowFilter.ShowPara(Me, mrsCondition)

        GoTo endHand
            
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��ָ������"
        
        Set rs = gclsPackage.GetBorrow(1, mlngTmp)
        If rs.BOF = True Then Exit Function
        
        intRow = mclsVsf(0).FindRow(mlngTmp, -1)
        With vsf(0)
            If intRow > 0 Then
                '�Ѽ���
                .Row = intRow
            Else
                'δ����
                If Val(.RowData(.Rows - 1)) > 0 Then
                    .Rows = .Rows + 1
                    mclsVsf(0).AppendRows = True
                End If
                .Row = .Rows - 1
            End If
            Call mclsVsf(0).LoadGridRow(.Row, rs)
        End With
        
        Call ExecuteCommand("��ȡ��������")
        Call ExecuteCommand("ˢ��״̬")
            
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ���뵥��"
                
        mclsVsf(0).ClearGrid
        
        Set rs = gclsPackage.GetBorrow(2, 0, ParamRead(mrsCondition, "��ʼ���ݺ�"), _
                                                ParamRead(mrsCondition, "�������ݺ�"), _
                                                ParamRead(mrsCondition, "������"), _
                                                ParamRead(mrsCondition, "��׼��"), _
                                                ParamRead(mrsCondition, "�ܾ���"), _
                                                IIf(Val(ParamRead(mrsCondition, "�µǼǵ���")) = 1, ParamRead(mrsCondition, "�Ǽǿ�ʼ����"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "�µǼǵ���")) = 1, ParamRead(mrsCondition, "�Ǽǽ�������"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "����׼����")) = 1, ParamRead(mrsCondition, "��׼��ʼ����"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "����׼����")) = 1, ParamRead(mrsCondition, "��׼��������"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "�Ѿܾ�����")) = 1, ParamRead(mrsCondition, "�ܾ���ʼ����"), ""), _
                                                IIf(Val(ParamRead(mrsCondition, "�Ѿܾ�����")) = 1, ParamRead(mrsCondition, "�ܾ���������"), ""), _
                                                (chk(0).Value = 1), (chk(1).Value = 1), (chk(2).Value = 1), (chk(3).Value = 1), _
                                                ParamRead(mrsCondition, "סԺ��"), _
                                                ParamRead(mrsCondition, "��������"), _
                                                ParamRead(mrsCondition, "��������"))
        If rs.BOF = False Then
            Call mclsVsf(0).LoadGrid(rs)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"
        
        With vsf(0)
            
            If IsPrivs(mstrPrivs, "�޸���������") Then
                Call mfrmCISBorrowEdit.RefreshData(.RowData(.Row), True, mblnBorrowAccount)
            Else
                Call mfrmCISBorrowEdit.RefreshData(.RowData(.Row), (Trim(.TextMatrix(.Row, .ColIndex("������"))) = UserInfo.����), mblnBorrowAccount)
            End If
            
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ���Ĳ���"
        
        Call mfrmChildPatientView.zlRefreshData(mrsCondition)
        Call mfrmChildPatientView.zlShowDocument
        
    '------------------------------------------------------------------------------------------------------------------
    Case "���ӽ�������"
        
        mbytMode = 1
        
        With vsf(0)
            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
            .Row = .Rows - 1
            If .Col = -1 Then .Col = 1
            .ShowCell .Row, .Col
        End With
        
        Call ExecuteCommand("��ȡ��������")
        
        Call mfrmCISBorrowEdit.NewData
        
        GoTo endHand
            
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ����������"
    
        With vsf(0)
            If Val(.RowData(.Row)) = 0 Then GoTo endHand
            
            If MsgBox("���Ƿ����Ҫɾ�����뵥��Ϊ��" & .TextMatrix(.Row, .ColIndex("No")) & "���ĵ��Ӳ�������������", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                strSQL = "zl_�������ļ�¼_Delete(" & Val(.RowData(.Row)) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                If SQLRecordExecute(rsSQL, Me.Caption) Then
                    ExecuteCommand = True
                    Call ExecuteCommand("�Ƴ���������")
                    If .Rows = 2 Then
                        Call mfrmCISBorrowEdit.ClearData
                    End If
                End If
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��׼��������"
        
        mbytMode = 3
        Call mfrmCISBorrowEdit.Aduit
        GoTo endHand
    
    Case "�黹����"
        
        mbytMode = 5
        Call mfrmCISBorrowEdit.Revert
        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ܾ���������"
    
        mbytMode = 4
        Call mfrmCISBorrowEdit.Refuse
        GoTo endHand
    '------------------------------------------------------------------------------------------------------------------
    Case "������׼����"
        With vsf(0)
            mlngTmp = Val(.RowData(.Row))
            If mlngTmp = 0 Then GoTo endHand

            If MsgBox("���Ƿ����Ҫ�������뵥��Ϊ��" & .TextMatrix(.Row, .ColIndex("No")) & "���ĵ��Ӳ��������������׼��", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                strSQL = "zl_�������ļ�¼_Rollback(" & Val(.RowData(.Row)) & ",1)"
                Call SQLRecordAdd(rsSQL, strSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "���˾ܾ�����"
        With vsf(0)
            mlngTmp = Val(.RowData(.Row))
            If mlngTmp = 0 Then GoTo endHand

            If MsgBox("���Ƿ����Ҫ�������뵥��Ϊ��" & .TextMatrix(.Row, .ColIndex("No")) & "���ĵ��Ӳ�����������ľܾ���", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                strSQL = "zl_�������ļ�¼_Rollback(" & Val(.RowData(.Row)) & ",2)"
                Call SQLRecordAdd(rsSQL, strSQL)
                ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
            End If
            GoTo endHand
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "�Ƴ���������"
        
        With vsf(0)
            If .Rows > 2 Then
                .RemoveItem .Row
                mclsVsf(0).AppendRows = True
            Else
                Call mclsVsf(0).ClearGrid
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "�ָ�����"
            
        If mfrmCISBorrowEdit.DataChanged Then
            With vsf(0)
                If Val(.RowData(.Row)) = 0 And .Rows > 2 Then
                    .Rows = .Rows - 1
                    .Row = .Rows - 1
                End If
            End With
            Call ExecuteCommand("��ȡ��������")
            mfrmCISBorrowEdit.DataChanged = False
        End If
        
        mbytMode = 2
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"
    
        '1.
        '--------------------------------------------------
        If mfrmCISBorrowEdit.DataChanged Then
            If mfrmCISBorrowEdit.ValidData(mblnBorrowReason) = False Then GoTo endHand
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
            
        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))
        
        '1.������ϸ����
        '--------------------------------------------------
        If mfrmCISBorrowEdit.DataChanged Then
            If mfrmCISBorrowEdit.SaveData(rsSQL, mlngTmp) = False Then GoTo endHand
        End If
        
        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)
        
        GoTo endHand
            
    '------------------------------------------------------------------------------------------------------------------
    Case "ǰһ��"
        With vsf(0)
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��һ��"
        With vsf(0)
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��ע���"
        
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����
            
            mstrFindKey = Trim(GetPara("��λ����", mlngModul, "No"))
            mclsVsf(0).LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)), ""))
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "дע���"
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
            'ʹ�ø��Ի�����
            Call SetPara("��λ����", mstrFindKey, mlngModul)
        End If
        Call SetPara("�ϴ�״̬", mintIndex, ģ���)
        Call SetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SetRegister(˽��ģ��, Me.Name, "������_" & TypeName(vsf(0)), mclsVsf(0).SaveStateToString)
        
    End Select

    ExecuteCommand = True
  
    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
endHand:
    

End Function

'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim lngLoop As Long
    
    Select Case Control.ID
    Case conMenu_File_Parameter
        Call frmCISBorrowPara.ShowEdit(Me, mstrPrivs)

    Case conMenu_File_BillPrintView                    'Ԥ����ǰ�ĵ�
        If Not mfrmChildDocumentView Is Nothing Then
        
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 1)
            
        End If

    Case conMenu_File_BillPrint                    '��ӡ��ǰ�ĵ�
        If Not mfrmChildDocumentView Is Nothing Then
            Call mfrmChildDocumentView.zlPrintDocument(cbsMain, 2)
        End If

    Case conMenu_File_BatPrint
        Dim blnDoctorAdvice As Boolean
        If zlDatabase.GetPara("סԺҽ����ӡ", ParamInfo.ϵͳ��, ParamInfo.ģ���, "����ҽ����", , IsPrivs(mstrPrivs, "��������")) = "����ҽ����" Then
            blnDoctorAdvice = False
        Else
            blnDoctorAdvice = True
        End If
        Call frmCISAduitPDF.ShowMe(Me, mfrmChildPatientView.VsfBody, 0, blnDoctorAdvice, False)
        
    Case conMenu_Edit_NewItem
        Call ExecuteCommand("���ӽ�������")

    Case conMenu_Edit_Delete                'ɾ����������
        Call ExecuteCommand("ɾ����������")

    Case conMenu_Edit_Audit                      '��׼��������
        Call ExecuteCommand("��׼��������")
    
    Case conMenu_Edit_Send                      '�黹����
        Call ExecuteCommand("�黹����")
        
    Case conMenu_Manage_Refuse                  '�ܾ���������
        Call ExecuteCommand("�ܾ���������")
        
    Case conMenu_Edit_Untread                   '������׼/�ܾ�
        With vsf(0)
            Select Case Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
            Case 2
                If ExecuteCommand("������׼����") Then
                    Call ExecuteCommand("ˢ��ָ������")
                End If
            Case 3
                If ExecuteCommand("���˾ܾ�����") Then
                    Call ExecuteCommand("ˢ��ָ������")
                End If
            End Select
        End With
    Case conMenu_Edit_Transf_Save                  '��������
    
        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then
                
                DataChanged = False
                
                Call ExecuteCommand("ˢ��ָ������")
                
            End If
        End If
    Case conMenu_Edit_Transf_Cancle                  '�ָ�����
        Call ExecuteCommand("�ָ�����")
        
    Case conMenu_View_Filter '����
        If ExecuteCommand("��������") Then
            Call ExecuteCommand("ˢ������")
        End If

    Case conMenu_View_Column
        If mintIndex = 0 Then
            If frmTemplateColumn.ShowColumn(Me, mclsVsf(0)) Then
                mclsVsf(0).AppendRows = True
            End If
        Else
            Call mfrmChildPatientView.zlColumnSelect
        End If

    Case conMenu_View_Refresh
        Call ExecuteCommand("ˢ������")
        
    Case conMenu_View_Forward
        Call ExecuteCommand("ǰһ��")

    Case conMenu_View_Backward
        Call ExecuteCommand("��һ��")

    Case conMenu_View_Option
        mobjFindKey.Execute

    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    Case conMenu_View_Location
        LocationObj txtLocation
        
    Case Else
        If Control.ID > 400 And Control.ID < 500 Then
            Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me, "ID=" & Val(vsf(0).RowData(vsf(0).Row)))
        Else
             '��ҵ���޹صĹ��ܣ������Ĺ���
            Call CommandBarExecutePublic(Control, Me, vsf(0), "���Ӳ����������뵥")
        End If
        
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand
    
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = (Val(.RowData(.Row)) > 0) And tbcTask.Selected.Index = 0
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_BillPrintView, conMenu_File_BillPrint, conMenu_File_BatPrint

            Control.Visible = IsPrivs(mstrPrivs, "��ӡԤ���ĵ�") And tbcTask.Selected.Index = 1
            Control.Enabled = (Control.Visible And tbcTask.Selected.Index = 1)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Parameter, conMenu_View_Filter, conMenu_View_Refresh, conMenu_View_Column
            Control.Enabled = DataChanged = False
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_EditPopup
            Control.Visible = (tbcTask.Selected.Index = 0)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem
            Control.Visible = IsPrivs(mstrPrivs, "�Ǽ�����") And tbcTask.Selected.Index = 0
            Control.Enabled = Control.Visible And DataChanged = False
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            Control.Visible = IsPrivs(mstrPrivs, "�Ǽ�����") And tbcTask.Selected.Index = 0
            With vsf(0)
                
                If IsPrivs(mstrPrivs, "�޸���������") Then
                    Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1
                Else
                    Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1 And Val(.TextMatrix(.Row, .ColIndex("������"))) = UserInfo.����
                End If
                
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Audit
                        
            Control.Visible = IsPrivs(mstrPrivs, "��������") And tbcTask.Selected.Index = 0
            
            With vsf(0)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Send
            
            Control.Visible = IsPrivs(mstrPrivs, "�黹����") And tbcTask.Selected.Index = 0
            
            With vsf(0)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 2
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Refuse
            Control.Visible = IsPrivs(mstrPrivs, "��������") And tbcTask.Selected.Index = 0
            With vsf(0)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            Control.Visible = IsPrivs(mstrPrivs, "��������") And tbcTask.Selected.Index = 0
            With vsf(0)
                Control.Visible = Control.Visible And (Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) > 1) And (Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) < 4)
                Control.Enabled = Control.Visible And DataChanged = False And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) > 1
                Control.Caption = IIf(Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 2, "������׼(&B)", "���˾ܾ�(&B)")
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
            
            Control.Visible = (IsPrivs(mstrPrivs, "�Ǽ�����") Or IsPrivs(mstrPrivs, "�޸���������") Or Trim(vsf(0).TextMatrix(.Row, vsf(0).ColIndex("������"))) = UserInfo.����) And tbcTask.Selected.Index = 0
            
            Control.Enabled = Control.Visible And DataChanged = True
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Forward
            
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (.Row > 1 And DataChanged = False)
            Case 1
                Control.Enabled = (mfrmChildPatientView.VsfBody.Row > 1)
                            
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Backward
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (.Row < .Rows - 1 And DataChanged = False)
            Case 1
                Control.Enabled = (mfrmChildPatientView.VsfBody.Row < mfrmChildPatientView.VsfBody.Rows - 1)
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_LocationItem        '
            Control.Checked = (mstrFindKey = Control.Parameter)
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (DataChanged = False)
            Case 1
                
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_Location
            Select Case tbcTask.Selected.Index
            Case 0
                Control.Enabled = (DataChanged = False)
            Case 1
                
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
End Sub

Private Sub chk_Click(Index As Integer)
        
    Call ExecuteCommand("ˢ������")
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)

    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Set mfrmCISBorrowEdit = New frmCISBorrowEdit
        Call mfrmCISBorrowEdit.InitData(Me, mlngModul, True, mstrPrivs, mblnBorrowAccount)
        Item.Handle = mfrmCISBorrowEdit.hWnd
    Case 3
        Set mfrmChildDocumentView = New frmChildDocumentView
        Call mfrmChildDocumentView.zlInitData(Me)
        Item.Handle = mfrmChildDocumentView.hWnd
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

    '------------------------------------------------------------------------------------------------------------------
errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mblnAllowClose = False

    mstrPrivs = UserInfo.ģ��Ȩ��
    mlngModul = ParamInfo.ģ���

    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��ע���")
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 1, 100, 100, 300, Me.ScaleHeight)
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
            
            On Error Resume Next

            Unload mfrmCISBorrowEdit
            Unload mfrmChildPatientView
            Unload mfrmChildDocumentView
        End If
    End If

End Sub

Private Sub mfrmChildPatientView_AfterDocumentChanged(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng�ύId As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)
    Call mfrmChildDocumentView.zlRefresh(lng����ID, lng��ҳID, strObject, strParam, strCaption, blnDataMove)
    
    mobjPrintView.Caption = "Ԥ��""" & mfrmChildPatientView.Title & """(&E)"
    mobjPrint.Caption = "��ӡ""" & mfrmChildPatientView.Title & """(&T)"
    With mfrmChildPatientView.VsfBody
        mobjPrintPatient.Caption = "��ӡ""" & .TextMatrix(.Row, .ColIndex("����")) & """�ĵ���(&B)"
    End With
    cbsMain.RecalcLayout
    
End Sub

Private Sub mfrmChildPatientView_DbClick()
'    Dim intRow As Integer
'    Dim strNo As String
'
'
'    With mfrmChildPatientView.VsfBody
'
'        strNo = .TextMatrix(.Row, .ColIndex("No"))
'
'        If strNo <> "" And DataChanged = False Then
'            tbcTask.Item(0).Selected = True
'            With vsf(0)
'                For intRow = 1 To .Rows - 1
'                    If strNo = .TextMatrix(intRow, .ColIndex("No")) Then
'                        .Row = intRow
'                        .ShowCell .Row, .Col
'                        Exit Sub
'                    End If
'                Next
'            End With
'        End If
'
'    End With
    '���ڶ�λ 20120326ȥ����λ����
    
End Sub

Private Sub mfrmChildPatientView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 Then
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub mfrmCISBorrowEdit_AfterDataChanged()
    Call ExecuteCommand("�ؼ�״̬")
End Sub

Private Sub mfrmCISBorrowEdit_ViewDocument(ByVal strNo As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    
    '�л������Ĳ��˵ĵ��Ӳ���״̬
    If strNo <> "" And lng����ID > 0 And lng��ҳID > 0 Then
        tbcTask.Item(1).Selected = True
        Call mfrmChildPatientView.zlLocationPatient(1, , , strNo, lng����ID, lng��ҳID)
    End If
    
End Sub

'�Զ�����̻���
'######################################################################################################################

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        tbcTask.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 1
        fra.Move fra.Left, fra.Top, picPane(Index).Width - fra.Left
        vsf(0).Move 0, vsf(0).Top, picPane(Index).Width, picPane(Index).Height - vsf(0).Top
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub tbcTask_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mintIndex = Item.Index
    
    Select Case Item.Index
    Case 0
        If dkpMain.Panes(2).Selected = False Then dkpMain.Panes(2).Select
        If dkpMain.Panes(3).Closed = False Then dkpMain.Panes(3).Close
    Case 1
        If dkpMain.Panes(3).Selected = False Then dkpMain.Panes(3).Select
        If dkpMain.Panes(2).Closed = False Then dkpMain.Panes(2).Close
    End Select
    
    Call ExecuteCommand("ˢ��״̬")
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim bytMatch As Byte
    
    If KeyAscii = vbKeyReturn Then
        If txtLocation.Text = "" Then Exit Sub
        lngRow = -1
        bytMatch = 2
        
        If tbcTask.Item(0).Selected Then
            intCol = mclsVsf(0).ColIndex(mstrFindKey)
            If intCol >= 0 Then
                lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch, vsf(0).Row + 1)
                If lngRow = -1 Then
                    lngRow = mclsVsf(0).FindRow(UCase(txtLocation.Text), intCol, bytMatch)
                End If
                If lngRow > 0 And vsf(0).Row <> lngRow Then
                    vsf(0).Row = lngRow
                    vsf(0).ShowCell vsf(0).Row, vsf(0).Col
                End If
            End If
        Else
            Call mfrmChildPatientView.zlLocationPatient(2, mstrFindKey, txtLocation.Text)
        End If
        
        Call LocationObj(txtLocation)
    Else
        If InStr(":��;��?��''||", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("��ȡ��������")

    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    
    If Button = 2 And Index = 0 Then
        Call SendLMouseButton(vsf(Index).hWnd, x, y)
        
        Set cbrPopupBar = CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        
        cbrPopupBar.ShowPopup
    End If
    
End Sub

