VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmItemConfigEdit 
   Caption         =   "#"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14595
   Icon            =   "frmItemConfigEdit.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   14595
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1785
      TabIndex        =   12
      Top             =   15
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5460
      Index           =   0
      Left            =   7890
      ScaleHeight     =   5460
      ScaleWidth      =   4830
      TabIndex        =   3
      Top             =   690
      Width           =   4830
      Begin VB.PictureBox picPane 
         BorderStyle     =   0  'None
         Height          =   2910
         Index           =   1
         Left            =   15
         ScaleHeight     =   2910
         ScaleWidth      =   4605
         TabIndex        =   8
         Top             =   2010
         Width           =   4605
         Begin VB.PictureBox picPane 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   0
            ScaleHeight     =   285
            ScaleWidth      =   2280
            TabIndex        =   9
            Top             =   0
            Width           =   2280
            Begin VB.ComboBox cbo 
               Height          =   300
               Left            =   -45
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   -30
               Width           =   2220
            End
         End
         Begin MSComctlLib.TreeView tvw 
            Height          =   1890
            Left            =   15
            TabIndex        =   11
            Top             =   330
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3334
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   476
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picPane 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   15
         ScaleHeight     =   375
         ScaleWidth      =   4605
         TabIndex        =   4
         Top             =   15
         Width           =   4605
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   6
            Top             =   45
            Width           =   2070
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   1
            Left            =   2955
            Picture         =   "frmItemConfigEdit.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   315
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����Ϣ"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   7
            Top             =   105
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3990
      Index           =   3
      Left            =   255
      ScaleHeight     =   3990
      ScaleWidth      =   6660
      TabIndex        =   0
      Top             =   1005
      Width           =   6660
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3645
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   150
         Width           =   5925
         _cx             =   10451
         _cy             =   6429
         Appearance      =   1
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
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
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
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
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4725
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   9225
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmItemConfigEdit.frx":685E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20876
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmItemConfigEdit.frx":70F2
      Left            =   525
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmItemConfigEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private Type Items
    �����Ϣ As String
End Type
Private usrSaveItem As Items
Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjToolBar As Object
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private WithEvents mclsVsf As zlVSFlexGrid.clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private WithEvents mclsParaVsf As zlVSFlexGrid.clsVsf
Attribute mclsParaVsf.VB_VarHelpID = -1
Private mlngModualCode As Long
Private mblnContiune As Boolean
Private mrsInfoTree As ADODB.Recordset
Private mrsTrigger As ADODB.Recordset
Private mstrEventDataKey As String
Private mblnStartUp As Boolean
Private mintMaxOutlineLevel As Integer
Private mintAryOutlineLevel(0 To 11) As Integer
Private mobjFindKey As CommandBarControl
Private mstrDataCode As String
Private mrsCondition As ADODB.Recordset

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object, ByVal lngModualCode As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngModualCode = lngModualCode
    
    
    InitDialog = True
    
End Function

Public Sub ConfigData(ByVal strDataCode As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 2
    mstrDataKey = strDataKey
    mstrDataCode = strDataCode
    
    Me.Caption = "���ݹ���"
        
    Call InitData
    Call InitGrid
    Call InitCommandBar
    Call InitDockPannel
    
    Call ReadData(mstrDataKey)
    
'    Call mclsVsf.LoadDataSource(gclsBusiness.EventMsgServerRead("����"))
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub CopyNewData(ByVal strEventDataKey As String, ByVal strReferMsgDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 4
    If strEventDataKey = "" Then Exit Sub
        
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "ҵ���¼�id", strEventDataKey)
    Call zlCommFun.SetParameter(mrsPara, "�ο���Ϣid", strReferMsgDataKey)
        
    If gclsBusiness.EventMsgEdit("Copy", mrsPara) Then
        ShowSimpleMsg "����ɸ�����Ϣ������"
        RaiseEvent AfterNewData("")
    End If
    
End Sub

Public Sub ModifyData(ByVal strEventDataKey As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mstrEventDataKey = strEventDataKey
    mbytMode = 2
    mstrDataKey = strDataKey
    
    Me.Caption = "�޸��¼���Ϣ"
        
    Call InitData
    Call InitGrid
    Call InitCommandBar
    Call InitDockPannel
    
    Call ReadData(mstrDataKey)
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 3
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "�¼���Ϣid", mstrDataKey)
        
    If gclsBusiness.EventMsgEdit("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
End Sub

'######################################################################################################################
Private Function InitData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    
    Set mrsCondition = zlCommFun.CreateCondition
    Set tvw.ImageList = gfrmPubResource.GetImageCtl
    mblnContiune = False
        
    cbo.AddItem "1 - �̶���Ϣ"
    cbo.ItemData(cbo.NewIndex) = 1
    cbo.AddItem "2 - ҵ����Ϣ"
    cbo.ItemData(cbo.NewIndex) = 2
    cbo.AddItem "3 - ��������"
    cbo.ItemData(cbo.NewIndex) = 3
    cbo.ListIndex = 0
    
    InitData = True
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '��ʼ����ؼ�
            
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(1), True, False, GetImageList(16))
        Call .ClearColumn

        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("parent_id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�ڵ����", 2100, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�ڵ�����", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("��������", 1080, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�ظ�Ƶ��", 810, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�����ظ�_Key", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�����ظ�", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("���ݸ�ֵ_Key", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("���ݸ�ֵ", 3000, flexAlignLeftCenter, flexDTString, , "", True)
                
        vsf(1).OutlineCol = .ColIndex("�ڵ����")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("�����ظ�"), True, vbVsfEditCombox)
        Call .InitializeEditColumn(.ColIndex("���ݸ�ֵ"), True, vbVsfEditText)
        
'        .AppendRows = True
        
    End With
        
    InitGrid = True
    
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "ID", strDataKey)
    
    mblnReading = True
        
    Set rsTmp = gclsBusiness.ItemRead("ID", rsCondition)
    If rsTmp.BOF = False Then
        
        txt(1).Text = zlCommFun.NVL(rsTmp("tab_title").Value)
        cmd(1).Tag = zlCommFun.NVL(rsTmp("tab_id").Value)
        
        usrSaveItem.�����Ϣ = txt(1).Text
                        
        Call GetRelationInfomation(cmd(1).Tag)
    
        Set mrsTrigger = New ADODB.Recordset
        With mrsTrigger
            .Fields.Append "����", adVarChar, 100
            .Fields.Append "���", adTinyInt
            .Open
        End With
        Call zlCommFun.SetCondition(rsCondition, "item_id", strDataKey)
        Set rsTmp = gclsBusiness.ItemFieldRead("item_id", rsCondition)
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                mrsTrigger.AddNew
                mrsTrigger("���").Value = zlCommFun.NVL(rsTmp("fld_order").Value)
                mrsTrigger("����").Value = zlCommFun.NVL(rsTmp("fld_title").Value)
                rsTmp.MoveNext
            Loop
'            If mrsTrigger.RecordCount > 0 Then
'                mclsVsf.DropColData(mclsVsf.ColIndex("�����ظ�")) = mclsVsf.DropColData(mclsVsf.ColIndex("�����ظ�")) & "|[T.��������.Count]"
'            End If
        End If
        
    End If
    '------------------------------------------------------------------------------------------------------------------
        
    With mclsVsf
        Call zlCommFun.SetCondition(rsCondition, "item_id", strDataKey)
        Call .LoadGrid(gclsBusiness.ItemConfigRead("item_id", rsCondition))
        Call vsf(1).AutoSize(.ColIndex("���ݸ�ֵ"), .ColIndex("���ݸ�ֵ"))
        mintMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"))
        For intLoop = mintMaxOutlineLevel To 1 Step -1
            If intLoop < 12 Then mintAryOutlineLevel(intLoop) = 1
            Call mclsVsf.OutLine(intLoop)
        Next
        
        Call OutlineExpand(mintMaxOutlineLevel)
        
    End With
                
    With vsf(1)
        .AutoSize .ColIndex("���ݸ�ֵ"), .ColIndex("���ݸ�ֵ")
    End With
    
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 300, 100, DockLeftOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, objPane)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
            
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

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
    Dim objFindKey As CommandBarControl
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    
    
    Set mobjToolBar = cbsMain.Add("������", xtpBarTop)
    mobjToolBar.ContextMenuPresent = False
    mobjToolBar.ShowTextBelowIcons = False
    mobjToolBar.EnableDocking xtpFlagStretched
        
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_Edit_Append, "��ʽ")
    objControl.IconId = conMenu_Edit_NewItem
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_Edit_Transf_Save, "����")
    
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    
   

    mstrFindKey = zlDataBase.GetPara("��λ����", ParamInfo.ϵͳ��, mlngModualCode, "����")
    If mstrFindKey = "" Then mstrFindKey = "����"

    Set mobjFindKey = zlCommFun.NewToolBar(mobjToolBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.����"): objControl.Parameter = "����"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.����"): objControl.Parameter = "����"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(mobjToolBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Filter, "����")
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Forward, "��һ��")
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Backward, "��һ��")
    
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlLabel, 2, "���:", True, , xtpButtonCaption)
    For intLoop = 1 To 10
        Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, 1, " " & intLoop & " ", , , xtpButtonCaption, "ѡ��չ����ǰ�㣬��ѡ������£��ǰ��")
        objControl.Parameter = intLoop
    Next
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, 1, "����", , , xtpButtonCaption, "ѡ��չ����ǰ�㣬��ѡ������£��ǰ��")
    
    
    Set objControl = zlCommFun.NewToolBar(mobjToolBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "������������", "������������"))
    objControl.IconId = conMenu_View_UnCheck
    objControl.flags = xtpFlagRightAlign
    
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    With cbsMain.KeyBindings
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub UpdateLevelCommandBar()
    Dim intLoop As Integer
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

'    Set objBar = cbsMain.Item(1)
'
'    objBar.Controls.DeleteAll
'
'    If mintMaxOutlineLevel > 0 Then
'        Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 0, "���:", , , xtpButtonCaption)
'        For intLoop = 1 To mintMaxOutlineLevel
'            Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, 1, " " & intLoop & " ", , , xtpButtonCaption, "ѡ��չ����ǰ�㣬��ѡ������£��ǰ��")
'            objControl.Parameter = intLoop
'        Next
'        objBar.Visible = True
'    Else
'        objBar.Visible = False
'    End If
    
End Sub

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intCurrentOrder As Integer
    Dim lngLoop As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim intLoop As Integer
    Dim strKey As String
    Dim strParentKey As String
    Dim varElement As Variant
    Dim varElementKey As Variant
    Dim intCount As Integer
    Dim strName As String
    
'    If Trim(txt(0).Text) = "" Then
'        ShowSimpleMsg "����ѡ��һ����Ϣ��"
'        Call LocationObj(txt(0))
'        Exit Function
'    End If
    
'    If Trim(cmd(1).Tag) = "" Then
'        ShowSimpleMsg "����ѡ��һ�������Ϣ��"
'        Call LocationObj(txt(1))
'        Exit Function
'    End If
        
    'ͬһ����·���в��ܴ���������ͬ�ġ���Ϣ��.Count"����
    With vsf(1)
        For lngLoop = 1 To .Rows - 1
            If InStr(UCase(.TextMatrix(lngLoop, .ColIndex("�����ظ�"))), UCase(".Count")) > 0 And .TextMatrix(lngLoop, .ColIndex("parent_id")) <> "" Then
                If CheckParentRepeat(.TextMatrix(lngLoop, .ColIndex("�����ظ�")), .TextMatrix(lngLoop, .ColIndex("parent_id"))) = False Then
                    Exit Function
                End If
            End If
        Next
        
    End With
    
    
    '���ݡ������ظ����͡����ݸ�ֵ��������Ӧ��Keyֵ�����ڲ�����Ϣ����
    With vsf(1)
        
        .Cell(flexcpText, 1, .ColIndex("�����ظ�_Key"), .Rows - 1, .ColIndex("�����ظ�_Key")) = ""
        .Cell(flexcpText, 1, .ColIndex("���ݸ�ֵ_Key"), .Rows - 1, .ColIndex("���ݸ�ֵ_Key")) = ""
        
        For lngLoop = 1 To .Rows - 1
            
            '�����ظ�
            If InStr(UCase(.TextMatrix(lngLoop, .ColIndex("�����ظ�"))), UCase(".Count")) > 0 Then
                strTemp = .TextMatrix(lngLoop, .ColIndex("�����ظ�"))
                
                If InStr(strTemp, "[S.") > 0 Then
                    strTemp = Replace(strTemp, "[S.", "[" & txt(1).Text & ".")
                    strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                    strTemp = Replace(strTemp, ".Count", "")
                    varTemp = Split(strTemp, ".")
                    strParentKey = ""
                    strKey = ""
                    For intLoop = 0 To UBound(varTemp)
                        mrsInfoTree.Filter = ""
                        mrsInfoTree.Filter = "�ϼ�id='" & strParentKey & "' And ����='" & varTemp(intLoop) & "' And ����=0"
                        If mrsInfoTree.RecordCount > 0 Then
                            strKey = mrsInfoTree("id").Value
                            strParentKey = mrsInfoTree("id").Value
                        Else
                            ShowSimpleMsg "��Ϣ��������"
                            .Row = lngLoop
                            .Col = .ColIndex("�����ظ�")
                            .ShowCell lngLoop, .ColIndex("�����ظ�")
                            .SetFocus
                            Exit Function
                        End If
                    Next
                    
                    If InStr(strKey, "R_") > 0 Then
                        strKey = Mid(strKey, InStr(strKey, "R_") + 2, 32)
                    Else
                        If InStr(strKey, "T_") > 0 Then
                            strKey = Mid(strKey, InStr(strKey, "T_") + 2, 32)
                        Else
                            strKey = ""
                        End If
                    End If
                    If strKey <> "" Then
                        .TextMatrix(lngLoop, .ColIndex("�����ظ�_Key")) = "[" & strKey & ".Count]"
                    End If
                Else
                    .TextMatrix(lngLoop, .ColIndex("�����ظ�_Key")) = strTemp
                End If
                
            End If
            
            
            '���ݸ�ֵ
            If .TextMatrix(lngLoop, .ColIndex("���ݸ�ֵ")) <> "" Then
                
                strTemp = .TextMatrix(lngLoop, .ColIndex("���ݸ�ֵ"))
                
                If InStr(strTemp, "[S.") > 0 Then
                    strTemp = Replace(strTemp, "[S.", "[" & txt(1).Text & ".")
                    varElement = GetElement(strTemp)
                    varElementKey = varElement
                    If IsEmpty(varElement) = False Then
                        For intCount = 0 To UBound(varElement)
                            
                            varTemp = Split(varElement(intCount), ".")
                            strParentKey = ""
                            strKey = ""
                            strName = ""
                            For intLoop = 0 To UBound(varTemp) - 1
                                mrsInfoTree.Filter = ""
                                If intLoop = UBound(varTemp) Then
                                    mrsInfoTree.Filter = "�ϼ�id='" & strParentKey & "' And ����='" & varTemp(intLoop) & "' And ����=1"
                                Else
                                    mrsInfoTree.Filter = "�ϼ�id='" & strParentKey & "' And ����='" & varTemp(intLoop) & "' And ����=0"
                                End If
                        
                                If mrsInfoTree.RecordCount > 0 Then
                                    strKey = mrsInfoTree("id").Value
                                    strParentKey = mrsInfoTree("id").Value
                                    strName = strName & "." & varTemp(intLoop)
                                Else
                                    ShowSimpleMsg "��Ϣ��������"
                                    .Row = lngLoop
                                    .Col = .ColIndex("���ݸ�ֵ")
                                    .ShowCell lngLoop, .ColIndex("���ݸ�ֵ")
                                    .SetFocus
                                    Exit Function
                                End If
                            Next
                            If strName <> "" Then strName = Mid(strName, 2)
                            
                            If InStr(strKey, "R_") > 0 Then
                                strKey = Mid(strKey, InStr(strKey, "R_") + 2, 32)
                            Else
                                If InStr(strKey, "T_") > 0 Then
                                    strKey = Mid(strKey, InStr(strKey, "T_") + 2, 32)
                                Else
                                    strKey = ""
                                End If
                            End If
                            
                            If strKey <> "" Then varElementKey(intCount) = Replace(varElementKey(intCount), strName & ".", strKey & ".")
    '                        If strKey <> "" Then strTemp = Replace(strTemp, strName & ".", strKey & ".")
                            
                        Next
                        
                        For intCount = 0 To UBound(varElement)
                            strTemp = Replace(strTemp, varElement(intCount), varElementKey(intCount))
                        Next
                        
                    End If
                End If
                
                .TextMatrix(lngLoop, .ColIndex("���ݸ�ֵ_Key")) = strTemp
                
            End If
            
        Next
        If Not (mrsInfoTree Is Nothing) Then mrsInfoTree.Filter = ""
    End With
        
    ValidData = True
End Function

Private Function GetElement(ByVal strExpress As String) As Variant
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim lngBeginVar As Long
    Dim lngEndVar As Long
    Dim strVar As String
    Dim strTemp As String
    Dim strChar As String
    
    lngCount = Len(strExpress)
    For lngLoop = 1 To lngCount
        strChar = Mid(strExpress, lngLoop, 1)
        Select Case strChar
        Case "["
            lngBeginVar = lngLoop
        Case "]"
            If lngBeginVar > 0 Then
                lngEndVar = lngLoop
                strTemp = Mid(strExpress, lngBeginVar + 1, lngEndVar - lngBeginVar - 1)
                
                If InStr("'" & strVar & "'", "'" & strTemp & "'") = 0 And InStr(strTemp, ".") > 0 Then
                    strVar = strVar & "'" & strTemp
                End If
                                
                lngBeginVar = 0
                lngEndVar = 0
            End If
        End Select
    Next
    If strVar <> "" Then strVar = Mid(strVar, 2)
    If strVar <> "" Then GetElement = Split(strVar, "'")
End Function

Private Function CheckTableOrder(ByVal intChildOrder As Integer, ByVal strCurrentConfig As String, ByVal strParentKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strParentConfig As String
    Dim intColIndex As Integer
    Dim lngRow As Long
    Dim intCurrentOrder As Integer
    
    With vsf(1)
        lngRow = .FindRow(strParentKey)
        
        intColIndex = .ColIndex("�����ظ�")
        
        If InStr(UCase(.TextMatrix(lngRow, intColIndex)), UCase(".Count")) > 0 Then
            strParentConfig = .TextMatrix(lngRow, intColIndex)
            
            mrsInfoTree.Filter = ""
            mrsInfoTree.Filter = "����='" & Replace(strParentConfig, ".Count", "") & "'"
            If mrsInfoTree.RecordCount > 0 Then
                intCurrentOrder = mrsInfoTree("���").Value
                If intCurrentOrder >= intChildOrder Then
                    ShowSimpleMsg "ͬһ���ͷ�֧·��������������������ϲ�ͬ�ġ�" & strCurrentConfig & "�����ã�����˳�����������Ϣ��������Ϣ����һ�£�"
                    .ShowCell lngRow, intColIndex
                    .Row = lngRow
                    .Col = intColIndex
                    .SetFocus
                    Exit Function
                End If
            End If
        End If
        strParentKey = .TextMatrix(lngRow, .ColIndex("parent_id"))
                
        If CheckTableOrder(intCurrentOrder, strCurrentConfig, strParentKey) = False Then
            Exit Function
        End If
        
    End With
                    
    CheckTableOrder = True
    
End Function

Private Function CheckParentRepeat(ByVal strCurrentConfig As String, ByVal strParentKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strParentConfig As String
    Dim intColIndex As Integer
    Dim lngRow As Long
    
    With vsf(1)
'        lngRow = .FindRow(strParentKey)
        lngRow = mclsVsf.FindRow(strParentKey, .ColIndex("ID"), 1)
        
        intColIndex = .ColIndex("�����ظ�")
        
        If InStr(UCase(.TextMatrix(lngRow, intColIndex)), UCase(".Count")) > 0 Then
            strParentConfig = .TextMatrix(lngRow, intColIndex)
            If strCurrentConfig = strParentConfig Then
                '��ʾ
                ShowSimpleMsg "ͬһ���ͷ�֧·���в��ܴ���������ͬ�ġ�" & strCurrentConfig & "�����ã�"
                .ShowCell lngRow, intColIndex
                .Row = lngRow
                .Col = intColIndex
                .SetFocus
                Exit Function
            End If
        End If
        
        strParentKey = .TextMatrix(lngRow, .ColIndex("parent_id"))
        If strParentKey <> "" Then
            If CheckParentRepeat(strCurrentConfig, strParentKey) = False Then
                Exit Function
            End If
        End If
        
    End With
                    
    CheckParentRepeat = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim strTemp As String
    Dim strLine As String
    Dim lngLoop As Long
    Dim aryTemp As Variant
    Dim lngCount As Long
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "item_id", strDataKey)
'    Call zlCommFun.SetParameter(rsPara, "��Ϣ��ʶ", Trim(txt(0).Text))
'    Call zlCommFun.SetParameter(rsPara, "��Ϣ�汾", Trim(txt(3).Text))
'    Call zlCommFun.SetParameter(rsPara, "item_id", mstrEventDataKey)
    Call zlCommFun.SetParameter(rsPara, "tab_id", cmd(1).Tag)
    Call zlCommFun.SetParameter(rsPara, "ʧ�ܴ���", 0)
'    Call zlCommFun.SetParameter(rsPara, "˵��", Trim(txt(2).Text))
        
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    lngCount = 0
    With vsf(1)
        For lngLoop = 1 To .Rows - 1
            
            strLine = ""
            strLine = Trim(.TextMatrix(lngLoop, .ColIndex("id")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("parent_id")))
            strLine = strLine & "," & lngLoop
            
            '1-Segment;2-Data;3-Composite;4-Group
            
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("�ڵ����")))
            Select Case Trim(.TextMatrix(lngLoop, .ColIndex("�ڵ�����")))
            Case "Complex"
                strLine = strLine & ",1"
            Case "Data"
                strLine = strLine & ",2"
            End Select
            
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("��������")))
            
            If .TextMatrix(lngLoop, .ColIndex("�ظ�Ƶ��")) <> "" Then
                aryTemp = Split(.TextMatrix(lngLoop, .ColIndex("�ظ�Ƶ��")), "��")
                strLine = strLine & "," & Trim(aryTemp(0))
                strLine = strLine & "," & Trim(aryTemp(1))
            Else
                strLine = strLine & ",0"
                strLine = strLine & ",0"
            End If
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("�����ظ�")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("���ݸ�ֵ")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("�����ظ�_Key")))
            strLine = strLine & "," & Trim(.TextMatrix(lngLoop, .ColIndex("���ݸ�ֵ_Key")))
                        
            If LenB(strTemp & ";" & strLine) > 3500 Then
                If strTemp <> "" Then
                    lngCount = lngCount + 1
                    strTemp = Mid(strTemp, 2)
                    Call zlCommFun.SetParameter(rsPara, "��Ϣ����_" & lngCount, strTemp)
                    strTemp = ""
                End If
            End If
            strTemp = strTemp & ";" & strLine
        Next
    End With
    
    If strTemp <> "" Then
        lngCount = lngCount + 1
        strTemp = Mid(strTemp, 2)
        Call zlCommFun.SetParameter(rsPara, "��Ϣ����_" & lngCount, strTemp)
    End If
    Call zlCommFun.SetParameter(rsPara, "��Ϣ���ö���", lngCount)
    
'    '------------------------------------------------------------------------------------------------------------------
'    strTemp = ""
'    With vsf(0)
'        For lngLoop = 1 To .Rows - 1
'            If Abs(Val(.TextMatrix(lngLoop, .ColIndex("ѡ��")))) = 1 Then
'                strTemp = strTemp & ";" & Trim(.TextMatrix(lngLoop, .ColIndex("id")))
'            End If
'        Next
'    End With
'    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
'    Call zlCommFun.SetParameter(rsPara, "Ͷ��Ŀ��", strTemp)
        
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '����
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.ItemConfigEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�޸�
        SaveData = gclsBusiness.ItemConfigEdit("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click()
    Dim objNode As Node
    
    tvw.Nodes.Clear
    
    Select Case cbo.ItemData(cbo.ListIndex)
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        tvw.Nodes.Add , , "K6", "��Ŀ��ʶ", "file", "file"
        tvw.Nodes.Add , , "K5", "��ǰʱ��", "file", "file"
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
        If Not (mrsInfoTree Is Nothing) Then
            
            mrsInfoTree.Filter = ""
            If mrsInfoTree.RecordCount > 0 Then
                mrsInfoTree.MoveFirst
                Do While Not mrsInfoTree.EOF
                    
                    If zlCommFun.NVL(mrsInfoTree("�ϼ�id").Value) = "" Then
                        Set objNode = tvw.Nodes.Add(, , "K_" & mrsInfoTree("id").Value, mrsInfoTree("����").Value)
                        objNode.Expanded = True
                    Else
                        Set objNode = tvw.Nodes.Add("K_" & mrsInfoTree("�ϼ�id").Value, tvwChild, "K_" & mrsInfoTree("id").Value, mrsInfoTree("����").Value)
                        objNode.Expanded = False
                    End If
                    
                    'constitute
                    
                    If Val(mrsInfoTree("��ϵ").Value) = 2 Then
                        objNode.Image = "constitute"
                    Else
                        objNode.Image = IIf(Val(mrsInfoTree("����").Value) = 0, "folder", "file")
                    End If
                    
                    mrsInfoTree.MoveNext
                Loop
            End If
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 3          '��������
        If mrsTrigger.RecordCount > 0 Then
            mrsTrigger.MoveFirst
            Do While Not mrsTrigger.EOF
                
                tvw.Nodes.Add , , "K" & mrsTrigger("���").Value, mrsTrigger("����").Value, "file", "file"
                
                mrsTrigger.MoveNext
            Loop
        End If
    End Select
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intLoop As Integer
    Dim lngRow As Long
    Dim intIndex As Integer
    Dim intMaxIndex As Integer
    Dim blnCancel As Boolean
    Dim strDataKey As String
    Dim strFile As String
    Dim objclsMessageSchema As clsMessageSchema
    Dim rsFormat As ADODB.Recordset
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append

        
        strFile = OpenHLDialog
        If strFile <> "" Then
            Set objclsMessageSchema = New clsMessageSchema
            If objclsMessageSchema.GetMessageFormat(strFile, rsFormat) Then

                If Not (rsFormat Is Nothing) Then
                    With mclsVsf
                        Call .LoadGrid(rsFormat)
                        mintMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"))
'                        mobjToolBar.Visible = (mintMaxOutlineLevel > 0)
                        For intLoop = mintMaxOutlineLevel To 1 Step -1
                            If intLoop < 12 Then mintAryOutlineLevel(intLoop) = 1
                            Call mclsVsf.OutLine(intLoop)
                        Next
                        
                        Call OutlineExpand(mintMaxOutlineLevel)
                        mblnDataChanged = True
                    End With
                End If

            End If
            Set objclsMessageSchema = Nothing
        End If
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save
        Call Save
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit
        Unload Me
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '��һ��
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '��һ��
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
    '--------------------------------------------------------------------------------------------------------------
    Case 1
        intIndex = Val(Control.Parameter)
        If mintAryOutlineLevel(intIndex) = 1 Then
            'չ��,ǰ����Զ�չ��
            Call OutlineExpand(intIndex)
        Else
            
            Call OutlineCollapsed(intIndex)
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append
        Call Fill(True)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        Call Fill(False)
    End Select
    
End Sub

Private Sub OutlineExpand(ByVal intIndex As Integer)

    Dim intLoop As Integer
    Dim lngRow As Long
    Dim intMaxIndex As Integer
    
    With vsf(1)
        If intIndex > 10 Then
            intMaxIndex = mintMaxOutlineLevel
        Else
            intMaxIndex = intIndex
        End If
        
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) = True And .RowOutlineLevel(lngRow) <= intMaxIndex Then
                .IsCollapsed(lngRow) = flexOutlineExpanded
            End If
        Next
    End With
    
    For intLoop = intIndex To 1 Step -1
        mintAryOutlineLevel(intLoop) = -1
    Next
            
End Sub

Private Sub OutlineCollapsed(ByVal intIndex As Integer)

    Dim intLoop As Integer
        
    If intIndex > 10 Then
        For intLoop = 11 To mintMaxOutlineLevel
            mclsVsf.OutLine intLoop
        Next
    Else
        mclsVsf.OutLine intIndex
    End If
    
    mintAryOutlineLevel(intIndex) = 1
            
End Sub


Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward
        Control.Visible = (mbytMode = 2)
        Control.Enabled = (Control.Visible And mblnDataChanged = False)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    '--------------------------------------------------------------------------------------------------------------
    Case 1
        
        Control.Checked = (mintAryOutlineLevel(Val(Control.Parameter)) = -1)
        Control.Visible = (Val(Control.Parameter) > 0 And Val(Control.Parameter) <= mintMaxOutlineLevel)
        
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim strFile As String
    Dim rsFormat As ADODB.Recordset
    Dim strMessageType As String
    Dim strMessageVer As String
    Dim rsData As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim strTemp As String
    Dim intLoop As Integer
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        
        Call zlCommFun.SetCondition(mrsCondition, "data_code", mstrDataCode)
        Set rsData = gclsBusiness.TableRead("SelectData", mrsCondition)
        
        If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "����,900,0,1;����,2400,0,0;˵��,1200,0,0", mfrmParent.Name & "\�����Ϣѡ��", "����±���ѡ��һ�������Ϣ", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
            
            If Trim(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                
                With vsf(1)
                    .Cell(flexcpText, 1, .ColIndex("�����ظ�"), .Rows - 1, .ColIndex("�����ظ�")) = ""
                    .Cell(flexcpText, 1, .ColIndex("�����ظ�"), .Rows - 1, .ColIndex("���ݸ�ֵ")) = ""
                End With
                
                txt(Index).Text = zlCommFun.NVL(rs("����").Value)
                txt(Index).Tag = ""
                usrSaveItem.�����Ϣ = txt(Index).Text
                cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                
                mblnDataChanged = True
                
                Call GetRelationInfomation(zlCommFun.NVL(rs("ID").Value))
                
            End If
'            Call ReEnabled
            Call LocationObj(txt(Index), True)
        Else
            txt(Index).Text = usrSaveItem.�����Ϣ
            txt(Index).Tag = ""
'            Call ReEnabled
            Call LocationObj(txt(Index), True)
            Exit Sub
        End If
        
    End Select
    
End Sub

Private Sub GetRelationInfomation(ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTemp As String
    Dim varTemp As Variant
    Dim intLoop As Integer
    Dim strParentKey As String
    Dim strElement As String
    
    Set mrsInfoTree = gclsBusiness.GetTableTree(strDataKey)
    mrsInfoTree.Filter = "��ϵ=2"
    If mrsInfoTree.RecordCount > 0 Then
        mrsInfoTree.MoveFirst
        Do While Not mrsInfoTree.EOF
            strTemp = strTemp & "|" & mrsInfoTree("����").Value & "." & mrsInfoTree("�ϼ�id").Value
            mrsInfoTree.MoveNext
        Loop
    End If
    
    
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '
    varTemp = Split(strTemp, "|")
    
    strTemp = ""
    For intLoop = 0 To UBound(varTemp)
        
        strElement = ""
        strParentKey = Mid(varTemp(intLoop), InStr(varTemp(intLoop), ".") + 1)
        strElement = Mid(varTemp(intLoop), 1, InStr(varTemp(intLoop), ".") - 1)
        If strParentKey = "" Then
            strElement = "[" & Mid(varTemp(intLoop), 1, InStr(varTemp(intLoop), ".") - 1) & ".Count]"
        Else
            Do While strParentKey <> ""
                mrsInfoTree.Filter = ""
                mrsInfoTree.Filter = "id='" & strParentKey & "'"
                If mrsInfoTree.RecordCount > 0 Then
                    '
                    strElement = mrsInfoTree("����").Value & "." & strElement
                    strParentKey = mrsInfoTree("�ϼ�id").Value
                Else
                    strParentKey = ""
                End If
            Loop
'            If strElement <> "" Then strElement = Mid(strElement, 1, Len(strElement) - 1)
            If strElement <> "" Then strElement = "[" & strElement & ".Count]"
        End If
        
        strTemp = strTemp & "|" & strElement
    Next
    
    strTemp = Replace(strTemp, "[" & txt(1).Text & ".", "[S.")
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    If strTemp <> "" Then
        strTemp = "0|1|" & strTemp
    Else
         strTemp = "0|1"
    End If
    
    mclsVsf.DropColData(mclsVsf.ColIndex("�����ظ�")) = strTemp
                
End Sub

Public Function OpenHLDialog() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strTmp As String
            
    With cmdlg
        .DialogTitle = "��ѡ����Ϣ��ʽ"
        .Filter = "��Ϣ��ʽ(*.xsd)|*.xsd"
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowOpen
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            OpenHLDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    ShowSimpleMsg "���ܴ��ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ�������!"
End Function

Private Sub Save()
    Dim strOldDataKey As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnDataChanged = True Then
        If ValidData = True Then
    
            If SaveData(mstrDataKey) = True Then
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    If mbytMode = 1 Then
                        mstrDataKey = ""
                        txt(0).Text = ""
                        txt(1).Text = ""
                        txt(2).Text = ""
                        txt(3).Text = ""
                        cmd(1).Tag = ""
                        mclsVsf.ClearGrid
                        Call LocationObj(txt(0))
                    Else
                        vsf(1).SetFocus
                    End If
                    
                    mblnDataChanged = False
                End If
                
            End If
        End If
    Else
        If mblnContiune = False Then Unload Me
    End If
    
End Sub

Private Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '�����˵�����
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    
    Select Case bytPlace
    Case 1  '
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Append, "׷��(&A)")
        cbrPopupItem.DefaultItem = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "�滻(&U)")
    
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(3).hWnd
    Case 2
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Call zlComLib.RestoreWinState(Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'
    Call zlCommFun.SetPaneRange(dkpMain, 2, 200, 15, 200, Me.ScaleHeight)
'
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    
    Call zlComLib.SaveWinState(Me)
    
    
    If Not (mclsVsf Is Nothing) Then Set mclsVsf = Nothing
    If Not (mrsInfoTree Is Nothing) Then Set mrsInfoTree = Nothing
    If Not (mrsTrigger Is Nothing) Then Set mrsTrigger = Nothing
    
End Sub

Private Sub mclsVsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    With vsf(1)
        .TextMatrix(Row, .ColIndex("�����ظ�")) = ""
        .TextMatrix(Row, .ColIndex("���ݸ�ֵ")) = ""
        mblnDataChanged = True
    End With
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        picPane(5).Move 15, 15, picPane(Index).Width - 30
'        picPane(4).Move 15, picPane(5).Top + picPane(5).Height + 15, picPane(5).Width
        picPane(1).Move 15, picPane(5).Top + picPane(5).Height + 15, picPane(5).Width, picPane(Index).Height - (picPane(5).Top + picPane(5).Height + 15) - 15
    Case 1
        picPane(2).Move 15, 15, picPane(Index).Width - 30
        tvw.Move 15, picPane(2).Top + picPane(2).Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (picPane(2).Top + picPane(2).Height + 15) - 15
    Case 2
        cbo.Move -30, -30, picPane(Index).Width + 60
    Case 3
        vsf(1).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 15
    Case 4
'        vsfPara.Move 15, 0, picPane(Index).Width - 30, picPane(Index).Height
'        mclsParaVsf.AppendRows = True
    Case 5
        txt(1).Move txt(1).Left, txt(1).Top, picPane(Index).Width - txt(1).Left - cmd(1).Width - 30
        cmd(1).Left = txt(1).Left + txt(1).Width + 15
        
    End Select
    
End Sub

Private Sub tvw_DblClick()
    
    Call Fill(True)
    
End Sub

Private Sub Fill(Optional ByVal blnAppend As Boolean = True)
    
    Dim strSelectText As String
    Dim strCellText As String
    Dim objNode As Node
        
    If tvw.SelectedItem.Child Is Nothing Then
        
        If mclsVsf.AllowColEdit(mclsVsf.ColIndex("���ݸ�ֵ")) = False Then GoTo EndPoint
        
        '��ȡ��ѡ�������
        Set objNode = tvw.SelectedItem
        Select Case cbo.ItemData(cbo.ListIndex)
        Case 1          '�̶�
            strSelectText = "[" & objNode.Text & "]"
        Case 2
            If Not (objNode.Parent Is Nothing) Then
                strSelectText = objNode.Text
                Do While Not (objNode.Parent Is Nothing)
                    strSelectText = objNode.Parent.Text & "." & strSelectText
                    Set objNode = objNode.Parent
                Loop
                If strSelectText <> "" Then strSelectText = "[" & strSelectText & "]"
            End If
            strSelectText = Replace(strSelectText, "[" & txt(1).Text & ".", "[S.")
        Case 3          '����
            strSelectText = "[T." & objNode.Text & "]"
        End Select
        
        If strSelectText = "" Then GoTo EndPoint
                
        With vsf(1)
                                            
            If blnAppend Then
                strCellText = .TextMatrix(.Row, .ColIndex("���ݸ�ֵ"))
                If Trim(strCellText) = "" Then
                    strCellText = strSelectText
                Else
                    strCellText = strCellText & "^" & strSelectText
                End If
                
            Else
                strCellText = strSelectText
            End If
            
            .TextMatrix(.Row, .ColIndex("���ݸ�ֵ")) = strCellText
            
            .AutoSize .ColIndex("���ݸ�ֵ"), .ColIndex("���ݸ�ֵ")
            If .TextMatrix(.Row, .ColIndex("���ݸ�ֵ")) <> "" And (.TextMatrix(.Row, .ColIndex("�����ظ�")) = "" Or .TextMatrix(.Row, .ColIndex("�����ظ�")) = "0") Then
                .TextMatrix(.Row, .ColIndex("�����ظ�")) = "1"
            End If
            mblnDataChanged = True
            
        End With
    End If
    
EndPoint:
    vsf(1).Col = vsf(1).ColIndex("���ݸ�ֵ")
    vsf(1).ShowCell vsf(1).Row, vsf(1).Col
    vsf(1).SetFocus
    
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call tvw_DblClick
    End If
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call zlCommFun.SendLMouseButton(tvw.hWnd, X, Y)

        Call ShowConetneMenu(1).ShowPopup

        
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
    
    Select Case Index
    Case 1
    
        txt(Index).Tag = "Changed"
            
    End Select
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            cmd(Index).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.�����Ϣ = ""
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
'            If cmd(index).Visible And cmd(index).Enabled Then Call cmd_Click(index)
        '--------------------------------------------------------------------------------------------------------------
        Case 1
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""

                Set rsCondition = zlCommFun.CreateCondition
                Call zlCommFun.SetCondition(rsCondition, "FilterText", txt(Index).Text)
                
                Set rsData = gclsBusiness.TableRead("FilterData", rsCondition)
                If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "����,900,0,1;����,2400,0,0;˵��,1200,0,0", mfrmParent.Name & "\�����Ϣ����", "����±���ѡ��һ�������", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
                    
                    If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                        txt(Index).Text = zlCommFun.NVL(rs("����").Value)
                        txt(Index).Tag = ""
                        usrSaveItem.�����Ϣ = txt(Index).Text
                        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                        mblnDataChanged = True
                    End If
                Else
                    txt(Index).Text = usrSaveItem.�����Ϣ
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            End If
        End Select
        
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 4
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Cancel Then Exit Sub

    Select Case Index
    Case 1
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.�����Ϣ
            txt(Index).Tag = ""
        End If
    End Select
    
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '�༭����
    Call mclsVsf.AfterEdit(Row, Col)
    mblnDataChanged = True
    
    With vsf(Index)
        Select Case Col
        Case .ColIndex("���ݸ�ֵ")
            If .TextMatrix(Row, .ColIndex("���ݸ�ֵ")) <> "" And (.TextMatrix(Row, .ColIndex("�����ظ�")) = "" Or .TextMatrix(Row, .ColIndex("�����ظ�")) = "0") Then
                .TextMatrix(Row, .ColIndex("�����ظ�")) = "1"
            End If
            .AutoSize .ColIndex("���ݸ�ֵ"), .ColIndex("���ݸ�ֵ")
        End Select
        
    End With
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnStartUp Then Exit Sub

    With vsf(Index)
        
        Select Case UCase(.TextMatrix(NewRow, .ColIndex("�ڵ�����")))
        Case UCase("Complex")
            mclsVsf.AllowColEdit(mclsVsf.ColIndex("���ݸ�ֵ")) = False
        Case Else
            mclsVsf.AllowColEdit(mclsVsf.ColIndex("���ݸ�ֵ")) = True
        End Select

    End With
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    Dim rsData As ADODB.Recordset
'    Dim strDataInfo As String
'
'    With vsf(Index)
'        Select Case Col
'        Case .ColIndex("���ݸ�ֵ")
'
'            strDataInfo = ShowPubSelect(Me, vsf(Index), mrsInfoTree, .TextMatrix(Row, Col), 7500, 6000)
'            .Cell(flexcpData, Row, Col, Row, Col) = strDataInfo
'            .TextMatrix(Row, Col) = strDataInfo
'            mblnDataChanged = True
'        End Select
'    End With
    
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim lngRow As Long
    
    Call mclsVsf.DbClick
    
    With vsf(Index)
        Select Case .Col
        Case .ColIndex("�����ظ�"), .ColIndex("���ݸ�ֵ")
        
        Case Else
            
            lngRow = .Row
            
            If .IsSubtotal(lngRow) = True Then
                .IsCollapsed(lngRow) = IIf(.IsCollapsed(lngRow) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
            End If
            
        End Select
    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngRow As Long
    
    With vsf(Index)
        If KeyAscii = vbKeySpace Then
            Select Case .Col
            Case .ColIndex("�����ظ�"), .ColIndex("���ݸ�ֵ")
            
            Case Else
                Call vsf_DblClick(Index)
            End Select
        End If
    End With
    
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub

