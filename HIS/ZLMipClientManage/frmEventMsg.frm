VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEventMsg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   0
      Left            =   960
      ScaleHeight     =   2685
      ScaleWidth      =   5220
      TabIndex        =   2
      Top             =   3705
      Width           =   5220
      Begin XtremeSuiteControls.TabControl tbc 
         Height          =   1590
         Left            =   450
         TabIndex        =   3
         Top             =   360
         Width           =   3165
         _Version        =   589884
         _ExtentX        =   5583
         _ExtentY        =   2805
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2685
      Index           =   1
      Left            =   960
      ScaleHeight     =   2685
      ScaleWidth      =   5220
      TabIndex        =   0
      Top             =   855
      Width           =   5220
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   210
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
         Appearance      =   0
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
      Bindings        =   "frmEventMsg.frx":0000
      Left            =   690
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEventMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'�ֲ�����
Private Enum Command
    ��ʼ�ؼ�
    ��ע���
    ������Ϣ
    ������Ϣ
    �޸���Ϣ
    ɾ����Ϣ
    Ͷ������
    Ͷ��Ӧ��
    ˢ������
    ˢ����Ϣ����
    ˢ����ϢͶ��
    ˢ��ָ����Ϣ
    �Ƴ�ָ����Ϣ
End Enum

Private mlngMoudalCode As Long
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mstrDataKey As String
Private mrsPara As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mfrmEventMsgEdit As frmEventMsgEdit
Attribute mfrmEventMsgEdit.VB_VarHelpID = -1
Private mfrmEventMsgConfig As frmEventMsgConfig
Private mfrmEventMsgDelivery As frmEventMsgDelivery
Private WithEvents mfrmEventMsgDeliveryEdit As frmEventMsgDeliveryEdit
Attribute mfrmEventMsgDeliveryEdit.VB_VarHelpID = -1
Private WithEvents mfrmEventMsgDeliveryApply As frmEventMsgDeliveryApply
Attribute mfrmEventMsgDeliveryApply.VB_VarHelpID = -1

'######################################################################################################################
'�ӿڷ���
Public Function InitForm(ByVal frmParent As Object, ByVal lngMoudalCode As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngMoudalCode = lngMoudalCode
    
    InitForm = True
    
End Function

Public Function ReadData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrDataKey = strDataKey
    
    Call ExecuteCommand(Command.ˢ������)
    
    ReadData = True
    
End Function

Public Sub Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Execute(Control)
End Sub

Public Sub Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call cbsMain_Update(Control)
End Sub

'######################################################################################################################
'˽�з���
Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '��ʼ����ؼ�
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("��Ϣ", 1500, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�汾", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("���", 1800, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("˵��", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .AppendRows = True
        
    End With
            
    InitGrid = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing)
    objPane.Title = "�¼���Ϣ"
    objPane.Options = PaneNoCaption
       
    
    Set objPane = dkpMain.CreatePane(2, 300, 200, DockBottomOf, objPane)
    objPane.Title = "��Ϣ���"
    objPane.Options = PaneNoCaption
        
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Sub InitTabControl()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    On Error GoTo errHand

    '------------------------------------------------------------------------------------------------------------------
    Call TabControlInit(tbc, xtpTabAppearancePropertyPage2003)
    With tbc

        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
        End With
        
        Set mfrmEventMsgConfig = New frmEventMsgConfig
        Call mfrmEventMsgConfig.InitForm(mfrmParent, mlngMoudalCode)
    
        Set mfrmEventMsgDelivery = New frmEventMsgDelivery
        Call mfrmEventMsgDelivery.InitForm(mfrmParent, mlngMoudalCode)
                
        .InsertItem 0, "��ֵ����", mfrmEventMsgConfig.hWnd, 0
        .InsertItem 1, "Ͷ��Ŀ��", mfrmEventMsgDelivery.hWnd, 0

        .Item(0).Selected = True
    End With

    Exit Sub

errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
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
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    
    
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "�¼�����Ϣ", , , xtpButtonIconAndCaption)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_CopyNewItem, "����")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Adjust, "Ͷ��Ŀ��", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ApplyTo, "Ͷ��Ӧ��")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "ˢ��", True)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

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
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "������Ϣ(&N)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "������Ϣ(&C)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "�޸���Ϣ(&M)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "ɾ����Ϣ(&D)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Adjust, "Ͷ��Ŀ��(&T)")
        cbrPopupItem.BeginGroup = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ApplyTo, "Ͷ��Ӧ��(&A)")
    
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        cbrPopupItem.BeginGroup = True
    
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant

    On Error GoTo errHand
            
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.��ʼ�ؼ�
        '
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ������
        
        Set mrsPara = zlCommFun.CreateCondition
        Call zlCommFun.SetCondition(mrsPara, "ҵ���¼�id", mstrDataKey)
    
                
        With vsf(0)
            mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
                    
            ExecuteCommand = mclsVsf(0).LoadDataSource(gclsBusiness.EventMsgRead("", mrsPara))
            
            Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
        End With
        
        Call ExecuteCommand(Command.ˢ����Ϣ����)
        Call ExecuteCommand(Command.ˢ����ϢͶ��)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ��ָ����Ϣ
        
        ExecuteCommand = LoadCustomData(Trim(varParam(0)))
        
        Call ExecuteCommand(Command.ˢ����Ϣ����)
        Call ExecuteCommand(Command.ˢ����ϢͶ��)
        
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ����Ϣ����
        
        With vsf(0)
            Call mfrmEventMsgConfig.ReadData(.TextMatrix(.Row, .ColIndex("id")))
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ����ϢͶ��
        
        With vsf(0)
            Call mfrmEventMsgDelivery.ReadData(.TextMatrix(.Row, .ColIndex("id")))
        End With
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.������Ϣ
    
        If mfrmEventMsgEdit Is Nothing Then
            Set mfrmEventMsgEdit = New frmEventMsgEdit
            Call mfrmEventMsgEdit.InitDialog(mfrmParent, mlngMoudalCode)
        End If
        
        With vsf(0)
            Call mfrmEventMsgEdit.NewData(mstrDataKey)
        End With
        
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.������Ϣ
        
        
        Set rsData = gclsBusiness.EventMsgRead("Select")
        
        If zlCommFun.ShowPubSelect(Me, Nothing, 3, "��Ϣ���,1200,0,1;�����Ϣ,3000,0,0;˵��,3000,0,0", mfrmParent.Name & "\��Ϣѡ��", "����±���ѡ��һ����Ҫ���Ƶ���Ϣ", rsData, rs) = 1 Then
            
            If mfrmEventMsgEdit Is Nothing Then
                Set mfrmEventMsgEdit = New frmEventMsgEdit
                Call mfrmEventMsgEdit.InitDialog(mfrmParent, mlngMoudalCode)
            End If
            Call mfrmEventMsgEdit.CopyNewData(mstrDataKey, rs("ID").Value)
            
'            If Trim(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
'
'                With vsf(1)
'                    .Cell(flexcpText, 1, .ColIndex("�����ظ�"), .Rows - 1, .ColIndex("�����ظ�")) = ""
'                    .Cell(flexcpText, 1, .ColIndex("�����ظ�"), .Rows - 1, .ColIndex("���ݸ�ֵ")) = ""
'                End With
'
'                txt(Index).Text = zlCommFun.NVL(rs("����").Value)
'                txt(Index).Tag = ""
'                usrSaveItem.�����Ϣ = txt(Index).Text
'                cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
'
'                mblnDataChanged = True
'
'                Call GetRelationInfomation(zlCommFun.NVL(rs("ID").Value))
                
'            End If
        End If
        
        DoEvents
        Me.SetFocus
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.�޸���Ϣ
    
        If mfrmEventMsgEdit Is Nothing Then
            Set mfrmEventMsgEdit = New frmEventMsgEdit
            Call mfrmEventMsgEdit.InitDialog(mfrmParent, mlngMoudalCode)
        End If
        
        With vsf(0)
            Call mfrmEventMsgEdit.ModifyData(mstrDataKey, .TextMatrix(.Row, .ColIndex("id")))
        End With
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ɾ����Ϣ
        
        If mfrmEventMsgEdit Is Nothing Then
            Set mfrmEventMsgEdit = New frmEventMsgEdit
            Call mfrmEventMsgEdit.InitDialog(mfrmParent, mlngMoudalCode)
        End If
        
        With vsf(0)
            If MsgBox("��ȷ��Ҫɾ����ǰ�¼���Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                Call mfrmEventMsgEdit.DeleteData(.TextMatrix(.Row, .ColIndex("id")))
            End If
        End With
        
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.Ͷ������
                
        
        If mfrmEventMsgDeliveryEdit Is Nothing Then
            Set mfrmEventMsgDeliveryEdit = New frmEventMsgDeliveryEdit
            Call mfrmEventMsgDeliveryEdit.InitDialog(mfrmParent, mlngMoudalCode)
        End If
        
        With vsf(0)
            Call mfrmEventMsgDeliveryEdit.ModifyData(mstrDataKey, .TextMatrix(.Row, .ColIndex("id")))
        End With
        DoEvents
        Me.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case Command.Ͷ��Ӧ��
        
        If mfrmEventMsgDeliveryApply Is Nothing Then
            Set mfrmEventMsgDeliveryApply = New frmEventMsgDeliveryApply
            
        End If
        
        With vsf(0)
            
            Call mfrmEventMsgDeliveryApply.ShowDialog(mfrmParent, .TextMatrix(.Row, .ColIndex("id")))

        End With
        DoEvents
        Me.SetFocus
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.�Ƴ�ָ����Ϣ
        
        With vsf(0)
            
            intRow = mclsVsf(0).FindRow(Trim(varParam(0)), .ColIndex("id"))
            
            If intRow > 0 Then
                If .Rows > 2 Then
                    .RemoveItem .Row
                    mclsVsf(0).AppendRows = True
                Else
                    Call mclsVsf(0).ClearGrid
                End If
            End If
        End With
    
    End Select
    
    
    GoTo EndHand

    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
End Function

Private Function LoadCustomData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim rsData As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "�¼���Ϣid", strDataKey)
    
    Set rsData = gclsBusiness.EventMsgRead("�¼���Ϣ", rsCondition)
    If rsData.BOF = True Then Exit Function
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(strDataKey, .ColIndex("id"))
        
        If intRow > 0 Then
            '�Ѽ���
            .Row = intRow
        Else
            'δ����
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("id"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
        
        Call mclsVsf(0).LoadGridRow(.Row, rsData)
    End With
    
    mclsVsf(0).AppendRows = True
    
    LoadCustomData = True
    
End Function

'######################################################################################################################
'�����¼�
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDataKey As String
    Dim blnCancel As Boolean
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem              '����
        
        Call ExecuteCommand(Command.������Ϣ)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_CopyNewItem
        
        Call ExecuteCommand(Command.������Ϣ)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify, conMenu_Edit_ModifyParent              '�޸�
        
        Call ExecuteCommand(Command.�޸���Ϣ)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete, conMenu_Edit_DeleteParent

        Call ExecuteCommand(Command.ɾ����Ϣ)
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Adjust
        
        Call ExecuteCommand(Command.Ͷ������)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ApplyTo
        
        Call ExecuteCommand(Command.Ͷ��Ӧ��)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               'ˢ��

        Call ExecuteCommand(Command.ˢ������)
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsf(0)
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem               '����
                                
            Control.Enabled = (mstrDataKey <> "")
        
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_CopyNewItem               '����
        
            Control.Enabled = (mstrDataKey <> "")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify, conMenu_Edit_ModifyParent              '�޸�
            
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
                    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete, conMenu_Edit_DeleteParent
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
            
        End Select
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Item.Handle = picPane(0).hWnd
    End Select
    
End Sub

Private Sub Form_Load()
    Call InitGrid
    Call InitDockPannel
    Call InitCommandBar
    Call InitTabControl
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 15, 30, Me.ScaleWidth, 150)
'    Call zlCommFun.SetPaneRange(dkpMain, 3, 15, 90, Me.ScaleWidth, 150)
    
    dkpMain.RecalcLayout
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
    If Not (mfrmEventMsgEdit Is Nothing) Then
        Unload mfrmEventMsgEdit
        Set mfrmEventMsgEdit = Nothing
    End If
    If Not (mfrmEventMsgConfig Is Nothing) Then
        Unload mfrmEventMsgConfig
        Set mfrmEventMsgConfig = Nothing
    End If
    
    If Not (mfrmEventMsgDelivery Is Nothing) Then
        Unload mfrmEventMsgDelivery
        Set mfrmEventMsgDelivery = Nothing
    End If
    
    If Not (mfrmEventMsgDeliveryEdit Is Nothing) Then
        Unload mfrmEventMsgDeliveryEdit
        Set mfrmEventMsgDeliveryEdit = Nothing
    End If
    
End Sub

Private Sub mfrmEventMsgDeliveryEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.ˢ��ָ����Ϣ, DataKey)
End Sub

Private Sub mfrmEventMsgDeliveryEdit_Backward(DataKey As String, Cancel As Boolean)
    Dim intRow As Integer
    
    With vsf(0)
    
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
        
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
    End With
End Sub

Private Sub mfrmEventMsgDeliveryEdit_Forward(DataKey As String, Cancel As Boolean)
    
    Dim intRow As Integer
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
                
        If .Row > 1 Then
            .Row = .Row - 1
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
    End With
End Sub

Private Sub mfrmEventMsgEdit_AfterDeleteData(ByVal DataKey As String)
    Call ExecuteCommand(Command.�Ƴ�ָ����Ϣ, DataKey)
End Sub

Private Sub mfrmEventMsgEdit_AfterModifyData(ByVal DataKey As String)
    Call ExecuteCommand(Command.ˢ��ָ����Ϣ, DataKey)
End Sub

Private Sub mfrmEventMsgEdit_AfterNewData(ByVal DataKey As String)
    If DataKey = "" Then
        Call ExecuteCommand(Command.ˢ������)
    Else
        Call ExecuteCommand(Command.ˢ��ָ����Ϣ, DataKey)
    End If
End Sub

Private Sub mfrmEventMsgEdit_Backward(DataKey As String, Cancel As Boolean)
    Dim intRow As Integer
    
    With vsf(0)
    
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
        
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
    End With
            
End Sub

Private Sub mfrmEventMsgEdit_Forward(DataKey As String, Cancel As Boolean)
    
    Dim intRow As Integer
    
    With vsf(0)
        
        intRow = mclsVsf(0).FindRow(DataKey, .ColIndex("id"))
        If intRow > 0 And .Row <> intRow Then .Row = intRow
                
        If .Row > 1 Then
            .Row = .Row - 1
            .ShowCell .Row, .Col
            DataKey = .TextMatrix(.Row, .ColIndex("id"))
        End If
    End With
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        tbc.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 1
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub


Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(0).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    If OldRow <> NewRow Then
        Call ExecuteCommand(Command.ˢ����Ϣ����)
        Call ExecuteCommand(Command.ˢ����ϢͶ��)
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(0).AppendRows = True
End Sub


Private Sub vsf_DblClick(Index As Integer)
    Dim objMenu As CommandBarControl
    
    Set objMenu = cbsMain.FindControl(, conMenu_Edit_Modify, False)
    If Not (objMenu Is Nothing) Then
        If objMenu.Enabled = True Then
            Call cbsMain_Execute(objMenu)
        End If
    End If
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick(Index)
    End If
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call zlCommFun.SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub
