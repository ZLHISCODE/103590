VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMipPollLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   8490
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   465
      Width           =   1305
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   -30
         TabIndex        =   7
         Top             =   -30
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   131137539
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   6480
      ScaleHeight     =   240
      ScaleWidth      =   1245
      TabIndex        =   4
      Top             =   390
      Width           =   1275
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   -30
         TabIndex        =   5
         Top             =   -30
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   131137539
         CurrentDate     =   41401
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   5100
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   390
      Width           =   1185
      Begin VB.ComboBox cboPeiord 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1215
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4140
      Index           =   1
      Left            =   105
      ScaleHeight     =   4140
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   480
      Width           =   4740
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3240
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   90
         Width           =   3030
         _cx             =   5345
         _cy             =   5715
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
         GridLines       =   8
         GridLinesFixed  =   8
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   5835
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMipPollLog.frx":0000
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMipPollLog.frx":059A
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMipPollLog.frx":0934
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMipPollLog.frx":1346
            Key             =   "info"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMipPollLog.frx":7BA8
            Key             =   "error"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   15
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMipPollLog.frx":85BA
      Left            =   375
      Top             =   15
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMipPollLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'��������
Private Enum Command
    ��ʼ�ؼ�
    ��ע���
    �����¼�
    �޸��¼�
    ɾ����־
    ˢ������
    ��ϸ��Ϣ
End Enum

Private mblnReading As Boolean
Private mclsMipServiceLog As clsMipServiceLog
Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mblnDataChanged As Boolean
Private mblnStartService As Boolean
Private mstrTitle As String

Public Event AfterClose(ByVal lngModual As Long)
Public Event AfterLoad(ByVal intIndex As Integer, ByVal strContent As String)

'######################################################################################################################
'�ӿڷ���
Public Function ShowForm()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnStartUp = True
    Call Form_Activate
End Function

'######################################################################################################################
'˽�з���

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As zlDataSQLite.SQLiteRecordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim rsCondition As ADODB.Recordset
    Dim strEditMode As String
    Dim blnMuliSelect As Boolean
    
    On Error GoTo errHand
            
    mblnReading = True
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.��ʼ�ؼ�
        Set mclsMipServiceLog = New clsMipServiceLog

        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel

        With cboPeiord
            .Clear
            .AddItem "��  ��"
            .AddItem "��  ��"
            .AddItem "ǰ����"
            .AddItem "��  ��"
            .AddItem "ǰһ��"
            .AddItem "ǰ����"
            .AddItem "��  ��"
            .AddItem "ǰһ��"
            .AddItem "ǰ����"
            .AddItem "��  ��"
            .AddItem "ǰ����"
            .AddItem "������"
            .AddItem "ǰ����"
            .AddItem "�Զ���"
        End With
        If cboPeiord.ListCount > 0 And cboPeiord.ListIndex = -1 Then cboPeiord.ListIndex = 0

        dtp(0).Value = Format(GetBasePeriod(cboPeiord.Text, 1), dtp(0).CustomFormat)
        dtp(1).Value = Format(GetBasePeriod(cboPeiord.Text, 2), dtp(1).CustomFormat)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ɾ����־
        
        With vsf(0)
        
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            If blnMuliSelect = True Then
                If MsgBox("��ȷ��Ҫɾ���Ѿ���ѡ����־��", vbQuestion + vbYesNo + vbDefaultButton2, mstrTitle) = vbYes Then
                    
                    If mclsMipServiceLog.OpenFile(App.Path & "\Data\zlMspPollServiceLog.db") = True Then
                        Set rsCondition = zlCommFun.CreateCondition
                        For intRow = 1 To .Rows - 1
                            If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 Then
                                Call zlCommFun.SetCondition(rsCondition, "ID", .TextMatrix(intRow, .ColIndex("ID")))
                                Call mclsMipServiceLog.EditRunLog("DeleteID", rsCondition)
                            End If
                        Next
                        Call ExecuteCommand(Command.ˢ������)
                        mclsMipServiceLog.CloseFile
                    End If
                    
                End If
            
            Else
                If MsgBox("��ȷ��Ҫɾ����ǰ�е���־��", vbQuestion + vbYesNo + vbDefaultButton2, mstrTitle) = vbYes Then
                    
                    If mclsMipServiceLog.OpenFile(App.Path & "\Data\zlMspPollServiceLog.db") = True Then
                        
                        Set rsCondition = zlCommFun.CreateCondition
                        Call zlCommFun.SetCondition(rsCondition, "ID", .TextMatrix(.Row, .ColIndex("ID")))
                                                                        
                        If mclsMipServiceLog.EditRunLog("DeleteID", rsCondition) Then
                            Call ExecuteCommand(Command.ˢ������)
                        End If
                        mclsMipServiceLog.CloseFile
                    End If
                    
                End If
            End If
            
        End With
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ������
        
        mclsVsf(0).ClearGrid
                        
        If mclsMipServiceLog.OpenFile(App.Path & "\Data\zlMspPollServiceLog.db") = True Then
            
            Set rsCondition = zlCommFun.CreateCondition
            Call zlCommFun.SetCondition(rsCondition, "��ʼʱ��", Format(dtp(0).Value, "yyyy-MM-dd") & " 00:00:00")
            Call zlCommFun.SetCondition(rsCondition, "����ʱ��", Format(dtp(1).Value, "yyyy-MM-dd") & " 23:59:59")
            
            rs = mclsMipServiceLog.GetRunLog("Filter", rsCondition)
            If rs.DataSet.BOF = False Then
                Call mclsVsf(0).LoadDataSource(rs.DataSet.DataSource)
                vsf(0).AutoSize 0, vsf(0).Cols - 1
            End If
            
            DataChanged = False
            
            mclsMipServiceLog.CloseFile
        End If
    End Select
    
    
    GoTo EndHand

    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    '------------------------------------------------------------------------------------------------------------------
EndHand:
    mblnReading = False
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '��ʼ����ؼ�
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, ils16)
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[���]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTBoolean, "", "[ѡ��]", False)
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        Call .AppendColumn("ID", 0, flexAlignLeftTop, flexDTString, , "ID", True, False, , True)
        Call .AppendColumn("ʱ��", 1890, flexAlignLeftTop, flexDTString, , "Log_Time", True, False)
        Call .AppendColumn("����", 600, flexAlignLeftTop, flexDTString, , "Log_Type", True, False)
        Call .AppendColumn("��Ϣ", 3000, flexAlignLeftTop, flexDTString, , "Log_Desc", True, False)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
                
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
        
    End With
            
    InitGrid = True
    
End Function

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
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "ȫѡ", True, , , , , "����ǰ�б��е�����������Ϊ��ѡ״̬")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "ȫ��", , , , , , "����ǰ�б��е�����������Ϊ�ǹ�ѡ״̬")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "���", True, , , , , "�����ǰ�л��߹�ѡ�е�������־")
    
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "ʱ�䷶Χ��", , , xtpButtonCaption)
    objControl.BeginGroup = True
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(2).hwnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "��", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(3).hwnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, 1, "��", , , xtpButtonCaption)
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = picBack(4).hwnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "ˢ��", True)
        
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyDelete, conMenu_Edit_Delete            '���
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll          'ȫѡ
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll       'ȫ��
    End With
    Exit Function
    
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

    Set objPane = dkpMain.CreatePane(1, 100, 300, DockTopOf, Nothing)
    objPane.Title = "��־"
    objPane.Options = PaneNoCaption
        
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim objPopup As CommandBarPopup
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
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ(&A)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��(&C)")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "���(&D)")
        cbrPopupItem.BeginGroup = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        cbrPopupItem.BeginGroup = True
    
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cboPeiord_Click()
    If mblnReading Then Exit Sub
    
    If cboPeiord.Text <> "�Զ���" Then
        dtp(0).Value = Format(GetBasePeriod(cboPeiord.Text, 1), dtp(0).CustomFormat)
        dtp(1).Value = Format(GetBasePeriod(cboPeiord.Text, 2), dtp(1).CustomFormat)
        
        Call ExecuteCommand(Command.ˢ������)
    Else
        DataChanged = True
    End If
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As Object
    
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 1
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
        End With
                
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        Call ExecuteCommand(Command.ˢ������)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Call ExecuteCommand(Command.ɾ����־)
        
    Case conMenu_File_Close
    '--------------------------------------------------------------------------------------------------------------
        Unload Me
        RaiseEvent AfterClose(mlngModualCode)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar         '״̬��
    
        cbsMain.StatusBar.Visible = Not cbsMain.StatusBar.Visible
        cbsMain.RecalcLayout
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
        
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh
        
'        Control.Enabled = DataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button            '������
        If cbsMain.Count >= 2 Then
            Control.Checked = cbsMain(2).Visible
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text              'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = cbsMain.Options.LargeIcons
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = cbsMain.StatusBar.Visible
    End Select
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case 1
        Item.Handle = picBack(1).hwnd
    End Select
End Sub

Private Sub dtp_Change(Index As Integer)
    '����ʱ�������Ϊ���Զ��塰
    mblnReading = True
    
    Select Case Index
    Case 0, 1
        Call zlControl.CboLocate(cboPeiord, "�Զ���")
    End Select
    
    mblnReading = False
    
    DataChanged = True
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    DoEvents
    mblnStartUp = False
    
    Call ExecuteCommand(Command.ˢ������)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1001
    mstrTitle = "��ѯ������"
    Call ExecuteCommand(Command.��ʼ�ؼ�)
    Call ExecuteCommand(Command.��ע���)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsMipServiceLog = Nothing
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        vsf(0).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    End Select
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf(Index).AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call zlCommFun.SendLMouseButton(vsf(Index).hwnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(Index).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).ValidateEdit(Col, Cancel)
End Sub
