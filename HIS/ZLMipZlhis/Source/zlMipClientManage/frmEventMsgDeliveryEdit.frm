VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEventMsgDeliveryEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ͷ��Ŀ��"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11280
   Icon            =   "frmEventMsgDeliveryEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   1
      Left            =   6690
      ScaleHeight     =   3195
      ScaleWidth      =   5205
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   5205
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   4095
         TabIndex        =   3
         Top             =   120
         Width           =   4095
         Begin VB.ComboBox cbo 
            Height          =   300
            Left            =   -45
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   -30
            Width           =   1905
         End
      End
      Begin MSComctlLib.TreeView tvw 
         Height          =   2505
         Left            =   1005
         TabIndex        =   5
         Top             =   495
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   4419
         _Version        =   393217
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ils16"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4830
      Index           =   3
      Left            =   150
      ScaleHeight     =   4830
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   690
      Width           =   5685
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1275
         Index           =   0
         Left            =   495
         TabIndex        =   1
         Top             =   465
         Width           =   6375
         _cx             =   11245
         _cy             =   2249
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
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   4725
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1140
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEventMsgDeliveryEdit.frx":000C
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEventMsgDeliveryEdit.frx":0166
            Key             =   "constitute"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEventMsgDeliveryEdit.frx":0700
            Key             =   "folder"
         EndProperty
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
      Bindings        =   "frmEventMsgDeliveryEdit.frx":0A9A
      Left            =   525
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEventMsgDeliveryEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type Items
    �����Ϣ As String
End Type
Private usrSaveItem As Items
Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mclsVsf(1) As zlVSFlexGrid.clsVsf
Private mlngModualCode As Long
Private mblnContiune As Boolean
Private mrsInfoTree As ADODB.Recordset
Private mstrEventDataKey As String

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

Public Sub ModifyData(ByVal strEventDataKey As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mstrEventDataKey = strEventDataKey
    mbytMode = 2
    mstrDataKey = strDataKey
    
    Me.Caption = "Ͷ��Ŀ������"
        
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
    
    mblnContiune = False
    
'    Set rsTmp = gclsBusiness.EventMsgStruct()
'    If Not (rsTmp Is Nothing) Then
'        txt(0).MaxLength = rsTmp("server_ip").DefinedSize
'        txt(1).MaxLength = rsTmp("server_port").Precision
'        txt(2).MaxLength = rsTmp("target_app").DefinedSize
'        txt(3).MaxLength = rsTmp("target_device").DefinedSize
'        txt(4).MaxLength = rsTmp("note").DefinedSize
'    End If
    
    cbo.AddItem "1 - ��ǰ��Ϣ"
    cbo.ItemData(cbo.NewIndex) = 1
    cbo.AddItem "2 - ��ǰ�¼�����Ϣ"
    cbo.ItemData(cbo.NewIndex) = 2
    cbo.AddItem "3 - ָ���¼�����Ϣ"
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
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "[ѡ��]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("����", 1800, flexAlignLeftCenter, flexDTString, , "", True)
        
        Call .AppendColumn("����", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�豸", 900, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�ӿ�", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("ע��", 1500, flexAlignLeftCenter, flexDTString, , "", True)
        
'        .IndicatorMode = 0
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
        
        
        .AppendRows = True
        
    End With
    
    InitGrid = True
    
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean

    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "�¼���Ϣid", strDataKey)
    
    mblnReading = True
        
        
    '------------------------------------------------------------------------------------------------------------------
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "�¼���Ϣid", strDataKey)
    
    Call mclsVsf(0).LoadDataSource(gclsBusiness.EventMsgServerRead("����", rsCondition))
    
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "�¼�"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "��ϸ"
    objPane.Options = PaneNoCaption
                
    objPane.Close
    
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
            
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "������������", "���������޸�"))
    objControl.IconId = conMenu_View_UnCheck
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "��һ��", True)
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "��һ��")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�", True)
    
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

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
       
    
    ValidData = True
    
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
    
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(lngLoop, .ColIndex("ѡ��")))) = 1 Then
                strTemp = strTemp & ";" & Trim(.TextMatrix(lngLoop, .ColIndex("id")))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    Call zlCommFun.SetParameter(rsPara, "Ͷ��Ŀ��", strTemp)
        
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�޸�
        SaveData = gclsBusiness.EventMsgEdit("TargetConfig", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
    Case conMenu_Edit_Transf_Save
        Call Save
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
    End Select
    
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '���������ؼ�Resize����
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward, 0
        Control.Visible = (mbytMode = 2)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
End Sub

'Private Sub cmd_Click(Index As Integer)
'    Dim strFile As String
'    Dim objclsHL7 As clsHL7V2EDI
'    Dim rsFormat As ADODB.Recordset
'    Dim strMessageType As String
'    Dim rsData As ADODB.Recordset
'    Dim rs As ADODB.Recordset
'    Dim strTemp As String
'
'    Select Case Index
'    '------------------------------------------------------------------------------------------------------------------
'    Case 0
'
'        'cmdlg
'        strFile = OpenHLDialog
'        If strFile <> "" Then
'
'            Set objclsHL7 = New clsHL7V2EDI
'
'            If objclsHL7.GetMessageFormat(strFile, strMessageType, rsFormat) Then
'
'                txt(0).Text = strMessageType
'                If Not (rsFormat Is Nothing) Then
'                    With mclsVsf(1)
''                        Call .LoadDataSource(rsFormat)
'                        Call .LoadGrid(rsFormat)
'                        Call .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"), True)
'                    End With
'                End If
'
'            End If
'
'            Set objclsHL7 = Nothing
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case 1
'
'
'        Set rsData = gclsBusiness.TableRead("SelectData")
'
'        If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "����,900,0,1;����,2400,0,0;˵��,1200,0,0", mfrmParent.Name & "\�����Ϣѡ��", "����±���ѡ��һ�������Ϣ", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
'
'            If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
'
'                txt(Index).Text = zlCommFun.NVL(rs("����").Value)
'                txt(Index).Tag = ""
'                usrSaveItem.�����Ϣ = txt(Index).Text
'                cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
'
'                mblnDataChanged = True
'
'                Call GetRelationInfomation(zlCommFun.NVL(rs("ID").Value))
'
'            End If
''            Call ReEnabled
'            Call LocationObj(txt(Index), True)
'        Else
'            txt(Index).Text = usrSaveItem.�����Ϣ
'            txt(Index).Tag = ""
''            Call ReEnabled
'            Call LocationObj(txt(Index), True)
'            Exit Sub
'        End If
'
'    End Select
'
'End Sub

'Private Sub GetRelationInfomation(ByVal strDataKey As String)
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim strTemp As String
'
'    strTemp = "0|1"
'    Set mrsInfoTree = gclsBusiness.GettableTree(strDataKey)
'    mrsInfoTree.Filter = "��ϵ=2"
'    If mrsInfoTree.RecordCount > 0 Then
'        mrsInfoTree.MoveFirst
'        Do While Not mrsInfoTree.EOF
'            strTemp = strTemp & "|" & mrsInfoTree("����").Value & ".RecordCount"
'            mrsInfoTree.MoveNext
'        Loop
'    End If
'    mclsVsf(1).DropColData(mclsVsf(1).ColIndex("�����ظ�")) = strTemp
'
'End Sub

'Public Function OpenHLDialog() As String
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'
'    Dim strTmp As String
'
'    With cmdlg
'        .DialogTitle = "��ѡ��HL7��Ϣ��׼"
'        .Filter = "��Ϣ��׼(*.config)|*.config"
'
'        On Error Resume Next
'
'        .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
'        .FileName = ""
'        .MaxFileSize = 32767
'        .CancelError = True
'        .ShowOpen
'
'        If Err.Number = 0 And .FileName <> "" Then
'
'            strTmp = .FileName
'
'            On Error GoTo errHand
'
'            OpenHLDialog = strTmp
'
'        Else
'            Err.Clear
'        End If
'    End With
'
'    Exit Function
'
'errHand:
'    ShowSimpleMsg "���ܴ��ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ�������!"
'End Function

'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub

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
                    End If
'                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
                
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(3).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 2, 200, 15, 200, Me.ScaleHeight)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Set mclsVsf(0) = Nothing
    Set mclsVsf(1) = Nothing
    
    Set mrsInfoTree = Nothing
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        picPane(2).Move 15, 15, picPane(Index).Width - 30
        tvw.Move 15, picPane(2).Top + picPane(2).Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (picPane(2).Top + picPane(2).Height + 15) - 15
    Case 2
        cbo.Move -30, -30, picPane(Index).Width + 60
    Case 3
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
    
End Sub

'Private Sub tvw_DblClick()
'
'    Dim strCellText As String
    
    
'    With vsf(1)
'        If Col = .ColIndex("���ݸ�ֵ") Then
'            Select Case UCase(.TextMatrix(Row, .ColIndex("�ڵ�����")))
'            Case UCase("Segment"), UCase("Group")
'                Cancel = True
'            Case Else
'
'                Cancel = False
'
'            End Select
'        End If
'    End With
    
    
'    If tvw.SelectedItem.Child Is Nothing Then
'
'        With vsf(1)
'            strCellText = .TextMatrix(.Row, .ColIndex("���ݸ�ֵ"))
'
'            If Trim(strCellText) <> "" Then
'
'                If tvw.SelectedItem.Parent Is Nothing Then
'                    strCellText = "[" & tvw.SelectedItem.Text & "]"
'                Else
'                    strCellText = strCellText & "^" & "[" & tvw.SelectedItem.Parent.Text & "." & tvw.SelectedItem.Text & "]"
'                End If
'            Else
'                If tvw.SelectedItem.Parent Is Nothing Then
'                    strCellText = "[" & tvw.SelectedItem.Text & "]"
'                Else
'                    strCellText = "[" & tvw.SelectedItem.Parent.Text & "." & tvw.SelectedItem.Text & "]"
'                End If
'            End If
'
'            .TextMatrix(.Row, .ColIndex("���ݸ�ֵ")) = strCellText
'        End With
'    End If
'End Sub

'Private Sub txt_Change(Index As Integer)
'
'    If mblnReading Then Exit Sub
'
'    mblnDataChanged = True
'
'    Select Case Index
'    Case 1
'
'        txt(Index).Tag = "Changed"
'
'    End Select
'
'End Sub
'
'Private Sub txt_GotFocus(Index As Integer)
'
'    zlControl.TxtSelAll txt(Index)
'
'    Select Case Index
'    Case 4
'        zlCommFun.OpenIme True
'    End Select
'
'End Sub
'
'Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    Select Case Index
'    Case 1
'        If KeyCode = vbKeyDelete Then
'            KeyCode = 0
'            txt(Index).Text = ""
'            cmd(Index).Tag = ""
'            txt(Index).Tag = ""
'            usrSaveItem.�����Ϣ = ""
'        End If
'    End Select
'End Sub
'
'Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'
'    Dim strText As String
'    Dim strTmp As String
'    Dim bytMode As Byte
'    Dim rsData As New ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'    Dim rsCondition As ADODB.Recordset
'
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        Select Case Index
'        '--------------------------------------------------------------------------------------------------------------
'        Case 0
''            If cmd(index).Visible And cmd(index).Enabled Then Call cmd_Click(index)
'        '--------------------------------------------------------------------------------------------------------------
'        Case 1
'            If txt(Index).Tag <> "" Then
'                txt(Index).Tag = ""
'
'                Set rsCondition = zlCommFun.CreateCondition
'                Call zlCommFun.SetCondition(rsCondition, "FilterText", txt(Index).Text)
'
'                Set rsData = gclsBusiness.TableRead("FilterData", rsCondition)
'                If zlCommFun.ShowPubSelect(Me, txt(Index), 2, "����,900,0,1;����,2400,0,0;˵��,1200,0,0", mfrmParent.Name & "\�����Ϣ����", "����±���ѡ��һ�������", rsData, rs, , , , Trim(cmd(Index).Tag), , True) = 1 Then
'
'                    If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
'                        txt(Index).Text = zlCommFun.NVL(rs("����").Value)
'                        txt(Index).Tag = ""
'                        usrSaveItem.�����Ϣ = txt(Index).Text
'                        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
'                        mblnDataChanged = True
'                    End If
'                Else
'                    txt(Index).Text = usrSaveItem.�����Ϣ
'                    txt(Index).Tag = ""
'                    Exit Sub
'                End If
'            End If
'        End Select
'
'        zlCommFun.PressKey vbKeyTab
'    Else
'        If Chr(KeyAscii) = "'" Then KeyAscii = 0
'    End If
'End Sub
'
'Private Sub txt_LostFocus(Index As Integer)
'
'    Select Case Index
'    Case 4
'        zlCommFun.OpenIme False
'    End Select
'
'End Sub
'
'Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 And txt(Index).Locked Then
'        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
'        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
'    End If
'End Sub
'
'Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 And txt(Index).Locked Then
'        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
'    End If
'End Sub
'
'Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
'
'    If Cancel Then Exit Sub
'
'    Select Case Index
'    Case 1
'        If (txt(Index).Tag = "Changed") Then
'            txt(Index).Text = usrSaveItem.�����Ϣ
'            txt(Index).Tag = ""
'        End If
'    End Select
'
'End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '�༭����
    Call mclsVsf(Index).AfterEdit(Row, Col)
    mblnDataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim lngRow As Long
    
    Call mclsVsf(Index).DbClick
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf(Index).KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPress(KeyAscii)
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(Index).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf(Index).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(Index).ValidateEdit(Col, Cancel)
End Sub

