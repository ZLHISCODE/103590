VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMain_�������󲡰���Ϣ 
   Caption         =   "ְ����ͨ���ﱨ��"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13905
   Icon            =   "frmMain_�������󲡰���Ϣ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   13905
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7560
      TabIndex        =   2
      ToolTipText     =   "��ݼ���F3"
      Top             =   0
      Width           =   1320
   End
   Begin VB.PictureBox picPTMZ 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   6570
      Left            =   1080
      ScaleHeight     =   6570
      ScaleWidth      =   11445
      TabIndex        =   0
      Top             =   420
      Width           =   11445
      Begin VSFlex8Ctl.VSFlexGrid vsfPTMZ 
         Height          =   4695
         Left            =   105
         TabIndex        =   1
         Top             =   135
         Width           =   10635
         _cx             =   18759
         _cy             =   8281
         Appearance      =   2
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmMain_�������󲡰���Ϣ.frx":6852
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   7
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
      Left            =   1770
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_�������󲡰���Ϣ.frx":6A6D
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_�������󲡰���Ϣ.frx":78BF
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   9855
      Width           =   13905
      _ExtentX        =   24527
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
            Object.Width           =   19659
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   End
End
Attribute VB_Name = "frmMain_�������󲡰���Ϣ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long      '�����ؼ�����ˢ��
Private mstrPrivs               As String               'Ȩ�޴�
Private mobjFindKey             As CommandBarPopup      '��ѯ
Private mstrFindKey             As String               '��ѯ��
Private mlngModule              As Long                 'ģ���
Private mstrSaveKey             As String               '������ϴεķ���ѡ��ؼ���
Private mRsPTMZ                 As ADODB.Recordset      '���ݼ�
Private mRsPTMZMX               As ADODB.Recordset      '���ݼ�
Private mRsPTMZBX               As ADODB.Recordset      '���ݼ�
Private mRsPTMZBXMX             As ADODB.Recordset      '���ݼ�
Private mstrSortID              As String               '����λ
Private mcbrPopupBar            As CommandBar           '��������
Private mintInsure              As Integer              '����
Dim cbrPopupItem                As CommandBarControl    '������

'��ӡģʽ
Private Enum gzlPrintModeS
    zlPrint = 1         '��ӡ
    zlView = 2          '�鿴
    zlExcel = 3         '�����Excel
End Enum
Private mzlPrintModeS           As gzlPrintModeS        '��ӡ

Private Const mstrPTMZ = "select A.RESIDENCE_NO As ID,A.Up As �Ƿ��ϴ�,A.UpMan AS �ϴ���,A.UpDateTime As �ϴ�ʱ��,A.StickID As ����ID,A.CnName As ����,A.Sex As �Ա�,A.IDENTITY_NUMBER As ���֤��,B.ҽ����,A.RESIDENCE_NO As סԺ��," & vbNewLine & _
                        "A.MEDICAL_RECORD_NO As ������,A.ADMISSION_DATE ��Ժ����,A.CONTACT_PERSON AS ��ϵ��,A.CONTACT_PHONE AS ��ϵ�绰,A.CONTACT_ADDRESS As ��ϵ��ַ" & vbNewLine & _
                        "from ���β�����Ϣ A,�����ʻ� B" & vbNewLine & _
                        "Where A.Stickid=B.����ID"
                          
Public Property Let intinsure(ByVal vNewValue As String)
    mintInsure = vNewValue
End Property

'==============================================================================
'=���ܣ� ��ʼ�˵�������
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo ErrH

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
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "�����Ǽ�(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸ĵǼ�(&E)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ���Ǽ�(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Send, "�ϴ�����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SendBack, "�����ϴ�(&J)")
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Find, "����(&F)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)

    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrSysName)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrSysName & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrSysName & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&E)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)
    
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16

    mstrFindKey = Trim(GetPara("��λ����"))
    If InStr("סԺ��,ҽ����", mstrFindKey) = 0 Then mstrFindKey = "סԺ��"

    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.��ˮ��", , , "��ˮ��")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.ҽ����", , , "ҽ����")

    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hwnd
    cbrCustom.flags = xtpFlagRightAlign

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��")
    objControl.flags = xtpFlagRightAlign
    objControl.Style = xtpButtonIcon

    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��")
    objControl.flags = xtpFlagRightAlign
    objControl.Style = xtpButtonIcon
    
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Send, "�ϴ�����")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SendBack, "�����ϴ�")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings
        
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '�޸�
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       'ɾ��
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add 0, vbKeyF4, conMenu_View_Option                'ѡ��λ����
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With
    '------------------------------------------------------------------------------------------------------------------
    '�����˵�����
    
    Set mcbrPopupBar = cbsMain.Add("������Ŀ�˵�", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)", True)
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&E)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "�ϴ�����(&S)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SendBack, "�����ϴ�(&J)")
    Set cbrPopupItem = mcbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    SaveFlexState vsfPTMZ, Me.Name
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "����", Me.WindowState
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "LEFT", Me.Left
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "TOP", Me.Top
End Sub

Private Sub picDetail_Resize()
'    tbcPage.Move picDetail.Left + 15, tbcPage.Top + 15
End Sub

'==============================================================================
'=��λ�õ�����ѡ��
'==============================================================================
Private Sub txtLocation_GotFocus()
    On Error GoTo ErrH
    Call zlControl.TxtSelAll(txtLocation)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ٶ�λ
'==============================================================================
Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo ErrH
    
    lngRow = 0
    If txtLocation.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        For lngLoop = vsfPTMZ.Row + 1 To vsfPTMZ.Rows - 1
            If InStr(UCase(vsfPTMZ.TextMatrix(lngLoop, vsfPTMZ.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '��ȡС�ڵ�ǰ�еļ�¼����
        If lngRow = 0 Then
            For lngLoop = 0 To vsfPTMZ.Row
                If InStr(UCase(vsfPTMZ.TextMatrix(lngLoop, vsfPTMZ.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If vsfPTMZ.Rows > 1 And lngRow >= 1 Then vsfPTMZ.Row = lngRow
        vsfPTMZ.ShowCell lngRow, vsfPTMZ.ColIndex(mstrFindKey)
        Call LocationObj(txtLocation)
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �����λ��¼ vsfPTMZ
'==============================================================================
Private Sub vsfPTMZ_AfterSort(ByVal COL As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
'    vsfSetRow vsfPTMZ, mstrSortID, "����ID"
    lngRow = vsfPTMZ.FindRow(mstrSortID, -1, vsfPTMZ.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfPTMZ.Row = lngRow
    vsfPTMZ.ShowCell lngRow, 1
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ǰ��¼����ID vsfPTMZ
'==============================================================================
Private Sub vsfPTMZ_BeforeSort(ByVal COL As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� vsfAuditItem
'==============================================================================
Private Sub vsfPTMZ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '�����˵�����
        
            Call SendLMouseButton(vsfPTMZ.hwnd, x, y)

            mcbrPopupBar.ShowPopup
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �ؼ���ʼ��
'==============================================================================
Private Sub InitControl()
    
    On Error GoTo ErrH
    
    Call InitCommandBar
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPTMZ_Resize()
    On Error Resume Next
    vsfPTMZ.Move 15, 15, picPTMZ.Width - 30, picPTMZ.Height - 30
End Sub

'==============================================================================
'=���ܣ� λ������
'==============================================================================
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_Edit_NewItem                       '������Ŀ
            Call NewPTMZ
        Case conMenu_Edit_Modify                        '�޸���Ŀ
            Call EditPTMZ
        Case conMenu_Edit_Delete                        'ɾ����Ŀ
            Call DeletePTMZ
        Case conMenu_Edit_Send                          '����
            Call UpdateCenter
        Case conMenu_Edit_SendBack                      '��������
            Call CancelUpdate
        Case conMenu_View_Find                          '��������
            Call FindPTMZ
        Case conMenu_File_Preview   'Ԥ��
            mzlPrintModeS = zlView
            Call ItemPrint
        Case conMenu_File_Print   '��ӡ
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel '�����&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
        Case conMenu_View_Forward
            Call ForwardPTMZ
        Case conMenu_View_Backward
            Call BackwardPTMZ
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txtLocation
        Case conMenu_View_Refresh               'ˢ��
            Call RefreshPTMZ
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '��ҵ���޹صĹ��ܣ������Ĺ���
                Call CommandBarExecutePublic(Control, Me, vsfPTMZ, "ְ����ͨ���ﱨ��")
            End If
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo ErrH

    With vsfPTMZ
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "�ϴ�"))
            Case conMenu_EditPopup
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
            Case conMenu_Edit_NewItem                    '������Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
            Case conMenu_Edit_Modify                        '�޸���Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "�ϴ�"))
            Case conMenu_Edit_Delete                  'ɾ��
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "�ϴ�"))
                If .Rows > 1 Then
                    Control.Enabled = .TextMatrix(.Row, .ColIndex("�Ƿ��ϴ�")) <> "1"
                End If
            Case conMenu_Edit_Send
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "�ϴ�"))
                If .Rows > 1 Then
                    Control.Enabled = .TextMatrix(.Row, .ColIndex("�Ƿ��ϴ�")) <> "1"
                End If
            Case conMenu_Edit_SendBack                      '��������
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
                Control.Enabled = ((.Rows > 1) And IsPrivs(mstrPrivs, "�ϴ�"))
                If .Rows > 1 Then
                    Control.Enabled = .TextMatrix(.Row, .ColIndex("�Ƿ��ϴ�")) = "1"
                End If
            Case conMenu_View_Refresh
                Control.Visible = IsPrivs(mstrPrivs, "�ϴ�")
            Case conMenu_View_Forward
                Control.Enabled = .Row > 1
            Case conMenu_View_Backward
                Control.Enabled = .Row + 1 < .Rows
            Case conMenu_View_Find, conMenu_View_Refresh
                Control.Enabled = True
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If

            Case Else
                Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ӡ ItemPrint
'==============================================================================
Private Sub ItemPrint()
    On Error GoTo ErrH
    subPrint (mzlPrintModeS)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim lngLoop         As Long
    Dim objControl      As Object
    Dim objPrint        As New zlPrint1Grd
    Dim objAppRow       As zlTabAppRow
    
    If vsfPTMZ Is Nothing Then Exit Sub
    LockWindowUpdate vsfPTMZ.hwnd
    vsfPTMZ.ColHidden(vsfPTMZ.ColIndex("ͼ��")) = True
    Call SearchPrintData(vsfPTMZ, frmPubResource.msfPrint)
    vsfPTMZ.ColHidden(vsfPTMZ.ColIndex("ͼ��")) = False
    LockWindowUpdate 0
    '���ô�ӡ��������
    Set objPrint.Body = frmPubResource.msfPrint
    objPrint.Title.Text = Me.Caption
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ�ˣ�" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ�䣺" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
    mstrPrivs = gstrPrivs
    Call InitControl
    If GetPersonSet Then
        'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
        RestoreWinState Me, App.ProductName
        RestoreFlexState vsfPTMZ, Me.Name
        Me.WindowState = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "����", 0)
        If Me.WindowState = 0 Then
            Me.Left = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "LEFT", Me.Left)
            Me.Top = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "TOP", Me.Top)
        End If
    End If
    gstrSQL = mstrPTMZ
    gstrSQL = gstrSQL & vbCrLf & " And nvl(A.Up,0)=0 "
    
    '��������
    Call DataLoadPTMZ(gstrSQL)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrH
    picPTMZ.Move Me.ScaleLeft, Me.ScaleTop + 800, Me.ScaleWidth, Me.ScaleHeight - stbThis.Height - 800
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DataLoadPTMZ(sSql As String)
    Dim strField        As String
    Dim strFieldWIDth   As String
    Dim varField        As Variant
    Dim varFieldWIDth   As Variant
    Dim i               As Integer
On Error GoTo ErrH

    Set mRsPTMZ = zlDatabase.OpenSQLRecord(sSql, Me.Caption)
    Set vsfPTMZ.DataSource = mRsPTMZ
    'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
    If GetPersonSet Then
        With vsfPTMZ
            strField = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrID", .Name & "����", "")
            strFieldWIDth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrID", .Name & "���", "")
            varField = Split(strField, ",")
            varFieldWIDth = Split(strFieldWIDth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" And Val(varFieldWIDth(i)) <> 0 Then
                    If .ColIndex(varField(i)) <> -1 Then
                         .ColPosition(.ColIndex(varField(i))) = i
                         .ColWidth(i) = Val(varFieldWIDth(i))
                    End If
                End If
            Next
        End With
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'����
Private Sub NewPTMZ()
    Dim str�Ǽ�id       As String
    Dim strWhere        As String
On Error GoTo ErrH
    With frmMain_�������󲡰���Ϣ�༭
        .Show vbModal
        str�Ǽ�id = .HospitalNumber
    End With
    Set frmMain_�������󲡰���Ϣ�༭ = Nothing
    If str�Ǽ�id = "" Then Exit Sub '���ȡ��
    'ˢ����ϸ����
    gstrSQL = mstrPTMZ
    strWhere = "And �Ǽ�ID='" & str�Ǽ�id & "'"
    Call DataLoadPTMZ(strWhere)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'�޸�
Private Sub EditPTMZ()
    Dim str�Ǽ�id       As String
    Dim strWhere        As String
On Error GoTo ErrH
    With frmMain_�������󲡰���Ϣ�༭
        str�Ǽ�id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
        .HospitalNumber = str�Ǽ�id
        .UpdateCenter = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("�Ƿ��ϴ�")) = "1"
        .Show vbModal
        str�Ǽ�id = .HospitalNumber
    End With
    Set frmMain_�������󲡰���Ϣ�༭ = Nothing
    If str�Ǽ�id = "" Then Exit Sub '���ȡ��
    'ˢ����ϸ����
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '���¶�λ
    vsfSetRow vsfPTMZ, str�Ǽ�id, "ID"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'ɾ��
Private Sub DeletePTMZ()
    Dim str�Ǽ�id       As String
    Dim strWhere        As String
On Error GoTo ErrH
    str�Ǽ�id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    If MsgBox("ȷ��ɾ��סԺ�š�" & str�Ǽ�id & "���Ĳ�����Ϣ��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    
    gstrSQL = "zl_���β�����Ϣ_Delete('" & str�Ǽ�id & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    'ˢ����ϸ����
    Call DataLoadPTMZ(mRsPTMZ.Source)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'ˢ��
Private Sub RefreshPTMZ()
    Dim str�Ǽ�id       As String
    Dim strErrMsg       As String
    Dim strWhere        As String
On Error GoTo ErrH
    If vsfPTMZ.Row > 1 Then
        str�Ǽ�id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    End If
    'ˢ����ϸ����
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '���¶�λ
    vsfSetRow vsfPTMZ, str�Ǽ�id, "ID"
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|RefreshPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'ǰһ��
Private Sub ForwardPTMZ()
    Dim strErrMsg       As String
On Error GoTo ErrH
    With vsfPTMZ
        If .Row > 1 Then
            .Row = .Row - 1
            .ShowCell .Row, .COL
        End If
    End With
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|ForwardPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'��һ��
Private Sub BackwardPTMZ()
    Dim strErrMsg       As String
On Error GoTo ErrH
    With vsfPTMZ
        If .Row < .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .COL
        End If
    End With
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|BackwardPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'��������
Private Sub FindPTMZ()
    Dim str�Ǽ�id       As String
    Dim strErrMsg       As String
    Dim strWhere        As String
On Error GoTo ErrH
    'ˢ����ϸ����
    With frmMain_�������󲡰���Ϣ����
        .Show vbModal
        strWhere = .strWhere
    End With
    Set frmMain_�������󲡰���Ϣ���� = Nothing
    If strWhere = "" Then Exit Sub

    Call DataLoadPTMZ(mstrPTMZ & strWhere)
    '���¶�λ
    vsfSetRow vsfPTMZ, str�Ǽ�id, "ID"
    Exit Sub
ErrH:
    strErrMsg = Me.Name & "|" & Me.Caption & "|FindPTMZ:" & vbCrLf & Err.Description
    MsgBox strErrMsg, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

'�ϴ�����������
Private Sub UpdateCenter()
    Dim str�Ǽ�id As String
    Dim rsTmp As ADODB.Recordset
    Dim cnTest As ADODB.Connection
    Dim strServer As String
    Dim strUser As String
    Dim strPwd As String
    Dim str����ֵ As String
    Dim strWhere As String
On Error GoTo ErrH
    str�Ǽ�id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    If MsgBox("ȷ���ϴ�סԺ�š�" & str�Ǽ�id & "���Ĳ�����Ϣ��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    '���Ӳ���������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_��������)
    Do Until rsTmp.EOF
        str����ֵ = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
        Select Case rsTmp("������")
            Case "�����û���"
                strUser = str����ֵ
            Case "�����û�����"
                strPwd = str����ֵ
            Case "����������"
                strServer = str����ֵ
        End Select
        rsTmp.MoveNext
    Loop
    
    Set cnTest = New ADODB.Connection
    If cnTest.State = adStateOpen Then cnTest.Close
    cnTest.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(strPwd) & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    cnTest.CursorLocation = adUseClient
    cnTest.Open
    If Err <> 0 Then
        MsgBox "��������������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '�ϴ�
    gstrSQL = "Select * From ���β�����Ϣ where RESIDENCE_NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�Ǽ�id)
    
   
'    With rsTmp
'        gstrSQL = "" & _
'        "INSERT INTO VIEW_MEDICAL_RECORD_INFO" & vbNewLine & _
'        "    (" & vbNewLine & _
'        "      HOSPITAL_NUMBER,RESIDENCE_NO,IN_COUNT,MEDICAL_RECORD_NO,MARITAL_STATUS,STATUS,BIRTH_ADDRESS,IDENTITY_NUMBER,UNIT_NAME,UNIT_ADDRESS," & vbNewLine & _
'        "      UNIT_PHONE,UNIT_ZIPCODE,REGISTER_ADDRESS,REGISTER_ZIPCODE,CONTACT_PERSON,RELATIONSHIP,CONTACT_ADDRESS,CONTACT_PHONE,ADMISSION_DATE,ADMISSION_DEPT," & vbNewLine & _
'        "      IN_DEPT_ZONE,DEPT_TRANSFERED_TO,DISCHARGE_DATE,DISCHARGE_DEPT,OUT_DEPT_ZONE,PAT_ADM_CONDITION,DIAGNOSIS_DATE,ALERGY_DRUGS,HBSAG,HCV_AB,HIV_AB," & vbNewLine & _
'        "      CLINIC_INHOSPITAL,IN_OUT,BEFORE_AFTER_TREATMENT,CLINIC_PATHOLOGY,EMIT_PATHOLOGY,EMER_TREAT_TIMES,ESC_EMER_TIMES,DIRECTOR,DIRECTOR_DOCTOR,ATTENDING_DOCTOR," & vbNewLine & _
'        "      INHOSPITAL_DOCTOR,REFRESH_DOCTOR,GRADUATE_DOCTOR,INTERM,CODE_NAME,MEDICAL_RECORD_MASS,CONTROL_DOCTOR,CONTROL_NURSE,BAL_DATE,BODY_EXAMINE_FLAG," & vbNewLine & _
'        "      FIRST_FLAG,FOLLOW_FLAG,FOLLOW_TERM,TEACH_MR_FLAG,BLOOD_TYPE,RH,BLOOD_TRAN_REACT_FLAG,ERYTHROCYTE,HEMOBLAST,PLASM," & vbNewLine & _
'        "      BLOOD,OTHER_BLOOD,HANDLE,HANDLE_DATE,IN_DIAGNOSIS_CODE,IN_DIAGNOSIS_NAME,IN_DIAGNOSIS_DATE,OUT_DIAGNOSIS_CODE1,OUT_DIAGNOSIS_NAME1,OUT_DIAGNOSIS_DATE1," & vbNewLine & _
'        "      TREAT_RESULT1,OUT_DIAGNOSIS_CODE2,OUT_DIAGNOSIS_NAME2,OUT_DIAGNOSIS_DATE2,TREAT_RESULT2,OUT_DIAGNOSIS_CODE3,OUT_DIAGNOSIS_NAME3,OUT_DIAGNOSIS_DATE3,TREAT_RESULT3,OPERATION_CODE1," & vbNewLine & _
'        "      OPERATION_NAME1,WOUND_GRADE1,HEAL1,OPERATING_DATE1,ANAESTHESIA_METHOD1,OPERATION_CODE2,OPERATION_NAME2,WOUND_GRADE2,HEAL2,OPERATING_DATE2," & vbNewLine & _
'        "      ANAESTHESIA_METHOD2,OPERATION_CODE3,OPERATION_NAME3,WOUND_GRADE3,HEAL3,OPERATING_DATE3,ANAESTHESIA_METHOD3" & vbNewLine & _
'        "      )"
'        gstrSQL = gstrSQL & vbNewLine & _
'        "VALUES" & vbNewLine & _
'        "  (" & vbNewLine & _
'        "   '" & !HOSPITAL_NUMBER & "' , '" & !RESIDENCE_NO & "' , '" & !IN_COUNT & "' , '" & !MEDICAL_RECORD_NO & "' , '" & !MARITAL_STATUS & "' , '" & !Status & "' , '" & !BIRTH_ADDRESS & "' , '" & !IDENTITY_NUMBER & "' , '" & !UNIT_NAME & "' , '" & !UNIT_ADDRESS & "' ," & vbNewLine & _
'        "   '" & !UNIT_PHONE & "' , '" & !UNIT_ZIPCODE & "' , '" & !REGISTER_ADDRESS & "' , '" & !REGISTER_ZIPCODE & "' , '" & !CONTACT_PERSON & "' , '" & !RELATIONSHIP & "' , '" & !CONTACT_ADDRESS & "' , '" & !CONTACT_PHONE & "' , to_date('" & !ADMISSION_DATE & "','yyyy-dd-mm hh24:mi:ss'), '" & !ADMISSION_DEPT & "' ," & vbNewLine & _
'        "   '" & !IN_DEPT_ZONE & "' , '" & !DEPT_TRANSFERED_TO & "' ,to_date( '" & !DISCHARGE_DATE & "','yyyy-dd-mm hh24:mi:ss'), '" & !DISCHARGE_DEPT & "' , '" & !OUT_DEPT_ZONE & "' , '" & !PAT_ADM_CONDITION & "' ,to_date( '" & !DIAGNOSIS_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !ALERGY_DRUGS & "' , '" & !HBSAG & "' , '" & !HCV_AB & "' , '" & !HIV_AB & "' ," & vbNewLine & _
'        "   '" & !CLINIC_INHOSPITAL & "' , '" & !IN_OUT & "' , '" & !BEFORE_AFTER_TREATMENT & "' , '" & !CLINIC_PATHOLOGY & "' , '" & !EMIT_PATHOLOGY & "' , '" & !EMER_TREAT_TIMES & "' , '" & !ESC_EMER_TIMES & "' , '" & !DIRECTOR & "' , '" & !DIRECTOR_DOCTOR & "' , '" & !ATTENDING_DOCTOR & "' ," & vbNewLine & _
'        "   '" & !INHOSPITAL_DOCTOR & "' , '" & !REFRESH_DOCTOR & "' , '" & !GRADUATE_DOCTOR & "' , '" & !INTERM & "' , '" & !CODE_NAME & "' , '" & !MEDICAL_RECORD_MASS & "' , '" & !CONTROL_DOCTOR & "' , '" & !CONTROL_NURSE & "' , to_date('" & !BAL_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !BODY_EXAMINE_FLAG & "' ," & vbNewLine & _
'        "   '" & !FIRST_FLAG & "' , '" & !FOLLOW_FLAG & "' , '" & !FOLLOW_TERM & "' , '" & !TEACH_MR_FLAG & "' , '" & !BLOOD_TYPE & "' , '" & !RH & "' , '" & !BLOOD_TRAN_REACT_FLAG & "' , '" & !ERYTHROCYTE & "' , '" & !HEMOBLAST & "' , '" & !PLASM & "' ," & vbNewLine & _
'        "   '" & !BLOOD & "' , '" & !OTHER_BLOOD & "' , '" & !Handle & "' , to_date('" & !HANDLE_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !IN_DIAGNOSIS_CODE & "' , '" & !IN_DIAGNOSIS_NAME & "' ,to_date( '" & !IN_DIAGNOSIS_DATE & "','yyyy-dd-mm hh24:mi:ss') , '" & !OUT_DIAGNOSIS_CODE1 & "' , '" & !OUT_DIAGNOSIS_NAME1 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE1 & "' ,'yyyy-dd-mm hh24:mi:ss')," & vbNewLine & _
'        "   '" & !TREAT_RESULT1 & "' , '" & !OUT_DIAGNOSIS_CODE2 & "' , '" & !OUT_DIAGNOSIS_NAME2 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE2 & "','yyyy-dd-mm hh24:mi:ss') , '" & !TREAT_RESULT2 & "' , '" & !OUT_DIAGNOSIS_CODE3 & "' , '" & !OUT_DIAGNOSIS_NAME3 & "' , to_date('" & !OUT_DIAGNOSIS_DATE3 & "','yyyy-dd-mm hh24:mi:ss') , '" & !TREAT_RESULT3 & "' , '" & !OPERATION_CODE1 & "' ," & vbNewLine & _
'        "   '" & !OPERATION_NAME1 & "' , '" & !WOUND_GRADE1 & "' , '" & !HEAL1 & "' , to_date('" & !OPERATING_DATE1 & "','yyyy-dd-mm hh24:mi:ss') , '" & !ANAESTHESIA_METHOD1 & "' , '" & !OPERATION_CODE2 & "' , '" & !OPERATION_NAME2 & "' , '" & !WOUND_GRADE2 & "' , '" & !HEAL2 & "' , to_date('" & !OPERATING_DATE2 & "','yyyy-dd-mm hh24:mi:ss') ," & vbNewLine & _
'        "   '" & !ANAESTHESIA_METHOD2 & "' , '" & !OPERATION_CODE3 & "' , '" & !OPERATION_NAME3 & "' , '" & !WOUND_GRADE3 & "' , '" & !HEAL3 & "' , to_date('" & !OPERATING_DATE3 & "','yyyy-dd-mm hh24:mi:ss'), '" & !ANAESTHESIA_METHOD3 & "'" & vbNewLine & _
'        "  )"
'    End With

With rsTmp
        gstrSQL = "" & _
        "INSERT INTO VIEW_MEDICAL_RECORD_INFO" & vbNewLine & _
        "    (" & vbNewLine & _
        "      HOSPITAL_NUMBER,RESIDENCE_NO,IN_COUNT,MEDICAL_RECORD_NO,MARITAL_STATUS,STATUS,BIRTH_ADDRESS,IDENTITY_NUMBER,UNIT_NAME,UNIT_ADDRESS," & vbNewLine & _
        "      UNIT_PHONE,UNIT_ZIPCODE,REGISTER_ADDRESS,REGISTER_ZIPCODE,CONTACT_PERSON,RELATIONSHIP,CONTACT_ADDRESS,CONTACT_PHONE,ADMISSION_DATE,ADMISSION_DEPT," & vbNewLine & _
        "      IN_DEPT_ZONE,DEPT_TRANSFERED_TO,DISCHARGE_DATE,DISCHARGE_DEPT,OUT_DEPT_ZONE,PAT_ADM_CONDITION,DIAGNOSIS_DATE,ALERGY_DRUGS,HBSAG,HCV_AB,HIV_AB," & vbNewLine & _
        "      CLINIC_INHOSPITAL,IN_OUT,BEFORE_AFTER_TREATMENT,CLINIC_PATHOLOGY,EMIT_PATHOLOGY,EMER_TREAT_TIMES,ESC_EMER_TIMES,DIRECTOR,DIRECTOR_DOCTOR,ATTENDING_DOCTOR," & vbNewLine & _
        "      INHOSPITAL_DOCTOR,REFRESH_DOCTOR,GRADUATE_DOCTOR,INTERM,CODE_NAME,MEDICAL_RECORD_MASS,CONTROL_DOCTOR,CONTROL_NURSE,BAL_DATE,BODY_EXAMINE_FLAG," & vbNewLine & _
        "      FIRST_FLAG,FOLLOW_FLAG,FOLLOW_TERM,TEACH_MR_FLAG,BLOOD_TYPE,RH,BLOOD_TRAN_REACT_FLAG,ERYTHROCYTE,HEMOBLAST,PLASM," & vbNewLine & _
        "      BLOOD,OTHER_BLOOD,HANDLE,HANDLE_DATE,IN_DIAGNOSIS_CODE,IN_DIAGNOSIS_NAME,IN_DIAGNOSIS_DATE,OUT_DIAGNOSIS_CODE1,OUT_DIAGNOSIS_NAME1,OUT_DIAGNOSIS_DATE1," & vbNewLine & _
        "      TREAT_RESULT1,OUT_DIAGNOSIS_CODE2,OUT_DIAGNOSIS_NAME2,OUT_DIAGNOSIS_DATE2,TREAT_RESULT2,OUT_DIAGNOSIS_CODE3,OUT_DIAGNOSIS_NAME3,OUT_DIAGNOSIS_DATE3,TREAT_RESULT3,OPERATION_CODE1," & vbNewLine & _
        "      OPERATION_NAME1,WOUND_GRADE1,HEAL1,OPERATING_DATE1,ANAESTHESIA_METHOD1,OPERATION_CODE2,OPERATION_NAME2,WOUND_GRADE2,HEAL2,OPERATING_DATE2," & vbNewLine & _
        "      ANAESTHESIA_METHOD2,OPERATION_CODE3,OPERATION_NAME3,WOUND_GRADE3,HEAL3,OPERATING_DATE3,ANAESTHESIA_METHOD3" & vbNewLine & _
        "      )"
        gstrSQL = gstrSQL & vbNewLine & _
        "VALUES" & vbNewLine & _
        "  (" & vbNewLine & _
        "   '" & !HOSPITAL_NUMBER & "' , '" & !RESIDENCE_NO & "' , '" & !IN_COUNT & "' , '" & !MEDICAL_RECORD_NO & "' , '" & !MARITAL_STATUS & "' , '" & !Status & "' , '" & !BIRTH_ADDRESS & "' , '" & !IDENTITY_NUMBER & "' , '" & !UNIT_NAME & "' , '" & !UNIT_ADDRESS & "' ," & vbNewLine & _
        "   '" & !UNIT_PHONE & "' , '" & !UNIT_ZIPCODE & "' , '" & !REGISTER_ADDRESS & "' , '" & !REGISTER_ZIPCODE & "' , '" & !CONTACT_PERSON & "' , '" & !RELATIONSHIP & "' , '" & !CONTACT_ADDRESS & "' , '" & !CONTACT_PHONE & "' , to_date('" & !ADMISSION_DATE & "','yyyy-mm-dd hh24:mi:ss'), '" & !ADMISSION_DEPT & "' ," & vbNewLine & _
        "   '" & !IN_DEPT_ZONE & "' , '" & !DEPT_TRANSFERED_TO & "' ,to_date( '" & !DISCHARGE_DATE & "','yyyy-mm-dd hh24:mi:ss'), '" & !DISCHARGE_DEPT & "' , '" & !OUT_DEPT_ZONE & "' , '" & !PAT_ADM_CONDITION & "' ,to_date( '" & !DIAGNOSIS_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !ALERGY_DRUGS & "' , '" & !HBSAG & "' , '" & !HCV_AB & "' , '" & !HIV_AB & "' ," & vbNewLine & _
        "   '" & !CLINIC_INHOSPITAL & "' , '" & !IN_OUT & "' , '" & !BEFORE_AFTER_TREATMENT & "' , '" & !CLINIC_PATHOLOGY & "' , '" & !EMIT_PATHOLOGY & "' , '" & !EMER_TREAT_TIMES & "' , '" & !ESC_EMER_TIMES & "' , '" & !DIRECTOR & "' , '" & !DIRECTOR_DOCTOR & "' , '" & !ATTENDING_DOCTOR & "' ," & vbNewLine & _
        "   '" & !INHOSPITAL_DOCTOR & "' , '" & !REFRESH_DOCTOR & "' , '" & !GRADUATE_DOCTOR & "' , '" & !INTERM & "' , '" & !CODE_NAME & "' , '" & !MEDICAL_RECORD_MASS & "' , '" & !CONTROL_DOCTOR & "' , '" & !CONTROL_NURSE & "' , to_date('" & !BAL_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !BODY_EXAMINE_FLAG & "' ," & vbNewLine & _
        "   '" & !FIRST_FLAG & "' , '" & !FOLLOW_FLAG & "' , '" & !FOLLOW_TERM & "' , '" & !TEACH_MR_FLAG & "' , '" & !BLOOD_TYPE & "' , '" & !RH & "' , '" & !BLOOD_TRAN_REACT_FLAG & "' , '" & !ERYTHROCYTE & "' , '" & !HEMOBLAST & "' , '" & !PLASM & "' ," & vbNewLine & _
        "   '" & !BLOOD & "' , '" & !OTHER_BLOOD & "' , '" & !Handle & "' , to_date('" & !HANDLE_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !IN_DIAGNOSIS_CODE & "' , '" & !IN_DIAGNOSIS_NAME & "' ,to_date( '" & !IN_DIAGNOSIS_DATE & "','yyyy-mm-dd hh24:mi:ss') , '" & !OUT_DIAGNOSIS_CODE1 & "' , '" & !OUT_DIAGNOSIS_NAME1 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE1 & "' ,'yyyy-mm-dd hh24:mi:ss')," & vbNewLine & _
        "   '" & !TREAT_RESULT1 & "' , '" & !OUT_DIAGNOSIS_CODE2 & "' , '" & !OUT_DIAGNOSIS_NAME2 & "' ,to_date( '" & !OUT_DIAGNOSIS_DATE2 & "','yyyy-mm-dd hh24:mi:ss') , '" & !TREAT_RESULT2 & "' , '" & !OUT_DIAGNOSIS_CODE3 & "' , '" & !OUT_DIAGNOSIS_NAME3 & "' , to_date('" & !OUT_DIAGNOSIS_DATE3 & "','yyyy-mm-dd hh24:mi:ss') , '" & !TREAT_RESULT3 & "' , '" & !OPERATION_CODE1 & "' ," & vbNewLine & _
        "   '" & !OPERATION_NAME1 & "' , '" & !WOUND_GRADE1 & "' , '" & !HEAL1 & "' , to_date('" & !OPERATING_DATE1 & "','yyyy-mm-dd hh24:mi:ss') , '" & !ANAESTHESIA_METHOD1 & "' , '" & !OPERATION_CODE2 & "' , '" & !OPERATION_NAME2 & "' , '" & !WOUND_GRADE2 & "' , '" & !HEAL2 & "' , to_date('" & !OPERATING_DATE2 & "','yyyy-mm-dd hh24:mi:ss') ," & vbNewLine & _
        "   '" & !ANAESTHESIA_METHOD2 & "' , '" & !OPERATION_CODE3 & "' , '" & !OPERATION_NAME3 & "' , '" & !WOUND_GRADE3 & "' , '" & !HEAL3 & "' , to_date('" & !OPERATING_DATE3 & "','yyyy-mm-dd hh24:mi:ss'), '" & !ANAESTHESIA_METHOD3 & "'" & vbNewLine & _
        "  )"
    End With
 

    cnTest.Execute gstrSQL
    '���±��ر�ʶ
    gstrSQL = "zl_���β�����Ϣ_UpServer('" & str�Ǽ�id & "','1','" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    'ˢ����ϸ����
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '���¶�λ
    vsfSetRow vsfPTMZ, str�Ǽ�id, "ID"
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

Private Sub CancelUpdate()
    Dim str�Ǽ�id As String
    Dim rsTmp As ADODB.Recordset
    Dim cnTest As ADODB.Connection
    Dim strServer As String
    Dim strUser As String
    Dim strPwd As String
    Dim str����ֵ As String
    Dim strWhere As String
On Error GoTo ErrH
    str�Ǽ�id = vsfPTMZ.TextMatrix(vsfPTMZ.Row, vsfPTMZ.ColIndex("ID"))
    If MsgBox("ȷ�ϳ����ϴ�סԺ�š�" & str�Ǽ�id & "���Ĳ�����Ϣ��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    '���Ӳ���������
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_��������)
    Do Until rsTmp.EOF
        str����ֵ = IIf(IsNull(rsTmp("����ֵ")), "", rsTmp("����ֵ"))
        Select Case rsTmp("������")
            Case "�����û���"
                strUser = str����ֵ
            Case "�����û�����"
                strPwd = str����ֵ
            Case "����������"
                strServer = str����ֵ
        End Select
        rsTmp.MoveNext
    Loop
    
    Set cnTest = New ADODB.Connection
    If cnTest.State = adStateOpen Then cnTest.Close
    cnTest.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(strPwd) & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    cnTest.CursorLocation = adUseClient
    cnTest.Open
    If Err <> 0 Then
        MsgBox "��������������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '�ϴ�
    gstrSQL = "Select * From VIEW_MEDICAL_RECORD_INFO where RESIDENCE_NO='" & str�Ǽ�id & "'"
    Set rsTmp = cnTest.Execute(gstrSQL)
    If Not ChkRsState(rsTmp) Then
        MsgBox "����סԺ�š�" & str�Ǽ�id & "��δ���������ز��ܳ����ϴ���", vbCritical, gstrSysName
        Exit Sub
    End If
    
    '���±��ر�ʶ
    gstrSQL = "zl_���β�����Ϣ_UpServer('" & str�Ǽ�id & "','0','" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    'ˢ����ϸ����
    Call DataLoadPTMZ(mRsPTMZ.Source)
    '���¶�λ
    vsfSetRow vsfPTMZ, str�Ǽ�id, "ID"
    Exit Sub
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub
