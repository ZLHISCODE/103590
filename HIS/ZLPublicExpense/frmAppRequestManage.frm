VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.dockingpane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppRequestManage 
   Caption         =   "ԤԼ�Ǽǹ���"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmAppRequestManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   11745
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7995
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      SimpleText      =   $"frmAppRequestManage.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppRequestManage.frx":05D1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   435
      Top             =   525
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAppRequestManage.frx":0E65
      Left            =   975
      Top             =   585
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppRequestManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As frmAppRequestMain
Private mfrmFilter As frmAppRequestFilter
Private mlngFaceBackColor As Long

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Select Case Control.ID
    Case conMenu_Edit_CancelRequest
        blnEnable = True
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then
            blnEnable = False
        Else
            If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then blnEnable = False
        End If
        Control.Enabled = blnEnable
    Case conMenu_Edit_ViewRequest
        blnEnable = True
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then
            blnEnable = False
        Else
            If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then blnEnable = False
        End If
        Control.Enabled = blnEnable
    End Select
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandle
    Call DefMainCommandBars
    Call InitPanel '��ʼ��dkpMain
    
    mlngFaceBackColor = cbsThis.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
    Me.BackColor = mlngFaceBackColor
    RestoreWinState mfrmMain, "frmAppRequestMain"
    RestoreWinState Me, "frmAppRequestManage"
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub zlDataPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim objVsf As VSFlexGrid
    
    Err = 0: On Error GoTo errHandle

    objOut.Title.Text = "ԤԼ�ǼǼ�¼���"
    Set objVsf = gobjControl.RPTCopyToVSF(mfrmMain.rptMain, objVsf)
    Set objOut.Body = objVsf
    
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Now 'Format(sys.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    
    Err = 0: On Error GoTo errHandle
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_View_StatusBar
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        Control.Checked = Not Control.Checked
        cbsThis(2).Visible = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not Control.Checked
        For Each objControl In cbsThis(2).Controls
            objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Control.Checked = Not Control.Checked
        cbsThis.Options.LargeIcons = Control.Checked
        cbsThis.RecalcLayout
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call gobjComlib.zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call gobjComlib.zlMailTo(Me.hWnd)
    Case conMenu_Help_About: Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Edit_AppRequest
        frmAppRequestEdit.ShowMe Me
        Call mfrmMain.RefreshData
    Case conMenu_View_Refresh
        Call mfrmMain.RefreshData
    Case conMenu_Edit_CancelRequest
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
        Call CancelRequest
        Call mfrmMain.RefreshData
    Case conMenu_Edit_ViewRequest
        If mfrmMain.rptMain.SelectedRows.Count = 0 Then Exit Sub
        If mfrmMain.rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
        Call frmAppRequestEdit.ReadBill(Me, Val(mfrmMain.rptMain.SelectedRows.Row(0).Record.Tag))
    Case Else
    End Select
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CancelRequest()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select 1 From ���˷�����Ϣ��¼ Where ID=[1] And ����ʱ�� Is Null"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mfrmMain.rptMain.SelectedRows.Row(0).Record.Tag))
    If rsTemp.EOF Then
        MsgBox "��ǰԤԼ�ǼǼ�¼�Ѿ�������,�޷�ȡ���Ǽ�!", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("�Ƿ�ȷ��ȡ������ԤԼ�ǼǼ�¼?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
    strSQL = "zl_���߷�������_����("
    strSQL = strSQL & mfrmMain.rptMain.SelectedRows.Row(0).Record.Tag & ",'"
    strSQL = strSQL & "ȡ���Ǽ�','"
    strSQL = strSQL & UserInfo.���� & "','"
    strSQL = strSQL & UserInfo.��� & "',"
    strSQL = strSQL & "Null,"
    strSQL = strSQL & 1 & ")"
    Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_SpecialColorChanged()
    Me.BackColor = cbsThis.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Public Sub RefreshRecord()
    With mfrmMain
        .mbln�Ǽ�ʱ�� = mfrmFilter.chkDate(0).Value
        .mbln����ʱ�� = mfrmFilter.chkDate(1).Value
        .mbln��ʾ���� = mfrmFilter.chkShowSet.Value
        .mdat����ʼ = mfrmFilter.dtpBegin(1).Value
        .mdat������� = mfrmFilter.dtpEnd(1).Value
        .mdat��ʼʱ�� = mfrmFilter.dtpBegin(0).Value
        .mdat����ʱ�� = mfrmFilter.dtpEnd(0).Value
        .mstr������ = NeedName(mfrmFilter.cbo������.Text)
        .mstr�Ǽ��� = NeedName(mfrmFilter.cbo�Ǽ���.Text)
        .mbyt���﷽ʽ = mfrmFilter.cbo���﷽ʽ.ListIndex
    End With
    Call mfrmMain.RefreshData
End Sub


Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo errHandle
    Set cbsThis.Icons = gobjCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Edit, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_Edit
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AppRequest, "ԤԼ�Ǽ�(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CancelRequest, "ȡ���Ǽ�(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ViewRequest, "�鿴�Ǽ�(&V)"):  cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        cbrSubControl.Checked = True
        Set cbrSubControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        cbrSubControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AppRequest, "ԤԼ�Ǽ�(&A)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 3003
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CancelRequest, "ȡ���Ǽ�(&C)")
        cbrControl.IconId = 3004
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyC, conMenu_Edit_ClsAll
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    DefMainCommandBars = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 230, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    Set mfrmFilter = New frmAppRequestFilter
    objPane.Handle = mfrmFilter.hWnd
    objPane.MaxTrackSize.Width = 265
    objPane.MinTrackSize.Width = 265
    mfrmFilter.SetForm Me
    
    Set objPane = dkpMain.CreatePane(2, 230, 300, DockRightOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    Set mfrmMain = New frmAppRequestMain
    objPane.Handle = mfrmMain.hWnd
    
    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState mfrmMain, "frmAppRequestMain"
    SaveWinState Me, "frmAppRequestManage"
End Sub
