VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPatinetAuditing 
   Caption         =   "�������"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   Icon            =   "frmPatinetAuditing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPatientList 
      Height          =   5265
      Left            =   360
      ScaleHeight     =   5205
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   870
      Width           =   4125
      Begin XtremeReportControl.ReportControl rptPatientList 
         Height          =   3945
         Left            =   750
         TabIndex        =   1
         Top             =   780
         Width           =   2415
         _Version        =   589884
         _ExtentX        =   4260
         _ExtentY        =   6959
         _StockProps     =   0
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cboʱ�� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4890
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   4890
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   149815297
         CurrentDate     =   40049
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   300
         Left            =   2790
         TabIndex        =   4
         Top             =   4890
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   149815297
         CurrentDate     =   40049
      End
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":6852
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":6DEC
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":7386
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":7920
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":7EBA
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":8454
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":87EE
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":8B88
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":8F22
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":92BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":FB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":16380
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":1CBE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":23444
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatinetAuditing.frx":29CA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   1980
      Top             =   270
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatinetAuditing.frx":30508
      Left            =   1260
      Top             =   270
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatinetAuditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mfrmWrite As frmLisStationWrite                  '������д����
Attribute mfrmWrite.VB_VarHelpID = -1
Private mlngPatienID As Long                                        '����ID
Private mstrPrivs   As String                                       'Ȩ��
Private mstrAuditingMan As String                                   '�����,���ʱ�������
Private Enum mCol
    ����ID
    ����
    �Ա�
    ����
    ����
    ����
    ��Դ
    ��ҳID
    ��ʶ��
    ��������
    ��λ
End Enum

Private Sub CreateCbs()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False
    

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
'    Me.cbrthis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&T)��"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    
    
    'conMenu_EditPopup
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&E)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������(&A)")
    End With



'    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

   

    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_Edit_Audit
        
        
    End With

    '���ò����ò˵�
'    With Me.cbrthis.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'        .AddHiddenCommand conMenu_File_Excel
'        .AddHiddenCommand conMenu_View_Jump
'        .AddHiddenCommand conMenu_View_Refresh
'    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�󱨸�"): cbrControl.BeginGroup = True
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    With cboʱ��
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "������"
        .AddItem "��  ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "�Զ���"
    End With
    cboʱ��.Text = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0)
End Sub
Private Sub CreateDockPane()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    
    Set mfrmWrite = New frmLisStationWrite                          '������д����
    mfrmWrite.mblnPatientFind = True
    mfrmWrite.fraComment.Tag = "����ʾ"
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(1, 150, 150, DockLeftOf, Nothing)
    Pane1.Title = "�����б�"
    Pane1.Handle = Me.picPatientList.hWnd
    Pane1.Options = PaneNoHideable Or PaneNoCloseable Or PaneNoFloatable

    Set Pane2 = dkpMain.CreatePane(2, 400, 600, DockRightOf, Nothing)
    Pane2.Title = "�����Ϣ"
    Pane2.Handle = mfrmWrite.hWnd
    Pane2.Options = PaneNoHideable Or PaneNoCloseable Or PaneNoFloatable
    
    Pane1.Select
    mfrmWrite.fraComment.Tag = "����ʾ"
    mfrmWrite.mblnPatientFind = True
End Sub

Private Sub cboʱ��_Click()
    zlDatabase.SetPara "�걾��Χ", cboʱ��.Text & ";" & Me.DTPDate & ";" & Me.dtpDateEnd, 100, 1208
    Me.DTPDate.Visible = (Me.cboʱ��.Text = "�Զ���")
    Me.dtpDateEnd.Visible = (Me.cboʱ��.Text = "�Զ���")
    'ˢ��
    If Me.Visible = True Then
        Call RefreshDate
    End If
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
        '------------------------------------------------------------------------------------------
        Case conMenu_File_PrintSet                                              '��ӡ����
        Case conMenu_File_Preview                                               'Ԥ��
        Case conMenu_File_Print                                                 '��ӡ
        Case conMenu_File_Exit                                                  '�˳�
            Unload Me
        '------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button                                                '��׼��ť
            Control.Checked = Not Control.Checked
            Me.cbrthis(2).Visible = Control.Checked
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                                  '�ı���ǩ
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Size                                                  '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
        
        
        '------------------------------------------------------------------------------------------
        Case conMenu_Edit_Audit                                                 '���
            If Not Me.rptPatientList.FocusedRow Is Nothing Then
                Call AuditingPatient(Me.rptPatientList.FocusedRow.Record(mCol.����ID).Value)
                Call RefreshDate
            End If
            
        '------------------------------------------------------------------------------------------
        Case conMenu_View_Refresh                                               'ˢ��
            RefreshDate
        '------------------------------------------------------------------------------------------
        Case conMenu_Help_Help                                                          '��������
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_Web                                                           'WEB�ϵ�
            Call zlHomePage(hWnd)
        
        Case conMenu_Help_Web_Home                                                      '��ҳ
            Call zlHomePage(Me.hWnd)
        
        Case conMenu_Help_Web_Mail                                                      '���ͷ���
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_Help_About                                                         '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button                                                    '��ʾ������
            Control.Checked = Me.cbrthis(2).Visible
        Case conMenu_View_ToolBar_Text                                                      '�Ƿ���ʾ����
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size                                                      '�Ƿ���ʾ��ͼ��
            Control.Checked = Me.cbrthis.Options.LargeIcons
    End Select
End Sub

Private Sub dkpMain_Resize()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Me.cbrthis.RecalcLayout
End Sub

Private Sub dtpDate_Change()
    zlDatabase.SetPara "�걾��Χ", cboʱ��.Text & ";" & Me.DTPDate & ";" & Me.dtpDateEnd, 100, 1208
    Call RefreshDate
End Sub

Private Sub dtpDateEnd_Change()
    zlDatabase.SetPara "�걾��Χ", cboʱ��.Text & ";" & Me.DTPDate & ";" & Me.dtpDateEnd, 100, 1208
    Call RefreshDate
End Sub

Private Sub Form_Load()
    Call CreateCbs
    Call CreateDockPane
    Call CreaterptListHead
    
    DTPDate.Value = Now
    dtpDateEnd.Value = Now
    
    
    Call RefreshDate
    
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.dkpMain.RecalcLayout
    Call mfrmWrite.zlRefreshPatient(mlngPatienID)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Unload mfrmWrite
End Sub

Private Sub picPatientList_Resize()
    On Error Resume Next
    Me.rptPatientList.Top = 10
    Me.rptPatientList.Left = 10
    Me.rptPatientList.Width = Me.picPatientList.ScaleWidth - 20
    Me.rptPatientList.Height = Me.picPatientList.ScaleHeight - Me.cboʱ��.Height - 40
    
    With Me.cboʱ��
        .Top = Me.picPatientList.ScaleHeight - .Height - 20
        .Left = 0
    End With
    With Me.DTPDate
        .Top = Me.cboʱ��.Top
        .Left = Me.cboʱ��.Left + Me.cboʱ��.Width + 20
    End With
    With Me.dtpDateEnd
        .Top = Me.cboʱ��.Top
        .Left = Me.DTPDate.Left + Me.DTPDate.Width + 20
    End With
End Sub
Private Sub CreaterptListHead()
    Dim Column As ReportColumn
    Dim i As Integer
    With Me.rptPatientList.Columns
        
        rptPatientList.AllowColumnRemove = False
        rptPatientList.ShowItemsInGroups = False
        rptPatientList.SetImageList Imglist
        With rptPatientList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        Set Column = .Add(mCol.����ID, "����ID", 45, False): Column.Visible = False
        Set Column = .Add(mCol.����, "����", 55, True)
        Set Column = .Add(mCol.�Ա�, "�Ա�", 40, True)
        Set Column = .Add(mCol.����, "����", 40, True)
        Set Column = .Add(mCol.��Դ, "��Դ", 40, True)
        Set Column = .Add(mCol.��ʶ��, "��ʶ��", 55, True)
        Set Column = .Add(mCol.����, "����", 40, True)
        Set Column = .Add(mCol.����, "����", 75, True)
        Set Column = .Add(mCol.��ҳID, "��ҳID", 55, True): Column.Visible = False
        Set Column = .Add(mCol.��������, "��������", 75, True)
        Set Column = .Add(mCol.��λ, "��λ", 120, True)
        
    End With
End Sub
Private Sub RefreshDate()
    'ˢ�²����б�����
    Dim rsTmp As New adodb.Recordset
    Dim strSQL As String
    Dim strStart As String
    Dim strEnd As String
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim blnSelect As Boolean
    
    On Error GoTo errH
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 2)
    
    If strStart = "�Զ���" Then
        strStart = Format(Me.DTPDate, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    End If
    
    
    strSQL = "Select distinct ����id, ����, �Ա�, ����, decode(������Դ,1,'����',2,'סԺ',3,'����',4,'����','����') as ������Դ, ��ʶ��, ����,b.���� as ���˿���, ��ҳid, ��������," & vbNewLine & _
            "Decode(a.����״̬, 1, '������', 2, '�Ѽ���') As ִ��״̬,decode(p.��Ŀ,'��������',p.����,'') as ��λ " & vbNewLine & _
            "From ����걾��¼ a,���ű� b,����ҽ������ P " & vbNewLine & _
            "Where ����ʱ�� between [1] And [2] " & vbNewLine & _
            " and a.�������ID =  B.ID(+) and a.ҽ��ID is not null And ����� is null and ΢����걾 is null and a.ҽ��id = P.ҽ��ID(+) "
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(strStart), CDate(strEnd))
    
    If Not Me.rptPatientList.FocusedRow Is Nothing Then intIndex = Me.rptPatientList.FocusedRow.Index
    
    Me.rptPatientList.Records.DeleteAll
    mfrmWrite.zlRefreshPatient (-1)
    
    Do While Not rsTmp.EOF
        With Me.rptPatientList
            Set Record = .Records.Add
            For intLoop = 0 To .Columns.Count - 1
                Record.AddItem ""
            Next
'            If rsTmp("ִ��״̬") = "�Ѽ���" Then
'                Record.Item(mCol.ִ��״̬).Value = "�Ѽ���"
'                Record.Item(mCol.ִ��״̬).Icon = 7
'            End If
            Record(mCol.����ID).Value = rsTmp("����ID")
            Record(mCol.����).Value = Nvl(rsTmp("����"))
            Record(mCol.����).Value = Nvl(rsTmp("����"))
            Record(mCol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
            Record(mCol.����).Value = Nvl(rsTmp("����"))
            Record(mCol.��Դ).Value = Nvl(rsTmp("������Դ"))
            Record(mCol.��ʶ��).Value = Nvl(rsTmp("��ʶ��"))
            Record(mCol.����).Value = Nvl(rsTmp("���˿���"))
            Record(mCol.��ҳID).Value = Nvl(rsTmp("��ҳid"))
            Record(mCol.��������).Value = Nvl(rsTmp("��������"))
            
'            If Nvl(rsTmp("��Ŀ")) = "��������" Then
                Record.Item(mCol.��λ).Value = Nvl(rsTmp("��λ"))
'            End If
            
            If mlngPatienID = rsTmp("����ID") Then
                blnSelect = True
                intIndex = Record.Index
            End If
        End With
        rsTmp.MoveNext
    Loop
    
    Me.rptPatientList.Populate
    
    If Me.rptPatientList.Rows.Count > 0 Then
        If blnSelect = True Then
            Me.rptPatientList.FocusedRow = Me.rptPatientList.Rows(intIndex)
        Else
            If intIndex >= Me.rptPatientList.Rows.Count Then
                Me.rptPatientList.FocusedRow = Me.rptPatientList.Rows(Me.rptPatientList.Rows.Count - 1)
            Else
                Me.rptPatientList.FocusedRow = Me.rptPatientList.Rows(intIndex)
            End If
            mlngPatienID = Me.rptPatientList.FocusedRow.Record(mCol.����ID).Value
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptPatientList_SelectionChanged()
    If Me.rptPatientList.FocusedRow Is Nothing Then Exit Sub
    mlngPatienID = Me.rptPatientList.FocusedRow.Record(mCol.����ID).Value
    mfrmWrite.zlRefreshPatient (mlngPatienID)
End Sub
Private Function AuditingPatient(lngPatientID As Long) As Boolean
    '----------------------------------------
    '����   ������Ϊ��λ�������
    '����   lngPatientID=����ID
    '--------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New adodb.Recordset
    Dim rs As New adodb.Recordset
    Dim lngKey As Long
    Dim strStart As String, strEnd As String
    Dim intPrivacy As Integer
    Dim blnRollBack As Boolean
    Dim strErrInfo As String
    Dim astrSQL() As String
    Dim intLoop As Integer
    ReDim astrSQL(0)
    
    If InStr(1, mstrPrivs, "��˱걾") <= 0 Then
        'û��Ȩ�޺������û���½ʱ�˳�
        MsgBox "��û��Ȩ�޽������,�����µ�½���������Ա�������!", vbInformation, gstrSysName
        Exit Function
    End If

            
    '11210 Ȩ�ޡ�δ�շ���ˡ�������˵�������ʱ��δ��Ч��
    If InStr(mstrPrivs, "δ�շ����") <= 0 Then
        If CheckChargeState(mlngKey, False) = False Then
            MsgBox "����δ�շѣ����ܽ�����ˣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(1), 2)
    
     If strStart = "�Զ���" Then
        strStart = Format(Me.DTPDate.Value, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd.Value, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    End If
    
    intPrivacy = zlDatabase.GetPara("���浥�Ƿ���ʾ��˽��Ŀ", 100, 1208, 0)
    
    strSQL = "select id,������ from ����걾��¼ where ����id = [1] and ����ʱ�� between [2] and [3] and ҽ��id is not null and ����� is null and ΢����걾 is null "
        
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID, CDate(strStart), CDate(strEnd))
    
    
    Do While Not rs.EOF
        lngKey = rs("ID")
        '21137 �ѹ鵵���治�����
        gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKey)
        If rsTmp.EOF = False Then
            MsgBox "���˱���סԺ�Ĳ������ύ��飬���ܽ�����ˣ�", vbInformation, Me.Caption
'            gcnOracle.RollbackTrans
'            blnRollBack = False
            Exit Function
        End If
                
        '���סԺ�����Ƿ��Ժ���л��۵�
        If CheckExesState(lngKey) = False Then
            MsgBox "��ǰסԺ���˻��л��۵�δ��ˣ����ѳ�Ժ��Ԥ��Ժ��", vbInformation, Me.Caption
'            gcnOracle.RollbackTrans
'            blnRollBack = False
            Exit Function
        End If
                
        
                
        '������˹����ж�
        strErrInfo = ""
        If VerifyAuditingRule(lngKey, strErrInfo) = 1 Then
            If Mid(strErrInfo, 1, 2) = "1|" And InStr(mstrPrivs, "ǿ����˹���") <= 0 Then
                strErrInfo = Mid(strErrInfo, 3)
                MsgBox "<" & strPatienName & ">�ļ��鵥���δͨ��!" & vbNewLine & strErrInfo
'                gcnOracle.RollbackTrans
'                blnRollBack = False
                Exit Function
            End If
            strErrInfo = Mid(strErrInfo, 3)
            If MsgBox("<" & strPatienName & ">�ļ��鵥���δͨ��!�Ƿ�����?" & vbNewLine & strErrInfo, _
                vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
'                gcnOracle.RollbackTrans
'                blnRollBack = False
                Exit Function
            End If
        End If
                
        
        'ǩ�����ɹ�ʱ�˳�
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "Signature;" & lngKey & ";" & mstrAuditingMan
        
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & lngKey & ",'" & UserInfo.���� & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        
        
        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
        astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & lngKey & "," & intPrivacy & ",'" & gstrUnitName & "')"             '��˺��������浥
        

        
        
        rs.MoveNext
    Loop
'    gcnOracle.BeginTrans
'    blnRollBack = True
    For intLoop = 1 To UBound(astrSQL)
        If UCase(Mid(astrSQL(intLoop), 1, 3)) = "ZL_" Then
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        Else
            If Signature(Val(Split(astrSQL(intLoop), ";")(1)), mstrAuditingMan) = False Then
'                gcnOracle.RollbackTrans
'                blnRollBack = False
                Exit Function
            End If
        End If
    Next
'    gcnOracle.CommitTrans
'    If blnAutoPrint Then ReportPrint True                                           '�Ƿ���ɺ�ֱ�Ӵ�ӡ����
    Exit Function
errH:
    If blnRollBack = True Then
        blnRollBack = False
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function
Private Sub AllReportPrint(lngPatient As Long, blnPrint As Boolean)
    '����           '�����˴�ӡ���浥
    '               lngPatient=����ID
    '               blnPrint =True��ӡ False=Ԥ��
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New adodb.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    Dim lngKey As Long
    Dim strҽ��ID As String                 'ҽ��ID�����ҽ��IDʹ��","�ָ���
    Dim str�걾ID As String                 '�걾ID, ����걾IDʹ��","�ָ���
    Dim strPrintCode As String              '���ݱ���
    Dim intItem As Integer
    Dim astrItem() As String
    Dim blnRollBack As Boolean                              '�Ƿ�ع�
   
    On Error GoTo errH
    
    
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "���ڴ�ӡ��ȴ�...", Me
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(1), 2)
    
    If strStart = "�Զ���" Then
        strStart = Format(Me.DTPDate, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    End If
    
    If strStart = "" Then strStart = GetDateTime("��  ��", 1)
    If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    
    strSQL = "Select id,ҽ��ID from ����걾��¼ where ����id = [1] and ����ʱ�� between [2] and [3] and ҽ��id is not null and ����� is null and ΢����걾 is null "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, CDate(strStart), CDate(strEnd))
    
    Do While Not rsTmp.EOF
        strҽ��ID = strҽ��ID & "," & rsTmp("ҽ��ID")
        str�걾ID = str�걾ID & "," & rsTmp("ID")
        rsTmp.MoveNext
    Loop
       
    If strҽ��ID <> "" Then strҽ��ID = Mid(strҽ��ID, 2)
    If str�걾ID <> "" Then str�걾ID = Mid(str�걾ID, 2)
    
    lngҽ��ID = Split(strҽ��ID, ",")(0)
    lngKey = Split(str�걾ID, ",")(0)
    
    '�ж����ʽʱ�õ���ʽ
    frmLabMainPrintFormat.ShowMe Me, strҽ��ID, strPrintCode
    
    strSQL = "select /*+ rule */ id from ����ͼ���� where �걾id In(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�걾ID)
    intLoop = 1
    Do Until rsTmp.EOF
        If intLoop < 9 Then
            strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
            Call LoadImageData(App.path, rsTmp("ID"))
            intLoop = intLoop + 1
        End If
        rsTmp.MoveNext
    Loop
    
    Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & strҽ��ID, _
                        "����ID=" & lng����ID, "�걾ID=" & str�걾ID, "���ҽ��=" & strҽ��ID, "����걾=" & str�걾ID, _
                        "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                        "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                        "ͼ��9=" & strChart(9), IIf(blnPrint, 2, 1))
    
   astrItem = Split(str�걾ID, ",")
   gcnOracle.BeginTrans
   blnRollBack = True
   For intLoop = 0 To UBound(astrItem)
        strSQL = "ZL_����걾��¼_�걾�ʿ�(" & astrItem(intLoop) & ",'',1)"
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
   Next
   gcnOracle.CommitTrans
    Me.MousePointer = 0
    zlCommFun.StopFlash
    
    On Error Resume Next
    'ɾ��ͼ���ļ�
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    Exit Sub
errH:
    If blnRollBack = True Then
        gcnOracle.RollbackTrans
    End If
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(objfrm As Object, strPrivs As String, strAuditingMan As String)
    '�򿪴��ڲ�����Ȩ��
    mstrPrivs = strPrivs
    mstrAuditingMan = strAuditingMan
    Me.Show vbModal, objfrm
End Sub
