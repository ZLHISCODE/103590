VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.MDIForm frmMDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "������"
   ClientHeight    =   10140
   ClientLeft      =   165
   ClientTop       =   60
   ClientWidth     =   16005
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '��Ļ����
   Begin ComCtl3.CoolBar cbarTool 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   1535
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   16005
      _CBHeight       =   870
      _Version        =   "6.7.9816"
      MinHeight1      =   285
      Width1          =   5880
      NewRow1         =   0   'False
      MinHeight2      =   525
      Width2          =   2880
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   780
         Top             =   225
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComDlg.CommonDialog DlgMain 
      Bindings        =   "frmMain.frx":1CFA
      Left            =   3885
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9765
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   661
      SimpleText      =   $"frmMain.frx":1D0E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":1D55
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23151
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   870
      Width           =   3135
      Begin XtremeSuiteControls.ShortcutBar sbFunc 
         Height          =   8640
         Left            =   120
         TabIndex        =   4
         Top             =   90
         Width           =   2550
         _Version        =   589884
         _ExtentX        =   4498
         _ExtentY        =   15240
         _StockProps     =   64
      End
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   4260
         Left            =   2625
         MousePointer    =   9  'Size W E
         ScaleHeight     =   4260
         ScaleWidth      =   45
         TabIndex        =   3
         Top             =   570
         Width           =   45
      End
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   7200
      Top             =   2280
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":25E7
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==��������
'==============================================================
Private mstrCurModule    As String          '��ǰѡ��ģ��
Public gstrLastModule   As String           '�ϴ�ѡ��ģ��
Public grsToolsMenu     As ADODB.Recordset  '�����߲˵�
Private mcllModuleBar   As Collection
Private mcllItems       As Collection       '�������Ӵ���
'==============================================================
'==�����ӿ�
'==============================================================
Public Sub RunByModule(ByVal strNo As String)
'���ܣ�ת��ִ��ģ��˵�
    Dim frmChild As Form, strTmp As String
    mstrCurModule = strNo
    strTmp = Mid(strNo, 1, 2)
    sbFunc.Tag = "ģ�����"
    If strTmp <> "" Then
        sbFunc.Selected = sbFunc.FindItem(Val(strTmp))
        If sbFunc.Tag <> "" Then 'ͬһ�����²��ᴥ���¼���ǿ�Ƶ���
            Call sbFunc_SelectedChanged(sbFunc.FindItem(Val(strTmp)))
        End If
    End If
    mstrCurModule = ""
End Sub

Public Function GetIcons() As ImageManager
    Set GetIcons = imgMain
End Function

'==============================================================
'=�ؼ��¼�
'==============================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer, strTmp As String
    
    On Error Resume Next
    Select Case Control.Id
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
        Case conMenu_File_Preview 'Ԥ��
            gfrmActive.SubPrint 2
        Case conMenu_File_Print '��ӡ
            gfrmActive.SubPrint 1
        Case conMenu_File_Excel '�����Excel
            gfrmActive.SubPrint 3
        Case conMenu_View_ToolBar_Button        '������
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
                
                cbarTool.Bands.Item(2).NewRow = Not cbarTool.Bands.Item(2).NewRow
                
                If cbarTool.Bands.Item(2).NewRow = True Then
                    If Me.cbsMain.Options.LargeIcons = True Then
                        cbarTool.Bands.Item(2).MinHeight = 520
                    Else
                        cbarTool.Bands.Item(2).MinHeight = 420
                    End If
                Else
                    cbarTool.Bands.Item(2).MinHeight = cbarTool.Bands(1).MinHeight
                End If
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text            '��ť����
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size        '��ͼ��
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            If Me.cbsMain.Options.LargeIcons = True Then
                cbarTool.Bands.Item(2).MinHeight = 520
            Else
                cbarTool.Bands.Item(2).MinHeight = 420
            End If
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolsList '�����б�
            Me.picFunc.Visible = Not Me.picFunc.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolsPwd
            Clipboard.Clear
            Clipboard.SetText gstrLoginUserPwd
        Case conMenu_Help_Help
            Select Case UCase(gfrmActive.name)
                Case UCase("frmAppMan")         'װж����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmAppStart")       'ϵͳװж����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppStart"
                Case UCase("frmAppUpgrade")     'ϵͳ��Ǩ
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppUpgrade"
                Case UCase("frmAppCheck")       '�������޸�
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppCheck"
                Case UCase("frmAppScript")      '�û���װ�ű�
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAppScript"
                Case UCase("frmDataMan")        '���ݹ���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmDataMove")       '���ݹ鵵ת��
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmDataMove"
                Case UCase("frmExp")            '���ݵ���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmExp"
                Case UCase("frmImp")            '���ݵ���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmImp"
                Case UCase("frmLoadOut")        '���ݵ���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmLoadOut"
                Case UCase("frmLoadIn")         '���ݵ���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmLoadIn"
                Case UCase("frmClearData")      '�������
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmClearData"
                Case UCase("frmRunMan")         '���й���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmRegist")         '�û�ע�����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmRegist"
                Case UCase("frmStatus")         '����״̬���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmStatus"
                Case UCase("frmAutoJobs")       '��̨��ҵ����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmAutoJobs"
                Case UCase("FrmRunLog")         '������־����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "FrmRunLog"
                Case UCase("frmParameters")
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmParameters"
                Case UCase("FrmErrLog")         '������־����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "FrmErrLog"
                Case UCase("FrmRunOption")      'ϵͳ����ѡ��
                    ShowHelp Me.hwnd, "zl9svrtools\" & "FrmRunOption"
                Case UCase("frmGrantMan")       'Ȩ�޹���
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmSvrCreate"
                Case UCase("frmRole")           '��ɫ��Ȩ����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmRole"
                Case UCase("frmUser")           '�û���Ȩ����
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmUser"
                Case UCase("frmMenu")           '�˵�����滮
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmMenu"
                Case UCase("frmMgrGrant") '��������Ȩ
                    ShowHelp Me.hwnd, "zl9svrtools\" & "frmMgrGrant"
                Case UCase("frmRptMan") '�������
                    ShowHelp Me.hwnd, "zlreport\main"
            End Select
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '����
            Call frmAbout.ShowAbout
        Case conMenu_File_RemoveTools 'ж�ع�����
            Call FileRemove
        Case conMenu_File_LogOut 'ע��
            Unload Me
            Call Main
        Case conMenu_File_Exit '�˳�
            Unload Me
        Case Else '�򿪲˵��е�ģ��
            strTmp = mcllModuleBar("K_" & Mid(Control.Id, 1, 3))
            If strTmp <> "" Then
                Call RunByModule(Mid(Control.Id, Len(strTmp) + 1))
            End If
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If gfrmActive Is Nothing Then
        blnEnabled = False
    Else
        blnEnabled = gfrmActive.SupportPrint()
    End If
    
    Select Case Control.Id
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel '��ӡ,Ԥ��,�����Excel
            Control.Enabled = blnEnabled
        Case conMenu_File_RemoveTools 'ж�ع����ߵĿ�����
            Control.Enabled = gblnDBA
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_ToolsList
            Control.Checked = Me.picFunc.Visible
    End Select
End Sub

Private Sub MDIForm_Load()
    Dim rsTmp           As ADODB.Recordset, strSQL      As String
    Dim rsTmpChild      As ADODB.Recordset
    Dim objFrmTmp       As frmItem, objFrmMain          As frmItem
    Dim sbItem          As ShortcutBarItem, sbItemMain  As ShortcutBarItem
    Dim objPopup        As CommandBarPopup, objPrarentPop As CommandBarPopup, objControl As CommandBarControl
    Dim lngControlID    As Long
    Dim strSort         As String
    
    
    On Error GoTo errH
    Call zl9PrintMode.IniPrintMode(gcnOracle, gstrUserName)
    Call InitCommandBar
    '����˵�ID
    Set mcllModuleBar = New Collection
    mcllModuleBar.Add conMenu_Tool_LoadAndUnload, "K_501" 'װж����
    mcllModuleBar.Add conMenu_Tool_DataMana, "K_502" '���ݹ���
    mcllModuleBar.Add conMenu_Tool_RunMana, "K_503" '���й���
    mcllModuleBar.Add conMenu_Tool_Popedom, "K_504" 'Ȩ�޹���
    mcllModuleBar.Add conMenu_Tool_Expert, "K_505" 'ר���
    mcllModuleBar.Add conMenu_Tool_DBA, "K_506" 'DBA����
    
    Set mcllItems = New Collection
    '������ʼ���Լ�һЩ�����Ľ�������
    Set gcbsMain = Me.cbsMain
    gblnSystemUser = gclsBase.IsStSystemUser(gstrLoginUserName)
    Me.Caption = Me.Caption & " [" & gstrLoginUserName & IIf(gstrServer = "", "", "@" & gstrServer) & "]"
    gstrSysName = gstrProductName & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    Call ApplyOEM(stbThis)
    Call ApplyOEM_Picture(Me, "Icon")
    '��ȡ�˵����ز˵�
    Call CheckProcManage    '��ʱ���\��ӱ䶯���̹������
    If CheckAndAdjustMustTable("Zlsvrtools", "����", False) Then
        strSort = "����,���"
    Else
        strSort = "���"
    End If
    strSQL = "Select * From Zlsvrtools" & IIf(gstrHaveProg <> "", " Where �ϼ� is null or instr('," & gstrHaveProg & ",' ,',' ||  ���  || ',' )>0", "")
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Set rsTmpChild = CopyNewRec(rsTmp)
    Set grsToolsMenu = CopyNewRec(rsTmp)
    '����ұ߿�ݷ����Լ��˵��µ�ģ��
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    rsTmp.Filter = "�ϼ�=NULL": rsTmp.Sort = "���"
    Do While Not rsTmp.EOF
        rsTmpChild.Filter = "�ϼ� = " & rsTmp!���
        rsTmpChild.Sort = strSort
        If rsTmpChild.RecordCount > 0 Then
            If rsTmpChild.RecordCount = 1 And rsTmpChild!��� = "0404" And Not gblnSystemUser Then
                '��ֻ�й�������Ȩģ�飬�Ҳ���ϵͳ�û�ʱ�����ټ��ط����Լ�����
            Else
                If mstrCurModule = "" Then mstrCurModule = rsTmpChild!���
                '��ȡ�˵���Ŀ,����Ӳ˵�
                If Not objControl Is Nothing Then
                    lngControlID = mcllModuleBar("K_5" & rsTmp!���)
                    Set objPrarentPop = objControl.CommandBar.Controls.Add(xtpControlButtonPopup, lngControlID, rsTmp!���� & IIf(rsTmp!��� & "" = "", "", "(&" & rsTmp!��� & ")"))
                    If Not objPrarentPop Is Nothing Then
                        Do While Not rsTmpChild.EOF
                            objPrarentPop.CommandBar.Controls.Add xtpControlButton, Val(lngControlID & rsTmpChild!���), rsTmpChild!���� & IIf(rsTmpChild!��� & "" = "", "", "(&" & rsTmpChild!��� & ")"), -1, False
                            rsTmpChild.MoveNext
                        Loop
                    End If
                End If
                '����ұ߿�ݵ���
                Set objFrmTmp = New frmItem
                objFrmTmp.gstrParentNo = rsTmp!��� & ""
                objFrmTmp.gstrParentCap = rsTmp!���� & ""
                Set sbItem = sbFunc.addItem(Val(rsTmp!���), rsTmp!����, objFrmTmp.hwnd)
                mcllItems.Add objFrmTmp, "K_" & rsTmp!��� & ""
                If objFrmMain Is Nothing Then Set objFrmMain = objFrmTmp
                If sbItemMain Is Nothing Then Set sbItemMain = sbItem
            End If
        End If
        rsTmp.MoveNext
    Loop
    '��λ��һ��ģ��
    Call sbFunc.Icons.AddIcons(imgMain.Icons)
    sbFunc.ExpandedLinesCount = sbFunc.ItemCount
    Call RunByModule(mstrCurModule)
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub MDIForm_Resize()

    On Error Resume Next
    If picVbar.Left < 2200 Then picVbar.Left = 2200
    If picVbar.Left > Width - 3000 Then picVbar.Left = Width - 3000
    picVbar.Top = 0
    picVbar.Height = picFunc.Height
    picFunc.Width = picVbar.Left + picVbar.Width
    
    sbFunc.Left = picFunc.ScaleLeft + 45
    sbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    sbFunc.Top = picFunc.ScaleTop
    sbFunc.Height = picFunc.ScaleHeight
    
    If stbThis.Panels(2) = "" Then
        '���⴦����Ȼ״̬���Ŀ�Ȳ���ȷ
        stbThis.Panels(2) = " "
        stbThis.Panels(2) = ""
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmChild As Form
    Set grsToolsMenu = Nothing
    Set mcllItems = Nothing
    Set mcllModuleBar = Nothing
    mstrCurModule = ""
    gstrLastModule = ""
    For Each frmChild In Forms
        Unload frmChild
    Next
End Sub

Private Sub picFunc_Resize()
    Call MDIForm_Resize
End Sub

Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        picVbar.Left = IIf(picVbar.Left + x < 2200, 2200, picVbar.Left + x)
        Call MDIForm_Resize
    End If
End Sub

Private Sub sbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub sbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If sbFunc.Tag <> "" Then
        If mstrCurModule <> gstrLastModule Then
            Call mcllItems("K_" & Format(Item.Id, "00")).RunByModule(mstrCurModule)
        End If
        sbFunc.Tag = ""
    Else
        Call mcllItems("K_" & Format(Item.Id, "00")).RunByModule
    End If
End Sub

'==============================================================
'=˽�з���
'==============================================================
Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_RemoveTools, "ж�ع�����(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_LogOut, "ע��(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    objMenu.Id = conMenu_ToolPopup
'    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_LoadAndUnload, "װж����(&I)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_DataMana, "���ݹ���(&D)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_RunMana, "���й���(&E)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Popedom, "Ȩ�޹���(&G)")
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Expert, "ר���(&R)")
'    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_View_ToolsList, "�����б�(&L)")
        If gstrUserName <> gstrLoginUserName Then
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolsPwd, gstrLoginUserName & "�����ݿ�����(�������):" & gstrLoginUserPwd)
        End If
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrWebSustainer)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrWebSustainer & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF12, conMenu_File_Parameter '��������
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyG, conMenu_View_Filter '����
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����һЩ�����Ĳ���������
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
        .AddHiddenCommand conMenu_File_Excel '�����Excel
    End With
End Sub

Private Sub FileRemove()
'���ܣ�ж�ع�����
    Dim rsTmp As ADODB.Recordset
    Dim blnReturn As Boolean
    
    '�ж��Ƿ����ж��
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    If rsTmp.RecordCount > 0 Then
        MsgBox "��ǰ�Ѿ���װ��Ӧ��ϵͳ������ɾ�������ߡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    If MsgBox("�������ǹ���Ӧ��ϵͳ�Ļ�����" & vbCrLf & "����ɾ����Ͳ��������κι�����������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    blnReturn = RemoveServer
    Me.MousePointer = vbDefault
    Me.Enabled = True
    If blnReturn Then Unload Me
End Sub

Private Function RemoveServer() As Boolean
'���ܣ���ж�������ʵ��
'-------------��ж�㷨-----------------
'   ɾ���û�
'   ɾ���ع���
'   ɾ����ռ�
'--------------------------------------
    Dim strSpaces As String, strFiles As String, strErrInfo As String
    Dim aryTbs() As String, aryFile() As String
    Dim rsTemp As New ADODB.Recordset, intVer As Integer
    Dim strSQL As String, intCount As Integer
    
    On Error GoTo 0
    With rsTemp
        .Open "select 1 from gv$session where USERNAME='ZLTOOLS'", gcnOracle
        If .EOF = False Then
            MsgBox "ZLTOOLS�û������ӵ����ݿ��ϣ��޷����ж�ز�����", vbExclamation, gstrSysName
            Exit Function
        End If
    End With
    
    '������ռ估�����ļ�
    strSpaces = "'ZLTOOLSTBS','ZLTOOLSTMP'"
    strFiles = ""
    With rsTemp
        strSQL = "select F.NAME from V$TABLESPACE T,V$DATAFILE F where T.TS#=F.TS# and T.NAME in (" & strSpaces & ")"
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        Do Until .EOF
            strFiles = strFiles & ";" & .Fields("NAME").value
            .MoveNext
        Loop
    End With
    If strFiles <> "" Then strFiles = Mid(strFiles, 2)

    On Error Resume Next
    
    '�ض�����,����ִ��ɾ�������߲���
    strSQL = "Truncate Table zltools.zlregaudit"
    gcnOracle.Execute strSQL
    If err.Number <> 0 Then Debug.Print err.Description
    
    strSQL = "Truncate Table zltools.zlRegFile"
    gcnOracle.Execute strSQL
    If err.Number <> 0 Then Debug.Print err.Description
   
    'ɾ����ϵͳ������
    
    stbThis.Panels(2).Text = "ɾ�������������ߡ�"
    DoEvents
    intCount = 0
    Do
        gcnOracle.Execute "Drop user ZLTOOLS cascade"
        If err.Number <> 0 Then Debug.Print err.Description
        With rsTemp
            If .State = adStateOpen Then .Close
            .Open "select * from all_users where username='ZLTOOLS'", gcnOracle
            If .EOF Then Exit Do
        End With
        intCount = intCount + 1
        DoEvents
        '���ɾ��100�������ʧ���˾Ͳ��ټ���
        If intCount > 100 Then
            MsgBox "����ɾ���û�ZLTOOLS������������״̬��δ�����", vbInformation, gstrSysName
            Exit Function
        End If
    Loop
    
    'ɾ���Ѿ������Ĺ���ͬ���
    stbThis.Panels(2).Text = "ɾ������ͬ��ʡ�"
    DoEvents
    If rsTemp.State = adStateOpen Then rsTemp.Close
    strSQL = "SELECT Synonym_Name FROM All_Synonyms WHERE owner='PUBLIC' AND table_owner='ZLTOOLS'"
    rsTemp.Open strSQL, gcnOracle, adOpenStatic
    Do Until rsTemp.EOF
        strSQL = "drop public Synonym  " & rsTemp("Synonym_Name")
        gcnOracle.Execute strSQL
        rsTemp.MoveNext
    Loop
    
    'ɾ�������ڱ�ռ��ϵĻع���
    stbThis.Panels(2).Text = "ɾ����ռ��еĻع��Ρ�"
    DoEvents
    With rsTemp
        If .State = adStateOpen Then .Close
        strSQL = "select SEGMENT_NAME from DBA_ROLLBACK_SEGS where tablespace_name in(" & strSpaces & ")"
        .Open strSQL, gcnOracle
        Do Until .EOF
            DoEvents
            gcnOracle.Execute "alter rollback segment " & .Fields(0).value & " offline"
            gcnOracle.Execute "drop rollback segment " & .Fields(0).value
            .MoveNext
        Loop
    End With
    
    'ɾ����ϵͳ���ݿռ�
    stbThis.Panels(2).Text = "ɾ�����ݱ�ռ䡭"
    DoEvents
    
    intVer = GetOracleVersion(, True)
    If intVer < 9 Then
        gcnOracle.Execute "alter rollback segment rbs_ZLTOOLS offline"
        gcnOracle.Execute "drop rollback segment rbs_ZLTOOLS"
    End If
    
    aryTbs = Split(strSpaces, ",")
    For intCount = LBound(aryTbs) To UBound(aryTbs)
        DoEvents
        strSpaces = Mid(aryTbs(intCount), 2, Len(aryTbs(intCount)) - 2)
        gcnOracle.Execute "alter tablespace " & strSpaces & " offline"
        gcnOracle.Execute "drop tablespace " & strSpaces & " including contents and datafiles cascade constraints"
    Next
    
    '��ͼɾ�����õ������ļ�
    stbThis.Panels(2).Text = "ɾ�����õ������ļ���"
    DoEvents
    aryFile = Split(strFiles, ";")
    For intCount = LBound(aryFile) To UBound(aryFile)
        err = 0
        Kill aryFile(intCount)
        If err <> 0 Then
            strErrInfo = strErrInfo & vbCr & "�ļ���" & aryFile(intCount)
        End If
    Next
    If strErrInfo <> "" Then
        MsgBox "�����߲�ж��ɣ����ֹ�ɾ���������ݣ�" & strErrInfo, vbExclamation, gstrSysName
    Else
        MsgBox "�����߲�ж���", vbExclamation, gstrSysName
    End If
    RemoveServer = True
End Function
