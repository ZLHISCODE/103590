VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmEInvoiceManage 
   Caption         =   "����Ʊ�ݹ���"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15120
   Icon            =   "frmEInvoiceManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFunc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8268
      Left            =   144
      ScaleHeight     =   8265
      ScaleWidth      =   3330
      TabIndex        =   1
      Top             =   600
      Width           =   3324
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   312
         ScaleHeight     =   6975
         ScaleWidth      =   2310
         TabIndex        =   2
         Top             =   264
         Width           =   2304
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   4320
            Left            =   48
            TabIndex        =   3
            Top             =   1272
            Width           =   2208
            _Version        =   589884
            _ExtentX        =   3895
            _ExtentY        =   7620
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   7608
         Left            =   24
         TabIndex        =   5
         Top             =   48
         Width           =   3000
         _Version        =   589884
         _ExtentX        =   5292
         _ExtentY        =   13414
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10584
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21590
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
   Begin XtremeCommandBars.ImageManager imgFunc 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmEInvoiceManage.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   864
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEInvoiceManage.frx":1EED4
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEInvoiceManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys As Long, mlngModule As Long, mstrDBUser As String
Private mblnFirst As Boolean
Private mstrEInvPrivs As String  '����Ʊ�ݲ���ģ��Ȩ��

Private Enum Panel_Index
    Pane_Fun = 1001
    Pane_Form = 1002
End Enum

Private marrFunc(1) As String
Private Enum FunID_Idex
    FunID_������������ = 101
    FunID_�վݷ�Ŀ���� = 102
    FunID_�շ��������� = 103
    FunID_֧�������� = 104
    FunID_��Ʊ������� = 105
    FunID_��������Ʊ�� = 201
    FunID_����Ʊ�ݴ�ӡ = 202
    FunID_����Ʊ�ݺ˶� = 203
End Enum

Private mWorkPan As Pane '��ǰ����
Private mfrmCurForm As Form '��ǰ���ܴ���
Private mfrmEInvoicePoint As frmEInvoicePoint
Private mfrmEInvoiceFees As frmEInvoiceFees
Private mfrmEInvoiceChannel As frmEInvoiceChannel
Private mfrmEInvoiceInsure As frmEInvoiceInsure
Private mfrmEInvoiceBalance As frmEInvoiceBalance
Private WithEvents mfrmEInvoiceCheck As frmEInvoiceCheck
Attribute mfrmEInvoiceCheck.VB_VarHelpID = -1
Private WithEvents mfrmEInvoiceCreate As frmEInvoiceCreate
Attribute mfrmEInvoiceCreate.VB_VarHelpID = -1
Private WithEvents mfrmEInvoicePrint As frmEInvoicePrint
Attribute mfrmEInvoicePrint.VB_VarHelpID = -1
Private mobjEInvoice As clsEInvoiceModule, mobjPubEInvoice As Object

Public Sub ShowMe(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, ByVal strDBUser As String, _
    objEInvoice As Object, Optional ByVal bytCheckTimeType As Byte)
    '�������
    '��Σ�
    '
    mlngSys = lngSys: mlngModule = lngModule
    mstrDBUser = strDBUser
    Set mobjEInvoice = objEInvoice
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    Call mobjEInvoice.zlGetEInvoiceProductName(Me, gstrProductName)
    
    On Error Resume Next
    Me.Show , frmMain
End Sub

Public Sub BHShowMe(ByVal lngMain As Long, ByVal lngSys As Long, ByVal lngModule As Long, ByVal strDBUser As String, _
    objEInvoice As Object)
    'BH���ó������
    mlngSys = lngSys: mlngModule = lngModule
    mstrDBUser = strDBUser
    Set mobjEInvoice = objEInvoice
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    Call mobjEInvoice.zlGetEInvoiceProductName(Me, gstrProductName)
    
    On Error Resume Next
    zlCommFun.ShowChildWindow Me.hWnd, lngMain
End Sub

Private Sub Form_Activate()
    If mblnFirst Then mblnFirst = False: Exit Sub
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    zlCommFun.ShowFlash "���ڼ������ݣ����Ե�...", Me
    mstrEInvPrivs = GetPrivFunc(mlngSys, 1145)     '����Ʊ�ݲ���Ȩ��
    
    Call DefMainCommandBars
    Call InitPanel '��ʼ��dkpMain
    Call InitFunPanel
    
    Call RestoreWinState(Me, App.ProductName)
    
    zlCommFun.StopFlash
End Sub

Private Sub InitFunPanel()
    Dim strCategory As String
    Dim objPic As PictureBox
    
    strCategory = "ҵ�����ݹ���,�������ݹ���"
    
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = "114,201,��������Ʊ��;5012,202,����Ʊ�ݴ�ӡ;3010,203,����Ʊ�ݺ˶�"
    marrFunc(1) = "100,101,������������;105,102,�վݷ�Ŀ����;102,103,�շ���������;104,104,֧��������;111,105,��Ʊ�������"
    
    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hWnd)
    Call scbFunc.Icons.AddIcons(imgFunc.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
End Sub

Public Sub InitSCBItem(ByRef scb As ShortcutBar, ByVal strItems As String, ByRef lngTPLhwnd As Long, Optional ByVal lngSelectedItem As Long = 1)
'���ܣ���ʼ��һ������������б�
'������
'      strItems         - ��������б����ƣ��Զ��ŷָ�,�����������ݳ�ʼ,���������,�ӿ�����
'      lngTPLhwnd       - �����б��ϰ󶨵�TaskPanel���ڵ���������������Picture��
'      lngSelectedItem  - ȱʡѡ��������,��1��ʼ
    Dim scbItem As ShortcutBarItem
    Dim i As Long
    Dim arrItem As Variant
    
    arrItem = Split(strItems, ",")
    For i = 0 To UBound(arrItem)
        Set scbItem = scb.AddItem(i + 1, arrItem(i), lngTPLhwnd)    'ͼ����ű�ָ����С1������Ҫ��1
        If i + 1 = lngSelectedItem Then Set scb.Selected = scbItem
    Next
    
    scb.ExpandedLinesCount = scb.ItemCount
End Sub

Public Sub InitTPLItem(ByRef scc As ShortcutCaption, ByRef tplFunc As TaskPanel, _
        ByVal strCategory As String, ByVal strItems As String, Optional ByVal lngSelectedItem As Long = 1)
'���ܣ���ʼ�����¼���һ����������б���һ�����飩
'������
'      strCategory      - ��ʾ��ShotcutCaption�ϵĵ�ǰ��������
'      strItems         - ���������������ƣ��Էֺŷָ�,�Զ��ŷָ�ͼ��ID���������鼰������������,����401,1,���ﻮ�۹���;412,2,�����շѹ���;......
'      lngSelectedItem  - ȱʡѡ��������,��1��ʼ
    Dim tplGroup As TaskPanelGroup
    Dim tplItem As TaskPanelGroupItem
    Dim arrItem As Variant
    Dim i As Long
    Dim lngImg As Long, lngID As Long
    Dim strItem As String
    Dim lngUbound As Long
    
    '����һ�����ط���
    scc.Caption = strCategory
    If tplFunc.Groups.Count = 0 Then
        Set tplGroup = tplFunc.Groups.Add(1, "����")
        tplGroup.CaptionVisible = False
        tplGroup.Expanded = True
        
        tplFunc.SetMargins 1, 2, 0, 2, 2
        tplFunc.SetIconSize 24, 24
        tplFunc.SelectItemOnFocus = True
    Else
        Set tplGroup = tplFunc.Groups(1)    'index�Ǵ�1��ʼ��
        tplGroup.Items.Clear
    End If
    
    arrItem = Split(strItems, ";")
    lngUbound = UBound(arrItem)
    For i = 0 To lngUbound
        lngImg = Split(arrItem(i), ",")(0) + 1  'ͼ����ű�ָ����С1������Ҫ��1
        lngID = Split(arrItem(i), ",")(1)       'ID����Ϊ�����ؼ�������Picture�����ţ�
        strItem = Split(arrItem(i), ",")(2)
        Set tplItem = tplGroup.Items.Add(lngID, strItem, xtpTaskItemTypeLink, lngImg)
        If i = lngUbound Then tplItem.SetMargins 0, 0, 0, 0 '��Ȼ���һ��ѡ��ʱ�Ŀ������ȫ��ס����
        If i + 1 = lngSelectedItem Then tplItem.Selected = True
    Next
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    If Val(tplFunc.Tag) = Item.ID Then Exit Sub
    tplFunc.Tag = Item.ID
    
    Select Case Item.ID
    Case FunID_������������
        If mfrmEInvoicePoint Is Nothing Then
            Set mfrmEInvoicePoint = New frmEInvoicePoint
            Call mfrmEInvoicePoint.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoicePoint
    Case FunID_�վݷ�Ŀ����
        If mfrmEInvoiceFees Is Nothing Then
            Set mfrmEInvoiceFees = New frmEInvoiceFees
            Call mfrmEInvoiceFees.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceFees
    Case FunID_�շ���������
        If mfrmEInvoiceChannel Is Nothing Then
            Set mfrmEInvoiceChannel = New frmEInvoiceChannel
            Call mfrmEInvoiceChannel.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceChannel
    Case FunID_֧��������
        If mfrmEInvoiceInsure Is Nothing Then
            Set mfrmEInvoiceInsure = New frmEInvoiceInsure
            Call mfrmEInvoiceInsure.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceInsure
    Case FunID_��Ʊ�������
        If mfrmEInvoiceBalance Is Nothing Then
            Set mfrmEInvoiceBalance = New frmEInvoiceBalance
            Call mfrmEInvoiceBalance.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser)
        End If
        Set mfrmCurForm = mfrmEInvoiceBalance
    Case FunID_����Ʊ�ݺ˶�
        If mfrmEInvoiceCheck Is Nothing Then
            Set mfrmEInvoiceCheck = New frmEInvoiceCheck
            Call mfrmEInvoiceCheck.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser, mobjEInvoice)
        End If
        Set mfrmCurForm = mfrmEInvoiceCheck
    Case FunID_��������Ʊ��
        If mfrmEInvoiceCreate Is Nothing Then
            Set mfrmEInvoiceCreate = New frmEInvoiceCreate
            Call mfrmEInvoiceCreate.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser, mstrEInvPrivs, mobjEInvoice, mobjPubEInvoice)
        End If
        Set mfrmCurForm = mfrmEInvoiceCreate
    Case FunID_����Ʊ�ݴ�ӡ
        If mfrmEInvoicePrint Is Nothing Then
            Set mfrmEInvoicePrint = New frmEInvoicePrint
            Call mfrmEInvoicePrint.InitCommVariable(Me, cbsThis, mlngSys, mlngModule, mstrDBUser, mstrEInvPrivs, mobjEInvoice, mobjPubEInvoice)
        End If
        Set mfrmCurForm = mfrmEInvoicePrint
    Case Else
        Exit Sub
    End Select
    
    mWorkPan.Handle = mfrmCurForm.hWnd
    'ˢ���Ӵ���˵���������
    Call DefSubCommandBars(mWorkPan)
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID�Ǵ�1��ʼ�ģ���ΪͬʱΪͼ����ţ�,�����Ǵ�0��ʼ
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub

Private Sub picFunc_Resize()
    On Error Resume Next
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - 45
    scbFunc.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    On Error Resume Next
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub

Private Sub InitPanel()
    Dim objPane As Pane

    On Error GoTo ErrHandler
    Set objPane = dkpMain.CreatePane(Pane_Fun, 120, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.Tag = Pane_Fun
    objPane.MinTrackSize.Width = 60
    objPane.MaxTrackSize.Width = 240

    Set mWorkPan = dkpMain.CreatePane(Pane_Form, 700, 400, DockRightOf, objPane)
    mWorkPan.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    mWorkPan.Tag = Pane_Form

    With dkpMain
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DefMainCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar

    On Error GoTo ErrHandler

    Set cbsThis.Icons = zlCommFun.GetPubIcons
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
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False: cbrControl.BeginGroup = True
    End With

    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With

    '���ò����ò˵�
    With cbsThis.Options
        '.AddHiddenCommand conMenu_File_PrintSet
        '.AddHiddenCommand conMenu_File_Excel
    End With

    DefMainCommandBars = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DefSubCommandBars(ByVal objItem As Pane)
    '���ܣ�ˢ���Ӵ���˵���������
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String

    On Error GoTo ErrHandler
    '��¼���в˵���ʽ
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsThis.Count >= 2 Then
        idx = GetFirstCommandBar(cbsThis(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsThis(2).Visible
            bytStyle = cbsThis(2).Controls(idx).Style
        End If
    End If

    'ˢ���Ӵ��ڲ˵�
    Call LockWindowUpdate(Me.hWnd)
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsThis.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsThis.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsThis.Count To 2 Step -1
        cbsThis(lngCount).Delete
    Next

    '���������¼���
    Call DefMainCommandBars

    '�Ӵ������¼���
    If Not mfrmCurForm Is Nothing Then
        Call mfrmCurForm.zlDefCommandBars
    End If

    '�ָ����̶���һЩ�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop + xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsThis.Count
        cbsThis(lngCount).ContextMenuPresent = False
        cbsThis(lngCount).ShowTextBelowIcons = False
        cbsThis(lngCount).EnableDocking xtpFlagStretched + xtpFlagHideWrap
        For Each objControl In cbsThis(lngCount).Controls
            If objControl.Type <> xtpControlLabel _
                And objControl.Type <> xtpControlEdit Then
                objControl.Style = bytStyle
            End If
        Next
        cbsThis(lngCount).Visible = blnShowBar
    Next

    '�������RecalcLayout����������
    Call LockWindowUpdate(0)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    stbThis.Top = Me.ScaleHeight - stbThis.Height
    stbThis.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
    Set mWorkPan = Nothing
    Set mfrmCurForm = Nothing
    
    'ж�����д���
    If Not mfrmEInvoicePoint Is Nothing Then Unload mfrmEInvoicePoint: Set mfrmEInvoicePoint = Nothing
    If Not mfrmEInvoiceCheck Is Nothing Then Unload mfrmEInvoiceCheck: Set mfrmEInvoiceCheck = Nothing
    If Not mfrmEInvoiceCreate Is Nothing Then Unload mfrmEInvoiceCreate: Set mfrmEInvoiceCreate = Nothing
    If Not mfrmEInvoicePrint Is Nothing Then Unload mfrmEInvoicePrint: Set mfrmEInvoicePrint = Nothing
    If Not mfrmEInvoiceFees Is Nothing Then Unload mfrmEInvoiceFees: Set mfrmEInvoiceFees = Nothing
    If Not mfrmEInvoiceChannel Is Nothing Then Unload mfrmEInvoiceChannel: Set mfrmEInvoiceChannel = Nothing
    If Not mfrmEInvoiceInsure Is Nothing Then Unload mfrmEInvoiceInsure: Set mfrmEInvoiceInsure = Nothing
    If Not mfrmEInvoiceBalance Is Nothing Then Unload mfrmEInvoiceBalance: Set mfrmEInvoiceBalance = Nothing
    Set mobjEInvoice = Nothing
    
    If Not mobjPubEInvoice Is Nothing Then
        Call mobjPubEInvoice.zlTerminate
        Set mobjPubEInvoice = Nothing
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case Else
        If Not mfrmCurForm Is Nothing Then
            Call mfrmCurForm.zlUpdateCommandBars(Control)
        End If
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    
    On Error GoTo ErrHandler
    Select Case Control.ID
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
    Case conMenu_Help_Help: Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((mlngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About: Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case Else
        If Not mfrmCurForm Is Nothing Then Call mfrmCurForm.zlExecuteCommandBars(Control)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Val(Item.Tag)
    Case Pane_Fun
        Item.Handle = picFunc.hWnd
    Case Pane_Form
        If Not mfrmCurForm Is Nothing Then
            Item.Handle = mfrmCurForm.hWnd
        End If
    End Select
End Sub

Private Sub mfrmEInvoiceCheck_ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    Call ShowPopupMenu(blnAddOutPutExcel)
End Sub

Private Sub mfrmEInvoiceCreate_ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    Call ShowPopupMenu(blnAddOutPutExcel)
End Sub

Private Sub mfrmEInvoicePrint_ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    Call ShowPopupMenu(blnAddOutPutExcel)
End Sub

Public Sub ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
    '�����Ҽ��˵�
    Dim objPopup As CommandBarPopup, cbCommandBar As CommandBar
    Dim cbrControl As CommandBarControl, cbrControlNew As CommandBarControl
    Dim i As Integer
    
    Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    Set cbCommandBar = cbsThis.Add("Popup", xtpBarPopup) '�����˵�
    If cbCommandBar Is Nothing Then Exit Sub
    
    For i = 1 To objPopup.CommandBar.Controls.Count
        Set cbrControl = objPopup.CommandBar.Controls(i)
        Call cbsThis_Update(cbrControl)   '�ж��Ƿ�ɼ�����Ϊ��һ��ʱ�˵���û��ִ��Update
        If cbrControl.Visible Then
            Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
            cbrControlNew.BeginGroup = cbrControl.BeginGroup
            cbrControlNew.IconId = cbrControl.IconId
            cbrControlNew.Enabled = cbrControl.Enabled
        End If
    Next
    
    If blnAddOutPutExcel Then
        Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_FilePopup, , True)
        If Not objPopup Is Nothing Then
            Set cbrControl = objPopup.CommandBar.Controls.Find(xtpControlButton, conMenu_File_Excel, , True)
            If cbrControl.Visible Then
                Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
                cbrControlNew.BeginGroup = True
                cbrControlNew.IconId = cbrControl.IconId
                cbrControlNew.Enabled = cbrControl.Enabled
            End If
        End If
    End If
    
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
End Sub

Private Sub mfrmEInvoiceCheck_ShowInfo(ByVal strInfo As String)
        Call ShowInfoInStatusBar(strInfo)
End Sub

Private Sub mfrmEInvoiceCreate_ShowInfo(ByVal strInfo As String)
        Call ShowInfoInStatusBar(strInfo)
End Sub

Private Sub mfrmEInvoicePrint_ShowInfo(ByVal strInfo As String)
        Call ShowInfoInStatusBar(strInfo)
End Sub

Private Sub ShowInfoInStatusBar(ByVal strInfo As String)
    stbThis.Panels(2).Text = strInfo
End Sub

Public Function GetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
'���ܣ���ȡ��������ӡԤ����ť��ĵ�һ����ť��index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_File_Preview Then
            idx = objControl.index + 1
        End If
    Next
    GetFirstCommandBar = idx
End Function
