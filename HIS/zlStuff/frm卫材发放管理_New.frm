VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frm���ķ��Ź���_New 
   Caption         =   "���ķ��Ź���"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   Icon            =   "frm���ķ��Ź���_New.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   2865
      Left            =   5505
      TabIndex        =   4
      Top             =   1230
      Visible         =   0   'False
      Width           =   2655
      _cx             =   4683
      _cy             =   5054
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm���ķ��Ź���_New.frx":08CA
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   2
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   135
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   1
      Top             =   1230
      Width           =   5070
      Begin VB.CheckBox Chk�嵥 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "��ʾ���й��̵���"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3105
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   1935
      End
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6060
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm���ķ��Ź���_New.frx":0917
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
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
      Left            =   0
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frm���ķ��Ź���_New.frx":11AB
      Left            =   615
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm���ķ��Ź���_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmδ����  As frm����δ�����嵥
Attribute mfrmδ����.VB_VarHelpID = -1
Private mfrm���ϻ��� As frm���ķ��ϻ���
Private mfrmȱ���嵥 As frm����ȱ���嵥
Private mfrm�ܷ��嵥 As frm���ľܷ����嵥
Private WithEvents mfrm�����嵥 As frm���������嵥
Attribute mfrm�����嵥.VB_VarHelpID = -1
Private WithEvents mfrmFilter As frm���ķ��Ź���
Attribute mfrmFilter.VB_VarHelpID = -1
Private mstrSelectTabItem As String
Private Const ID_PANE_SEARCH = 201
Private Const conMenu_Popu_סԺ�� = 1011
Private Const conMenu_Popu_���� = 1012
Private Const conMenu_Popu_���� = 1013
Private Const conMenu_Popu_����ID = 1014
Private Const conMenu_Popu_����� = 1015
Private Const conMenu_Popu_IC���� = 1016

Private mArrFilter As Variant   '��������
Private mcbrControl As CommandBarControl
Private mcbrMenuBar As CommandBarPopup
Private mcbrToolBar As CommandBar
Private mrsNotPayStuff As ADODB.Recordset   '�������ݼ�
Private mrsChargeOff As New ADODB.Recordset                   '������ʾ���������¼
Private mrsBakStuff As ADODB.Recordset
Private mstrPrivs As String
Private mlngModule As Long
Private mintUnit As Integer     '��λ:0-ɢװ��λ;1-��װ��λ
Private mint�ֺ� As Integer     '��ǰ�ֺ�:0-9,1-12,2-15
Private mblnҳǩ As Boolean
Private mintҳǩ As Integer

Private Enum mPage
    pag_δ���嵥 = 0
    pag_���ܷ��� = 1
    pag_ȱ���嵥 = 2
    pag_�ܷ��嵥 = 3
    pag_�����嵥 = 4
End Enum
'------------------------------------------------------------------------------------------------------------
'�Ӳ��ϴ�����ҩ�������Ĳ���
Private mblnTrans As Boolean            'True��ʾ�Ӳ��ϴ�����ҩ���ڵ���
Private mstrNo  As String               '���ݺţ������ڶ�λ
Private mlng�ⷿid As Long              '��ҩ�ⷿID��һ��ͷ��ϲ���һ��
Private mlng����id As Long              'mlng����id
Private mstrStuffStartDate As String     '���ϵ��ݿ�ʼʱ��
Private mstrStuffEndDate As String       '���ϵ��ݽ���ʱ��

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString

Private mint����ģʽ As Integer

Private mobjPlugIn As Object             '��ҽӿڶ���
'----------------------------------------------------------------------------------------------------------
Public Sub ShowList(ByVal frmMain As Form, ByVal lng����id As Long, ByVal strNo As String, ByVal lng�ⷿID As Long, ByVal strStartDate As String, ByVal strEndDate As String)
    '-----------------------------------------------------------------------------------------------------------
    '����:��ҩ��Ʒ����
    '���:frmMain-������
    '     lng����ID-����ID
    '     strNo-������
    '     lng�ⷿid-�ⷿID
    '     strStartDate-��ʼ����
    '     strEndDate-��������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 21:51:18
    '-----------------------------------------------------------------------------------------------------------
    
    mlng����id = lng����id
    mstrNo = strNo
    mlng�ⷿid = lng�ⷿID
    mstrStuffStartDate = strStartDate
    mstrStuffEndDate = strEndDate
    mblnTrans = True
    Me.Show , frmMain
    Me.ZOrder 0
End Sub

Private Sub initLocalPara()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��ʼ�����ز�������ֵ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 21:11:08
    '-----------------------------------------------------------------------------------------------------------
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mint�ֺ� = zlDatabase.GetPara("�����ֺ�", glngSys, mlngModule, "0")
    mblnҳǩ = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "���ķ��Ź���", "������һ�δ���ر�ʱ��ҳǩ", 0)) = 1)
    mintҳǩ = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "���ķ��Ź���", "��ǰҳǩ", 0))
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    
    mint����ģʽ = 0
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0 Then
        mint����ģʽ = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "���ķ��Ź���", "����ģʽ", "0"))
        If mint����ģʽ < 0 Then
            mint����ģʽ = 0
        End If
    End If
End Sub
 Private Sub InitPage()
    '------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:
    '����:���˺�
    '����:2007/08/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim stdfnt As StdFont
    Dim objItem As TabControlItem
    
    
    Set mfrmδ���� = New frm����δ�����嵥
    Set objItem = tbPage.InsertItem(mPage.pag_δ���嵥, "δ�����嵥", mfrmδ����.hwnd, 0)
    objItem.Tag = mPage.pag_δ���嵥
    Set mfrm���ϻ��� = New frm���ķ��ϻ���
    Set objItem = tbPage.InsertItem(mPage.pag_���ܷ���, "���ܷ���", mfrm���ϻ���.hwnd, 0)
    objItem.Tag = mPage.pag_���ܷ���
    
    Set mfrmȱ���嵥 = New frm����ȱ���嵥
    Set objItem = tbPage.InsertItem(mPage.pag_ȱ���嵥, "ȱ���嵥", mfrmȱ���嵥.hwnd, 0)
    objItem.Tag = mPage.pag_ȱ���嵥
    
    Set mfrm�ܷ��嵥 = New frm���ľܷ����嵥
    Set objItem = tbPage.InsertItem(mPage.pag_�ܷ��嵥, "�ܷ����嵥", mfrm�ܷ��嵥.hwnd, 0)
    objItem.Tag = mPage.pag_�ܷ��嵥
    
    Set mfrm�����嵥 = New frm���������嵥
    Set objItem = tbPage.InsertItem(mPage.pag_�����嵥, "�����嵥", mfrm�����嵥.hwnd, 0)
    objItem.Tag = mPage.pag_�����嵥
    Call mfrmFilter_zlRefreshCon(mArrFilter)
 
    With tbPage
        If mintҳǩ <> 0 And mblnҳǩ Then
            .Item(mintҳǩ).Selected = True
        Else
            .Item(0).Selected = True
        End If
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        Set stdfnt = Me.Font
        Set .PaintManager.Font = stdfnt
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function InitComandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/1/9
    '----------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl

    Dim panThis As Pane
    err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    
    
'    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.Id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrint, "��ӡ���ϵ���(&B)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrintView, "��ӡ����֪ͨ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.Id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)"):        mcbrControl.IconId = 3010
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "�ܷ�(&H)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "�ָ�(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Untread, "����(&T)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "����(&S)"):  mcbrControl.IconId = 21905:  mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CfPay, "����������(&C)"):  mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BillPay, "��Ʊ�ݺŷ���(&B)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_BillBackPay, "����������(&N)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_OtherPay, "�������ⷿ����(&N)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StopPay, "ֹͣ���ϱ��(&S)")
        mcbrControl.Visible = zlStr.IsHavePrivs(mstrPrivs, "ֹͣ����") Or zlStr.IsHavePrivs(mstrPrivs, "�ָ�����")
        mcbrControl.BeginGroup = mcbrControl.Visible
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.Id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
    
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_FontSize, "����(&F)")
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_FontSize_1, "С����(&S)", -1, False)
        If mint�ֺ� = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_FontSize_2, "������(&M)", -1, False)
        If mint�ֺ� = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = mcbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_FontSize_3, "������(&B)", -1, False)
        If mint�ֺ� = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Column, "������(&C)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.Id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With

    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With

    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3010
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "�ܷ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "�ָ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Untread, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "����"): mcbrControl.IconId = 21905: mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    InitComandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitPanel()
    Dim PaneSearch As Pane
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frm���ķ��Ź���
    mfrmFilter.Set���ϴ������� mblnTrans, mstrNo, mstrStuffStartDate, mstrStuffEndDate, mlng����id, mlng�ⷿid
    Set mArrFilter = mfrmFilter.GetFilterCon
    
    With dkpMan
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        
        '.Options.HideClient = True
        Set PaneSearch = .CreatePane(ID_PANE_SEARCH, 400, 100, DockTopOf, Nothing)
        PaneSearch.Title = "����"
       ' PaneSearch.Options = PaneNoCloseable
        '.ImageList = imlPaneIcons '
        .SetCommandBars cbsThis
    End With
End Function
Private Function SendBillPay(ByVal blnƱ�ݺŷ��� As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-06 14:55:11
    '-----------------------------------------------------------------------------------------------------------
    With Frm�����ŷ���
        .In_���� = 0
        .In_����IN = mArrFilter("����")
        .In_���ϲ���id = Val(mArrFilter("���ϲ���ID"))
        .In_����� = GetCheckPara()
        .In_����δ���Ϸ��� = 1
        .In_Ȩ�� = mstrPrivs
        .mstr������ = gstrUserName
        .��Ʊ�ݺŷ��� = blnƱ�ݺŷ���
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    SendBillPay = True
    Call mfrmFilter_zlRefreshCon(mArrFilter)
    
End Function
Private Function SendBackPay() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����ݺ�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-06 15:07:05
    '-----------------------------------------------------------------------------------------------------------
    Set Frm����������.In_PlugIn = mobjPlugIn
    If Frm����������.ShowCard(Me, Val(mArrFilter("���ϲ���ID")), mstrPrivs) = False Then Exit Function
    SendBackPay = True
    Call mfrmFilter_zlRefreshCon(mArrFilter)
End Function
Private Sub StopPayStuffFlag()
    '-----------------------------------------------------------------------------------------------------------
    '����:ֹͣ���ϱ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-06 15:09:00
    '-----------------------------------------------------------------------------------------------------------
    'ֹͣ����
    '��ҩ��ʽ=-1
    Dim frmFlag As New Frm���ٷ�ҩ������־
    frmFlag.In_����� = GetCheckPara
    
    '--50313��zdt��:��ֹͣ���ϵķ��ϲ���id���и�ֵ
    frmFlag.In_�ⷿid = Val(mArrFilter("���ϲ���id"))
    frmFlag.In_�������� = Val(mArrFilter("��������"))
    
    frmFlag.gstrParentName = Replace(Me.Name, "_New", "")
    frmFlag.Show vbModal
    Call mfrmFilter_zlRefreshCon(mArrFilter)
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long
    Dim blnAskPring As Boolean
    
    Dim cllFind As Collection
    
    '------------------------------------
    Select Case Control.Id
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    
    Case conMenu_File_BillPrint
        '���ϵ���ӡ
        If tbPage.Selected Is Nothing Then
            blnAskPring = True
        Else
            blnAskPring = Not (Val(tbPage.Selected.Tag) = mPage.pag_�����嵥)
        End If
        Call mfrm�����嵥.zlPrintBill(True, "", IIf(mfrmδ����.cboEdit(0).ListIndex = -1, 1, mfrmδ����.cboEdit(0).ListIndex + 1), mstrPrivs, blnAskPring)
        
    Case conMenu_File_BillPrintView
        '��ӡ����֪ͨ��
        If tbPage.Selected Is Nothing Then
            blnAskPring = True
        Else
            blnAskPring = Not (Val(tbPage.Selected.Tag) = mPage.pag_�����嵥)
        End If
        Call mfrm�����嵥.zlPrintBill(False, "", 1, mstrPrivs, blnAskPring)
    Case conMenu_File_Parameter:
        '��������
        If frmPayExitParaSet.ShowSetPara(Me, mlngModule, mstrPrivs) = False Then Exit Sub
        Call initLocalPara
        Set mArrFilter = mfrmFilter.GetFilterCon
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_NewItem:       '����
        Set mfrmδ����.In_PlugIn = mobjPlugIn
        If mfrmδ����.zlPayStuff = False Then Exit Sub
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Manage_Discard:  '�ܷ�
        '�ܷ�
        If Save�ܷ� = False Then Exit Sub
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Manage_Recall  '�ָ�
        If mfrm�ܷ��嵥.zlRestorePayStuff = False Then Exit Sub
        Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Manage_Untread     '����
        Set mfrm�����嵥.In_PlugIn = mobjPlugIn
       If mfrm�����嵥.zlBackPayStuff = False Then Exit Sub
       Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_Edit_ChargeDelAudit    '����
        Set frm��������.In_PlugIn = mobjPlugIn
        If frm��������.ShowList(Me, mstrPrivs, mlngModule, Val(mArrFilter("���ϲ���ID")), mintUnit) = False Then Exit Sub
        
    Case conMenu_Edit_CfPay    '����������
        Call SendBillPay(False)
    Case conMenu_Edit_BillPay    '��Ʊ�ݺŷ���
        Call SendBillPay(True)
    Case conMenu_Edit_BillBackPay    '����������
        Call SendBackPay
    Case conMenu_Edit_OtherPay       '�������ⷿ����
        Call SendOtherPay
    Case conMenu_Edit_StopPay    'ֹͣ���ϱ��
        Call StopPayStuffFlag
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each mcbrControl In Me.cbsThis(2).Controls
            mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_FontSize_1, conMenu_View_FontSize_2, conMenu_View_FontSize_3
        mint�ֺ� = Val(Control.Parameter)
        Call SetFontSize
        Call zlDatabase.SetPara("�����ֺ�", mint�ֺ�, glngSys, mlngModule)
    Case conMenu_View_Refresh   'ˢ��
        mstrSelectTabItem = "," & Val(tbPage.Selected.Tag)
       Set mArrFilter = mfrmFilter.GetFilterCon
       Call mfrmFilter_zlRefreshCon(mArrFilter)
    Case conMenu_View_Column '������
        Call LoadFulltoColSel
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
'    Case conMenu_Popu_סԺ��
'        mfrmFilter.PatiTittle = 0
'    Case conMenu_Popu_����
'        mfrmFilter.PatiTittle = 1
'    Case conMenu_Popu_����
'        mfrmFilter.PatiTittle = 2
'    Case conMenu_Popu_����ID    '
'        mfrmFilter.PatiTittle = 3
'    Case conMenu_Popu_�����
'        mfrmFilter.PatiTittle = 4
'    Case conMenu_Popu_���￨��
'        mfrmFilter.PatiTittle = 5
'    Case conMenu_Popu_IC����
'        mfrmFilter.PatiTittle = 6
    Case Else
        If Control.Id > 401 And Control.Id < 499 Then
            '��ر���ִ��
            Call OpenRpt(Control)
        End If
        
        '�����˵�
        If Control.Id >= conMenu_Popu_סԺ�� And Control.Id <= conMenu_Popu_סԺ�� + 6 + gintCardCount Then
            mint����ģʽ = Control.Id - conMenu_Popu_סԺ��
'            mfrmFilter.PatiTittle = mint����ģʽ
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub SetFontSize()
    '-----------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-06 17:24:45
    '-----------------------------------------------------------------------------------------------------------
    Dim curFontSize As Currency
    Dim stdfnt As StdFont
    
    curFontSize = Decode(mint�ֺ�, 1, 11, 2, 15, 9)
    mfrm���ϻ���.zlSetFontSize curFontSize
    mfrm�ܷ��嵥.zlSetFontSize curFontSize
    mfrmȱ���嵥.zlSetFontSize curFontSize
    mfrm�����嵥.zlSetFontSize curFontSize
    mfrmδ����.zlSetFontSize curFontSize
     If Not tbPage.PaintManager.Font Is Nothing Then
        With tbPage
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = curFontSize
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = curFontSize
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim sngStatusHeight As Single
    On Error Resume Next
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbsThis.GetClientRect Left, Top, Right, Bottom '
    With picList
        .Left = Left
        .Top = Top
        .Width = Right - Left
        .Height = Bottom - Top
    End With
End Sub
Private Function ISHaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�Ƿ�����ص�����
    '���:
    '����:
    '����:��,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-02 00:22:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim vsList As VSFlexGrid
    Dim lngRow As Long
    err = 0: On Error GoTo ErrHand:
    If tbPage.Selected Is Nothing Then Exit Function
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_���ܷ���
       ISHaveData = mfrm���ϻ���.zlHaveData
    Case mPage.pag_�ܷ��嵥
       ISHaveData = mfrm�ܷ��嵥.zlHaveData
    Case mPage.pag_δ���嵥
       ISHaveData = mfrmδ����.zlHaveData
    Case mPage.pag_ȱ���嵥
       ISHaveData = mfrmȱ���嵥.zlHaveData
    Case mPage.pag_�����嵥
       ISHaveData = mfrm�����嵥.zlHaveData
    End Select
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Ȩ�޿���(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'Ȩ�޿���
  
  Select Case Control.Id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
    Case conMenu_File_BillPrint, conMenu_File_BillPrintView
    Case conMenu_File_Parameter:
        '��������
       ' Control.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��������")
    Case conMenu_Edit_NewItem, conMenu_Edit_CfPay, conMenu_Edit_BillPay:     '����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������Ϸ���")
    Case conMenu_Edit_StopPay
         Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ֹͣ����") Or zlStr.IsHavePrivs(mstrPrivs, "�ָ�����")
    Case conMenu_Manage_Discard:  '�ܷ�
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������Ͼܷ�")
    Case conMenu_Manage_Recall  '�ָ�����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������ϻָ�")
    Case conMenu_Manage_Untread, conMenu_Edit_BillBackPay    '����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "������������")
    Case conMenu_EditPopup
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�������Ϸ���") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "������������") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "�������Ͼܷ�") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "�������ϻָ�") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "������������") Or _
                          zlStr.IsHavePrivs(mstrPrivs, "������������")
    Case conMenu_Edit_ChargeDelAudit    '����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "������������")
    End Select
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '���ÿؼ����������
   Call Ȩ�޿���(Control)
   Select Case Control.Id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = ISHaveData
    Case conMenu_File_BillPrintView ' conMenu_File_BillPrint,
        If tbPage.Selected Is Nothing Then
            Control.Enabled = False
        Else
            Control.Enabled = Val(tbPage.Selected.Tag) = mPage.pag_�����嵥
        End If
    Case conMenu_File_Parameter:
        '��������
       ' Control.Enabled = zlStr.IsHavePrivs(mstrPrivs, "��������")
    Case conMenu_Edit_NewItem:       '����
        Control.Enabled = mfrmδ����.zlHaveSel���� And (Val(tbPage.Selected.Tag) = mPage.pag_δ���嵥 Or Val(tbPage.Selected.Tag) = mPage.pag_���ܷ���)
    Case conMenu_Manage_Discard:  '�ܷ�
        '        mfrmδ����.zl
        Control.Enabled = mfrmδ����.zlHaveSel�ܷ� And (Val(tbPage.Selected.Tag) = mPage.pag_δ���嵥)
    Case conMenu_Manage_Recall
        '�ָ�
        Control.Enabled = mfrm�ܷ��嵥.zlHaveSel�ָ� And (Val(tbPage.Selected.Tag) = mPage.pag_�ܷ��嵥)
    Case conMenu_Manage_Untread     '����
        Control.Enabled = mfrm�����嵥.zlHaveSel���� And (Val(tbPage.Selected.Tag) = mPage.pag_�����嵥)
    Case conMenu_Edit_ChargeDelAudit    '����
    Case conMenu_View_FontSize_1, conMenu_View_FontSize_2, conMenu_View_FontSize_3
        Control.Checked = Val(Control.Parameter) = mint�ֺ�
    
    End Select
End Sub

Private Sub Chk�嵥_Click()
     mfrm�����嵥.zl��ʾ�����̵��� = Chk�嵥.Value = 1
End Sub

Private Sub Form_Activate()
    If mfrmFilter.CheckDept = False Then
        ShowMsgBox "����Ӧ������һ�����з��ϲ������ʻ���" & vbCrLf & "�㲻�Ƿ��ϲ��ŵĹ�����Ա,��鿴���Ź���"
        Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    mlngModule = glngModul

    'һ��ͨ�ӿ�
    On Error Resume Next
    Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not gobjSquareCard Is Nothing Then
        If gobjSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle) = False Then
            Set gobjSquareCard = Nothing
        Else
            gstrCardType = gobjSquareCard.zlGetIDKindStr
            
            'ȡ�����￨������֮������ѿ�
            gstrCardType = Mid(gstrCardType, InStr(1, gstrCardType, "��|���￨"))
        End If
    End If
    
    err.Clear: On Error GoTo 0
    
    Call initLocalPara
    Call InitComandBars
    Call InitPanel
    Call InitPage
    Call SetFontSize
    RestoreWinState Me, App.ProductName
    '2008-03-12:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '�ָ�¼��״̬
'    mfrmFilter.PatiTittle = mint����ģʽ

    '��ҩҵ����Ҳ���
    Call zlPlugIn_Ini(glngSys, mlngModule, mobjPlugIn)
End Sub


Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case ID_PANE_SEARCH
        Item.Handle = mfrmFilter.hwnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnTrans = False
    
    'ж��һ��ͨ�ӿ�
    gstrCardType = ""
    Set gobjSquareCard = Nothing
    
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter
    If Not mfrm���ϻ��� Is Nothing Then Unload mfrm���ϻ���
    If Not mfrm�ܷ��嵥 Is Nothing Then Unload mfrm�ܷ��嵥
    If Not mfrmȱ���嵥 Is Nothing Then Unload mfrmȱ���嵥
    If Not mfrm�����嵥 Is Nothing Then Unload mfrm�����嵥
    If Not mfrmδ���� Is Nothing Then Unload mfrmδ����
    
    Set mfrmFilter = Nothing
    Set mfrm���ϻ��� = Nothing
    Set mfrm�ܷ��嵥 = Nothing
    Set mfrmȱ���嵥 = Nothing
    Set mfrm�����嵥 = Nothing
    Set mfrmδ���� = Nothing
    mstrNo = "": mstrStuffStartDate = "": mstrStuffEndDate = "": mlng����id = 0: mlng�ⷿid = 0
    
    Call SaveWinState(Me, App.ProductName)
    
    '��������ģʽ
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "���ķ��Ź���", "����ģʽ", mint����ģʽ)
    If mblnҳǩ Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\���ķ��Ź���", "��ǰҳǩ", tbPage.Selected.Index)
    End If
    
    'ж����ҽӿ�
    Call zlPlugIn_Unload(mobjPlugIn)
End Sub

Private Sub mfrmFilter_zlPopupMenus(ByVal x As Long, ByVal Y As Long)
    '�����˵�
'    Dim intType As Integer
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim intCount As Integer
    Dim strCardName As String
    
  '  If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
'    mint����ģʽ = mfrmFilter.PatiTittle
    With cbrPopupBar
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_סԺ��, "סԺ��(&A)")
        If mint����ģʽ = 0 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_����, "����(&C)")
        If mint����ģʽ = 1 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_����, "����(&N)")
        If mint����ģʽ = 2 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_����ID, "����ID(&I)")
        If mint����ģʽ = 3 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_�����, "�����(&M)")
        If mint����ģʽ = 4 Then cbrPopupItem.Checked = True
        Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_IC����, "IC����(&K)")
        If mint����ģʽ = 5 Then cbrPopupItem.Checked = True
        
        '��̬ȡ�������ѿ�
        If gstrCardType <> "" Then
            gintCardCount = UBound(Split(gstrCardType, ";")) + 1
            For intCount = 0 To UBound(Split(gstrCardType, ";"))
                'ȡ���п�����
                strCardName = Split(Split(gstrCardType, ";")(intCount), "|")(1)
                
                Set cbrPopupItem = .Controls.Add(xtpControlButton, conMenu_Popu_IC���� + intCount + 1, strCardName & "(&" & intCount + 1 & ")")
                
                If mint����ģʽ = conMenu_Popu_IC���� + intCount + 1 Then
                    cbrPopupItem.Checked = True
                End If
                
                '���濨��Ϣ
                cbrPopupItem.Parameter = Split(gstrCardType, ";")(intCount)
                
                If intCount = 0 Then
                    cbrPopupItem.BeginGroup = True
                End If
            Next
        End If
        
    End With
    cbrPopupBar.ShowPopup
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    
    '���������˸ı�
    Set mfrm�����嵥.zlArrFilter = mArrFilter
    Call mfrmδ����.zlFullData(Me, mstrPrivs, mlngModule, mintUnit, arrFilter)
    Call mfrm�ܷ��嵥.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_�����嵥
        mfrm�����嵥.zlRefreshData Me, mstrPrivs, mlngModule, mintUnit, mArrFilter
    Case mPage.pag_�ܷ��嵥
        mfrm�ܷ��嵥.zlFullData mrsNotPayStuff
    Case mPage.pag_���ܷ���
        If mfrm���ϻ���.zlFullData(mintUnit, mrsNotPayStuff, mrsChargeOff) = False Then Exit Sub
    Case Else
    End Select
    
End Sub

Private Sub mfrm�����嵥_zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset)
    Set mrsBakStuff = rsNotStuffStuff
    stbThis.Panels(2).Text = "����" & mrsBakStuff.RecordCount & "����¼,�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
End Sub

Private Sub mfrmδ����_zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset, ByVal rsChargeOff As ADODB.Recordset)
    Set mrsNotPayStuff = rsNotStuffStuff
    Set mrsChargeOff = rsChargeOff
    stbThis.Panels(2).Text = "����" & mrsNotPayStuff.RecordCount & "�������ϼ�¼,�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
End Sub

Private Sub picList_Resize()
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Width = .ScaleWidth
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
        Chk�嵥.Top = tbPage.Top
        Chk�嵥.Left = .ScaleWidth - Chk�嵥.Width - 100
    End With
End Sub

Private Function Save�ܷ�() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�ܷ�δ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-24 11:59:21
    '-----------------------------------------------------------------------------------------------------------
    Dim cllProc As Collection
    Set cllProc = New Collection
    
    With mrsNotPayStuff
        .Filter = "ִ��״̬=2"
        If .RecordCount = 0 Then
            ShowMsgBox "��������صľܷ��������ѡ��ܷ�����,������ֹ!"
            Exit Function
        End If
        .Sort = "����id Asc"
        .MoveFirst
        Do While Not .EOF '
            If !ִ��״̬ = 2 Then
                'Zl_�������Ϸ���_�ܷ�(Id_In In �����շ���¼.ID%Type)
                gstrSQL = "Zl_�������Ϸ���_�ܷ�(" & NVL(!Id) & ")"
                AddArray cllProc, gstrSQL
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, Me.Caption
    Save�ܷ� = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    
End Function
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, ObjAppRow As zlTabAppRow
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_δ���嵥
        Set objPrint.Body = mfrmδ����.vsGrid
    Case mPage.pag_���ܷ���
        Set objPrint.Body = mfrm���ϻ���.vsGrid
    Case mPage.pag_�ܷ��嵥
        Set objPrint.Body = mfrm�ܷ��嵥.vsGrid
    Case mPage.pag_ȱ���嵥
        Set objPrint.Body = mfrmȱ���嵥.vsGrid
    Case mPage.pag_�����嵥
        Set objPrint.Body = mfrm�����嵥.vsGrid
    Case Else
        Exit Sub
    End Select
    
    objPrint.Title.Text = tbPage.Selected.Caption & "���"
    Set ObjAppRow = New zlTabAppRow
    Call ObjAppRow.Add("��ӡ��:" & gstrUserName)
    Call ObjAppRow.Add("��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD��"))
    Call objPrint.BelowAppRows.Add(ObjAppRow)
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Function OpenRpt(ByVal Control As XtremeCommandBars.ICommandBarControl) As Boolean
    '------------------------------------------------------------------------------
    '����:�򿪱���
    '����:Control-ִ�б���Ŀؼ�
    '����:
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim arrData As Variant
    Dim strNo As String, intRecodeSta As Integer, lng���ϲ���ID As Long, lngλ�� As Long
    Dim lng����ID As Long, lng����ID As Long, strסԺ�� As String, lng���� As Long
    
    'CStr(gobjComLib.zlCommFun.NVL(rsTmp!ϵͳ, 0) &  "," & rsTmp!���)
    arrData = Split(Control.Parameter, ",")
    'Set mrs�ܷ� = zldatabase.OpenSQLRecord(gstrsql, Me.Caption, _
           Val(mArrFilter("���ϲ���ID")), _
           CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
           CStr("," & mArrFilter("����") & ","), _
           Val(mArrFilter("��������ID")), _
           CStr(mArrFilter("���ݺ�")(0)), CStr(mArrFilter("���ݺ�")(1)), _
           Val(mArrFilter("����ID")), Val(mArrFilter("סԺ��")), _
           CStr(mArrFilter("����")))
    
    strסԺ�� = Val(mArrFilter("סԺ��"))
    lng���ϲ���ID = Val(mArrFilter("��������ID"))
        
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_δ���嵥
        With mfrmδ����.vsGrid
            lngλ�� = Val(.Cell(flexcpData, .Row, .ColIndex("���ݺ�")))
            mrsNotPayStuff.Find "λ��=" & lngλ��
            With mrsNotPayStuff
                If Not mrsNotPayStuff.EOF Then
                    lng����ID = Val(NVL(!����ID))
                    intRecodeSta = 1
                    lng����ID = Val(NVL(!����ID))
                    strסԺ�� = NVL(!סԺ��)
                    lng���� = NVL(!����)
                End If
            End With
        End With
    Case mPage.pag_���ܷ���
        With mfrm���ϻ���.vsGrid
            lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("��������")))
        End With
    Case mPage.pag_�ܷ��嵥
        With mfrm�ܷ��嵥.vsGrid
            lng���� = Val(.Cell(flexcpData, .Row, .ColIndex("��������")))
            lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("��������")))
            strסԺ�� = .TextMatrix(.Row, .ColIndex("סԺ��"))
            lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("���ݺ�")))
            intRecodeSta = 1
        End With
    Case mPage.pag_ȱ���嵥
        With mfrmȱ���嵥.vsGrid
            lngλ�� = Val(.Cell(flexcpData, .Row, .ColIndex("���ݺ�")))
            mrsNotPayStuff.Find "λ��=" & lngλ��
            With mrsNotPayStuff
                If Not mrsNotPayStuff.EOF Then
                    lng����ID = Val(NVL(!����ID))
                    intRecodeSta = 1
                    lng����ID = Val(NVL(!����ID))
                    strסԺ�� = NVL(!סԺ��)
                    lng���� = NVL(!����)
                End If
            End With
        End With
    Case mPage.pag_�����嵥
        With mfrm�����嵥.vsGrid
            lngλ�� = Val(.Cell(flexcpData, .Row, .ColIndex("���ݺ�")))
            mrsBakStuff.Find "λ��=" & lngλ��
            With mrsBakStuff
                If Not mrsBakStuff.EOF Then
                    lng����ID = Val(NVL(!����ID))
                    intRecodeSta = 1
                    lng����ID = Val(NVL(!����ID))
                    strסԺ�� = NVL(!סԺ��)
                    lng���� = NVL(!����)
                End If
            End With
        End With
    End Select
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Val(arrData(0)), arrData(1), Me, "NO=" & strNo, "��¼״̬=" & intRecodeSta, _
        "���ϲ���=" & Val(mArrFilter("���ϲ���ID")), "��������=" & lng����, _
        "����=" & lng����ID, "����=" & lng����ID, "סԺ��=" & strסԺ��, _
        "��ʼ����=" & mArrFilter("���ڷ�Χ")(0), "��������=" & mArrFilter("���ڷ�Χ")(1))
End Function
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
        
        Chk�嵥.Visible = False
        stbThis.Panels(2) = ""
        Select Case Val(Item.Tag)
        Case mPage.pag_���ܷ���
            If mfrm���ϻ���.zlFullData(mintUnit, mrsNotPayStuff, mrsChargeOff) = False Then Exit Sub
            stbThis.Panels(2).Text = "�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
        Case mPage.pag_�ܷ��嵥
            If mfrm�ܷ��嵥.zlFullData(mrsNotPayStuff) = False Then Exit Sub
            stbThis.Panels(2).Text = "�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
        Case mPage.pag_ȱ���嵥
            If mfrmȱ���嵥.zlFullData(mintUnit, mrsNotPayStuff) = False Then Exit Sub
            stbThis.Panels(2).Text = "�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
        Case mPage.pag_�����嵥
            Chk�嵥.Visible = True
            Call mfrm�����嵥.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
            If mrsBakStuff Is Nothing Then Exit Sub
            If mrsBakStuff.State = 1 Then
                stbThis.Panels(2).Text = "����" & mrsBakStuff.RecordCount & "������¼ ,�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
            End If
        Case mPage.pag_δ���嵥
            If mrsNotPayStuff Is Nothing Then Exit Sub
            If mrsNotPayStuff.State = 1 Then
                stbThis.Panels(2).Text = "����" & mrsNotPayStuff.RecordCount & "�������ϼ�¼,�ϴλ��ܷ��Ϻ�Ϊ:" & mfrmδ����.zl_�ϴλ��ܷ��Ϻ�
            End If
        End Select
        
End Sub


Private Function GetCheckPara() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-06 15:02:49
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mArrFilter("���ϲ���id")))
    With rsTemp
        If Not .EOF Then
            GetCheckPara = NVL(!�����, 0)
        End If
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function LoadFulltoColSel() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-09 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    Dim sngFrmHeight As Single, sngSelSumHeight As Single
    
    Select Case Val(Me.tbPage.Selected.Tag)
    Case mPage.pag_���ܷ���
        Set vsGrid = mfrm���ϻ���.vsGrid
    Case mPage.pag_�ܷ��嵥
        Set vsGrid = mfrm�ܷ��嵥.vsGrid
    Case mPage.pag_ȱ���嵥
        Set vsGrid = mfrmȱ���嵥.vsGrid
    Case mPage.pag_�����嵥
        Set vsGrid = mfrm�����嵥.vsGrid
    Case mPage.pag_δ���嵥
        Set vsGrid = mfrmδ����.vsGrid
    End Select
    vsColSet.Clear 1
    vsColSet.Rows = 2
    With vsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If Trim(.ColKey(i)) <> "" And (.ColData(i) = 1 Or .ColData(i) = 0) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("����")) = .ColKey(i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("ѡ��")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = .ColData(i)
                If .ColData(i) = 1 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.Rows = vsColSet.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.Rows > 2 Then vsColSet.Rows = vsColSet.Rows - 1
    SetParent vsColSet.hwnd, vsGrid.Parent.hwnd
    sngFrmHeight = vsGrid.Parent.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.Rows) + 60
        
        .Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000001
        .Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000005
        .BackColorSel = &H8000000D
        .Row = 1
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .Left = vsGrid.Left + .Cell(flexcpWidth, 0, 0, 0, 0) + 30
        .Top = vsGrid.Top + vsGrid.RowHeight(0) + 15
        sngFrmHeight = sngFrmHeight - .Top
        
        If sngFrmHeight > sngSelSumHeight Then
            .Height = sngSelSumHeight
        Else
            .Height = IIf(sngFrmHeight < 0, 0, sngFrmHeight)
        End If
        .SetFocus
    End With
End Function
Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-09 17:31:22
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    Select Case Val(Me.tbPage.Selected.Tag)
    Case mPage.pag_���ܷ���
        Set vsGrid = mfrm���ϻ���.vsGrid
    Case mPage.pag_�ܷ��嵥
        Set vsGrid = mfrm�ܷ��嵥.vsGrid
    Case mPage.pag_ȱ���嵥
        Set vsGrid = mfrmȱ���嵥.vsGrid
    Case mPage.pag_�����嵥
        Set vsGrid = mfrm�����嵥.vsGrid
    Case mPage.pag_δ���嵥
        Set vsGrid = mfrmδ����.vsGrid
    End Select
    With vsGrid
        
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
    End With
    
End Function
Private Sub vsColSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�޸ĺ�
    Dim strColKey As String, blnShow As Boolean
    With vsColSet
        Select Case Col
        Case .ColIndex("ѡ��")
            blnShow = GetVsGridBoolColVal(vsColSet, Row, .ColIndex("ѡ��"))
            Call SetVsGridCol(.TextMatrix(Row, .ColIndex("����")), blnShow)
        Case Else
        End Select
    End With
End Sub
Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("ѡ��")
            'rowdata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If .RowData(Row) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vsColSet_LostFocus()
    vsColSet.Visible = False
End Sub

Private Function SendOtherPay() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-06 14:55:11
    '-----------------------------------------------------------------------------------------------------------
    With frm������
        .In_���� = 0
        .In_����IN = mArrFilter("����")
        .In_���ϲ���id = Val(mArrFilter("���ϲ���ID"))
        .In_����� = GetCheckPara()
        .In_����δ���Ϸ��� = 1
        .In_Ȩ�� = mstrPrivs
        .mstr������ = gstrUserName
        Set .In_PlugIn = mobjPlugIn
        .Show 1, Me
    End With
    SendOtherPay = True
    Call mfrmFilter_zlRefreshCon(mArrFilter)
    
End Function
