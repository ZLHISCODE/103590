VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.dockingpane.9600.ocx"
Begin VB.Form frmGroupAndPesons 
   Caption         =   "�ɿ���Ա����"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frmGroupAndPesons.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11745
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList igl16 
      Left            =   6075
      Top             =   240
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
            Picture         =   "frmGroupAndPesons.frx":058A
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":0B24
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":10BE
            Key             =   "Group"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPersons 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   375
      ScaleHeight     =   5685
      ScaleWidth      =   7185
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3360
      Width           =   7185
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   555
         TabIndex        =   13
         Top             =   435
         Width           =   3870
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   555
         TabIndex        =   11
         Top             =   75
         Width           =   3900
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "����"
         Height          =   300
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   60
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "�Ƴ�"
         Height          =   300
         Index           =   4
         Left            =   5175
         TabIndex        =   15
         Top             =   60
         Width           =   570
      End
      Begin MSComctlLib.ListView lvwPerson 
         Height          =   6435
         Left            =   0
         TabIndex        =   16
         Top             =   915
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   11351
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "igl32"
         SmallIcons      =   "igl16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "���"
            Object.Tag             =   "���"
            Text            =   "���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "��������"
            Object.Tag             =   "��������"
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "�Ա�"
            Object.Tag             =   "�Ա�"
            Text            =   "�Ա�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "�칫�ҵ绰"
            Object.Tag             =   "�칫�ҵ绰"
            Text            =   "�칫�ҵ绰"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "�����ʼ�"
            Object.Tag             =   "�����ʼ�"
            Text            =   "�����ʼ�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "����ְ��"
            Object.Tag             =   "����ְ��"
            Text            =   "����ְ��"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ��"
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   12
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   135
         Width           =   360
      End
   End
   Begin VB.PictureBox picGroup 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   465
      ScaleHeight     =   7860
      ScaleWidth      =   4935
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton cmdFucn 
         Caption         =   "ɾ��"
         Height          =   300
         Index           =   2
         Left            =   3900
         TabIndex        =   8
         Top             =   855
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "�޸�"
         Height          =   300
         Index           =   1
         Left            =   3285
         TabIndex        =   7
         Top             =   855
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "����"
         Height          =   300
         Index           =   0
         Left            =   2685
         TabIndex        =   6
         Top             =   855
         Width           =   570
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   645
         TabIndex        =   3
         Top             =   465
         Width           =   3900
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   645
         TabIndex        =   5
         Top             =   855
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   645
         TabIndex        =   1
         Top             =   75
         Width           =   3900
      End
      Begin MSComctlLib.ListView lvwGroups 
         Height          =   6510
         Left            =   60
         TabIndex        =   9
         Top             =   1275
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   11483
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "igl32"
         SmallIcons      =   "igl16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "������"
            Object.Tag             =   "������"
            Text            =   "������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "������"
            Object.Tag             =   "������"
            Text            =   "������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "˵��"
            Object.Tag             =   "˵��"
            Text            =   "˵��"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   2
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   915
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   135
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   8085
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmGroupAndPesons.frx":1658
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
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
   Begin MSComctlLib.ImageList igl32 
      Left            =   6930
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":1EEC
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":27C6
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroupAndPesons.frx":30A0
            Key             =   "Group"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStructure 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   4260
      ScaleHeight     =   5685
      ScaleWidth      =   7185
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1215
      Width           =   7185
      Begin MSComctlLib.ListView lvwStructure 
         Height          =   6435
         Left            =   0
         TabIndex        =   24
         Top             =   555
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   11351
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "igl32"
         SmallIcons      =   "igl16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "���"
            Object.Tag             =   "���"
            Text            =   "���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "��������"
            Object.Tag             =   "��������"
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "�Ա�"
            Object.Tag             =   "�Ա�"
            Text            =   "�Ա�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "����"
            Object.Tag             =   "����"
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "�칫�ҵ绰"
            Object.Tag             =   "�칫�ҵ绰"
            Text            =   "�칫�ҵ绰"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "�����ʼ�"
            Object.Tag             =   "�����ʼ�"
            Text            =   "�����ʼ�"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "����ְ��"
            Object.Tag             =   "����ְ��"
            Text            =   "����ְ��"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "�Ƴ�"
         Height          =   300
         Index           =   6
         Left            =   5160
         TabIndex        =   23
         Top             =   45
         Width           =   570
      End
      Begin VB.CommandButton cmdFucn 
         Caption         =   "����"
         Height          =   300
         Index           =   5
         Left            =   4560
         TabIndex        =   22
         Top             =   60
         Width           =   570
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   555
         TabIndex        =   21
         Top             =   75
         Width           =   3900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�鳤"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   135
         Width           =   360
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmGroupAndPesons.frx":397A
      Left            =   555
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmGroupAndPesons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************************************************************
'����:�ɿ���Ա����
'����:���˺�
'����:2010-11-23 15:42:13
'˵��:
'    33633
'********************************************************************************************************************************************
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private WithEvents mfrmFilter As frmBillInFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mlngModul As Long, mstrPrivs As String
Private mblnFirst As Boolean  '��һ�μ��ش���
Private mstrKey As String, mstrPreGroupKey As String    '��һ�εļ�¼
Private Enum mPaneID
    Pane_Group = 1    '
    Pane_Persons = 3
    Pane_Structure = 2
End Enum
Private mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Private mintSucess As Integer '>0��ʾֻ�ٸ�����һ��ֵ��
Private mintGroupColumn As Integer, mintPersonColumn As Integer
Private mintStructureColumn As Integer
Private mblnEdit As Boolean '�Ƿ��ڱ༭״̬
Private mblnStartDrop As Boolean '��ʼ�϶�
Private mstrSelect As String
Private mblnReSel As Boolean '�Ƿ�����ѡ��
Private mblnItemClick As Boolean  '�Ƿ�㵽��Ŀ
Private Enum mTxtIdx
    idx_������ = 0
    idx_������ = 1
    idx_��˵�� = 2
    idx_��Ա = 3
    idx_ԭ�� = 4
    idx_�鳤 = 6
End Enum
Private Enum mCmdIdx
    idx_������ = 0
    idx_���޸� = 1
    idx_��ɾ�� = 2
    idx_��Ա���� = 3
    idx_��Ա�Ƴ� = 4
    idx_�鳤���� = 5
    idx_�鳤ɾ�� = 6
End Enum
Private Sub cmdFucn_Click(Index As Integer)
    Select Case Index
    Case mCmdIdx.idx_������
        Call AddGroups(0)
    Case mCmdIdx.idx_���޸�
        Call AddGroups(1)
    Case mCmdIdx.idx_��ɾ��
        Call DeleteGroup
    Case mCmdIdx.idx_��Ա����
        Call AddPerson
    Case mCmdIdx.idx_��Ա�Ƴ�
        If lvwPerson.SelectedItem Is Nothing Then Exit Sub
        If lvwGroups.SelectedItem Is Nothing Then Exit Sub
        Call PersonFromGroupToOtherGroup(Mid(lvwPerson.SelectedItem.Key, 2), Trim(txtEdit(mTxtIdx.idx_��Ա)), _
            Val(txtEdit(mTxtIdx.idx_ԭ��).Tag), Trim(txtEdit(mTxtIdx.idx_ԭ��)))
    Case mCmdIdx.idx_�鳤����
        Call AddStructure
    Case mCmdIdx.idx_�鳤ɾ��
        If lvwStructure.SelectedItem Is Nothing Then Exit Sub
        If lvwGroups.SelectedItem Is Nothing Then Exit Sub
        Call DeleteStructure(Mid(lvwStructure.SelectedItem.Key, 2), Trim(lvwStructure.SelectedItem.Text))
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mstrPreGroupKey = ""
    Call LoadGroups
    RestoreListViewState lvwPerson, Me.Name, ""
    RestoreListViewState lvwGroups, Me.Name, ""
    RestoreListViewState lvwStructure, Me.Name, ""
End Sub
Public Function ShowGroups(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����ó�Ա����(�������)
    '���:frmMain-����
    '       lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:
    '����:���������һ��ɹ���,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-11-23 15:45:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModule: mstrPrivs = strPrivs: mintSucess = 0
    Me.Show 1, frmMain
    ShowGroups = mintSucess > 0
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    
    mblnFirst = True
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars '��ʼ�˵���������
    Call InitPanel
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    mblnEdit = zlStr.IsHavePrivs(mstrPrivs, "��Ա����")
    Call SetCtrlVisible
End Sub
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���visible����
    '����:���˺�
    '����:2010-11-23 16:38:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error Resume Next
    For i = 0 To txtEdit.UBound
        txtEdit(i).Visible = mblnEdit
    Next
    For i = 0 To lbl.UBound
        lbl(i).Visible = mblnEdit
    Next
    For i = 0 To cmdFucn.UBound
        cmdFucn(i).Visible = mblnEdit
    Next
End Sub
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_View_LargeICO   '"��ͼ��(&G)"
            Call SetIcoShow(0)
    Case conMenu_View_MinICO ' "Сͼ��(&M)")
            Call SetIcoShow(1)
    Case conMenu_View_ListICO  '"�б�(&L)"
            Call SetIcoShow(2)
    Case conMenu_View_DetailsICO '"��ϸ����(&D)"
            Call SetIcoShow(3)
    Case conMenu_View_Refresh   'ˢ��
        Call LoadGroups
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Function GetViewShow(ByVal bytShow As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�б�ǰ����ʾ��ʽ
    '���:bytShow(0-��ͼ��;1-Сͼ��;2-�б�;3-��ϸ����
    '����:
    '����:�����ʾ��ʽ���������ķ�ʽһ��,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-11-23 16:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.ActiveControl Is Me.lvwGroups Then
        GetViewShow = (lvwGroups.View = bytShow)
    Else
        GetViewShow = (lvwPerson.View = bytShow)
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
    Case conMenu_View_LargeICO   '"��ͼ��(&G)"
        Control.Checked = GetViewShow(0)
    Case conMenu_View_MinICO ' "Сͼ��(&M)")
        Control.Checked = GetViewShow(1)
    Case conMenu_View_ListICO  '"�б�(&L)"
        Control.Checked = GetViewShow(2)
    Case conMenu_View_DetailsICO '"��ϸ����(&
        Control.Checked = GetViewShow(3)
    Case conMenu_View_Refresh   'ˢ��
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
           ' Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1502" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1502"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.ID
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Parameter     '��������
        Case Else   '�����������ܵ���
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If Me.Visible = False Then Exit Sub

    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub
Private Sub Form_Initialize()
  Call InitCommonControls
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTemp As String
    Err = 0: On Error Resume Next
   SaveWinState Me, App.ProductName
   zlSaveDockPanceToReg Me, dkpMan, "����"
   SaveListViewState lvwPerson, Me.Name, ""
   SaveListViewState lvwGroups, Me.Name, ""
   SaveListViewState lvwStructure, Me.Name, ""
End Sub

Private Sub lvwGroups_DragDrop(Source As Control, x As Single, y As Single)
    Dim objList As ListItem, strMoveID As String, strMoveName As String, i As Long
    If Source Is lvwPerson And Not lvwGroups.DropHighlight Is Nothing Then
        'Set lvwGroups.SelectedItem = lvwGroups.DropHighlight
        mblnStartDrop = False: mstrSelect = "": mblnReSel = False
        With lvwPerson
            strMoveID = "": strMoveName = "": i = 1
            For Each objList In .ListItems
                If objList.Selected Then
                    strMoveID = strMoveID & "," & Mid(objList.Key, 2)
                    If i > 3 Then
                       If i = 4 Then strMoveName = strMoveName & "..."
                    Else
                        strMoveName = strMoveName & "," & objList.Text
                    End If
                End If
            Next
            If strMoveName <> "" Then strMoveName = Mid(strMoveName, 2)
            If strMoveID <> "" Then strMoveID = Mid(strMoveID, 2)
            If strMoveID = "" Then Exit Sub
            Call PersonFromGroupToOtherGroup(strMoveID, strMoveName, _
                Mid(lvwGroups.SelectedItem.Key, 2), lvwGroups.SelectedItem.Text, _
                Mid(lvwGroups.DropHighlight.Key, 2), lvwGroups.DropHighlight.Text, False)
            Set lvwGroups.DropHighlight = Nothing
            lvwGroups.SelectedItem.EnsureVisible
            Call ClearDropVariable
        End With
    End If
End Sub

Private Sub lvwGroups_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim objOver As ListItem
    If Source Is lvwPerson Then
        Set objOver = lvwGroups.HitTest(x, y)
        If Not objOver Is Nothing Then
            If objOver.Key <> lvwGroups.SelectedItem.Key Then
                Set lvwGroups.DropHighlight = objOver
                lvwGroups.DropHighlight.EnsureVisible
            Else
                Set lvwGroups.DropHighlight = Nothing
            End If
        Else
            Set lvwGroups.DropHighlight = Nothing
        End If
    End If
End Sub

Private Sub lvwGroups_GotFocus()
    '
    Call SetGoupsEnable
End Sub

Private Sub lvwPerson_DragDrop(Source As Control, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False: mblnItemClick = False
End Sub

Private Sub lvwPerson_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim objList As ListItem
    If mblnReSel = True Then Exit Sub
    '��ѡʱ,Ҫ����ѡ��
    If Source Is lvwPerson Then
        mblnReSel = True
        If InStr(1, mstrSelect, ",") = 0 Then Exit Sub
        With lvwPerson
            For Each objList In .ListItems
                If InStr("," & mstrSelect & ",", "," & objList.Key & ",") > 0 And objList.Selected = False Then objList.Selected = True
            Next
        End With
    End If
End Sub

Private Sub lvwStructure_DragDrop(Source As Control, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False: mblnItemClick = False
End Sub

Private Sub lvwStructure_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim objList As ListItem
    If mblnReSel = True Then Exit Sub
    '��ѡʱ,Ҫ����ѡ��
    If Source Is lvwStructure Then
        mblnReSel = True
        If InStr(1, mstrSelect, ",") = 0 Then Exit Sub
        With lvwStructure
            For Each objList In .ListItems
                If InStr("," & mstrSelect & ",", "," & objList.Key & ",") > 0 And objList.Selected = False Then objList.Selected = True
            Next
        End With
    End If
End Sub

Private Sub lvwPerson_GotFocus()
    Call SetPersonEnable
End Sub

Private Sub lvwStructure_GotFocus()
    Call SetStructureEnable
End Sub

Private Sub lvwPerson_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtEdit(mTxtIdx.idx_��Ա).Text = Item.Text
    txtEdit(mTxtIdx.idx_��Ա).Tag = Mid(Item.Key, 2)
    txtEdit(mTxtIdx.idx_ԭ��).Text = Me.lvwGroups.SelectedItem.Text
    txtEdit(mTxtIdx.idx_ԭ��).Tag = Mid(Me.lvwGroups.SelectedItem.Key, 2)
    cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Tag = 1
    Call SetPersonEnable
    mblnItemClick = True
End Sub

Private Sub lvwStructure_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtEdit(mTxtIdx.idx_�鳤).Text = Item.Text
    txtEdit(mTxtIdx.idx_�鳤).Tag = Mid(Item.Key, 2)
    cmdFucn(mCmdIdx.idx_�鳤ɾ��).Tag = 1
    Call SetStructureEnable
    mblnItemClick = True
End Sub

Private Sub SetPersonEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ա�ı༭����
    '����:���˺�
    '����:2010-11-24 17:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDelete As Boolean
    blnDelete = Val(cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Tag) > 0
    cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Enabled = blnDelete
    cmdFucn(mCmdIdx.idx_��Ա����).Enabled = Not blnDelete
    txtEdit(mTxtIdx.idx_ԭ��).Enabled = False
    txtEdit(mTxtIdx.idx_ԭ��).BackColor = Me.BackColor
End Sub
Private Sub SetGoupsEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ı༭����
    '����:���˺�
    '����:2010-11-24 17:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnModify As Boolean
    blnModify = Not Me.lvwGroups.SelectedItem Is Nothing
    cmdFucn(mCmdIdx.idx_��ɾ��).Enabled = blnModify
    cmdFucn(mCmdIdx.idx_���޸�).Enabled = blnModify
    If Not blnModify Then
        cmdFucn(mCmdIdx.idx_������).Enabled = Trim(txtEdit(mTxtIdx.idx_������).Text) <> ""
    Else
        cmdFucn(mCmdIdx.idx_������).Enabled = Trim(txtEdit(mTxtIdx.idx_������).Text) <> "" And lvwGroups.SelectedItem.Text <> Trim(txtEdit(mTxtIdx.idx_������))
    End If
End Sub

Private Sub SetStructureEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ı༭����
    '����:���˺�
    '����:2010-11-24 17:50:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnModify As Boolean
    blnModify = Not Me.lvwStructure.SelectedItem Is Nothing
    cmdFucn(mCmdIdx.idx_�鳤ɾ��).Enabled = blnModify
    If Not blnModify Then
        cmdFucn(mCmdIdx.idx_�鳤����).Enabled = Trim(txtEdit(mTxtIdx.idx_�鳤).Text) <> ""
    Else
        cmdFucn(mCmdIdx.idx_�鳤����).Enabled = Trim(txtEdit(mTxtIdx.idx_�鳤).Text) <> "" And lvwStructure.SelectedItem.Text <> Trim(txtEdit(mTxtIdx.idx_�鳤))
    End If
End Sub

Private Sub lvwPerson_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False
End Sub

Private Sub lvwStructure_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnStartDrop = False: mstrSelect = "": mblnReSel = False
End Sub

Private Sub lvwPerson_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objList As ListItem, i As Long
    If Button <> 1 Then Exit Sub
    If mblnEdit = False Then Exit Sub '���ܱ༭��,�����϶�
    If mblnStartDrop Then Exit Sub
    If mblnItemClick = False Then Exit Sub
    
    '�϶���ʼ
    With lvwPerson
        If .ListItems.Count = 0 Then Exit Sub
        For Each objList In .ListItems
            If objList.Selected Then mstrSelect = mstrSelect & "," & objList.Key
        Next
    End With
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    If InStr(1, mstrSelect, ",") > 0 Then
        Set lvwPerson.DragIcon = igl32.ListImages("Group").Picture
    Else
        Set lvwPerson.DragIcon = lvwPerson.SelectedItem.CreateDragImage
    End If
    lvwPerson.Drag 1
    mblnStartDrop = True: mblnReSel = False
End Sub

Private Sub lvwStructure_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objList As ListItem, i As Long
    If Button <> 1 Then Exit Sub
    If mblnEdit = False Then Exit Sub '���ܱ༭��,�����϶�
    If mblnStartDrop Then Exit Sub
    If mblnItemClick = False Then Exit Sub
    
    '�϶���ʼ
    With lvwStructure
        If .ListItems.Count = 0 Then Exit Sub
        For Each objList In .ListItems
            If objList.Selected Then mstrSelect = mstrSelect & "," & objList.Key
        Next
    End With
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    If InStr(1, mstrSelect, ",") > 0 Then
        Set lvwStructure.DragIcon = igl32.ListImages("Group").Picture
    Else
        Set lvwStructure.DragIcon = lvwStructure.SelectedItem.CreateDragImage
    End If
    lvwStructure.Drag 1
    mblnStartDrop = True: mblnReSel = False
End Sub

Private Sub picGroup_Resize()
    Dim sngLeft As Single
    Dim sngTop As Single
    Err = 0: On Error Resume Next
    With picGroup
        
        sngLeft = .ScaleWidth - (txtEdit(mTxtIdx.idx_������).Width + txtEdit(mTxtIdx.idx_������).Left)
        sngLeft = sngLeft - (cmdFucn(mCmdIdx.idx_������).Width + 10) * 3
        sngTop = txtEdit(mTxtIdx.idx_������).Top
        If sngLeft < 0 Then
            sngLeft = txtEdit(mTxtIdx.idx_������).Left
            sngTop = txtEdit(mTxtIdx.idx_������).Top + txtEdit(mTxtIdx.idx_������).Height + 100
        ElseIf sngLeft < (txtEdit(mTxtIdx.idx_������).Width + txtEdit(mTxtIdx.idx_������).Left) Or sngLeft > (txtEdit(mTxtIdx.idx_������).Width + txtEdit(mTxtIdx.idx_������).Left) Then
                sngLeft = (txtEdit(mTxtIdx.idx_������).Width + txtEdit(mTxtIdx.idx_������).Left) + 100
        End If
        
        cmdFucn(mCmdIdx.idx_������).Left = sngLeft
        cmdFucn(mCmdIdx.idx_���޸�).Left = cmdFucn(mCmdIdx.idx_������).Left + cmdFucn(mCmdIdx.idx_������).Width + 10
        cmdFucn(mCmdIdx.idx_��ɾ��).Left = cmdFucn(mCmdIdx.idx_���޸�).Left + cmdFucn(mCmdIdx.idx_���޸�).Width + 10
        cmdFucn(mCmdIdx.idx_���޸�).Top = sngTop
        cmdFucn(mCmdIdx.idx_��ɾ��).Top = sngTop
        cmdFucn(mCmdIdx.idx_������).Top = sngTop
        txtEdit(mTxtIdx.idx_������).Width = .ScaleWidth - txtEdit(mTxtIdx.idx_������).Left - 100
        txtEdit(mTxtIdx.idx_��˵��).Width = .ScaleWidth - txtEdit(mTxtIdx.idx_��˵��).Left - 100
        If mblnEdit = False Then
             lvwGroups.Top = .ScaleTop
        Else
            lvwGroups.Top = sngTop + cmdFucn(mCmdIdx.idx_������).Height + 50
        End If
        lvwGroups.Left = .ScaleLeft
        lvwGroups.Width = .ScaleWidth
        lvwGroups.Height = .ScaleHeight - lvwGroups.Top - 50
    End With
End Sub

Private Sub picPersons_Resize()
    Dim sngLeft As Single, sngTop As Single
    Err = 0: On Error Resume Next
    With picPersons
        If .ScaleWidth - txtEdit(mTxtIdx.idx_��Ա).Left > 3900 Then
            txtEdit(mTxtIdx.idx_��Ա).Width = 3900
        Else
            txtEdit(mTxtIdx.idx_��Ա).Width = .ScaleWidth - txtEdit(mTxtIdx.idx_��Ա).Left - 50
        End If
        txtEdit(mTxtIdx.idx_ԭ��).Width = txtEdit(mTxtIdx.idx_��Ա).Width
        
        lvwPerson.Left = .ScaleLeft
        sngLeft = txtEdit(mTxtIdx.idx_ԭ��).Left + txtEdit(mTxtIdx.idx_ԭ��).Width + (cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Width + 10) * 2
        sngLeft = .ScaleWidth - sngLeft
        If sngLeft < 0 Then
            sngTop = txtEdit(mTxtIdx.idx_ԭ��).Top + txtEdit(mTxtIdx.idx_ԭ��).Height + 50
            sngLeft = txtEdit(mTxtIdx.idx_ԭ��).Left
        Else
            sngLeft = txtEdit(mTxtIdx.idx_ԭ��).Left + txtEdit(mTxtIdx.idx_ԭ��).Width + 50
            sngTop = txtEdit(mTxtIdx.idx_ԭ��).Top
        End If
        cmdFucn(mCmdIdx.idx_��Ա����).Left = sngLeft
        cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Left = cmdFucn(mCmdIdx.idx_��Ա����).Left + cmdFucn(mCmdIdx.idx_��Ա����).Width + 10
        cmdFucn(mCmdIdx.idx_��Ա����).Top = sngTop
        cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Top = sngTop
        sngTop = sngTop + cmdFucn(mCmdIdx.idx_��Ա����).Height + 50
        If mblnEdit Then
            lvwPerson.Top = sngTop
        Else
            lvwPerson.Top = .ScaleTop
        End If

        
        lvwPerson.Width = .ScaleWidth
        lvwPerson.Height = .ScaleHeight - lvwPerson.Top
    End With
End Sub
Private Sub SetIcoShow(ByVal bytShow As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͼ����ʾ��ʽ
    '���:bytShow-(0-��ͼ��;1- Сͼ��;2-�б�;3-��ϸ����
    '����:���˺�
    '����:2010-11-23 16:54:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLvw As Object
    If Me.ActiveControl Is Me.lvwGroups Then
        Set objLvw = lvwGroups
    Else
        Set objLvw = lvwPerson
    End If
    With objLvw
        .View = bytShow
    End With
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '���:
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-15 11:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
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
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��(&G)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����(&D)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2010-11-15 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, lngWidth As Long
    Dim lngHeight As Long
    With dkpMan
        Set objPane = .CreatePane(mPaneID.Pane_Group, 400, 400, DockLeftOf, Nothing)
        objPane.Title = "�ɿ���Ա������Ϣ": objPane.Options = PaneNoCloseable Or PaneNoHideable
        objPane.Handle = picGroup.hWnd
        objPane.Tag = mPaneID.Pane_Group
        
        Set objPane = .CreatePane(mPaneID.Pane_Structure, 400, 200, DockRightOf)
        objPane.Title = "�鳤������Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picStructure.hWnd
        objPane.Tag = mPaneID.Pane_Structure
        
        Set objPane = .CreatePane(mPaneID.Pane_Persons, 400, 400, DockBottomOf, objPane)
        objPane.Title = "���Ա��Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picPersons.hWnd
        objPane.Tag = mPaneID.Pane_Persons
        
        
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
'    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Function
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPaneID.Pane_Persons
        Item.Handle = picPersons.hWnd
    Case mPaneID.Pane_Group
        Item.Handle = picGroup.hWnd
    End Select
End Sub

Private Function LoadGroups() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ݸ�����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-15 14:54:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, lngPreID As Long, objItem As ListItem
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    
    Err = 0: On Error GoTo errHandle:
    gstrSQL = "" & _
    "   Select A.Id, A.������,A.����, A.˵��, A.������id, A.ɾ������,B.���� as ������  " & _
    "   From ����ɿ���� A,��Ա�� B " & _
    "   Where A.������ID=B.Id(+) And (A.ɾ������>Sysdate Or A.ɾ������ Is Null)"
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    With lvwGroups
        .ListItems.Clear
        Do While Not rsTemp.EOF
            Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!ID), NVL(rsTemp!������), "Group", "Group")
            objItem.SubItems(1) = NVL(rsTemp!������)
            objItem.SubItems(2) = NVL(rsTemp!˵��)
            objItem.Tag = NVL(rsTemp!������id)
            If mstrPreGroupKey = objItem.Key Then objItem.Selected = True: objItem.EnsureVisible
            rsTemp.MoveNext
        Loop
        If Not .SelectedItem Is Nothing Then
            Call lvwGroups_ItemClick(.SelectedItem)
        End If
    End With
    LoadGroups = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub LoadGroupStructure(ByVal lng��ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ա����
    '���:lng��ID-��ID
    '����:���˺�
    '����:2010-11-23 17:36:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey  As String, objItem As ListItem, i As Integer
    Dim strIco As String
    
    On Error GoTo ErrHand
    
    gstrSQL = " " & _
    "   Select A.��Id ,A.�鳤ID, B.���,B.����,B.����,b.��������,B.���֤��,B.�Ա�,B.����,B.�칫�ҵ绰,B.�����ʼ�,B.����ְ�� " & _
    "   From �������鳤���� A,��Ա�� B " & _
    "   Where A.�鳤ID=B.Id And A.��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng��ID)
    
   With lvwStructure
        If Not .SelectedItem Is Nothing Then strKey = .SelectedItem.Key
        .ListItems.Clear
        Do While Not rsTemp.EOF
            If InStr(1, NVL(rsTemp!�Ա�), "��") > 0 Then
                strIco = "Man"
            ElseIf InStr(1, NVL(rsTemp!�Ա�), "Ů") > 0 Then
                strIco = "Woman"
            Else
                strIco = "Man" ' "Other"
            End If
            Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!�鳤ID), NVL(rsTemp!����), strIco, strIco)
            i = 1
            objItem.SubItems(i) = NVL(rsTemp!���): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
            objItem.SubItems(i) = Format(rsTemp!��������, "yyyy-mm-dd"): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!�Ա�): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!�칫�ҵ绰): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!�����ʼ�): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!����ְ��): i = i + 1
            If strKey = objItem.Key Then
                objItem.Selected = True: objItem.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
        If Not .SelectedItem Is Nothing Then
            Call lvwStructure_ItemClick(.SelectedItem)
        End If
    End With
    mstrSelect = "": mblnItemClick = False: mblnStartDrop = False: mblnReSel = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub LoadGroupPersons(ByVal lng��ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ա����
    '���:lng��ID-��ID
    '����:���˺�
    '����:2010-11-23 17:36:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strKey  As String, objItem As ListItem, i As Integer
    Dim strIco As String
    
    On Error GoTo ErrHand
    
    gstrSQL = " " & _
    "   Select A.��Id ,A.��ԱID, B.���,B.����,B.����,b.��������,B.���֤��,B.�Ա�,B.����,B.�칫�ҵ绰,B.�����ʼ�,B.����ְ�� " & _
    "   From �ɿ��Ա��� A,��Ա�� B " & _
    "   Where A.��ԱID=B.Id And A.��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng��ID)
    
   With lvwPerson
        If Not .SelectedItem Is Nothing Then strKey = .SelectedItem.Key
        .ListItems.Clear
        Do While Not rsTemp.EOF
            If InStr(1, NVL(rsTemp!�Ա�), "��") > 0 Then
                strIco = "Man"
            ElseIf InStr(1, NVL(rsTemp!�Ա�), "Ů") > 0 Then
                strIco = "Woman"
            Else
                strIco = "Man" ' "Other"
            End If
            Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!��ԱID), NVL(rsTemp!����), strIco, strIco)
            i = 1
            objItem.SubItems(i) = NVL(rsTemp!���): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
            objItem.SubItems(i) = Format(rsTemp!��������, "yyyy-mm-dd"): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!�Ա�): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!�칫�ҵ绰): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!�����ʼ�): i = i + 1
            objItem.SubItems(i) = NVL(rsTemp!����ְ��): i = i + 1
            If strKey = objItem.Key Then
                objItem.Selected = True: objItem.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
        If Not .SelectedItem Is Nothing Then
            Call lvwPerson_ItemClick(.SelectedItem)
        End If
    End With
    mstrSelect = "": mblnItemClick = False: mblnStartDrop = False: mblnReSel = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
 Private Sub zlRptPrint(ByVal bytFunc As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2010-11-23 17:55:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrint As Object, objLvw As Object
    Dim objRow As New zlTabAppRow
    Dim str��λ As String
        
    Set objPrint = New zlPrintLvw
    If Me.ActiveControl Is lvwGroups Then
        objPrint.Title.Text = GetUnitName & "�����嵥"
        Set objLvw = lvwGroups
    Else
        If lvwGroups Is Nothing Then Exit Sub
        objPrint.Title.Text = GetUnitName & lvwGroups.SelectedItem.Text & "��Ա���"
        Set objLvw = lvwPerson
    End If
    Set objPrint.Body.objData = objLvw
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytFunc
    End If
End Sub
Private Sub DeleteGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ����
    '����:���˺�
    '����:2010-11-23 17:59:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTittle As String, intIndex As Integer
    Dim rsTemp As ADODB.Recordset
    With lvwGroups
        If .SelectedItem Is Nothing Then Exit Sub
        lngID = Val(Mid(.SelectedItem.Key, 2))
        strTittle = .SelectedItem.Text
    End With
    If lngID = 0 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ��������Ϊ�� " & strTittle & "���ķ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    
    gstrSQL = "" & _
    "   Select Count(distinct A.��ԱID) as ��Ա��,Sum(nvl(C.���,0)) as ��� " & _
    "   From  �ɿ��Ա��� A,��Ա�� B,��Ա�ɿ���� C " & _
    "   where A.��ԱID=B.id and B.����=C.�տ�Ա(+) and C.����(+)=1 and  A.��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If Val(NVL(rsTemp!��Ա��)) <> 0 Then
        If Val(NVL(rsTemp!���)) = 0 Then
            If MsgBox("������Ϊ�� " & strTittle & "���»���" & Val(NVL(rsTemp!��Ա��)) & "����Ա,���Ƿ�Ҫ��ɢ���飿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Me.MousePointer = 0
                Exit Sub
            End If
        Else
            If MsgBox("������Ϊ�� " & strTittle & "���»���" & Val(NVL(rsTemp!��Ա��)) & "��" & vbCrLf & "��Ա,���һ������ݴ��,���Ƿ�Ҫ��ɢ���飿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'Zl_����ɿ����_Delete(Id_In In ����ɿ����.ID%Type) Is
    gstrSQL = "Zl_����ɿ����_Delete(" & lngID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = 0
    With lvwGroups
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            lvwGroups_ItemClick .SelectedItem
        Else
            Call lvwGroups_GotFocus
        End If
    End With
    Call SetStructureEnable
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.MousePointer = 0
End Sub



Private Sub PersonFromGroupToOtherGroup(ByVal str��ԱID As String, ByVal str��Ա���� As String, _
    lngԭ��ID As Long, strԭ������ As String, _
    Optional lng����ID As Long = -1, Optional str�������� As String = "", _
    Optional blnFromOtherGroupMoveCur As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ա��һ���ƶ�����һ����,�����Ƴ�ĳһ��
    '���:str��ԱID-ָ���ĳ�Ա(���ʱ,�ö��ŷ���)
    '       lngԭ��ID-ԭ���ID
    '       lng����ID-�����ID(Ϊ-1��ʾ�Ƴ�)
    '       blnFromOtherGroupMoveCur-���������ƶ�����ǰ��;����ӵ�ǰ���ƶ���������
    '����:
    '����:
    '����:���˺�
    '����:2010-11-24 10:51:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTemp As String, intIndex As Integer, strID As String
    Dim rsTemp As ADODB.Recordset, cllPro As Collection
    Dim varData As Variant, i As Long
    
    If str��ԱID = "" Then Exit Sub
    
    If lng����ID < 0 Then
        If MsgBox("��ȷ��Ҫ����Ա�� " & str��Ա���� & "���ӡ� " & strԭ������ & "�����Ƴ���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("��ȷ��Ҫ����Ա�� " & str��Ա���� & "���� �� " & strԭ������ & "����" & vbCrLf & "�Ƶ� �� " & str�������� & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    Set cllPro = New Collection
    
    ' Zl_�ɿ��Ա���_Move
    '  ��Աid_In Varchar2,
    '  ԭ��id_In In �ɿ��Ա���.��id%Type,
    '  ����id_In In �ɿ��Ա���.��id%Type := -1
     If Len(str��ԱID) < 2000 Then
        gstrSQL = "Zl_�ɿ��Ա���_Move('" & str��ԱID & "'," & lngԭ��ID & "," & lng����ID & ")"
        AddArray cllPro, gstrSQL
     Else
        varData = Split(str��ԱID, ",")
        strTemp = ""
        For i = 0 To UBound(varData)
            If varData(i) <> "" Then
                If Len(strTemp) >= 1980 Then
                    strTemp = Mid(strTemp, 2)
                    gstrSQL = "Zl_�ɿ��Ա���_Move('" & strTemp & "'," & lngԭ��ID & "," & lng����ID & ")"
                    AddArray cllPro, gstrSQL
                    strTemp = ""
                End If
                strTemp = strTemp & "," & varData(i)
            End If
        Next
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            gstrSQL = "Zl_�ɿ��Ա���_Move('" & strTemp & "'," & lngԭ��ID & "," & lng����ID & ")"
            AddArray cllPro, gstrSQL
        End If
    End If
    Err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo errHandle:
    Dim objItem As ListItem
    Me.MousePointer = 0
    
    With lvwPerson
        intIndex = .SelectedItem.Index
        If blnFromOtherGroupMoveCur = False Or lng����ID <= 0 Then   '�ӵ�ǰ���Ƶ�������ʱ,��Ҫ�Ƴ�����
            varData = Split(str��ԱID, ",")
            For i = 0 To UBound(varData)
                If varData(i) <> "" Then .ListItems.Remove "K" & varData(i)
            Next
        Else     '���������ƶ�����ǰ��ʱ,��Ҫ��������
            varData = Split(str��ԱID, ",")
            For i = 0 To UBound(varData)
                LoadLocalPerson Val(varData(i))
            Next
        End If
        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            lvwPerson_ItemClick .SelectedItem
        Else
            Call lvwGroups_GotFocus
        End If
    End With
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrHand:
    Me.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DeleteStructure(ByVal str��ԱID As String, ByVal str��Ա���� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ���鳤
    '���:str��ԱID-ָ���ĳ�Ա(���ʱ,�ö��ŷ���)
    '����:
    '����:
    '����:���˺�
    '����:2010-11-24 10:51:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strTemp As String, intIndex As Integer, strID As String
    Dim rsTemp As ADODB.Recordset, cllPro As Collection
    Dim varData As Variant, i As Long
    
    If str��ԱID = "" Then Exit Sub
    
    If MsgBox("��ȷ��Ҫ���鳤�� " & str��Ա���� & "���Ƴ���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Err = 0: On Error GoTo errHandle:
    Me.MousePointer = 11
    Set cllPro = New Collection
    
    ' Zl_�ɿ��Ա���_Move
    '  ��Աid_In Varchar2,
    '  ԭ��id_In In �ɿ��Ա���.��id%Type,
    '  ����id_In In �ɿ��Ա���.��id%Type := -1
     If Len(str��ԱID) < 2000 Then
        gstrSQL = "Zl_�ɿ��Ա���_Move('" & str��ԱID & "'," & Mid(lvwGroups.SelectedItem.Key, 2) & ",Null" & ",1)"
        AddArray cllPro, gstrSQL
     Else
        varData = Split(str��ԱID, ",")
        strTemp = ""
        For i = 0 To UBound(varData)
            If varData(i) <> "" Then
                If Len(strTemp) >= 1980 Then
                    strTemp = Mid(strTemp, 2)
                    gstrSQL = "Zl_�ɿ��Ա���_Move('" & strTemp & "'," & Mid(lvwGroups.SelectedItem.Key, 2) & ",Null" & ",1)"
                    AddArray cllPro, gstrSQL
                    strTemp = ""
                End If
                strTemp = strTemp & "," & varData(i)
            End If
        Next
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            gstrSQL = "Zl_�ɿ��Ա���_Move('" & strTemp & "'," & Mid(lvwGroups.SelectedItem.Key, 2) & ",Null" & ",1)"
            AddArray cllPro, gstrSQL
        End If
    End If
    Err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption
    
    Err = 0: On Error GoTo errHandle:
    Dim objItem As ListItem
    Me.MousePointer = 0
    
    With lvwStructure
        intIndex = .SelectedItem.Index

        varData = Split(str��ԱID, ",")
        For i = 0 To UBound(varData)
            If varData(i) <> "" Then .ListItems.Remove "K" & varData(i)
        Next

        If .ListItems.Count > 0 Then
            intIndex = IIf(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
            lvwStructure_ItemClick .SelectedItem
        Else
            Call lvwGroups_GotFocus
        End If
    End With
    Call SetStructureEnable
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrHand:
    Me.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ClearDropVariable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����϶�����ֵ
    '����:���˺�
    '����:2010-11-26 16:38:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnItemClick = False: mstrSelect = "": mblnStartDrop = False: mblnReSel = False
End Sub

Private Function CheckGroupInput(ByVal lngID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�,����true,���򷵼�False
    '����:���˺�
    '����:2010-11-24 16:15:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If Trim(txtEdit(mTxtIdx.idx_������)) = "" Then
        ShowMsgbox "�����Ʊ�������,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_������): Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mTxtIdx.idx_������)), 50, 0, "������") = False Then
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_������): Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txtEdit(mTxtIdx.idx_��˵��)), 50, 0, "��˵��") = False Then
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��˵��): Exit Function
    End If
    If Val(txtEdit(mTxtIdx.idx_������).Tag) = 0 Then
        ShowMsgbox "�����˱�����������벻�Ϸ�,��ѡ��!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_������): Exit Function
    End If
    gstrSQL = "Select 1 From ����ɿ���� where ������=[1] and ɾ������>=sysdate and ID+0<>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtEdit(mTxtIdx.idx_������)), lngID)
    If Not rsTemp.EOF Then
        ShowMsgbox "�������Ѿ�����,����������!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_������): Exit Function
    End If
    CheckGroupInput = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddGroups(bytType As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ϣ
    '���:bytType:0-����;1-�޸�
    '����:���˺�
    '����:2010-11-24 16:04:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    Dim objList As ListItem
    
    If bytType <> 0 Then
        With lvwGroups
            If .SelectedItem Is Nothing Then Exit Sub
            lngID = Val(Mid(.SelectedItem.Key, 2))
        End With
    End If
    
    If CheckGroupInput(lngID) = False Then Exit Sub
    
    On Error GoTo errHandle
    If bytType = 0 Then
        lngID = zlDatabase.GetNextId("����ɿ����")
    End If
   ' Zl_����ɿ����_Update
   gstrSQL = "Zl_����ɿ����_Update("
    '  Id_In       In ����ɿ����.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  ������_In   In ����ɿ����.������%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_������).Text) & "',"
    '  ����_In     In ����ɿ����.����%Type,
    gstrSQL = gstrSQL & "'" & Left(Trim(zlCommFun.SpellCode(Trim(mTxtIdx.idx_������))), 20) & "',"
    '  ˵��_In     In ����ɿ����.˵��%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_��˵��).Text) & "',"
    '  ������id_In In ����ɿ����.������id%Type,
    gstrSQL = gstrSQL & "" & Val(txtEdit(mTxtIdx.idx_������).Tag) & ","
    '  �޸ı�־_In Integer:=0
    gstrSQL = gstrSQL & "" & bytType & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    If bytType = 0 Then
        Set objList = lvwGroups.ListItems.Add(, "K" & lngID, Trim(txtEdit(mTxtIdx.idx_������).Text), "Group", "Group")
    Else
        Set objList = lvwGroups.SelectedItem
    End If
    objList.Text = Trim(txtEdit(mTxtIdx.idx_������).Text)
    objList.Tag = Trim(txtEdit(mTxtIdx.idx_������).Tag)
    objList.SubItems(1) = Trim(txtEdit(mTxtIdx.idx_������).Text)
    objList.SubItems(2) = Trim(txtEdit(mTxtIdx.idx_��˵��).Text)
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CheckPersonInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ա�������Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�,����true,���򷵼�False
    '����:���˺�
    '����:2010-11-24 16:15:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem
    
    On Error GoTo errHandle
    If lvwGroups.SelectedItem Is Nothing Then Exit Function
    
    If Val(txtEdit(mTxtIdx.idx_��Ա).Tag) = 0 Then
        ShowMsgbox "���Ա����ѡ��,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��Ա): Exit Function
    End If
    If Val(Mid(lvwGroups.SelectedItem.Key, 2)) <> Val(txtEdit(mTxtIdx.idx_ԭ��).Tag) Then
        If Val(txtEdit(mTxtIdx.idx_ԭ��).Tag) <> 0 Then
            gstrSQL = "Select sum(nvl(���,0) ) as ��� From ��Ա�ɿ���� Where ����=1 and �տ�Ա=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtEdit(mTxtIdx.idx_��Ա)))
            If Val(NVL(rsTemp!���)) > 0 Then
                ShowMsgbox "��Ա������Ϊ" & txtEdit(mTxtIdx.idx_������) & "�л������ݴ��,�����ƶ�������!"
                zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��Ա): Exit Function
            End If
        End If
    Else
        For Each objItem In lvwPerson.ListItems
            If Val(txtEdit(mTxtIdx.idx_��Ա).Tag) = Val(Mid(objItem.Key, 2)) Then
                ShowMsgbox "��Ա" & txtEdit(mTxtIdx.idx_��Ա) & "�Ѿ��ڸ����д���,û��Ҫ������,����!"
                zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��Ա): Exit Function
                Exit Function
            End If
        Next
    End If
    CheckPersonInput = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckStructureInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ա�������Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�,����true,���򷵼�False
    '����:���˺�
    '����:2010-11-24 16:15:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem
    Dim strSQL As String
    
    On Error GoTo errHandle
    If lvwGroups.SelectedItem Is Nothing Then Exit Function
    
    If Val(txtEdit(mTxtIdx.idx_�鳤).Tag) = 0 Then
        ShowMsgbox "�鳤����ѡ��,����!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�鳤): Exit Function
    End If

    For Each objItem In lvwStructure.ListItems
        If Val(txtEdit(mTxtIdx.idx_�鳤).Tag) = Val(Mid(objItem.Key, 2)) Then
            ShowMsgbox "�鳤" & txtEdit(mTxtIdx.idx_�鳤) & "�Ѿ��ڸ����д���,û��Ҫ������,����!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�鳤): Exit Function
            Exit Function
        End If
    Next
    
    CheckStructureInput = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddPerson()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ա��Ϣ
    '����:���˺�
    '����:2010-11-24 16:04:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    If CheckPersonInput = False Then Exit Sub
    On Error GoTo errHandle
    If Val(txtEdit(mTxtIdx.idx_ԭ��).Tag) <> 0 Then
        Call PersonFromGroupToOtherGroup(txtEdit(mTxtIdx.idx_��Ա).Tag, txtEdit(mTxtIdx.idx_��Ա), txtEdit(mTxtIdx.idx_ԭ��).Tag, txtEdit(mTxtIdx.idx_ԭ��), Mid(lvwGroups.SelectedItem.Key, 2), lvwGroups.SelectedItem.Text)
        'Call LoadLocalPerson(Val(txtEdit(mTxtIdx.idx_��Ա).Tag))
        Exit Sub
    End If
    'Zl_�ɿ��Ա���_Insert
    gstrSQL = "Zl_�ɿ��Ա���_Insert("
    '  ��id_In   In �ɿ��Ա���.��id%Type,
    gstrSQL = gstrSQL & "" & Mid(lvwGroups.SelectedItem.Key, 2) & ","
    '  ��Աid_In In �ɿ��Ա���.��Աid%Type
    gstrSQL = gstrSQL & "" & Val(txtEdit(mTxtIdx.idx_��Ա).Tag) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Call LoadLocalPerson(Val(txtEdit(mTxtIdx.idx_��Ա).Tag))
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_��Ա)
    Call SetGoupsEnable
    Call SetPersonEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddStructure()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ա��Ϣ
    '����:���˺�
    '����:2010-11-24 16:04:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    If CheckStructureInput = False Then Exit Sub
    On Error GoTo errHandle
    'Zl_�ɿ��Ա���_Insert
    gstrSQL = "Zl_�ɿ��Ա���_Insert("
    '  ��id_In   In �ɿ��Ա���.��id%Type,
    gstrSQL = gstrSQL & "" & Mid(lvwGroups.SelectedItem.Key, 2) & ","
    '  ��Աid_In In �ɿ��Ա���.��Աid%Type
    gstrSQL = gstrSQL & "" & Val(txtEdit(mTxtIdx.idx_�鳤).Tag) & ",1)"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Call LoadLocalStructure(Val(txtEdit(mTxtIdx.idx_�鳤).Tag))
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_�鳤)
    Call SetGoupsEnable
    Call SetPersonEnable
    Call SetStructureEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadLocalPerson(ByVal lng��ԱID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����Ա����Ϣ��ListView
    '����:���˺�
    '����:2010-11-24 16:52:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem, i As Long
    Dim strIco As String
    On Error GoTo errHandle
    
    gstrSQL = " " & _
    "   Select  B.ID  ,B.���,B.����,B.����,b.��������,B.���֤��,B.�Ա�,B.����,B.�칫�ҵ绰,B.�����ʼ�,B.����ְ�� " & _
    "   From  ��Ա�� B " & _
    "   Where   B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng��ԱID)
    If rsTemp.EOF Then Exit Sub
    
   With lvwPerson
        If InStr(1, NVL(rsTemp!�Ա�), "��") > 0 Then
            strIco = "Man"
        ElseIf InStr(1, NVL(rsTemp!�Ա�), "Ů") > 0 Then
            strIco = "Woman"
        Else
            strIco = "Man" '"Other"
        End If
        Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!ID), NVL(rsTemp!����), strIco, strIco)
        i = 1
        objItem.SubItems(i) = NVL(rsTemp!���): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
        objItem.SubItems(i) = Format(rsTemp!��������, "yyyy-mm-dd"): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!�Ա�): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!�칫�ҵ绰): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!�����ʼ�): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!����ְ��): i = i + 1
        objItem.Selected = True: objItem.EnsureVisible
        Call lvwPerson_ItemClick(objItem)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadLocalStructure(ByVal lng��ԱID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����Ա����Ϣ��ListView
    '����:���˺�
    '����:2010-11-24 16:52:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, objItem As ListItem, i As Long
    Dim strIco As String
    On Error GoTo errHandle
    
    gstrSQL = " " & _
    "   Select  B.ID  ,B.���,B.����,B.����,b.��������,B.���֤��,B.�Ա�,B.����,B.�칫�ҵ绰,B.�����ʼ�,B.����ְ�� " & _
    "   From  ��Ա�� B " & _
    "   Where   B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng��ԱID)
    If rsTemp.EOF Then Exit Sub
    
   With lvwStructure
        If InStr(1, NVL(rsTemp!�Ա�), "��") > 0 Then
            strIco = "Man"
        ElseIf InStr(1, NVL(rsTemp!�Ա�), "Ů") > 0 Then
            strIco = "Woman"
        Else
            strIco = "Man" '"Other"
        End If
        Set objItem = .ListItems.Add(, "K" & NVL(rsTemp!ID), NVL(rsTemp!����), strIco, strIco)
        i = 1
        objItem.SubItems(i) = NVL(rsTemp!���): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
        objItem.SubItems(i) = Format(rsTemp!��������, "yyyy-mm-dd"): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!�Ա�): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!����): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!�칫�ҵ绰): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!�����ʼ�): i = i + 1
        objItem.SubItems(i) = NVL(rsTemp!����ְ��): i = i + 1
        objItem.Selected = True: objItem.EnsureVisible
        Call lvwStructure_ItemClick(objItem)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ��������
    '����:���˺�
    '����:2010-11-15 17:54:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Me.ActiveControl Is lvwGroups Then
        With lvwGroups
            zlIsHaveData = .ListItems.Count > 0: Exit Function
        End With
    Else
        With lvwPerson
            zlIsHaveData = .ListItems.Count > 0: Exit Function
        End With
    End If
End Function
 
Private Sub lvwGroups_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintGroupColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwGroups.SortOrder = IIf(lvwGroups.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintGroupColumn = ColumnHeader.Index - 1
        lvwGroups.SortKey = mintGroupColumn
        lvwGroups.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwGroups_DblClick()
    If mblnEdit Then
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_������)
    End If
End Sub
 
Private Sub lvwGroups_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub lvwGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
       Set objPopup = cbsThis.FindControl(, conMenu_ViewPopup, , True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
Private Sub lvwGroups_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrPreGroupKey = Item.Key Then Exit Sub
    mstrPreGroupKey = Item.Key
    txtEdit(mTxtIdx.idx_������).Text = Item.Text
    txtEdit(mTxtIdx.idx_������).Text = Item.SubItems(1)
    txtEdit(mTxtIdx.idx_������).Tag = Item.Tag
    txtEdit(mTxtIdx.idx_��˵��).Text = Item.SubItems(2)
    Call LoadGroupPersons(Val(Mid(Item.Key, 2)))  '���س�Ա��Ϣ
    Call LoadGroupStructure(Val(Mid(Item.Key, 2)))
    Call SetGoupsEnable
End Sub


Private Sub lvwPerson_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintPersonColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwPerson.SortOrder = IIf(lvwPerson.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintPersonColumn = ColumnHeader.Index - 1
        lvwPerson.SortKey = mintPersonColumn
        lvwPerson.SortOrder = lvwAscending
    End If
End Sub

  
Private Sub lvwPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lvwStructure_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lvwPerson_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim objPopup As CommandBarPopup
    mblnStartDrop = False
    If Button = 2 Then
       Set objPopup = cbsThis.FindControl(, conMenu_ViewPopup, , True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
    mblnItemClick = False
End Sub

Private Sub lvwStructure_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintStructureColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwStructure.SortOrder = IIf(lvwStructure.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintStructureColumn = ColumnHeader.Index - 1
        lvwStructure.SortKey = mintStructureColumn
        lvwStructure.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwStructure_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim objPopup As CommandBarPopup
    mblnStartDrop = False
    If Button = 2 Then
       Set objPopup = cbsThis.FindControl(, conMenu_ViewPopup, , True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
    mblnItemClick = False
End Sub


Private Sub picStructure_Resize()
    With lvwStructure
        .Height = picStructure.ScaleHeight - .Top - 30
        .Width = picStructure.ScaleWidth
    End With
End Sub

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = "":
    If mTxtIdx.idx_��Ա = Index Then
        cmdFucn(mCmdIdx.idx_��Ա�Ƴ�).Tag = ""
        Call SetPersonEnable
    Else
        If Index = mTxtIdx.idx_�鳤 Then
            Call SetStructureEnable
        Else
            Call SetGoupsEnable
            Call SetPersonEnable
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = mTxtIdx.idx_��Ա Or idx_������ = Index Then
        zlCommFun.OpenIme False
    Else
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lng��ԱID As Long, rsTemp As ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errHandle
    
    Select Case Index
    Case mTxtIdx.idx_������
       If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index)), , lng��ԱID) = False Then Exit Sub
       txtEdit(Index).Tag = lng��ԱID
    Case mTxtIdx.idx_��Ա
       If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index)), , lng��ԱID, , "", "����Һ�Ա,�����շ�Ա,Ԥ���տ�Ա,סԺ����Ա,��Ժ�Ǽ�Ա,�����Ǽ���") = False Then Exit Sub
       txtEdit(Index).Tag = lng��ԱID
       '��ȡ��ص�����Ϣ
        gstrSQL = "" & _
        "   Select A.Id, A.������ " & _
        "   From ����ɿ���� A,�ɿ��Ա��� B " & _
        "   Where A.ID=B.��Id and B.��ԱID=[1] And (A.ɾ������>Sysdate Or A.ɾ������ Is Null)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng��ԱID)
        If Not rsTemp.EOF Then
            txtEdit(mTxtIdx.idx_ԭ��).Text = NVL(rsTemp!������)
            txtEdit(mTxtIdx.idx_ԭ��).Tag = NVL(rsTemp!ID)
        Else
            txtEdit(mTxtIdx.idx_ԭ��).Text = ""
        End If
    Case mTxtIdx.idx_�鳤
        If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index)), , lng��ԱID) = False Then Exit Sub
        txtEdit(Index).Tag = lng��ԱID
    Case Else
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode������
    '����:���˺�
    '����:2010-11-15 17:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��ID As Long
    If Not Me.lvwGroups.SelectedItem Is Nothing Then
        lng��ID = Val(Mid(Me.lvwGroups.SelectedItem.Key, 2))
    End If
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "��ID=" & lng��ID)
End Sub


Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub
