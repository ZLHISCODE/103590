VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoleGrant 
   Caption         =   "��ɫ��Ȩ"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14025
   Icon            =   "frmRoleGrant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14025
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Frame fraHSplit 
      Height          =   30
      Left            =   240
      MousePointer    =   7  'Size N S
      TabIndex        =   31
      Top             =   4800
      Width           =   9615
   End
   Begin VB.Frame fraVSplit 
      Height          =   7095
      Left            =   6360
      MousePointer    =   9  'Size W E
      TabIndex        =   18
      Top             =   480
      Width           =   30
   End
   Begin VB.CommandButton cmdUnSel 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   3960
      TabIndex        =   30
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   3000
      TabIndex        =   29
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "ȫ��չ��(&D)"
      Height          =   350
      Left            =   1680
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "������Ȩ��(&V)"
      Height          =   350
      Left            =   9240
      TabIndex        =   27
      Top             =   7680
      Width           =   1695
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Index           =   1
      Left            =   6480
      ScaleHeight     =   6975
      ScaleWidth      =   5175
      TabIndex        =   19
      Top             =   600
      Width           =   5175
      Begin MSComctlLib.TreeView tvwMenu 
         Height          =   3690
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   6509
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgTreeview"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lvwFunc 
         Height          =   2415
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   4560
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   4260
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15724768
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "˵��"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwModRelas 
         Height          =   4890
         Left            =   3600
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   8625
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgTreeview"
         Appearance      =   1
      End
      Begin VB.Label lblNotice 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         Height          =   180
         Left            =   1080
         TabIndex        =   26
         Top             =   180
         Width           =   90
      End
      Begin VB.Label lblFuncNote 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         Caption         =   "����ģ�鹦��"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   4320
         Width           =   1080
      End
      Begin VB.Label lblMenuNote 
         AutoSize        =   -1  'True
         BackColor       =   &H00EFF0E0&
         Caption         =   "����ģ��"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   7155
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   4035
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -60
      TabIndex        =   11
      Top             =   585
      Width           =   11070
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   12750
      TabIndex        =   13
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   11520
      TabIndex        =   12
      Top             =   7680
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgTreeview 
      Left            =   5280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":000C
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":0E5E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":76C0
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":859A
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":EDFC
            Key             =   "Function"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":FC4E
            Key             =   "Optional"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleGrant.frx":164B0
            Key             =   "Fixed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   8108
      Visible         =   0   'False
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   8064
      Width           =   14028
      _ExtentX        =   24739
      _ExtentY        =   635
      SimpleText      =   $"frmRoleGrant.frx":1CD12
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRoleGrant.frx":1CD59
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21828
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
   Begin MSComctlLib.ListView lvwTmp 
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      View            =   1
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Index           =   0
      Left            =   240
      ScaleHeight     =   6975
      ScaleWidth      =   6135
      TabIndex        =   17
      Top             =   600
      Width           =   6135
      Begin VB.CheckBox chkShowDisReport 
         Caption         =   "��ʾͣ�ñ���(&R)"
         Height          =   345
         Left            =   4800
         TabIndex        =   32
         Top             =   105
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.CheckBox chkOnlyShow 
         Caption         =   "������Ȩ(&G)"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   150
         Width           =   1290
      End
      Begin VB.CheckBox chkVirtual 
         Caption         =   "������ģ��(&M)"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   150
         Width           =   1545
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   120
         Width           =   1530
      End
      Begin MSComctlLib.TreeView tvwMenu 
         Height          =   3690
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   6509
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imgTreeview"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvwFunc 
         Height          =   2415
         Index           =   0
         Left            =   15
         TabIndex        =   8
         Top             =   4560
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   4260
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "˵��"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.Label lblMenuNote 
         AutoSize        =   -1  'True
         Caption         =   "ģ��˵�"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblFuncNote 
         AutoSize        =   -1  'True
         Caption         =   "ģ�鹦��"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label lblSearch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1155
         TabIndex        =   3
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ��ϵͳ(&S)"
      Height          =   180
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.Image imgRoleGrant 
      Height          =   480
      Left            =   120
      Picture         =   "frmRoleGrant.frx":1D5ED
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�Խ�ɫ�������շ�Ա������Ȩ����(Ctrl+F����,F3��һ��)"
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuState 
         Caption         =   "������Ŀ(&1)"
         Index           =   0
      End
      Begin VB.Menu mnuPopuState 
         Caption         =   "������ϸ(&2)"
         Checked         =   -1  'True
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRoleGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���β˵�ö��
Private Enum MenuType
    MT_ģ�� = 0
    MT_����ģ�� = 1
End Enum
'������ʽ
Private Enum ShowFaceType
    SFT_Ӧ��ϵͳ = 0 'Ӧ��ϵͳ����ʽ
    SFT_�Զ��屨�� = 1 '���������Լ��Զ��屨��
    SFT_���� = 2 '�����������������
End Enum
'��ȡ��������
Private Enum ReadDataType
    RDT_Menu = 0
    RDT_Module = 1
    RDT_Function = 2
    RDT_Table = 3
    RDT_Systems = 4
    RDT_ModRelas = 5
End Enum
'��տؼ�����
Private Enum ClearType
    CT_�������� = 0 '�л�����ģ�飬��Ҫ��չ�������
    CT_����ģ�� = 1 '�л�������ģ���빦�ܼ���л���Ҫ��չ���ģ��
    CT_���� = 2 'ģ���л�
    CT_Sys = 3 'ϵͳ�л�
End Enum

Private Const BLN_TEST = False
Private mblnOk As Boolean
Private mstrRole As String

Private WithEvents mclsPrivilege As clsPrivilege
Attribute mclsPrivilege.VB_VarHelpID = -1

Private mrsTree As ADODB.Recordset        'ģ��˵�
Private mrsModule As ADODB.Recordset      'ģ�����������¼
'ģ����Ȩ���������������Լ�ģ����Ȩ�����Ļ��棬�ڱ�����Ȩʱ�Ὣ��Ϣ���µ�mrsModule��
Private mrsModsInfo As ADODB.Recordset
Private mrsTable  As ADODB.Recordset      '��������������¼
Private mrsFunction  As ADODB.Recordset   '�������������¼
Private mrsSys   As ADODB.Recordset   '���а�װϵͳ
Private mrsModRelas As ADODB.Recordset   'ģ���ϵ����
Private mrsRelasTree As ADODB.Recordset  '����ģ������

Private mrsRelas As ADODB.Recordset 'Ȩ�޹�ϵ(����)
Private mrsRelExcl As ADODB.Recordset ' Ȩ�޹�ϵ(����)
Private mrsGroup As ADODB.Recordset 'Ȩ�޷���
Private mblnVirtual As Boolean

Private mintUpdate As Integer
Private mblnItem As Boolean
'��ݼ����Ʊ���
Private mblnExpanded As Boolean '�Ƿ�ȫ��չ��
'��������
Private mlngSys As Long
Private msftStyle As ShowFaceType  '������ʽ
Private mcllHaveSys As Collection '����ģ��˵���ϵ��ϵͳ,key=ϵͳ,ֵ=1
Private mcllKeyModule As Collection 'ģ��˵���ϵ��Key=ϵͳ_ģ��,ֵ=�˵�Key1,�˵�Key2...
Private mcllCodeModule As Collection '��ţ�ģ���ϵ��Key=���,ֵ=ϵͳ_ģ��
Private mstrFind As String '�����ַ���
Private mlngCurPos As Long '��ǰ����λ��
Private mblnClear As Boolean
Private mcllTip As Collection '������ʾ
Private mstrCurRelas As String
Private mblnReturn As Boolean
Private mintActive As Integer
Private mblnUnRefresh As Boolean
Private mblnSaveClick As Boolean '��¼�Ƿ񱣴��޸�

Private mrsRptGroups As ADODB.Recordset      '��¼�����������
Private mrsReports As ADODB.Recordset        '��¼��������
Private mrsGroups As ADODB.Recordset         '��¼����������

Public Function GrantToRole(ByVal strRole As String) As Boolean
    mblnOk = False
    If gstrע���� = "" Then
        MsgBox "���û���ע������Ч��������ע�ᡣ", vbExclamation, gstrSysName
        Exit Function
    End If
    mstrRole = strRole
    Set mcllHaveSys = New Collection
    Set mcllKeyModule = New Collection
    Set mcllCodeModule = New Collection
    Set mrsModule = ReadData(RDT_Module)
    If mrsModule.RecordCount = 0 Then
        Set mrsFunction = ReadData(RDT_Function)
        If mrsFunction.RecordCount = 0 Then
            Set mrsTable = ReadData(RDT_Table)
            If mrsTable.RecordCount = 0 Then
                MsgBox "�㲻���п������Ȩ�������ܽ�����Ȩ������", vbInformation, gstrSysName
                On Error Resume Next
                Unload Me
                err.Clear: On Error GoTo 0
                Exit Function
            End If
        End If
    End If
    Me.Show vbModal, frmMDIMain
    GrantToRole = mblnSaveClick
    mblnSaveClick = False
End Function

Private Sub chkOnlyShow_Click()
    
    LockWindowUpdate Me.hwnd
    If msftStyle <> SFT_���� Then Call AdjustRelasTree
    mlngCurPos = 0
    Call SetOnlyShow
    If tvwMenu(MT_ģ��).Nodes.Count = 0 Then
        Call ClearFace(CT_Sys)
    End If
    LockWindowUpdate 0
End Sub

Private Sub chkShowDisReport_Click()
    Dim objNode As Node
    
    LockWindowUpdate Me.hwnd
    mrsModsInfo.Filter = "ϵͳ=0 And ��� >=100"
    mrsModsInfo.Sort = "���"
    tvwMenu(MT_ģ��).Nodes.Clear
    tvwMenu(MT_ģ��).Nodes.Add , , "K_0", "���б���", "����", "����_ѡ��"
    If mrsGroups.RecordCount <> 0 Then mrsGroups.MoveFirst
    With mrsGroups
        Do While Not .EOF
            If IsNull(!�ϼ�id) = True Then
                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_0", tvwChild, "K_" & !Id, !����, "����", "����_ѡ��")
            Else
                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & !�ϼ�id, tvwChild, "K_" & !Id, !����, "����", "����_ѡ��")
            End If
            .MoveNext
        Loop
    End With
    With mrsModsInfo
        Do While Not .EOF
            mrsRptGroups.Filter = IIf(chkShowDisReport.value, "", "�Ƿ�ͣ�� = 0 and ") & "����id = " & !���
            mrsReports.Filter = IIf(chkShowDisReport.value, "", "�Ƿ�ͣ�� = 0 and ") & "����id = " & !���
            If mrsRptGroups.RecordCount = 1 Then
                '��ģ��Ϊ�����鷢����
                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Nvl(mrsRptGroups!����id, 0), tvwChild, "M_0_" & !���, "��" & Format(!���, "000000") & "��" & !����, "Module")
                objNode.Checked = !��Ȩ�� = 1
                mrsModsInfo.Update "ģ������", 2
            ElseIf mrsReports.RecordCount = 1 Then
                '��ģ��Ϊ��������
                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Nvl(mrsReports!����id, 0), tvwChild, "M_0_" & !���, "��" & Format(!���, "000000") & "��" & !����, "Module")
                objNode.Checked = !��Ȩ�� = 1
                mrsModsInfo.Update "ģ������", 2
            End If
            .MoveNext
        Loop
    End With

    'ɾ��û��������ļ���
    Call DeleteNodes
    LockWindowUpdate 0
End Sub

Private Sub chkVirtual_Click()
    Dim objNode As Node
    mlngCurPos = 0
    If Not mblnVirtual Then Call SetVirtualVisual(chkVirtual.value <> 0)
    '��������������ȹ�ѡ����ģ�飬ѡ������ģ�飬��ȡ����ѡ��ѡ�д��ڵ�ģ�飬
    '����Ҫ������ǰѡ�еĽڵ㣨�����Ѿ������ڣ���ȡ����ǰѡ��ڵ�ļӴ�״̬(��������ᱨ����������������
    On Error Resume Next
    Set objNode = tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag)
    If err.Number <> 0 Then
        err.Clear: tvwMenu(MT_ģ��).Tag = ""
        If tvwMenu(MT_ģ��).Nodes.Count <> 0 Then
            Call tvwMenu_NodeClick(MT_ģ��, tvwMenu(MT_ģ��).Nodes(1))
        End If
    End If
End Sub

Private Sub cmbSystem_Click()
    Dim strPre As String
    Dim objNode As Node, objNodeChild As Node
    Dim blnHaveSys As Boolean
    Dim strTMp As String
    Dim strFirstNode As String
    Dim strPreKey As String
    Dim i As Long
    
    mblnUnRefresh = True
    '�ж��Ƿ��л�ϵͳ���л��˲�ˢ������
    If Val(cmbSystem.Tag) <> cmbSystem.ListIndex Then
        LockWindowUpdate Me.hwnd
        Call ClearFace(CT_Sys)
        chkShowDisReport.Visible = False
        mlngSys = Val(cmbSystem.ItemData(cmbSystem.ListIndex))
        cmbSystem.Tag = cmbSystem.ListIndex
        If mlngSys = 0 Then '��Ӧ��ϵͳ��Ȩ
            chkVirtual.Visible = False
            Select Case cmbSystem.Text
                Case "��������"
                    msftStyle = SFT_����
                    If mrsTable Is Nothing Then Set mrsTable = ReadData(RDT_Table)
                    If glngSysNo <> -1 Then
                        mrsTable.Filter = "ϵͳ = " & glngSysNo
                    Else
                        mrsTable.Filter = ""
                    End If
                    mrsTable.Sort = "ϵͳ,����"
                    With mrsTable
                        Do While Not .EOF
                            If strPre <> !ϵͳ & "" Then
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add(, , "K_" & !ϵͳ, "��" & !ϵͳ & "��" & !ϵͳ��, "����", "����_ѡ��")
                                strPre = !ϵͳ & "": objNode.Checked = True
                            End If
                            Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & !ϵͳ, tvwChild, "T_" & !ϵͳ & "_" & !����, RPAD("��" & !���� & "��", 20) & !˵��, "Table")
                            objNode.Checked = !��Ȩ�� = 1
                            'Ĭ�Ϲ�ѡ������һ���Ӽ�����ѡ����ȡ����ѡ
                            If Not objNode.Checked Then objNode.Parent.Checked = False
                            '��Ĭ�Ϲ�ѡ�ڵ㣬��ֹ�ٶȽ���
'                            '��ȡ��һ���ڵ����һ����ѡ�ڵ�
'                            If objNode.Checked And tvwMenu(MT_ģ��).Tag = "" Then
'                                tvwMenu(MT_ģ��).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            .MoveNext
                        Loop
                    End With
                Case "ȡ������"
                    msftStyle = SFT_����
                    If mrsFunction Is Nothing Then Set mrsFunction = ReadData(RDT_Function)
                    If glngSysNo <> -1 Then
                        mrsFunction.Filter = "ϵͳ = " & glngSysNo
                    Else
                        mrsFunction.Filter = ""
                    End If
                    mrsFunction.Sort = "ϵͳ,������"
                    With mrsFunction
                        Do While Not .EOF
                            If strPre <> !ϵͳ & "" Then
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add(, , "K_" & !ϵͳ, "��" & !ϵͳ & "��" & !ϵͳ��, "����", "����_ѡ��")
                                strPre = !ϵͳ & "": objNode.Checked = True
                            End If
                            Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & !ϵͳ, tvwChild, "F_" & !ϵͳ & "_" & !������, RPAD("��" & !������ & "����������" & !������ & "����", 52) & !˵��, "Function")
                            objNode.Checked = !��Ȩ�� = 1
                            '��Ĭ�Ϲ�ѡ�ڵ㣬��ֹ�ٶȽ���
                            'Ĭ�Ϲ�ѡ������һ���Ӽ�����ѡ����ȡ����ѡ
                            If Not objNode.Checked Then objNode.Parent.Checked = False
'                            '��ȡ��һ���ڵ����һ����ѡ�ڵ�
'                            If objNode.Checked And tvwMenu(MT_ģ��).Tag = "" Then
'                                tvwMenu(MT_ģ��).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            .MoveNext
                        Loop
                    End With
                Case "��������"
                    If mrsModsInfo Is Nothing Then Set mrsModsInfo = GetModuleInfo
                    mrsModsInfo.Filter = "ϵͳ=0 And ���<100"
                    mrsModsInfo.Sort = "���"
                    With mrsModsInfo
                        Do While Not .EOF
                            Set objNode = tvwMenu(MT_ģ��).Nodes.Add(, , "M_0_" & !���, "��" & Format(!���, "000000") & "��" & !����, "Module")
                            objNode.Checked = !��Ȩ�� = 1
                            '���ģ������
                            mrsModsInfo.Update "ģ������", 2
                            '��Ĭ�Ϲ�ѡ�ڵ㣬��ֹ�ٶȽ���
'                            '��ȡ��һ���ڵ����һ����ѡ�ڵ�
'                            If objNode.Checked And tvwMenu(MT_ģ��).Tag = "" Then
'                                tvwMenu(MT_ģ��).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            .MoveNext
                        Loop
                    End With
                Case "�Զ��屨��"
                    msftStyle = SFT_�Զ��屨��
                    chkShowDisReport.Visible = True
                    On Error GoTo errH
                    If mrsModsInfo Is Nothing Then Set mrsModsInfo = GetModuleInfo
                    mrsModsInfo.Filter = "ϵͳ=0 And ��� >=100"
                    mrsModsInfo.Sort = "���"
                    
                    gstrSQL = "Select Id, �ϼ�id, ����, ˵��" & vbNewLine & _
                                "From (Select Id, �ϼ�id, ����, ˵�� From Zlrptclasses)" & vbNewLine & _
                                "Start With �ϼ�id Is Null" & vbNewLine & _
                                "Connect By Prior Id = �ϼ�id"
                    Set mrsGroups = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "���ҷ���")
                    gstrSQL = "Select Id, ����id, ����, ����id, Nvl(�Ƿ�ͣ��, 0) �Ƿ�ͣ�� From Zlrptgroups Where ϵͳ Is Null And ����id Is Not Null"
                    Set mrsRptGroups = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "�����ѷ����ķ���")
                    gstrSQL = "Select ����id, ����id, Nvl(�Ƿ�ͣ��, 0) �Ƿ�ͣ��  From Zlreports Where ϵͳ Is Null And ����id Is Not Null"
                    Set mrsReports = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "�����ѷ����ı���")
                    
                    tvwMenu(MT_ģ��).Nodes.Add , , "K_0", "���б���", "����", "����_ѡ��"
                    
                    With mrsGroups
                        Do While Not .EOF
                            If IsNull(!�ϼ�id) = True Then
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_0", tvwChild, "K_" & !Id, !����, "����", "����_ѡ��")
                            Else
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & !�ϼ�id, tvwChild, "K_" & !Id, !����, "����", "����_ѡ��")
                            End If
                            .MoveNext
                        Loop
                    End With
                    With mrsModsInfo
                        Do While Not .EOF
                            mrsRptGroups.Filter = "�Ƿ�ͣ�� = 0 and ����id = " & !���
                            mrsReports.Filter = "�Ƿ�ͣ�� = 0 and ����id = " & !���
                            If mrsRptGroups.RecordCount = 1 Then
                                '��ģ��Ϊ�����鷢����
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Nvl(mrsRptGroups!����id, 0), tvwChild, "M_0_" & !���, "��" & Format(!���, "000000") & "��" & !����, "Module")
                                objNode.Checked = !��Ȩ�� = 1
                                '���ģ������
                                mrsModsInfo.Update "ģ������", 2
                            ElseIf mrsReports.RecordCount = 1 Then
                                '��ģ��Ϊ��������
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Nvl(mrsReports!����id, 0), tvwChild, "M_0_" & !���, "��" & Format(!���, "000000") & "��" & !����, "Module")
                                objNode.Checked = !��Ȩ�� = 1
                                '���ģ������
                                mrsModsInfo.Update "ģ������", 2
                            End If
                            .MoveNext
                        Loop
                    End With
                    
                    'ɾ��û��������ļ���
                    Call DeleteNodes
            End Select
        Else 'Ӧ��ϵͳ��Ȩ
            msftStyle = SFT_Ӧ��ϵͳ
            '�жϽڵ�ģ���ϵ�Ƿ�洢
            On Error Resume Next
            strTMp = mcllHaveSys("S_" & mlngSys)
            If err.Number <> 0 Then err.Clear: strTMp = ""
            blnHaveSys = strTMp <> ""
            On Error GoTo errH
            If mrsModsInfo Is Nothing Then Set mrsModsInfo = GetModuleInfo
            If mrsTree Is Nothing Then Set mrsTree = ReadData(RDT_Menu)
            With mrsTree
                .Filter = "ϵͳ=" & mlngSys
                Do While Not .EOF
                    If !ģ�� = 0 Then
                        If !�ϼ� = 0 Then
                            Set objNode = tvwMenu(MT_ģ��).Nodes.Add(, , "K_" & Format(!���, "000000"), !����, "����", "����_ѡ��")
                        Else
                            Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Format(!�ϼ�, "000000"), tvwChild, "K_" & Format(!���, "000000"), !����, "����", "����_ѡ��")
                        End If
                    Else
                        mrsModsInfo.Filter = "ϵͳ=" & mlngSys & " And ���=" & !ģ��
                        If Not mrsModsInfo.EOF Then 'û���ҵ���Ӧģ������ʾ
                            If !�ϼ� = 0 Then
                                Set objNode = tvwMenu(MT_ģ��).Nodes.Add(, , "M_" & Format(!���, "000000") & "_" & !ģ��, "��" & Format(!ģ��, "000000") & "��" & !����, "Module")
                                If Not blnHaveSys Then mcllCodeModule.Add objNode.Key, "K_" & !���
                            Else
                                '���Ӷ�ģ����¼��ڵ�ͬ����ģ���֧��
                                On Error Resume Next
                                strPreKey = mcllCodeModule("K_" & !�ϼ�)
                                If err.Number <> 0 Then
                                    err.Clear
                                    Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Format(!�ϼ�, "000000"), tvwChild, "M_" & Format(!���, "000000") & "_" & !ģ��, "��" & Format(!ģ��, "000000") & "��" & !����, "Module")
                                Else
                                    Set objNode = tvwMenu(MT_ģ��).Nodes.Add(strPreKey, tvwChild, "M_" & Format(!���, "000000") & "_" & !ģ��, "��" & Format(!ģ��, "000000") & "��" & !����, "Module")
                                End If
                                On Error GoTo errH
                                If Not blnHaveSys Then mcllCodeModule.Add objNode.Key, "K_" & !���
                            End If
                            
                            objNode.Checked = mrsModsInfo!��Ȩ�� = 1
                            '��Ĭ�Ϲ�ѡ�ڵ㣬��ֹ�ٶȽ���
'                            '��ȡ��һ���ڵ����һ����ѡ�ڵ�
'                            If objNode.Checked And tvwMenu(MT_ģ��).Tag = "" Then
'                                tvwMenu(MT_ģ��).Tag = objNode.Key
'                            End If
'                            If strFirstNode = "" Then
'                                strFirstNode = objNode.Key
'                            End If
                            '���ģ�����ͣ��洢�ڵ�ģ���ϵ
                            If Not blnHaveSys Then
                                mrsModsInfo.Update "ģ������", 1
                                On Error Resume Next
                                strTMp = mcllKeyModule("K_" & mlngSys & "_" & !ģ��)
                                If err.Number <> 0 Then
                                    err.Clear: strTMp = ""
                                    strTMp = objNode.Key
                                Else
                                    mcllKeyModule.Remove "K_" & mlngSys & "_" & !ģ��
                                    strTMp = strTMp & "," & objNode.Key
                                End If
                                mcllKeyModule.Add strTMp, "K_" & mlngSys & "_" & !ģ��
                                On Error GoTo errH
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End With
            '�ж��Ƿ�չʾ����ģ�飬��ȡ����ģ��ĵ�һ���ڵ㣬���һ����ѡ�Ľڵ�
            Call SetVirtualVisual(tvwMenu(MT_ģ��).Nodes.Count = 0 Or chkVirtual.value <> 0, strFirstNode)
            '��Ǹ�ϵͳ��ģ��˵���ϵ�Ѿ���¼
            If Not blnHaveSys Then
                mcllHaveSys.Add "1", "S_" & mlngSys
            End If
            '�������年ѡ״̬
            Call CheckNode(tvwMenu(MT_ģ��))
        End If
        '��Ĭ�Ϲ�ѡ�ڵ㣬��ֹ�ٶȽ���
'        'û�й�ѡ�ڵ�����һ����һ�ڵ㣬��ѡ���һ�ڵ�
'        If tvwMenu(MT_ģ��).Tag = "" And strFirstNode <> "" Then
'            tvwMenu(MT_ģ��).Tag = strFirstNode
'        End If
'        'չ����ǽڵ�
'        If tvwMenu(MT_ģ��).Tag <> "" Then
'            strTmp = tvwMenu(MT_ģ��).Tag: tvwMenu(MT_ģ��).Tag = ""
'            Call SetNodeExpand(tvwMenu(MT_ģ��), strTmp) 'չ���ڵ�
'            Call tvwMenu_NodeClick(MT_ģ��, tvwMenu(MT_ģ��).Nodes(strTmp))
'        End If
        If chkOnlyShow.value = 1 Then
            Call SetOnlyShow
        End If
        Call Form_Resize
        LockWindowUpdate 0
        mblnUnRefresh = False
        Call RefreshState
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "cmbSystem_Click:" & err.Description, vbInformation, Me.Caption
End Sub

Private Function DeleteNodes(Optional ByVal strKey As String) As Boolean
'ɾ��tvwMenu(MT_ģ��)ָ���ڵ������пսڵ�
'���:tvwMenu(MT_ģ��)ָ���ڵ��keyֵ�����Ϊ�գ���Ĭ��Ϊ���ڵ�
    Dim objNode As Node
    Dim strDelKeys As String
    Dim arrTmp As Variant, i As Long
    Dim blnNotDel As Boolean

    '��ȡ��ʼ�ڵ㣬�Ա�ѭ��
    If tvwMenu(MT_ģ��).Nodes.Count = 0 Then Exit Function
    If strKey = "" Then
        Set objNode = tvwMenu(MT_ģ��).Nodes(1)
    ElseIf tvwMenu(MT_ģ��).Nodes(strKey).Children <> 0 Then
        Set objNode = tvwMenu(MT_ģ��).Nodes(strKey).Child
    End If
    '��ȡ����ɾ���Ľڵ�
    Do While Not objNode Is Nothing
        '���Ӽ���ѡ�У��򸸼�ѡ��
        If objNode.Key Like "M*" Then
            blnNotDel = True
        ElseIf Not DeleteNodes(objNode.Key) Then
            strDelKeys = strDelKeys & "|" & objNode.Key
        Else
            blnNotDel = True
        End If
        Set objNode = objNode.Next
    Loop
    
    'ɾ������ɾ���Ľڵ�
    arrTmp = Split(Mid(strDelKeys, 2), "|")
    For i = LBound(arrTmp) To UBound(arrTmp)
        tvwMenu(MT_ģ��).Nodes.Remove arrTmp(i)
    Next
    DeleteNodes = blnNotDel
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    frmModuleCheck.ShowMe IIf(glngSysNo = -1, 0, glngSysNo)
End Sub

Private Sub cmdExp_Click()
    Call Form_KeyDown(vbKeyD, vbCtrlMask)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "ZL9Svrtools\" & Me.name
End Sub

Private Sub cmdOK_Click()
    'Dim sglTimer0 As Single
    Dim str������() As String, i As Long
    
    mblnSaveClick = True
    mrsSys.Filter = "": mrsSys.Sort = "���"
    ReDim str������(mrsSys.RecordCount - 1)
    For i = LBound(str������) To UBound(str������)
        str������(i) = mrsSys!������ & ""
        mrsSys.MoveNext
    Next
    MousePointer = 13
    'sglTimer0 = Timer
    If mrsTable Is Nothing Then Set mrsTable = ReadData(RDT_Table)
    If mrsFunction Is Nothing Then Set mrsFunction = ReadData(RDT_Function)
    If mrsRelasTree Is Nothing Then Call GetRelasTree(True)
    '�������ֵ�������ģ����Ȩ���������
    '#�����ݸ��µ�����
    Call UpdateGrantState
    pgb.Visible = True
    Set mclsPrivilege = New clsPrivilege
    Call mclsPrivilege.InitOracle(gcnOracle)
    If mclsPrivilege.InitPrivilege(str������, mstrRole, mrsModule, mrsTable, mrsFunction) Then
        'һ�����**********************************************************************************************************
        If mclsPrivilege.RevokePrivilege Then
        End If
        '������Ȩ**********************************************************************************************************
        If mclsPrivilege.GrantPrivilege Then
            '������Ҫ������־
            Call SaveAuditLog(2, "��ɫ��Ȩ", "�޸Ľ�ɫ��" & Split(mstrRole, "_")(1) & "����Ȩ��")
        End If
    End If
'    MsgBox Timer - sglTimer0
    MousePointer = 0
    mblnOk = True
    If mclsPrivilege.FailInfo <> "" Then
        MsgBox "������Ȩģ����󲻴��ڻ�Ȩ�޴����ԭ��" & vbCr & "����Ȩ��δ�������裺" & mclsPrivilege.FailInfo, vbExclamation, gstrSysName
    End If
    Set mclsPrivilege = Nothing
    Unload Me
End Sub

Private Sub cmdSelAll_Click()
    Dim objCur As Object
    If mintActive > 1 Then
        Set objCur = lvwFunc(mintActive Mod 2)
    Else
        Set objCur = tvwMenu(mintActive Mod 2)
    End If
    Call SetSel(objCur, True)
End Sub

Private Sub cmdUnSel_Click()
    Dim objCur As Object
    If mintActive > 1 Then
        Set objCur = lvwFunc(mintActive Mod 2)
    Else
        Set objCur = tvwMenu(mintActive Mod 2)
    End If
    Call SetSel(objCur, False)
End Sub

Private Sub Form_Activate()
    Dim strSql As String
    Dim lStyle As Long
    
    Call ApplyOEM(stbThis)
    lblNote.Caption = "�Խ�ɫ��" & Mid(mstrRole, 4) & "������Ȩ����"
    cmbSystem.Tag = "-1" '���㴦��
    Call SendMessage(tvwMenu(MT_����ģ��).hwnd, TVM_SETBKCOLOR, 0, ByVal &HEFF0E0)
    lStyle = GetWindowLong(tvwMenu(MT_����ģ��).hwnd, GWL_STYLE)
    Call SetWindowLong(tvwMenu(MT_����ģ��).hwnd, GWL_STYLE, lStyle - TVS_HASLINES)
    Call SetWindowLong(tvwMenu(MT_����ģ��).hwnd, GWL_STYLE, lStyle)
    mblnExpanded = False
    If mrsRelas Is Nothing Then
    '--- ��ʼ��Ȩ�޹�ϵ��¼��
        strSql = "Select ϵͳ, ���, ���, ����, nvl(����,0) as ���� From zlProgrelas Where ����=1"
        Set mrsRelas = gcnOracle.Execute(strSql)
        
        strSql = "Select ϵͳ, ���, ���, ����, ���� From zlProgrelas Where ��ϵ=1"
        Set mrsRelExcl = gcnOracle.Execute(strSql)
        
        strSql = "Select ϵͳ, ���, ����, ���, ��ϵ, ����, �����ϵ From Zlprogrelas"
        Set mrsGroup = gcnOracle.Execute(strSql)
    End If
    Call FillSystem
    If cmbSystem.ListCount < 1 Then
        Unload Me
        Exit Sub
    Else
        If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
    End If
    If cmbSystem.ListCount = 1 Then
        cmbSystem.Enabled = False
    End If
    Call LvwFlatColumnHeader(lvwFunc(MT_����ģ��))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, objNode As Node, lst As ListItem
    Dim strKey As String, strAll As String, strGrant As String
    Dim arrTmp As Variant
    
    If KeyCode = vbKeyF3 Then '������һ��
        mlngCurPos = FindModule(mlngCurPos)
    ElseIf KeyCode = vbKeyReturn Then   '������һ��
        If (TypeOf Me.ActiveControl Is TreeView) Then
            If Me.ActiveControl.Index = MT_ģ�� Then
                mlngCurPos = FindModule(mlngCurPos)
            End If
        End If
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then '�۵�չ������
        LockWindowUpdate Me.hwnd
        Call SynchronizeState '��ͬ��״̬
        If cmdExp.Caption = "ȫ���۵�(&D)" Then cmdExp.Tag = 1
        mblnExpanded = Not (mblnExpanded)
        For i = 0 To IIf(msftStyle = SFT_Ӧ��ϵͳ, 1, 0)
            For Each objNode In tvwMenu(i).Nodes
                objNode.Expanded = mblnExpanded
            Next
            If tvwMenu(i).Nodes.Count > 0 Then
                tvwMenu(i).Nodes(1).Selected = True
                tvwMenu(i).Nodes(1).EnsureVisible
            End If
        Next
        cmdExp.Tag = 0
        cmdExp.Caption = IIf(mblnExpanded, "ȫ���۵�(&D)", "ȫ��չ��(&D)")
        LockWindowUpdate 0
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then 'ȫѡ
        Call SetSel(Me.ActiveControl, True)
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then 'ȫ��
        Call SetSel(Me.ActiveControl, False)
    End If
End Sub

Private Sub SetSel(ByRef objCur As Object, Optional ByVal blnSelALl As Boolean = True)
    Dim i As Long, objNode As Node, lst As ListItem
    Dim strKey As String, strAll As String, strGrant As String
    Dim arrTmp As Variant
    
    LockWindowUpdate Me.hwnd
    If blnSelALl Then
        If (TypeOf objCur Is TreeView) Then '��λ�����β˵���ֻ�Ը����β�ȡȫѡ��ȡ��Ȩ��
            For Each objNode In objCur.Nodes
                objNode.Checked = True
                strKey = GetUpdateKey(objNode.Key)
                If strKey <> "" Then Call UpdateGrantState(strKey, True)
            Next
            If msftStyle <> SFT_���� Then Call SynchronizeState 'ͬ��״̬
            If objCur.Tag <> "" Then
                On Error Resume Next
                Set objNode = objCur.Nodes(objCur.Tag)
                If err.Number = 0 Then
                    Call tvwMenu_NodeClick(objCur.Index, objCur.Nodes(objCur.Tag))
                Else
                    err.Clear: objCur.Tag = ""
                End If
                On Error GoTo 0
            End If
        ElseIf (TypeOf objCur Is ListView) Then   '��λ�����β˵���ֻ�Ը����β�ȡȫѡ��ȡ��Ȩ��
            If objCur.Enabled Then
                For Each lst In objCur.ListItems
                    lst.Checked = True
                    strGrant = strGrant & "," & lst.Text
                Next
                strGrant = Mid(strGrant, 2)
                strKey = GetUpdateKey(tvwMenu(objCur.Index).Tag)
                If strKey <> "" Then
                    arrTmp = Split(strKey, "_")
                    strGrant = CheckFunc(Val(arrTmp(1)), Val(arrTmp(2)), objCur, strGrant)
                    Call UpdateGrantState(strKey, True, strGrant, 1)
                End If
                If msftStyle <> SFT_���� Then Call SynchronizeState '�ٴ�ͬ��״̬
            End If
        End If
    Else
        If (TypeOf objCur Is TreeView) Then '��λ�����β˵���ֻ�Ը����β�ȡȫѡ��ȡ��Ȩ��
            For Each objNode In objCur.Nodes
                objNode.Checked = False
                strKey = GetUpdateKey(objNode.Key)
                If strKey <> "" Then Call UpdateGrantState(strKey, False)
            Next
            If msftStyle <> SFT_���� Then Call SynchronizeState   'ͬ��״̬
            If objCur.Tag <> "" Then
                On Error Resume Next
                Set objNode = objCur.Nodes(objCur.Tag)
                If err.Number = 0 Then
                    Call tvwMenu_NodeClick(objCur.Index, objCur.Nodes(objCur.Tag))
                Else
                    err.Clear: objCur.Tag = ""
                End If
                On Error GoTo 0
            End If
        ElseIf (TypeOf objCur Is ListView) Then   '��λ�����β˵���ֻ�Ը����β�ȡȫѡ��ȡ��Ȩ��
            If objCur.Enabled Then
                For Each lst In objCur.ListItems
                    lst.Checked = False
                Next
                strKey = GetUpdateKey(tvwMenu(objCur.Index).Tag)
                If strKey <> "" Then
                    arrTmp = Split(strKey, "_")
                    strGrant = CheckFunc(Val(arrTmp(1)), Val(arrTmp(2)), objCur)
                    Call UpdateGrantState(strKey, True, strGrant, 1)
                End If
                If msftStyle <> SFT_���� Then Call SynchronizeState 'ͬ��״̬
            End If
        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub InitTips()
    Dim ObjTip  As clsTipSwap
    Dim i As Integer
    
    Set mcllTip = New Collection
    For i = 0 To 1
        Set ObjTip = New clsTipSwap
        Set ObjTip.ParentControl = lvwFunc(i)
        ObjTip.Icon = TTIconInfo
        ObjTip.Style = TTBalloon
        ObjTip.Create
        mcllTip.Add ObjTip, "T_" & i
    Next

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim lngHeight As Long, lngTop As Long
    On Error Resume Next
    '�����������߶�
    If Me.Height < 7000 Then Me.Height = 7000
    If Me.Width < 9300 Then Me.Width = 9300
    '���ý�ɫ��Ȩ�Ϸ�����
    cmbSystem.Left = Me.ScaleWidth - cmbSystem.Width - 60
    lblSys.Left = cmbSystem.Left - lblSys.Width - 30
    fraLine.Width = Me.Width + 100
    '���ý�ɫ��Ȩ�·�����λ��
    lngHeight = Me.ScaleHeight - stbThis.Height
    pgb.Top = lngHeight + (stbThis.Height - pgb.Height) / 2
    pgb.Left = stbThis.Panels(2).Left + Me.TextWidth("��") * 12
    pgb.Width = stbThis.Panels(2).Left + stbThis.Panels(2).Width - pgb.Left - 100
    cmdCancel.Top = lngHeight - cmdCancel.Height - 60
    cmdOK.Top = cmdCancel.Top
    cmdHelp.Top = cmdCancel.Top
    cmdExp.Top = cmdCancel.Top
    cmdSelAll.Top = cmdCancel.Top
    cmdUnSel.Top = cmdCancel.Top
    cmdCheck.Top = cmdCancel.Top
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    cmdCheck.Left = cmdOK.Left - cmdCheck.Width - 500
    lngHeight = cmdCancel.Top - 60 - picMenu(MT_ģ��).Top
    If fraHSplit.Tag = "" Then 'û���϶���6-4����
        fraHSplit.Top = picMenu(MT_ģ��).Top + lngHeight * 0.6
    End If
    If fraHSplit.Top - fraLine.Top - fraLine.Height < 2000 Then fraHSplit.Top = fraLine.Top + 2000 + fraLine.Height
    fraHSplit.Width = Me.ScaleWidth + 300
    If msftStyle = SFT_���� Then
        fraHSplit.Visible = False
        lngTop = lngHeight
    Else
        fraHSplit.Visible = True
        lngTop = fraHSplit.Top - picMenu(MT_ģ��).Top
    End If
    
    '���ý�ɫ��Ȩ�м�����
    If msftStyle = SFT_Ӧ��ϵͳ Then
        picMenu(MT_����ģ��).Visible = True
        fraVSplit.Visible = True
        lvwFunc(MT_ģ��).Visible = True
        lblFuncNote(MT_ģ��).Visible = True
        If fraVSplit.Tag = "" Then 'û���϶���7-3����
            fraVSplit.Left = Me.ScaleWidth * 0.7
        ElseIf Me.ScaleWidth - fraVSplit.Left < 2000 Then
            fraVSplit.Left = Me.ScaleWidth - 2000
        End If
        picMenu(MT_ģ��).Width = fraVSplit.Left - 15 - picMenu(MT_ģ��).Left
        picMenu(MT_����ģ��).Left = fraVSplit.Left + fraVSplit.Width + 15
        picMenu(MT_����ģ��).Width = Me.ScaleWidth - picMenu(MT_����ģ��).Left
        For i = 0 To 1
            picMenu(i).Height = lngHeight
            tvwMenu(i).Width = picMenu(i).ScaleWidth - tvwMenu(i).Left
            lvwFunc(i).Width = tvwMenu(i).Width
            tvwMenu(i).Height = lngTop - tvwMenu(i).Top - 30
            lblFuncNote(i).Top = lngTop + fraHSplit.Height + 30
            lvwFunc(i).Top = lblFuncNote(i).Top + lblFuncNote(i).Height + 30
            lvwFunc(i).Height = picMenu(i).ScaleHeight - lvwFunc(i).Top - 30
        Next
        fraVSplit.Top = fraLine.Top - 120
        fraVSplit.Height = cmdCancel.Top - 60 - fraVSplit.Top
    Else
        picMenu(MT_����ģ��).Visible = False
        fraVSplit.Visible = False
        picMenu(MT_ģ��).Width = Me.ScaleWidth - 30 - picMenu(MT_ģ��).Left
        picMenu(MT_ģ��).Height = cmdCancel.Top - 60 - picMenu(MT_ģ��).Top
        tvwMenu(MT_ģ��).Width = picMenu(MT_ģ��).ScaleWidth - tvwMenu(MT_ģ��).Left
        If msftStyle = SFT_�Զ��屨�� Then
            lvwFunc(MT_ģ��).Visible = True
            lblFuncNote(MT_ģ��).Visible = True
            lvwFunc(MT_ģ��).Width = tvwMenu(MT_ģ��).Width
            tvwMenu(MT_ģ��).Height = lngTop - tvwMenu(MT_ģ��).Top - 30
            lblFuncNote(MT_ģ��).Top = lngTop + fraHSplit.Height + 30
            lvwFunc(MT_ģ��).Top = lblFuncNote(MT_ģ��).Top + lblFuncNote(MT_ģ��).Height + 30
            lvwFunc(MT_ģ��).Height = picMenu(MT_ģ��).ScaleHeight - lvwFunc(MT_ģ��).Top - 30
        Else
             lvwFunc(MT_ģ��).Visible = False
             lblFuncNote(MT_ģ��).Visible = False
             tvwMenu(MT_ģ��).Height = picMenu(MT_ģ��).ScaleHeight - tvwMenu(MT_ģ��).Top - 30
        End If
    End If
    '˵��������
    If lvwFunc(0).View = lvwReport Then
        For i = 0 To 1
            lvwFunc(i).ColumnHeaders(3).Width = lvwFunc(i).Width - lvwFunc(i).ColumnHeaders(1).Width - lvwFunc(i).ColumnHeaders(2).Width
        Next
    End If
    '����״̬��չʾ���й�����ϵ
    If BLN_TEST Then
        If picMenu(MT_����ģ��).Visible Then
            tvwModRelas.Visible = True
            tvwModRelas.Left = tvwMenu(MT_����ģ��).Left + tvwMenu(MT_����ģ��).Width * 0.5
            tvwModRelas.Width = tvwMenu(MT_����ģ��).Width * 0.5
            tvwModRelas.Top = tvwMenu(MT_����ģ��).Top
            tvwModRelas.Height = tvwMenu(MT_����ģ��).Height
            tvwModRelas.ZOrder
        End If
    End If
    Exit Sub
'ErrH:
'    If 0 = 1 Then
'        Resume
'    End If
'    MsgBox "From_Resize:" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ClearDataAndVar(True)
End Sub

Private Sub fraHSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then fraHSplit.Top = fraHSplit.Top + Y
End Sub

Private Sub fraHSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fraHSplit.Top - fraLine.Top < 2500 Then fraHSplit.Top = fraLine.Top + 2500
    If fraHSplit.Top > picMenu(0).Height + picMenu(0).Top - 2000 Then fraHSplit.Top = picMenu(0).Height + picMenu(0).Top - 2000
    fraHSplit.Tag = "�϶�"
    Call Form_Resize
End Sub

Private Sub fraVSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 1 Then fraVSplit.Left = fraVSplit.Left + X
End Sub

Private Sub fraVSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fraVSplit.Left < 3500 Then fraVSplit.Left = 3500
    If fraVSplit.Left > Me.ScaleWidth - 3500 Then fraVSplit.Left = Me.ScaleWidth - 3500
    fraVSplit.Tag = "�϶�"
    Call Form_Resize
End Sub

Private Function ReadData(ByVal rdtInput As ReadDataType) As ADODB.Recordset
'���ܣ���ȡ����
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    Select Case rdtInput
        Case RDT_Menu
            strSql = "Select Level As ���, Id As ���, Nvl(�ϼ�id, 0) As �ϼ�, ����, Decode(Nvl(�̱���, '��'), '��', ����, �̱���) As �̱���, ���, ˵��," & vbNewLine & _
                            "       Nvl(ģ��, 0) As ģ��, Nvl(ϵͳ, 0) As ϵͳ, Nvl(ͼ��, 0) As ͼ��" & vbNewLine & _
                            "From Zlmenus" & vbNewLine & _
                            "Where ��� In ('ȱʡ', '����')" & vbNewLine & _
                            "Start With �ϼ�id Is Null" & vbNewLine & _
                            "Connect By Prior Id = �ϼ�id"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "ģ��˵���ȡ", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "���,���,�ϼ�,����,ģ��,ϵͳ,ͼ��")
        Case RDT_Module
            If gblnInIDE Then '���Ի������������Ȩ
                strSql = "Select F.���, F.����, F.ϵͳ, F.����, F.ȱʡֵ, Decode(R.����, Null, 0, 1) As ��Ȩ��" & vbNewLine & _
                            "From (Select C.���, C.����, Nvl(B.ϵͳ, 0) ϵͳ, B.����, Nvl(B.ȱʡֵ, 0) ȱʡֵ" & vbNewLine & _
                            "       From Zlprogfuncs b, Zlprograms c" & vbNewLine & _
                            "       Where Nvl(C.ϵͳ, 0) = Nvl(B.ϵͳ, 0) And C.��� = B.���) f," & vbNewLine & _
                            "     (Select Nvl(A.ϵͳ, 0) ϵͳ, A.���, A.���� From Zlrolegrant a Where A.��ɫ = [1]) r" & vbNewLine & _
                            "Where F.ϵͳ = R.ϵͳ(+) And F.��� = R.���(+) And F.���� = R.����(+)" & vbNewLine & _
                            "Order By F.���"
            Else
                strSql = "Select G.���, G.����, Nvl(G.ϵͳ,0) ϵͳ, F.����, Nvl(F.ȱʡֵ, 0) As ȱʡֵ, Decode(R.����, Null, 0, 1) As ��Ȩ��" & vbNewLine & _
                                "From Zlprograms  g," & vbNewLine & _
                                "     (Select ϵͳ, ���, ����, ȱʡֵ From Zlprogfuncs Where ϵͳ Is Null Or (��� Between 10000 And 19999)" & vbNewLine & _
                                "       Union" & vbNewLine & _
                                "       Select A.ϵͳ, A.����id As ���, A.����, 1 As ȱʡֵ From Zlreports b, Zlrptputs a Where A.����id = B.Id And B.ϵͳ Is Null" & vbNewLine & _
                                "       Union" & vbNewLine & _
                                "       Select F.ϵͳ, F.���, F.����, F.ȱʡֵ From Zlprogfuncs f, Zlregfunc r" & vbNewLine & _
                                "       Where Trunc(F.ϵͳ / 100) = R.ϵͳ And F.��� = R.��� And F.���� = R.���� And" & vbNewLine & _
                                "             1 = (Select 1 From Zlregaudit a Where A.��Ŀ = '��Ȩ֤��')) f," & vbNewLine & _
                                "     (Select Nvl(ϵͳ, 0) ϵͳ, ���, ��ɫ, ���� From Zlrolegrant Where ��ɫ = [1]) r" & vbNewLine & _
                                "Where Nvl(G.ϵͳ, 0) = Nvl(F.ϵͳ, 0) And G.��� = F.��� And F.��� = R.���(+) And F.���� = R.����(+) And Nvl(F.ϵͳ, 0) = R.ϵͳ(+)" & vbNewLine & _
                                "Order By ���"
            End If
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "��Ȩģ���ȡ", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "���,����,ϵͳ,����,��Ȩ��,ȱʡֵ")
        Case RDT_Function
            strSql = "Select S.���� ϵͳ��, S.��� ϵͳ, S.������, Upper(F.������) As ������, F.˵��,F.������, Decode(U.Table_Name, Null, 0, 1) ��Ȩ��" & vbNewLine & _
                        "From Zlsystems s, Zlfunctions f," & vbNewLine & _
                        "     (Select Table_Schema As Owner, Grantee, Table_Name From All_Tab_Privs Where Table_Schema = User) u," & vbNewLine & _
                        "     (Select Table_Schema As ������, Table_Name As ����, Privilege As Ȩ��" & vbNewLine & _
                        "       From All_Tab_Privs" & vbNewLine & _
                        "       Where Privilege = 'EXECUTE' And Grantable = 'YES'" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select Owner, Object_Name, 'EXECUTE' From All_Objects Where Owner = User And Object_Type = 'FUNCTION') r" & vbNewLine & _
                        "Where F.ϵͳ = S.��� And S.������ = R.������ And Upper(F.������) = R.���� And U.Grantee(+) = [1] And U.Owner(+) = R.������ And" & vbNewLine & _
                        "      U.Table_Name(+) = R.����"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "����������ȡ", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "������,������,˵��,ϵͳ,ϵͳ��,������,��Ȩ��,��Ȩ�� �����Ȩ")
        Case RDT_Table
            strSql = "Select T.ϵͳ, T.ϵͳ��, T.������, T.����, T.˵��, Decode(R.Table_Name, Null, 0, 1) ��Ȩ��" & vbNewLine & _
                        "From (Select S.���� ϵͳ��, S.��� ϵͳ, S.������, B.����, B.˵�� From Zlsystems s, Zlbasecode b Where B.ϵͳ = S.���) t," & vbNewLine & _
                        "     (Select Table_Schema As ������, Table_Name As ����" & vbNewLine & _
                        "       From All_Tab_Privs" & vbNewLine & _
                        "       Where Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE') And Grantable = 'YES' And" & vbNewLine & _
                        "             (Table_Schema, Table_Name) In" & vbNewLine & _
                        "             (Select S.������, B.���� From Zlsystems s, Zlbasecode b Where B.ϵͳ = S.��� And S.������ = User)" & vbNewLine & _
                        "       Group By Table_Schema, Table_Name" & vbNewLine & _
                        "       Having Count(Privilege) = 4" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select User, Object_Name" & vbNewLine & _
                        "       From User_Objects" & vbNewLine & _
                        "       Where Object_Type = 'TABLE' And" & vbNewLine & _
                        "             Object_Name In (Select B.���� From Zlsystems s, Zlbasecode b Where B.ϵͳ = S.��� And S.������ = User)) g," & vbNewLine & _
                        "     (Select Grantor As Owner, Table_Name" & vbNewLine & _
                        "       From All_Tab_Privs" & vbNewLine & _
                        "       Where Grantor = User And Grantee = [1] And Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
                        "       Group By Grantor, Table_Name" & vbNewLine & _
                        "       Having Count(Privilege) = 4) r" & vbNewLine & _
                        "Where T.������ = G.������ And T.���� = G.���� And T.������ = R.Owner(+) And T.���� = R.Table_Name(+)"
'            "Select T.ϵͳ,T.ϵͳ�� ,T.������, T.����,T.˵��, Decode(R.Table_Name, Null, 0, 1) ��Ȩ��" & vbNewLine & _
'                            "From (Select S.���� ϵͳ��,S.���  ϵͳ, S.������, B.����, B.˵��" & vbNewLine & _
'                            "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
'                            "       Where B.ϵͳ = S.���) t," & vbNewLine & _
'                            "     (Select ������, ����" & vbNewLine & _
'                            "       From (Select Table_Schema As ������, Table_Name As ����, Privilege As Ȩ�� From All_Tab_Privs Where Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE') And Grantable = 'YES'  Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'DELETE' From All_Objects Where Owner = User And Object_Type = 'TABLE' Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'INSERT' From All_Objects Where Owner = User And Object_Type = 'TABLE' Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'SELECT' From All_Objects Where Owner = User And Object_Type = 'TABLE' Union" & vbNewLine & _
'                            "              Select Owner, Object_Name, 'UPDATE' From All_Objects Where Owner = User And Object_Type = 'TABLE')" & vbNewLine & _
'                            "       Group By ������, ����" & vbNewLine & _
'                            "       Having Count(Ȩ��) = 4) g," & vbNewLine & _
'                            "     (Select Grantor As Owner, Table_Name" & vbNewLine & _
'                            "       From All_Tab_Privs" & vbNewLine & _
'                            "       Where Grantor = User And Grantee =[1] And Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
'                            "       Group By Grantor, Table_Name" & vbNewLine & _
'                            "       Having Count(Privilege) = 4) r" & vbNewLine & _
'                            "Where T.������ = G.������ And T.���� = G.���� And T.������ = R.Owner(+) And T.���� = R.Table_Name(+)"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "���������ȡ", mstrRole)
            Set ReadData = CopyNewRec(rsTmp, , "����,˵��,ϵͳ,ϵͳ��,������,��Ȩ��,��Ȩ�� �����Ȩ")
        Case RDT_Systems
            If gblnInIDE Then
                rsTmp.CursorLocation = adUseClient
                '��ʾ�������е�ϵͳ
                Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
                Set ReadData = rsTmp
            Else
                Set ReadData = zlGetRegSystems
            End If
        Case RDT_ModRelas
            strSql = "Select ϵͳ, ģ��, ����, ���ϵͳ, ���ģ��, ����, ȱʡֵ, �������, �����Ϣ, ����" & vbNewLine & _
                        "From (With a As (Select A.ϵͳ, A.ģ��, A.����, Nvl(A.���ϵͳ, 0) ���ϵͳ, A.���ģ��, B.����, A.ȱʡֵ, A.�������," & vbNewLine & _
                        "                        A.���� || ',' || A.��ع��� || ',' || A.������� || ',' || A.ȱʡֵ �����Ϣ" & vbNewLine & _
                        "                 From Zlmodulerelas a, Zlprograms b" & vbNewLine & _
                        "                 Where Nvl(A.ϵͳ, 0) = Nvl(B.ϵͳ, 0) And A.ģ�� = B.���)" & vbNewLine & _
                        "       Select A.ϵͳ, A.ģ��, A.����, A.���ϵͳ, A.���ģ��, A.����, Decode(Sum(A.ȱʡֵ), 0, 0, 1) ȱʡֵ, Decode(Sum(A.�������), 0, 0, 1) �������," & vbNewLine & _
                        "              F_List2str(Cast(Collect(A.�����Ϣ) As T_Strlist), ';') �����Ϣ, 0 ����" & vbNewLine & _
                        "       From a" & vbNewLine & _
                        "       Group By A.ϵͳ, A.ģ��, A.����, A.���ϵͳ, A.���ģ��, A.����" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select A.ϵͳ, A.ģ��, Null ����, A.���ϵͳ, A.���ģ��, A.����, Decode(Sum(A.ȱʡֵ), 0, 0, 1) ȱʡֵ," & vbNewLine & _
                        "              Decode(Sum(A.�������), 0, 0, 1) �������, F_List2str(Cast(Collect(A.�����Ϣ) As T_Strlist), ';') �����Ϣ, 1 ����" & vbNewLine & _
                        "       From a" & vbNewLine & _
                        "       Group By A.ϵͳ, A.ģ��, A.���ϵͳ, A.���ģ��, A.����)"
            Set ReadData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "ģ���ϵ��ȡ")
    End Select
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "ReadData:" & err.Description, vbInformation, Me.Caption
End Function

Private Sub ClearDataAndVar(Optional ByVal blnFormUnload As Boolean)
'���ܣ����һЩ����
    Set mclsPrivilege = Nothing
    Set mrsModule = Nothing
    Set mrsTable = Nothing
    Set mrsFunction = Nothing
    Set mrsModsInfo = Nothing
    Set mrsTree = Nothing
    Set mcllHaveSys = Nothing
    Set mcllKeyModule = Nothing
    Set mcllTip = Nothing
    Set mclsPrivilege = Nothing
    Set mrsRelasTree = Nothing
    'ģ���ϵ���Լ����ܹ�ϵ��ϵͳ�б���Բ���գ��Ա��´�ʹ��
End Sub

Private Sub FillSystem()
    '��ʾ�������е�ϵͳ
    If mrsSys Is Nothing Then Set mrsSys = ReadData(RDT_Systems)
    If glngSysNo <> -1 Then mrsSys.Filter = "��� = " & glngSysNo
    cmbSystem.Clear
    mrsSys.Sort = "���"
    With mrsSys
        Do Until .EOF
            cmbSystem.addItem RPAD(!���� & "��" & !��� & "��", 25) & " v" & !�汾��
            cmbSystem.ItemData(cmbSystem.NewIndex) = !��� & ""
            If !������ & "" = UCase(gstrUserName) And cmbSystem.ListIndex < 0 Then
                cmbSystem.ListIndex = cmbSystem.NewIndex
            End If
            .MoveNext
        Loop
    End With
    '������ϵͳ�ǳ���̶���
    If (gobjRegister.zlRegTool And 2) = 2 Then cmbSystem.addItem "�Զ��屨��"
    cmbSystem.addItem "��������"
    cmbSystem.addItem "ȡ������"
    cmbSystem.addItem "��������"
    If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
End Sub

Private Function GetModuleInfo() As ADODB.Recordset
'����:��ȡģ����Ϣ
    Dim rsReturn As ADODB.Recordset
    Dim strGrant As String, strDefault As String, intGrant As Integer, strAll As String
    Dim lngSys As Long, lng��� As Long, str���� As String
    
    Dim strPre As String
    
    On Error GoTo errH
    If mrsModule Is Nothing Then Set mrsModule = ReadData(RDT_Module)
    'ģ������=0:����ģ��
    '              =1:ʵ��ģ��
    '              =2:�Զ��屨��ģ����������
    
    'ģ�����=0=���������ģ�鲻������
    '               1=���������ģ�鱻����
    '               2=�������ģ�鲻������
    '               3=�������ģ�鱻����
    Set rsReturn = CopyNewRec(mrsModule, True, "���,����,ϵͳ,��Ȩ��", Array("��Ȩ�ı�", adInteger, 1, 0, "�ı�����", adInteger, 1, 0, _
                                                        "��Ȩ����", adVarChar, 2000, Empty, "Ĭ�Ϲ���", adVarChar, 2000, Empty, _
                                                        "���й���", adVarChar, 2000, Empty, "ģ������", adInteger, 1, 0, "ģ�����", adInteger, 1, 0))
    mrsModule.Filter = ""
    mrsModule.Sort = "ϵͳ,���,����"
    With mrsModule
        Do While Not .EOF
            If strPre <> !ϵͳ & "_" & !��� Then
                If strPre <> "" Then
                    rsReturn.AddNew Array("���", "����", "ϵͳ", "��Ȩ��", "��Ȩ����", "Ĭ�Ϲ���", "���й���", "ģ������", "��Ȩ�ı�", "�ı�����", "ģ�����"), _
                                                    Array(lng���, str����, lngSys, intGrant, Mid(strGrant, 2), Mid(strDefault, 2), Mid(strAll, 2), 0, 0, 0, 0)
                End If
                strGrant = "": strDefault = "": intGrant = 0: strAll = ""
                lngSys = !ϵͳ: str���� = !���� & "": lng��� = !���
                strPre = !ϵͳ & "_" & !���
            End If
            If !��Ȩ�� = 1 Then
                intGrant = 1
            End If
            If !���� & "" <> "����" Then
                If !��Ȩ�� = 1 Then strGrant = strGrant & "," & !����
                If !ȱʡֵ = 1 Then strDefault = strDefault & "," & !����
                strAll = strAll & "," & !����
            End If
            .MoveNext
        Loop
        '���һ��ģ��ļ���
        If strPre <> "" Then
            rsReturn.AddNew Array("���", "����", "ϵͳ", "��Ȩ��", "��Ȩ����", "Ĭ�Ϲ���", "���й���", "ģ������", "��Ȩ�ı�", "�ı�����", "ģ�����"), _
                                            Array(lng���, str����, lngSys, intGrant, Mid(strGrant, 2), Mid(strDefault, 2), Mid(strAll, 2), 0, 0, 0, 0)
        End If
    End With
    Set GetModuleInfo = rsReturn
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "GetModuleInfo:" & err.Description, vbInformation, Me.Caption
End Function

Private Function SetVirtualVisual(ByVal blnVisual As Boolean, Optional ByRef strFirstNode As String) As Boolean
'���ܣ���������ģ��ɼ���
'������blnVisual=�Ƿ�ɼ�
'���أ���һ���ڵ��Key
    Dim objNode As Node
    
    On Error GoTo errH
    If msftStyle <> SFT_Ӧ��ϵͳ Then
        chkVirtual.Visible = False: Exit Function
    End If
    chkVirtual.Visible = True '�ظ��ɼ�,���ݸ�ϵͳ״������������
    mblnVirtual = True
    '������ʵ��ģ�飬��ֻ��ʾ����ģ�飬�����ظ�ѡ��
    mrsModsInfo.Filter = "ϵͳ=" & mlngSys & " And ģ������=1" & IIf(chkOnlyShow.value = 1, " And ��Ȩ��=1", "")
    If mrsModsInfo.EOF Then
        tvwMenu(MT_ģ��).Nodes.Clear '������нڵ�
        tvwMenu(MT_ģ��).Tag = ""
        strFirstNode = ""
        blnVisual = True
        chkVirtual.value = 1
        chkVirtual.Visible = False
    End If
    '����������ģ�飬�����ظ�ѡ��
    mrsModsInfo.Filter = "ϵͳ=" & mlngSys & " And ģ������=0  " & IIf(chkOnlyShow.value = 1, " And ��Ȩ��=1", "")
    If mrsModsInfo.EOF Then
        blnVisual = False
        chkVirtual.value = 0
        chkVirtual.Visible = False
    End If
    mblnVirtual = False
    On Error Resume Next
    tvwMenu(MT_ģ��).Nodes.Remove "V_" & mlngSys
    err.Clear: On Error GoTo errH
    If blnVisual Then
        Set objNode = tvwMenu(MT_ģ��).Nodes.Add(, , "V_" & mlngSys, "����ģ��", "����", "����_ѡ��")
        If chkOnlyShow.value = 1 Then objNode.Checked = True
        With mrsModsInfo
            Do While Not .EOF
                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("V_" & mlngSys, 4, "M_000000_" & !���, "��" & Format(!���, "000000") & "��" & !����, "Module")
                objNode.Checked = !��Ȩ�� = 1
                '��Ĭ�Ϲ�ѡ�ڵ㣬��ֹ�ٶȽ���
'                '��ȡ��һ�ڵ����һ��ѡ�ڵ�
'                If objNode.Checked And tvwMenu(MT_ģ��).Tag = "" Then
'                    tvwMenu(MT_ģ��).Tag = objNode.Key
'                End If
'                If strFirstNode = "" Then
'                    strFirstNode = objNode.Key
'                End If
                .MoveNext
            Loop
        End With
    End If
    SetVirtualVisual = True '���ص�һ���ڵ�
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "SetVirtualVisual:" & err.Description, vbInformation, Me.Caption
End Function

Private Function SetOnlyShow(Optional ByVal strKey As String) As Boolean
'���ܣ�ֻ��ʾѡ����
'������strKey=�ж�ĳ���ڵ㣬Ϊ�����ж���������
'���أ�True=���Ӽ��ڵ��û���Ӽ��ڵ��ҵ�ǰ�ڵ㱻��ѡ��False=�������
    Dim objNode As Node
    Dim strDelKeys As String, strTMp As String
    Dim arrTmp As Variant, i As Long
    
    If chkOnlyShow.value = 1 Then
        If tvwMenu(MT_ģ��).Nodes.Count = 0 Then Exit Function
        '��ȡ��ʼ�ڵ㣬�Ա�ѭ��
        If strKey = "" Then
            If GetUpdateKey(tvwMenu(MT_ģ��).Tag) <> "" Then
                If Not tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag).Checked Then
                    tvwMenu(MT_ģ��).Tag = ""
                End If
            Else
                tvwMenu(MT_ģ��).Tag = ""
            End If
            Set objNode = tvwMenu(MT_ģ��).Nodes(1)
        ElseIf tvwMenu(MT_ģ��).Nodes(strKey).Children <> 0 Then
            Set objNode = tvwMenu(MT_ģ��).Nodes(strKey).Child
        End If
        '��ȡ����ɾ���Ľڵ�
        Do While Not objNode Is Nothing
            '���Ӽ���ѡ�У��򸸼�ѡ��
            objNode.Checked = SetOnlyShow(objNode.Key)
            If Not objNode.Checked Then
                strDelKeys = strDelKeys & "|" & objNode.Key
            Else
                If tvwMenu(MT_ģ��).Tag = "" And GetUpdateKey(objNode.Key) <> "" Then
                    tvwMenu(MT_ģ��).Tag = objNode.Key
                End If
            End If
            Set objNode = objNode.Next
        Loop
        'ɾ������ɾ���Ľڵ�
        arrTmp = Split(Mid(strDelKeys, 2), "|")
        For i = LBound(arrTmp) To UBound(arrTmp)
            tvwMenu(MT_ģ��).Nodes.Remove arrTmp(i)
        Next
        If strKey <> "" Then
            '����ǰ�����ڵ���δɾ�����Ӽ��ڵ㣬���ñ��������ڵ�
            If tvwMenu(MT_ģ��).Nodes(strKey).Children <> 0 Then
                Set objNode = tvwMenu(MT_ģ��).Nodes(strKey).Child
                SetOnlyShow = True
            Else '�������ڵ�û���ӽڵ㣬�жϽڵ��Ƿ�ѡ��
                SetOnlyShow = tvwMenu(MT_ģ��).Nodes(strKey).Checked
            End If
        End If
        If strKey = "" Then
            Call SetVirtualVisual(chkVirtual.value <> 0)
            'չ����ǽڵ�
            If tvwMenu(MT_ģ��).Tag <> "" Then
                strTMp = tvwMenu(MT_ģ��).Tag: tvwMenu(MT_ģ��).Tag = ""
                Call SetNodeExpand(tvwMenu(MT_ģ��), strTMp) 'չ���ڵ�
                Call tvwMenu_NodeClick(MT_ģ��, tvwMenu(MT_ģ��).Nodes(strTMp))
            End If
        End If
    Else
        cmbSystem.Tag = "-1"
        Call cmbSystem_Click
    End If
End Function

Private Function FindModule(Optional ByVal intCurPosition As Long, Optional ByVal blnSmart As Boolean) As Long
'���ܣ�����ģ�����
'������intCurPosition=��ǰλ�ã�<=1��ʾ��ͷ��β��ʼ���ң�����ӵ�ǰλ�ÿ�ʼ����
'          blnSmart=True-��������������������Ӧ,False-��������Ӧ������
'���أ�ƥ����Ŀλ��
    Dim i As Integer
    Dim blnFind As Boolean
    Dim strLike As String, strKeyLike As String
    Dim objNode As Node
    Dim strMsg As String
    Dim strName As String
    
    On Error Resume Next
    If intCurPosition < 0 Then FindModule = -1: Exit Function
    '��ʼλ�ô���
    If intCurPosition >= tvwMenu(MT_ģ��).Nodes.Count Then
        intCurPosition = 0
    End If
    '�����ַ�������
    strName = cmbSystem.Text
    If msftStyle = SFT_���� Then
        strLike = "*" & mstrFind & "*"
        If cmbSystem.Text = "��������" Then
            strKeyLike = "T_*_*"
        Else 'ȡ������
            strKeyLike = "F_*_*"
        End If
    Else
        If msftStyle <> SFT_�Զ��屨�� Then
            strName = "ģ��"
        End If
        If IsNumeric(mstrFind) Then '����Ų���
            strLike = "��*" & mstrFind & "*��*"
        Else '�����Ʋ��
            strLike = "��*��*" & mstrFind & "*"
        End If
        strKeyLike = "M_*_*"
    End If
    '���в���
    For i = intCurPosition + 1 To tvwMenu(MT_ģ��).Nodes.Count
        Set objNode = tvwMenu(MT_ģ��).Nodes(i)
        If objNode.Key Like strKeyLike Then
            If objNode.Text Like strLike Then
                objNode.Expanded = True
                objNode.Selected = True: blnFind = True
                Exit For
            End If
        End If
    Next
    'δ���ҵ�ԭ����ʾ
    If Not blnFind Then
        If mblnReturn Then
            If mlngCurPos <= 1 Then
                If chkOnlyShow.value = 1 Then
                    strMsg = "�������ҵ�" & strName & "δ��ѡ����ȡ����ѡ""������Ȩ""�ٽ��в��ң�"
                ElseIf chkVirtual.Visible And chkVirtual.value = 0 Then
                    strMsg = "�������ҵ�" & strName & "����������ģ�飬�빴ѡ""������ģ��""�ٽ��в��ң�"
                Else
                    strMsg = "δ�ҵ�ƥ���" & strName & "��"
                End If
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, Me.Caption
                End If
                mlngCurPos = -1
                '��ʾ�Ƿ��ͷ��ʼ����
            Else
                If MsgBox("δ�ҵ�ƥ���" & strName & "���Ƿ����½��в���", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    mlngCurPos = 0
                    mlngCurPos = FindModule(mlngCurPos)
                    FindModule = mlngCurPos
                Else
                    FindModule = -1
                End If
            End If
        Else
            FindModule = -1
        End If
    Else
        FindModule = i
        Call tvwMenu_NodeClick(MT_ģ��, objNode)
    End If
End Function

Private Sub lvwFunc_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvwFunc(Index).SortKey = ColumnHeader.Index - 1 Then
        lvwFunc(Index).SortOrder = IIf(lvwFunc(Index).SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwFunc(Index).SortKey = ColumnHeader.Index - 1
        lvwFunc(Index).SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwFunc_DblClick(Index As Integer)
    Dim objList As ListItem
    On Error Resume Next
    Set objList = lvwFunc(Index).ListItems(lvwFunc(Index).Tag)
    err.Clear: On Error GoTo 0
    If Not objList Is Nothing Then
        objList.Checked = Not objList.Checked
        Call lvwFunc_ItemCheck(Index, objList)
    End If
End Sub

Private Sub lvwFunc_GotFocus(Index As Integer)
    mintActive = Index + 2
End Sub

Private Sub lvwFunc_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim arrTmp As Variant
    Dim lngϵͳ As Long, lng��� As Long
    Dim strPrivs As String, strTMp As String
    Dim objItem As ListItem
    Dim blnChange As Boolean
    
    arrTmp = Split(tvwMenu(Index).Tag, "_")
    If Index = MT_ģ�� Then
        lngϵͳ = mlngSys
        lng��� = Val(arrTmp(2))
    Else
        lngϵͳ = Val(arrTmp(3))
        lng��� = Val(arrTmp(4))
    End If
    If mblnItem Then Exit Sub
    If Item.Checked Then
        For Each objItem In lvwFunc(Index).ListItems
            If objItem.Checked = True Then
                strPrivs = strPrivs & "," & objItem.Text
            End If
        Next
        If strPrivs <> "" Then strPrivs = Mid(strPrivs, 2)
        '�����ϵ,�ڴ˴���
        mrsRelExcl.Filter = "ϵͳ = " & lngϵͳ & " And ��� = " & lng��� & " And ���� = '" & Item.Text & "'"
        If Not mrsRelExcl.EOF Then
            mrsGroup.Filter = "ϵͳ = " & lngϵͳ & " And ��� = " & lng��� & " And ��� = " & mrsRelExcl!���
            If Not mrsGroup.EOF Then
                strPrivs = setExcl(mrsGroup, mrsRelExcl!���� & "_" & mrsRelExcl!���, Item.Checked, lvwFunc(Index), strPrivs)
            End If
        End If
    End If
    '���ӹ�ϵ,�ڴ˴���
    strTMp = strPrivs
    strPrivs = CheckFunc(lngϵͳ, lng���, lvwFunc(Index), strPrivs)
    blnChange = strTMp <> strPrivs
    '������Ȩ���
    Call UpdateGrantState("M_" & lngϵͳ & "_" & lng���, True, strPrivs, 1)
    '��һ��ͬ����ѡģ�����Ȩ������ᷢ���������
    '��ȡ��ģ��̶�����ģ�����ұߵĶ�Ӧģ�飬��ֱ��ȡ����ģ����Ȩ������ȡ���Ĺ̶�����ģ���ֱ���ѡ��
    Call lvwFunc_ItemClick(Index, Item)
    If Index = MT_ģ�� Then
        blnChange = blnChange Or CheckNode(tvwMenu(MT_����ģ��))
    End If
    '������Ȩ�ı䣬��ͬ����¼��
    If blnChange Then Call SynchronizeState
End Sub

Private Sub lvwFunc_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim strKey As String, arrTmp As Variant
    
    If lvwFunc(Index).Tag <> Item.Key Or mstrCurRelas Like "*_*_" Then '�ϴ�ѡ��ģ��
        lvwFunc(Index).Tag = Item.Key
        If Index = MT_ģ�� Then  'ģ������
            '���ع���ģ��
            strKey = mlngSys & "_" & Val(Split(tvwMenu(Index).Tag, "_")(2)) & "_" & Item.Text
            Call FillRelasModule(strKey)
        End If
    End If
    Item.Selected = True
End Sub

Private Sub lvwFunc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objItem As ListItem, strTip As String
    
    Set objItem = lvwFunc(Index).HitTest(X, Y)
    If mcllTip Is Nothing Then Call InitTips
    If Not objItem Is Nothing Then
        strTip = objItem.SubItems(2)
        If strTip = "" Then strTip = "<��˵����Ϣ>"
        mcllTip("T_" & Index).TipText = SwapText(strTip)
        mcllTip("T_" & Index).Title = objItem.Text
    Else
        mcllTip("T_" & Index).TipText = ""
        mcllTip("T_" & Index).Title = ""
    End If
End Sub

Private Function SwapText(ByVal strTxt As String) As String
    
    Dim strReturn As String, strTMp As String, i As Integer
    strReturn = strTxt
    If InStr(strTxt, ";") > 0 Then
        strReturn = SwapWord(strReturn, ";")
    End If
    If InStr(strTxt, "��") > 0 Then
        strReturn = SwapWord(strReturn, "��")
    End If
    If InStr(strTxt, ".") > 0 Then
        strReturn = SwapWord(strReturn, ".")
    End If
    If InStr(strTxt, "��") > 0 Then
        strReturn = SwapWord(strReturn, "��")
    End If
    
    If strReturn = strTxt Then
        strReturn = swapLine("����" & strTxt)
    End If
    '--
    strReturn = Replace(strReturn, " ", "")
    strReturn = Replace(strReturn, "��", "")
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR];[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR].[CR]", "[CR]")
    
    '���ڻ���
    Dim varLine As Variant
    
    varLine = Split(strReturn, "[CR]")
    For i = 0 To UBound(varLine)
        strTMp = strTMp & swapLine("����" & varLine(i)) & vbNewLine
    Next
    
    If strTMp <> "" Then
        strReturn = strTMp
    End If
    '--���������Ŀ���
    strReturn = ClearLine(strReturn)
    SwapText = strReturn
End Function

Private Function ClearLine(strTxt) As String
    Dim i As Integer, Y As Integer
    Dim varLine As Variant
    Dim strReturn As String
    varLine = Split(strTxt, vbNewLine)
    For i = 0 To UBound(varLine)
        If InStr(",.;?!])}%>���������������ݣ�����������", Mid(varLine(i), 1, 1)) > 0 Then
            strReturn = Mid(strReturn, 1, Len(strReturn) - 4) & Mid(varLine(i), 1, 1) & "[CR]" & Mid(varLine(i), 2) & "[CR]"
        Else
            strReturn = strReturn & varLine(i) & "[CR]"
        End If
    Next
    
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR];[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR].[CR]", "[CR]")
    
    strReturn = Replace(strReturn, "[CR][CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]", vbNewLine)
    ClearLine = strReturn
End Function

Private Function SwapWord(ByVal strTxt As String, strWord As String) As String
    Dim varLine As Variant
    Dim strReturn As String
    Dim i As Integer
    Dim strTxtTmp As String
    
    strTxtTmp = strTxt
    If Mid(strTxt, Len(strTxt), 1) = strWord Then
        strTxtTmp = Mid(strTxt, 1, Len(strTxt) - 1)
    End If
    
    If InStr(strTxtTmp, strWord) > 0 Then
        varLine = Split(strTxtTmp, strWord)
        For i = 0 To UBound(varLine)
            If varLine(i) <> "" Then
                'varLine(i) = swapLine("����" & varLine(i))
                If varLine(i) & strWord <> strWord Then
                    strReturn = strReturn & varLine(i) & strWord & "[CR]"
                End If
            End If
        Next
    End If
    'If Mid(strTxtTmp, Len(strTxtTmp), 1) <> strWord Then strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
    If strReturn <> "" Then
        SwapWord = strReturn
    Else
        SwapWord = strTxt
    End If
End Function

Private Function swapLine(ByVal strTxt As String) As String
    Dim strTMp As String
    strTMp = strTxt
    
    If Len(strTxt) > 18 Then
        swapLine = Mid(strTMp, 1, 18) & vbNewLine
        strTMp = Mid(strTMp, 19)
        swapLine = swapLine & swapLine(strTMp)
    Else
        swapLine = strTxt
    End If
End Function

Private Sub lvwFunc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call PopupMenu(Me.mnuPopu, 2)
    End If
End Sub

Private Sub mclsPrivilege_AfterProgress()
    stbThis.Panels(2).Text = ""
    pgb.value = 0
    DoEvents
End Sub

Private Sub mclsPrivilege_BeforeProgress(ByVal Title As String, ByVal Max As Long)
    stbThis.Panels(2).Text = Title
    pgb.Max = Max
    DoEvents
End Sub

Private Sub mclsPrivilege_Progressing(ByVal Progress As Long)
    pgb.value = Progress
End Sub

Private Sub mnuPopuState_Click(Index As Integer)
    Dim i As Integer
    mnuPopuState(0).Checked = (Index = 0)
    mnuPopuState(1).Checked = (Index = 1)
    For i = 0 To 1
        lvwFunc(i).View = IIf(Index = 0, lvwSmallIcon, lvwReport)
    Next
End Sub

Private Sub tvwMenu_Click(Index As Integer)
    Dim blnDo As Boolean, objNode As Node
    '�̶�������ϵ����ȡ����ѡ
    If Index = MT_����ģ�� And tvwMenu(MT_����ģ��).Tag <> "" Then
        Set objNode = tvwMenu(MT_����ģ��).Nodes(tvwMenu(MT_����ģ��).Tag)
        blnDo = RelasCanSet(objNode.Key)
    Else
        blnDo = True
        If Index = MT_ģ�� Then
            On Error Resume Next
            Set objNode = tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag)
            If err.Number = 0 Then
                On Error GoTo 0
                Call tvwMenu_NodeClick(Index, objNode)
            Else
                On Error GoTo 0
            End If
        End If
    End If
    If Not blnDo Then objNode.Checked = Not objNode.Checked
End Sub

Private Sub tvwMenu_Collapse(Index As Integer, ByVal Node As MSComctlLib.Node)
    If Val(cmdExp.Tag) = 1 Then Exit Sub
    Call tvwMenu_NodeClick(Index, Node)
End Sub

Private Sub tvwMenu_DblClick(Index As Integer)
    Dim objNode As Node
    Dim blnDo As Boolean
    On Error Resume Next
    Set objNode = tvwMenu(Index).Nodes(tvwMenu(Index).Tag)
    err.Clear: On Error GoTo 0
    If Not objNode Is Nothing Then
        '�̶�������ϵ����ȡ����ѡ
        If Index = MT_����ģ�� Then
            blnDo = RelasCanSet(objNode.Key, True)
        Else
            blnDo = True
        End If
        If objNode.Children = 0 And blnDo Then
            objNode.Checked = Not objNode.Checked
            Call tvwMenu_NodeCheck(Index, objNode)
        End If
    End If
End Sub

Private Sub tvwMenu_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
    Node.ExpandedImage = "����_ѡ��"
End Sub

Private Sub tvwMenu_GotFocus(Index As Integer)
    mintActive = Index
End Sub

Private Sub tvwMenu_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)
    Dim blnDo As Boolean
    Dim blnUpdate As Boolean, blnTmp As Boolean
    Dim arrTmp As Variant
    Dim strKey As String
    '���ýڵ���״̬
    If Index = MT_����ģ�� Then
        blnDo = RelasCanSet(Node.Key)
    Else
        blnDo = True
    End If
    '��һ��ͬ����ѡģ�����Ȩ������ᷢ���������
    '��ȡ��ģ��̶�����ģ�����ұߵĶ�Ӧģ�飬��ֱ��ȡ����ģ����Ȩ������ȡ���Ĺ̶�����ģ���ֱ���ѡ��
    If blnDo Then
        strKey = GetUpdateKey(Node.Key)
        If strKey <> "" Then
            If strKey Like "M*" Then blnUpdate = True
            Call UpdateGrantState(strKey, Node.Checked)
        End If
    End If
    Call tvwMenu_NodeClick(Index, Node)
    '��ѡ�ڵ����������
    If blnDo Then
        '������Ȩ
        If Index = MT_ģ�� And Node.Key Like "M*" Then
            blnTmp = CheckNode(tvwMenu(MT_����ģ��))
            blnTmp = blnTmp Or CheckNode(tvwMenu(Index), Node.Key)
        Else
            blnTmp = CheckNode(tvwMenu(Index), Node.Key)
        End If
        If blnUpdate Or blnTmp Then Call SynchronizeState
    End If
End Sub

Private Sub AddModuleNode(ByVal strKey As String, ByVal strName As String, ByVal intType As Integer)
    '��Σ��ڵ�keyֵ���ڵ����ƣ��ڵ����ͣ�1��ʵ��ģ�飬0������ģ�飩
    '���ܣ����ָ���ڵ�
    Dim objNode As Node
    Dim colNodes As Collection
    Dim arrNodes() As String
    Dim strNodeKye As String
    Dim i As Long, j As Long, lngRelative As Long
    
    If intType = 0 Then
        'λ������ģ��
        Set objNode = tvwMenu(MT_ģ��).Nodes.Add("V_" & mlngSys, tvwChild, "M_000000" & "_" & strKey, "��" & Format(strKey, "000000") & "��" & strName, "Module")
        objNode.Checked = True
    Else
        'λ�ڷ�����ģ��
        mrsTree.Filter = "ģ�� = " & strKey
        mrsTree.Filter = "��� = " & mrsTree!�ϼ�
        With mrsTree
            On Error Resume Next
            Set objNode = tvwMenu(MT_ģ��).Nodes("K_" & Format(!���, "000000"))
            If err.Number <> 0 Then
                err.Clear
                '���ҽڵ���
                Set colNodes = New Collection
                Call FindModulePath(mrsTree!���, colNodes)
                '��ӽڵ���
                For i = colNodes.Count To 1 Step -1
                    arrNodes = Split(colNodes(i), "?")
                    On Error Resume Next
                    '�������ж����ڵ㲢����
                    strNodeKye = FindNodePosition(arrNodes, lngRelative)
                    Set objNode = tvwMenu(MT_ģ��).Nodes.Add(strNodeKye, lngRelative, "K_" & Format(arrNodes(0), "000000"), arrNodes(2), "����", "����_ѡ��")
                    objNode.Checked = True
                    colNodes.Remove i
                    If err.Number <> 0 Then err.Clear
                Next
                Call AddModuleNode(strKey, "", 1)
            Else
                mrsTree.Filter = "ģ�� = " & strKey
                Set objNode = tvwMenu(MT_ģ��).Nodes.Add("K_" & Format(!�ϼ�, "000000"), tvwChild, "M_" & Format(!���, "000000") & "_" & !ģ��, "��" & Format(!ģ��, "000000") & "��" & !����, "Module")
                objNode.Checked = True
            End If
        End With
    End If
End Sub

Private Function FindNodePosition(arrNodes() As String, lngRelative As Long) As String
    '��Σ��洢�ڵ���Ϣ������
    '       arrNodes():�ڵ�����
    '       FindNodePosition:Ҫ����Ľڵ�����ڽڵ�
    '       lngRelative:Ҫ����Ľڵ�����ڽڵ�����λ��
    '���Σ�����Ҫ����ڵ�����ڽڵ�����ǵ����λ��
    '���ܣ����ҽڵ��������е����λ��
    Dim objNode As Node
    Dim j As Long
    
    If arrNodes(1) = 0 Then
        Set objNode = tvwMenu(MT_ģ��).Nodes(1).FirstSibling
    Else
        Set objNode = tvwMenu(MT_ģ��).Nodes("K_" & Format(arrNodes(1), "000000")).Child
        If objNode Is Nothing Then
            FindNodePosition = "K_" & Format(arrNodes(1), "000000")
            lngRelative = tvwChild
            Exit Function
        End If
    End If
    Do While Not objNode Is Nothing
        If arrNodes(0) < Val(Split(objNode.Key, "_")(1)) And Split(objNode.Key, "_")(0) = "K" Then
            FindNodePosition = objNode.Key
            lngRelative = tvwPrevious
            Exit Function
        End If
        If objNode.Next Is Nothing Or objNode.Next.Key = "V_" & mlngSys Then
            FindNodePosition = objNode.Key
            lngRelative = tvwNext
            Exit Function
        Else
            Set objNode = objNode.Next
        End If
    Loop
End Function

Private Sub FindModulePath(ByVal strNum As String, colNodes As Collection)
    '���ܣ���ȡһ���ڵ㵽���ڵ��·��
    '��Σ��ڵ�ı�ţ��洢�ڵ�·���ļ��϶���
    Dim objNode As Node

    mrsTree.Filter = "��� = " & strNum
    On Error Resume Next
    Set objNode = tvwMenu(MT_ģ��).Nodes("K_" & Format(strNum, "000000"))
    If err.Number <> 0 Then
        err.Clear
        colNodes.Add mrsTree!��� & "?" & mrsTree!�ϼ� & "?" & mrsTree!����
        If mrsTree!�ϼ� <> 0 Then
            Call FindModulePath(mrsTree!�ϼ�, colNodes)
        End If
    End If
End Sub

Private Sub tvwMenu_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
'��Ҫ���û�ý���ڵ��״̬
    Dim arrTmp As Variant
    Dim ctTmp As ClearType
    Dim strKey As String
    
    lblNotice.Caption = ""
    '����ϴ��ǹ��ܵ����ģ�飬�����¼������ģ��
    If tvwMenu(Index).Tag <> Node.Key Or mstrCurRelas Like "*_*_?*" Then
        If tvwMenu(Index).Tag <> "" And Index = MT_ģ�� And chkOnlyShow.value = False Then
            tvwMenu(Index).Nodes(tvwMenu(Index).Tag).Bold = False
        End If
        ctTmp = -1
        If Node.Key Like "M_*" And Index = MT_ģ�� Then
            strKey = mlngSys & "_" & Val(Split(Node.Key, "_")(2)) & "_"
        End If
        If tvwMenu(Index).Tag <> Node.Key And Index = MT_ģ�� Then
            ctTmp = CT_����
        ElseIf tvwMenu(Index).Tag <> Node.Key Then
            ctTmp = CT_��������
        ElseIf Index <> MT_����ģ�� Then
            If strKey <> "" And mstrCurRelas Like strKey & "*" Then
            Else
                ctTmp = CT_����ģ��
            End If
        End If
        tvwMenu(Index).Tag = Node.Key
        Call ClearFace(ctTmp) '��ս���
        If strKey <> "" Then
            Call FillRelasModule(strKey)
        End If
        If Node.Key Like "M_*" Then
            If ctTmp = CT_���� Then
                Call FillFunc(lvwFunc(MT_ģ��))
            End If
            If ctTmp >= CT_����ģ�� Then
            ElseIf ctTmp >= CT_�������� Then
                Call FillFunc(lvwFunc(MT_����ģ��))
            End If
        End If
    End If
    If Index = MT_ģ�� Then
        lvwFunc(MT_ģ��).Enabled = Node.Checked
        If Not Node.Checked Then
            Set lvwFunc(MT_ģ��).SelectedItem = Nothing
        End If
        lvwFunc(MT_ģ��).BackColor = IIf(lvwFunc(MT_ģ��).Enabled, &H80000005, &H8000000F)
    ElseIf RelasCanSet(Node.Key) Then
         lvwFunc(MT_����ģ��).Enabled = Node.Checked
         If Not Node.Checked Then
            Set lvwFunc(MT_����ģ��).SelectedItem = Nothing
         End If
         lvwFunc(MT_����ģ��).BackColor = IIf(lvwFunc(MT_����ģ��).Enabled, &HEFF0E0, &H8000000F)
    End If
    If Index = MT_ģ�� Then Node.Bold = True
    Node.Selected = True
End Sub

Private Sub ClearFace(Optional ByVal ctInput As ClearType)
'���ܣ���ս���Ĳ�������
    'CT_Sys��Ҫ�ر��������
    If ctInput = CT_Sys Then
        tvwMenu(MT_ģ��).Nodes.Clear: tvwMenu(MT_ģ��).Tag = ""
        mblnClear = True: txtSearch.Text = "": mblnClear = False
        mblnExpanded = False: mintActive = 0
        cmdExp.Caption = IIf(mblnExpanded, "ȫ���۵�&D)", "ȫ��չ��(&D)")
    End If
    'CT_Sys��CT_������Ҫ�ر��������
    If ctInput >= CT_���� Then
        lvwFunc(MT_ģ��).ListItems.Clear: lvwFunc(MT_ģ��).Tag = ""
        mstrCurRelas = ""
    End If
    'CT_Sys��CT_����ģ�飬CT_���ܣ���Ҫ�ر��������
    If ctInput >= CT_����ģ�� Then
        tvwMenu(MT_����ģ��).Nodes.Clear: tvwMenu(MT_����ģ��).Tag = ""
    End If
    'CT_Sys��CT_����ģ�飬CT_���ܣ�CT_����������Ҫ�ر��������
    If ctInput >= CT_�������� Then
        lvwFunc(MT_����ģ��).ListItems.Clear: lvwFunc(MT_����ģ��).Tag = ""
    End If
End Sub

Private Sub txtSearch_Change()
    mlngCurPos = 0
    mstrFind = txtSearch.Text
    mblnReturn = False
    If mstrFind <> "" And Not mblnClear Then
        mlngCurPos = FindModule(mlngCurPos)
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "%" Or Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "*" Or Chr(KeyAscii) = "_" Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        If Not mblnReturn Then
            mblnReturn = True
            mlngCurPos = 0
            mstrFind = txtSearch.Text
            mlngCurPos = FindModule(mlngCurPos, True)
        End If
    End If
End Sub

Private Function CheckNode(ByRef tvwInput As TreeView, Optional ByVal strKey As String, Optional ByVal lngLevel As Long, Optional ByRef lngCount As Long) As Boolean
'���ܣ�����ģ��˵���ѡ״̬
'���أ���ѡ״̬�Ƿ����仯
    Dim blnCheck As Boolean, blnChildCheck As Boolean
    Dim objNode As Node, objParent As Node
    Dim arrTmp As Variant
    Dim strKeyTmp As String
    Dim strGrant As String
    
    On Error GoTo errH
    With tvwInput
        If tvwInput.Index = MT_ģ�� Then
            If .Nodes.Count = 0 Then Exit Function
            If strKey = "" Then '��ʼ���ù�ѡ״̬
                Set objNode = .Nodes(1)
                Do While Not objNode Is Nothing
                    blnCheck = CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount) '�ݹ�����
                    strKeyTmp = GetUpdateKey(objNode.Key)
                    If objNode.Checked <> blnCheck And strKeyTmp <> "" Then
                        lngCount = lngCount + 1
                        '������Ȩ
                        Call UpdateGrantState(strKeyTmp, blnCheck)
                    End If
                    objNode.Checked = blnCheck
                    Set objNode = objNode.Next
                Loop
            Else
                Set objParent = .Nodes(strKey)
                If lngLevel <= 0 Then '�ֹ���ѡ
                    '���¼���ʼ
                    blnCheck = objParent.Checked
                    Set objNode = objParent.Child
                    Do While Not objNode Is Nothing
                        '��ѡ״̬�븸����ͬ���������Ȩ��¼��
                        strKeyTmp = GetUpdateKey(objNode.Key)
                        If strKeyTmp <> "" And objNode.Checked <> blnCheck Then
                            lngCount = lngCount + 1
                            '������Ȩ
                            Call UpdateGrantState(strKeyTmp, blnCheck)
                        End If
                        objNode.Checked = blnCheck '���ý�㹴ѡ״̬
                        If objNode.Children <> 0 Then '������ӽڵ㣬��ݹ�
                            Call CheckNode(tvwInput, objNode.Key, -1, lngCount)
                        End If
                        Set objNode = objNode.Next
                    Loop
                    If lngLevel = 0 Then
                        '��������
                        Set objParent = .Nodes(strKey)
                        Do While Not objParent.Parent Is Nothing
                            Set objParent = objParent.Parent
                            blnChildCheck = True
                            Set objNode = objParent.Child
                            Do While Not objNode Is Nothing
                                If Not objNode.Checked Then
                                    blnChildCheck = False '�Ӽ���һ��δ��ѡ����������ѡ
                                    Exit Do
                                End If
                                Set objNode = objNode.Next
                            Loop
                            '��ѡ״̬�븸����ͬ���������Ȩ��¼��
                            strKeyTmp = GetUpdateKey(objParent.Key)
                            If strKeyTmp <> "" And objParent.Checked <> blnChildCheck Then
                                lngCount = lngCount + 1
                                '������Ȩ
                                Call UpdateGrantState(strKeyTmp, blnCheck)
                            End If
                            objParent.Checked = blnChildCheck
                            '������һ��ѭ��
                        Loop
                    End If
                Else
                    Set objParent = .Nodes(strKey)
                    If objParent.Children <> 0 Then '���Ӽ��Ĳ��ж�
                        Set objNode = objParent.Child
                        blnChildCheck = True
                        Do While Not objNode Is Nothing
                            If objNode.Children <> 0 Then
                                blnCheck = CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount)
                                strKeyTmp = GetUpdateKey(objNode.Key)
                                If strKeyTmp <> "" And objNode.Checked <> blnCheck Then
                                    lngCount = lngCount + 1
                                    Call UpdateGrantState(strKeyTmp, blnCheck, , 1)
                                End If
                                objNode.Checked = blnCheck
                            End If
                            blnChildCheck = blnChildCheck And objNode.Checked
                            '�Ӽ���һ��δ��ѡ����������ѡ
                            If Not blnChildCheck Then Exit Do
                            Set objNode = objNode.Next
                        Loop
                        CheckNode = blnChildCheck
                    Else
                        CheckNode = objParent.Checked
                    End If
                End If
            End If
        Else '���ģ�������
            If .Nodes.Count = 0 Then Exit Function
            If strKey = "" Then '��ʼ���ù�ѡ״̬����Ҫ����Ϊ�ڼ���ʱ�Ѿ�����
                arrTmp = Split(mstrCurRelas, "_")
                If arrTmp(2) <> "" Then
                    blnCheck = lvwFunc(MT_ģ��).ListItems("F_" & arrTmp(2)).Checked
                Else
                    blnCheck = tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag).Checked
                End If
                Set objNode = .Nodes(1)
                Do While Not objNode Is Nothing
                    arrTmp = Split(objNode.Key, "_")
                    mrsModsInfo.Filter = "ϵͳ=" & arrTmp(3) & " And ���=" & arrTmp(4)
                    blnChildCheck = mrsModsInfo!��Ȩ��
                    If Not blnChildCheck Then
                        blnChildCheck = blnCheck And objNode.Tag Like "1^*"
                    End If
                    If objNode.Checked <> blnChildCheck Then
                        lngCount = lngCount + 1
                        If blnChildCheck And mrsModsInfo!��Ȩ�� = 0 Then 'δ��Ȩ��������Ҫ������Ȩ
                            strGrant = mrsModsInfo!Ĭ�Ϲ���
                            mrsModsInfo.Filter = "ϵͳ=" & arrTmp(1) & " And ���=" & arrTmp(2)
                            strGrant = GetGrantByRelasInfo(mrsModsInfo!��Ȩ���� & "", strGrant, Split(objNode.Tag, "^")(1))
                        End If
                        '������Ȩ
                        Call UpdateGrantState("M_" & arrTmp(3) & "_" & arrTmp(4), blnChildCheck, strGrant)
                    End If
                    objNode.Checked = blnChildCheck
                    If objNode.Children <> 0 Then
                        Call CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount) '�ݹ�����
                    End If
                    Set objNode = objNode.Next
                Loop
            Else
                Set objParent = .Nodes(strKey)
                blnCheck = objParent.Checked
                Set objNode = objParent.Child
                Do While Not objNode Is Nothing
                    If objNode.Children <> 0 Then
                        Call CheckNode(tvwInput, objNode.Key, lngLevel + 1, lngCount)
                    End If
                    arrTmp = Split(objNode.Key, "_")
                    mrsModsInfo.Filter = "ϵͳ=" & arrTmp(3) & " And ���=" & arrTmp(4)
                    '�ж��Ӽ���ѡ
                    If blnCheck Then
                        blnChildCheck = objNode.Tag Like "1^*"
                        If Not blnChildCheck Then
                            blnChildCheck = mrsModsInfo!��Ȩ��
                        End If
                    Else
                        blnChildCheck = blnCheck
                    End If
                    If objNode.Checked <> blnChildCheck Then
                        lngCount = lngCount + 1
                        If TypeName(arrTmp) <> "String()" Then
                            arrTmp = Split(objNode.Key, "_")
                            mrsModsInfo.Filter = "ϵͳ=" & arrTmp(3) & " And ���=" & arrTmp(4)
                        End If
                        If blnChildCheck And mrsModsInfo!��Ȩ�� = 0 Then
                            strGrant = mrsModsInfo!Ĭ�Ϲ���
                            mrsModsInfo.Filter = "ϵͳ=" & arrTmp(1) & " And ���=" & arrTmp(2)
                            strGrant = GetGrantByRelasInfo(mrsModsInfo!��Ȩ���� & "", strGrant, Split(objNode.Tag, "^")(1))
                        End If
                        '������Ȩ
                        Call UpdateGrantState("M_" & arrTmp(3) & "_" & arrTmp(4), blnChildCheck, strGrant)
                    End If
                    objNode.Checked = blnChildCheck
                    Set objNode = objNode.Next
                Loop
            End If
        End If
    End With
    If lngLevel = 0 Then
        '#ADD#��ӽ���Ȩ���ͬ��������,�������μ�¼��״̬
        CheckNode = lngCount <> 0
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "CheckNode:" & err.Description, vbInformation, Me.Caption
End Function

Private Function GetUpdateKey(ByVal strKey As String) As String
'���ܣ���ȡ���ݻ�����µ�Key
    Dim arrTmp As Variant
    
    If strKey Like "M*" Then
        arrTmp = Split(strKey, "_")
        If UBound(arrTmp) = 2 Then
            GetUpdateKey = "M_" & mlngSys & "_" & arrTmp(2)
        ElseIf UBound(arrTmp) = 4 Then
            GetUpdateKey = "M_" & arrTmp(3) & "_" & arrTmp(4)
        End If
    ElseIf strKey Like "T*" Then
        GetUpdateKey = strKey
    ElseIf strKey Like "F*" Then
        GetUpdateKey = strKey
    End If
End Function

Private Function RelasCanSet(ByVal strKey As String, Optional ByVal blnDblClick As Boolean) As Boolean
'���ܣ��жϹ���ģ��ڵ��Ƿ���Թ�ѡ�Լ�ȡ����ѡ
    Dim objNode As Node
    Dim arrTmp As Variant, blnCheck As Boolean
    Dim strTMp As String
    RelasCanSet = True
    Set objNode = tvwMenu(MT_����ģ��).Nodes(strKey)
    If Not objNode.Checked And Not blnDblClick Or blnDblClick And objNode.Checked Then
        If objNode.Tag Like "1^*" Then
            If Not objNode.Parent Is Nothing Then
                RelasCanSet = Not objNode.Parent.Checked
                If Not RelasCanSet Then lblNotice.Caption = objNode.Text & "���ϼ�ģ��" & objNode.Parent.Text & "�ǹ̶�������ϵ���ϼ�ģ��δȡ����Ȩ���ò���ȡ����Ȩ��"
            Else
                arrTmp = Split(mstrCurRelas, "_")
                strTMp = tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag).Text
                If arrTmp(2) <> "" Then
                    blnCheck = lvwFunc(MT_ģ��).ListItems("F_" & arrTmp(2)).Checked
                    strTMp = strTMp & "�ġ�" & arrTmp(2) & "������"
                Else
                    blnCheck = tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag).Checked
                End If
                RelasCanSet = Not blnCheck
                If Not RelasCanSet Then lblNotice.Caption = objNode.Text & "���ϼ�ģ��" & strTMp & "�ǹ̶�������ϵ���ϼ�" & IIf(arrTmp(2) <> "", "����", "ģ��") & "δȡ����Ȩ����ģ�鲻��ȡ����Ȩ��"
            End If
        End If
    Else
        '�ϼ�δ��ѡ�����ܹ�ѡ����Ӹù�����Ҫ��������ѭ����ԭ��
        If Not objNode.Parent Is Nothing Then
            RelasCanSet = objNode.Parent.Checked
            If Not RelasCanSet Then lblNotice.Caption = objNode.Text & "���ϼ�ģ��" & objNode.Parent.Text & "δ������Ȩ���ϼ�ģ��δ��Ȩ����ģ�鲻�ܽ�����Ȩ��"
        End If
    End If
End Function


Private Sub SynchronizeState(Optional ByVal blnFinaly As Boolean, Optional ByVal lngTimes As Long)
'���ܣ�����Ȩ��¼��ͬ�������μ�¼��
'������lngChange=��Ȩ�����仯����
'          blnFinaly=�Ƿ���������Ȩ
    Dim strKey As String
    Dim arrTmp As Variant, strTMp As String
    Dim i As Long
    Dim blnHaveChange As Boolean, blnCheck As Boolean
    Dim objNode As Node
    
    On Error GoTo errH
    '�����ڼ������νṹʱͬʱ��������Ȩ�˼�¼���������Ҫ����Ȩ������Ϣ���µ�����
    With mrsModsInfo
        .Filter = "�ı�����=1"
        '���������ģ�鱻���
        '�������ģ�鱻����
        Do While Not .EOF
            If Not mrsRelasTree Is Nothing Then
                If !ģ����� = 1 Or !ģ����� = 3 Then
                    strKey = !ϵͳ & "_" & !���
                    mrsRelasTree.Filter = "MainKey='" & strKey & "' And ��Ȩ��<>" & !��Ȩ��
                    Do While Not mrsRelasTree.EOF
                        If mrsRelasTree!TreeName & "" = mstrCurRelas Then
                            blnHaveChange = True
                        End If
                        mrsRelasTree.Update Array("��Ȩ��", "��Ȩ����"), Array(!��Ȩ��, !��Ȩ����)
                        mrsRelasTree.MoveNext
                    Loop
                End If
            End If
            'ͬ�����˵�
            If !ϵͳ = mlngSys Then
                If !ģ������ = 1 Then
                    strTMp = mcllKeyModule("K_" & mlngSys & "_" & !���)
                    arrTmp = Split(strTMp, ",")
                    For i = LBound(arrTmp) To UBound(arrTmp) '���ܲ˵������ڣ���˴�������
                        On Error Resume Next
                        Set objNode = tvwMenu(MT_ģ��).Nodes(arrTmp(i))
                        If err.Number = 0 Then
                            objNode.Checked = !��Ȩ�� = 1
                            'ɾ����ȡ����Ȩ�Ľڵ�
                            If objNode.Checked = False And chkOnlyShow = 1 Then
                                Do While (Not objNode.Parent Is Nothing) And objNode.Previous Is Nothing And objNode.Next Is Nothing
                                    Set objNode = objNode.Parent
                                    tvwMenu(MT_ģ��).Nodes.Remove objNode.Child.Key
                                Loop
                                If objNode.Children = 0 Then
                                    tvwMenu(MT_ģ��).Nodes.Remove objNode.Key
                                End If
                            End If
                        Else
                            '��Ӹ�ʵ��ģ��ڵ���
                            Call AddModuleNode(!���, !����, 1)
                        End If
                    Next
                ElseIf chkVirtual.value Then
                    '�������ģ��
                    On Error Resume Next
                    Set objNode = tvwMenu(MT_ģ��).Nodes("M_000000_" & !���)
                    If err.Number = 0 Then
                        objNode.Checked = !��Ȩ�� = 1
                        'ɾ����ȡ����Ȩ�Ľڵ�
                        If objNode.Checked = False And chkOnlyShow = 1 Then
                            tvwMenu(MT_ģ��).Nodes.Remove "M_000000_" & !���
                        End If
                    Else
                        Call AddModuleNode(!���, !����, 0)
                    End If
                End If
                err.Clear: On Error GoTo errH
            End If
            '�޸����εȴ���һ���ж�
            .Update "�ı�����", 0
            .MoveNext
        Loop
    End With
    '����ģ���ͬ��
    If blnHaveChange And Not blnFinaly Then
        With mrsRelasTree
            .Filter = "TreeName='" & mstrCurRelas & "'"
            .Sort = "Level,���"
            arrTmp = Split(mstrCurRelas, "_")
            If arrTmp(2) <> "" Then
                blnCheck = lvwFunc(MT_ģ��).ListItems("F_" & arrTmp(2)).Checked
            Else
                blnCheck = tvwMenu(MT_ģ��).Nodes(tvwMenu(MT_ģ��).Tag).Checked
            End If
            Do While Not .EOF
                Set objNode = tvwMenu(MT_����ģ��).Nodes("M_" & !Key)
                If !Level = 1 Then
                   objNode.Checked = !��Ȩ�� = 1 Or !������� = 1 And blnCheck
                Else
                   objNode.Checked = (!��Ȩ�� = 1 Or !������� = 1) And objNode.Parent.Checked
                End If
                '�����ܸ��µ���Ȩ��¼��
                If !��Ȩ�� = 0 And objNode.Checked Then
                    Call UpdateGrantState("M_" & !MainKey, True, !Ĭ�Ϲ���) '��ʱ������Ȩ�仯ֻ��¼����������ͬ��
                End If
                .MoveNext
            Loop
            mrsModsInfo.Filter = "�ı�����=1"
            If Not .EOF And lngTimes < 4 Then  '�ٴ�ͬ��
                Call SynchronizeState(False, lngTimes + 1)
            End If
        End With
    End If
    Call RefreshState
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "SynchronizeState:" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub RefreshState()
    Dim strMsg As String, strTMp As String
    Dim lngModule As Long, lngBase As Long, lngReport As Long
    Dim lngFunc As Long, lngTable As Long
    Dim lngCurSys As Long, lngVirtual As Long, lngNotVirtual As Long
    
    If mblnUnRefresh Then Exit Sub
    If msftStyle <> SFT_���� Then
        mrsModsInfo.Filter = "��Ȩ��=1 And ϵͳ<>0"
        lngModule = mrsModsInfo.RecordCount
        If mlngSys <> 0 Then
            mrsModsInfo.Filter = "��Ȩ��=1 And ϵͳ=" & mlngSys
            lngCurSys = mrsModsInfo.RecordCount
            mrsModsInfo.Filter = "��Ȩ��=1 And ϵͳ=" & mlngSys & " And ģ������=1"
            lngNotVirtual = mrsModsInfo.RecordCount
            lngVirtual = lngCurSys - lngNotVirtual
        End If
        mrsModsInfo.Filter = "��Ȩ��=1 And ϵͳ=0 And ���<100"
        lngBase = mrsModsInfo.RecordCount
        mrsModsInfo.Filter = "��Ȩ��=1 And ϵͳ=0 And ���>=100"
        lngReport = mrsModsInfo.RecordCount
        If lngModule <> 0 Then
            strMsg = "����ϵͳ�Ѿ���Ȩ��" & lngModule
            If lngCurSys <> 0 Then
                strMsg = strMsg & "����ǰϵͳ����Ȩ��" & lngCurSys
                If lngNotVirtual <> 0 Then
                    strTMp = "��ʵ��ģ�飺" & lngNotVirtual & IIf(lngVirtual = 0, "��", "")
                End If
                If lngVirtual <> 0 Then
                    If strTMp = "" Then
                        strTMp = "������ģ�飺" & lngVirtual & "��"
                    Else
                        strTMp = strTMp & "������ģ�飺" & lngVirtual & "��"
                    End If
                End If
                strMsg = strMsg & strTMp
            End If
        End If
        strTMp = ""
        If lngBase <> 0 Then
            strTMp = "������������Ȩ��" & lngBase
        End If
        If lngReport <> 0 Then
            strTMp = strTMp & IIf(strTMp = "", "", "��") & "�Զ��屨������Ȩ��" & lngReport
        End If
    Else
        If cmbSystem.Text = "ȡ������" Then
            If mrsFunction Is Nothing Then Set mrsFunction = ReadData(RDT_Function)
            mrsFunction.Filter = "��Ȩ��=1"
            lngFunc = mrsFunction.RecordCount
            If lngFunc <> 0 Then
                strMsg = "ȡ����������Ȩ��" & lngFunc
            End If
        Else
            If mrsTable Is Nothing Then Set mrsTable = ReadData(RDT_Table)
            mrsTable.Filter = "��Ȩ��=1"
            lngTable = mrsTable.RecordCount
            If lngTable <> 0 Then
                strMsg = "������������Ȩ��" & lngTable
            End If
        End If
    End If
    If strMsg <> "" Or strTMp <> "" Then
        strMsg = strMsg & IIf(strMsg = "", "", IIf(strTMp = "", "", "��")) & strTMp
    End If
    stbThis.Panels(2).Text = strMsg
End Sub

Private Sub SetNodeExpand(ByVal tvwInput As TreeView, ByVal strKey As String)
'���ܣ�չ��ĳ���ڵ�
    Dim objNode As Node
    Set objNode = tvwInput.Nodes(strKey)
    objNode.Expanded = True
    Do While Not objNode.Parent Is Nothing
        objNode.Parent.Expanded = True
        Set objNode = objNode.Parent
    Loop
    tvwInput.Nodes(strKey).EnsureVisible
End Sub

Private Sub FillFunc(ByVal lvwInput As ListView)
    Dim lngϵͳ As Long, lng��� As Long
    Dim arrTmp As Variant
    Dim strSql As String, rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim strAll As String, strGant As String, strDefault As String
    Dim blnAddALL As Boolean
    Dim objNode As Node

    If tvwMenu(lvwInput.Index).Tag = "" Then Exit Sub
    Set objNode = tvwMenu(lvwInput.Index).Nodes(tvwMenu(lvwInput.Index).Tag)
    arrTmp = Split(objNode.Key, "_")
    If lvwInput.Index = MT_ģ�� Then
        lngϵͳ = mlngSys
        lng��� = Val(arrTmp(2))
    Else
        lngϵͳ = Val(arrTmp(3))
        lng��� = Val(arrTmp(4))
    End If
    If lngϵͳ = 0 Then
        '�Զ��屨��򹤾�
        strSql = "Select ����, ����, ˵�� From zlProgFuncs Where ϵͳ Is Null And ��� = " & lng��� & " And ���� <> '����'"
        Set rsTmp = gcnOracle.Execute(strSql)
    Else
        '�����Ӧ��ϵͳ
        If gblnInIDE Then
            strSql = "Select a.����, to_char(Nvl(a.����,999),'000') as ����, a.˵�� " & _
                     "         From zlProgFuncs A " & _
                     "         Where A.ϵͳ = " & lngϵͳ & _
                     " And A.��� = " & lng��� & " And A.���� <> '����'" & _
                     " Order By to_char(a.����,'000')"
            Set rsTmp = gcnOracle.Execute(strSql)
        Else
            strSql = "Select Distinct ����,to_Char(Nvl(����,999),'000') as ����,˵�� From (Select A.����, A.����, A.˵�� " & _
                     "         From zlProgFuncs A, Zlregfunc B " & _
                     "         Where Trunc(A.ϵͳ / 100) = B.ϵͳ And A.��� = B.��� And A.���� = B.���� And A.ϵͳ = " & lngϵͳ & " And A.��� = " & lng��� & " And " & _
                     "               A.���� <> '����' " & _
                     "         Union " & _
                     "         Select ����, ����, ˵�� From zlProgFuncs  a Where ���� <> '����' And ��� Between 10000 And 19999 And A.ϵͳ = " & lngϵͳ & " And A.��� = " & lng��� & _
                     "         Union " & _
                     "         Select  A.����, A.����, A.˵�� " & _
                     "         From zlProgFuncs A, zlRPTPuts B " & _
                     "         Where A.ϵͳ = B.ϵͳ And A.��� = B.����ID And A.���� = B.���� And A.ϵͳ = " & lngϵͳ & " And A.��� = " & lng��� & ")" & _
                     "  Order By to_char(����,'000')"
            Set rsTmp = gcnOracle.Execute(strSql)
        End If
    End If
    With rsTmp
        Do While Not .EOF
            strAll = strAll & "," & !����
            Set objItem = lvwInput.ListItems.Add(, "F_" & !����, !����)
            objItem.SubItems(1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(2) = IIf(IsNull(!˵��), "", !˵��)
            .MoveNext
        Loop
        If strAll <> "" Then strAll = Mid(strAll, 2)
    End With
    Call UpdateAllFunc(lngϵͳ, lng���, strAll)
    lvwInput.SortKey = 1
    mblnItem = True
    mrsModsInfo.Filter = "ϵͳ=" & lngϵͳ & " And ���=" & lng���
    strGant = IIf(mrsModsInfo!��Ȩ�� = 1, mrsModsInfo!��Ȩ����, mrsModsInfo!Ĭ�Ϲ���)
    strGant = CheckFunc(lngϵͳ, lng���, lvwInput, strGant, strAll)
    If mrsModsInfo!��Ȩ�� = 1 Then Call UpdateGrantState("M_" & lngϵͳ & "_" & lng���, True, strGant, 1)
    mblnItem = False
End Sub

Private Function CheckFunc(ByVal lngϵͳ As Long, ByVal lng��� As Long, Optional ByRef lvwFunInput As ListView, Optional ByVal strGrant As String, Optional ByVal strAll As String) As String
'���ܣ����ù���ѡ��״̬
'������lngϵͳ=ϵͳ��
'           lng���=ģ���
'           lvwFunInput=���й������õ��б�
'           strGrant=��Ȩ����
'           strALl=���й���
'           blnClick=���ܹ�ѡ����
'���أ�lvwFunInputΪ��ʱ��������Ȩ����
'˵������lvwFunInputΪ�գ��������strALl
    Dim lst As ListItem, lstTmp As ListItem
    Dim arrTmp As Variant, i As Long
    Dim strTMp As String
    
    On Error GoTo errH
    '���������б�ı���,������ԭ�����б�ѡ״̬
    lvwTmp.ListItems.Clear: lvwTmp.Tag = ""
    If Not lvwFunInput Is Nothing Then
        For Each lst In lvwFunInput.ListItems
            If strGrant <> "" Then '��ʼ���ù��ܹ�ѡ״̬
                If InStr("," & strGrant & ",", "," & lst.Text & ",") > 0 Then
                    lst.Checked = True
                End If
            End If
            Set lstTmp = lvwTmp.ListItems.Add(, lst.Key, lst.Text)
            lstTmp.Checked = lst.Checked
            If lst.Checked Then strTMp = strTMp & "," & lst.Text
        Next
        lvwTmp.Tag = Mid(strTMp, 2)
    Else
        arrTmp = Split(strAll, ",")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If arrTmp(i) <> "����" Then
                Set lstTmp = lvwTmp.ListItems.Add(, "F_" & arrTmp(i), arrTmp(i))
                If strGrant <> "" Then
                    If InStr("," & strGrant & ",", "," & lstTmp.Text & ",") > 0 Then
                        lstTmp.Checked = True
                    End If
                End If
            End If
        Next
        lvwTmp.Tag = strGrant
    End If
    Call CheckFuncRelas(lngϵͳ, lng���, lvwTmp)
    CheckFunc = lvwTmp.Tag
    '�ӱ���״̬�ظ�ԭ�б�
    If Not lvwFunInput Is Nothing Then
        For Each lstTmp In lvwTmp.ListItems
            lvwFunInput.ListItems(lstTmp.Key).Checked = lstTmp.Checked
        Next
    End If
    Exit Function
errH:
    MsgBox "CheckFunc:" & err.Description, vbInformation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Function

Private Sub CheckFuncRelas(ByVal lngϵͳ As Long, ByVal lng��� As Long, ByRef lvwInput As ListView)
'���ܣ����Ȩ�޹�ϵ
'������lngϵͳ=ϵͳ��
'           lng���=ģ���
'           lvwInput=�������õĵ��б�
'���أ����ܵ���Ȩ���
    Dim i As Integer, intUpdate As Integer
    Dim lst As ListItem
    
    On Error GoTo errHand
    mintUpdate = 1
    Do While mintUpdate >= 1
        intUpdate = 0
        For Each lst In lvwInput.ListItems
            '���ӹ�ϵ
            '���ÿ��ѡ��Ĺ����Ƿ����Ȩ�޼�Ĺ�ϵ,ֻ������ĿΪ��������,��ĿΪ�����,���������״̬��������
            mrsRelas.Filter = "ϵͳ = " & lngϵͳ & " And ��� = " & lng��� & " And ���� = '" & lst.Text & "'"
            Do Until mrsRelas.EOF
                mrsGroup.Filter = "ϵͳ = " & lngϵͳ & " And ��� = " & lng��� & " And ��� = " & mrsRelas!���
                If Not mrsGroup.EOF Then
                    mintUpdate = 1
                    Call setState(mrsGroup, mrsRelas!���� & "_" & mrsRelas!���, lst.Checked, lvwInput)
                    If mintUpdate > 0 Then
                        intUpdate = intUpdate + 1
                    End If
                End If
                mrsRelas.MoveNext
            Loop
            '�����ϵ(����ѡģ��ʱ������),
            mrsRelExcl.Filter = "ϵͳ = " & lngϵͳ & " And ��� = " & lng��� & " And ���� = '" & lst.Text & "'"
            If Not mrsRelExcl.EOF Then
                mrsGroup.Filter = "ϵͳ = " & lngϵͳ & " And ��� = " & lng��� & " And ��� = " & mrsRelExcl!���
                If Not mrsGroup.EOF Then
                    mintUpdate = 1
                    Call setState(mrsGroup, mrsRelExcl!���� & "_" & mrsRelExcl!���, lst.Checked, lvwInput)
                    If mintUpdate > 0 Then
                        intUpdate = intUpdate + 1
                    End If
                End If
            End If
        Next
        mintUpdate = intUpdate
    Loop
    Exit Sub
errHand:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbQuestion, gstrSysName
End Sub

Private Sub setState(ByVal rsTmp As ADODB.Recordset, ByVal strKey As String, ByVal blnCheck As Boolean, ByRef lvwInput As ListView)
    '����һ������Ȩ�޵�״̬��
    Dim objRelas As clsRelas, lst As ListItem, intUpdate As Integer
    Dim objGroup As clsRGroup, strPrivs As String
    Set objGroup = New clsRGroup
    
    Do Until rsTmp.EOF
        Set objRelas = New clsRelas
        objRelas.���� = rsTmp!���
        objRelas.���� = rsTmp!����
        objRelas.��ϵ = rsTmp!��ϵ
        objRelas.������ = rsTmp!����
        objRelas.�����ܹ�ϵ = rsTmp!�����ϵ
        objRelas.Checked = InStr("," & lvwInput.Tag & ",", "," & rsTmp!���� & ",") > 0
        objRelas.Key = rsTmp!���� & "_" & rsTmp!���
        Call objGroup.Add(objRelas, objRelas.Key)
        rsTmp.MoveNext
    Loop

    Call objGroup.RelasCheck(strKey, blnCheck)
    intUpdate = mintUpdate
    For Each lst In lvwInput.ListItems
        For Each objRelas In objGroup
            If objRelas.���� = lst.Text And lst.Checked <> objRelas.Checked Then
               If lst.Checked <> objRelas.Checked Then
                    lst.Checked = objRelas.Checked
                   mintUpdate = mintUpdate + 1
               End If
            End If
        Next
    Next
    
    If intUpdate = mintUpdate Then
        mintUpdate = mintUpdate - 1
        If mintUpdate < 0 Then mintUpdate = 0
    End If
    
    strPrivs = ""
    For Each lst In lvwInput.ListItems
        If lst.Checked = True Then
            strPrivs = strPrivs & "," & lst.Text
        End If
    Next
    If strPrivs <> "" Then strPrivs = Mid(strPrivs, 2)
    lvwInput.Tag = strPrivs
End Sub

Private Function setExcl(ByVal rsTmp As ADODB.Recordset, ByVal strKey As String, ByVal blnCheck As Boolean, ByRef lvwInput As ListView, ByVal strGrant As String) As String
    '����һ�黥��Ȩ�޵�״̬��
    Dim objRelas As clsRelas, lst As ListItem, intUpdate As Integer
    Dim objGroup As clsRGroup, strPrivs As String
    Set objGroup = New clsRGroup
    
    Do Until rsTmp.EOF
        Set objRelas = New clsRelas
        objRelas.���� = rsTmp!���
        objRelas.���� = rsTmp!����
        objRelas.��ϵ = rsTmp!��ϵ
        objRelas.������ = rsTmp!����
        objRelas.�����ܹ�ϵ = rsTmp!�����ϵ
        objRelas.Checked = InStr("," & strGrant & ",", "," & rsTmp!���� & ",") > 0
        objRelas.Key = rsTmp!���� & "_" & rsTmp!���
        Call objGroup.Add(objRelas, objRelas.Key)
        rsTmp.MoveNext
    Loop

    Call objGroup.RelasCheck(strKey, blnCheck)
    For Each lst In lvwInput.ListItems
        For Each objRelas In objGroup
            If objRelas.���� = lst.Text And lst.Checked <> objRelas.Checked Then
                lvwInput.ListItems(lst.Index).Checked = objRelas.Checked
            End If
        Next
    Next
    
    strPrivs = ""
    For Each lst In lvwInput.ListItems
        If lst.Checked = True Then
            strPrivs = strPrivs & "," & lst.Text
        End If
    Next
    If strPrivs <> "" Then strPrivs = Mid(strPrivs, 2)
    setExcl = strPrivs
End Function

Private Sub UpdateAllFunc(ByVal lngSys As Long, ByVal lngModule As Long, ByVal strFuncs As String)
'������Щ�Ǳ����ܣ������Ҫ�������й���
    mrsModsInfo.Filter = "ϵͳ=" & lngSys & " And ���=" & lngModule
    If Not mrsModsInfo.EOF Then
        mrsModsInfo.Update "���й���", strFuncs
    End If
End Sub

Private Sub UpdateGrantState(Optional ByVal strKey As String, Optional ByVal blnGrant As Boolean = True, Optional ByVal strFuncs As String, Optional ByVal intGrantType As Integer)
'���ܣ�����ģ����Ȩ��Ϣ
'������strKey=ϵͳ_ģ��,Ϊ��ʱ��ʾ��ģ����Ȩͳ����Ϣ���µ�ģ����Ȩ��ϸ��Ϣ��¼����
'          strFuncs=��Ȩ����
'          blnGrant=True-��Ȩ,False-ȡ����Ȩ
'          intGrantType=0-Ĭ�ϱ�����Ȩ,1-������Ȩ
    Dim arrTmp As Variant
    Dim strGrant As String
    Dim blnChange As Boolean
    Dim strObj As String, i As Long
    
    On Error GoTo errH
    '��ģ����Ȩ��������Ϣ���µ�ģ����Ȩ��ϸ��Ϣ��¼����
    If strKey = "" Then
        Call AdjustRelasTree '�̶������Զ���Ȩ����
        With mrsModsInfo
            .Filter = "��Ȩ�ı�=1"
            .Sort = "ϵͳ,���"
            Do While Not .EOF
                
                mrsModule.Filter = "ϵͳ=" & !ϵͳ & " And ���=" & !���
                If !��Ȩ�� = 0 Then 'ȡ����Ȩ
                    Do While Not mrsModule.EOF
                        mrsModule.Update "��Ȩ��", 0
                        mrsModule.MoveNext
                    Loop
                Else '������Ȩ
                    'Ȩ�޹�ϵ���
                    strGrant = CheckFunc(!ϵͳ, !���, , !��Ȩ����, !���й���)
                    Do While Not mrsModule.EOF
                        '��ģ����Ȩ�����������һ����Ȩ
                        If mrsModule!���� & "" = "����" Then
                            mrsModule.Update "��Ȩ��", 1
                        ElseIf InStr("," & strGrant & ",", "," & mrsModule!���� & ",") > 0 Then
                            mrsModule.Update "��Ȩ��", 1
                        Else
                            mrsModule.Update "��Ȩ��", 0
                        End If
                        mrsModule.MoveNext
                    Loop
                End If
                .MoveNext
            Loop
        End With
        Exit Sub
    End If
    '��ĳһ��ģ������ȡ����Ȩ����Ȩ
    arrTmp = Split(strKey, "_")
    strObj = arrTmp(2)
    If UBound(arrTmp) > 2 Then
        strObj = Mid(strKey, Len(arrTmp(1)) + 4)
    End If
    Select Case arrTmp(0)
        Case "M"
            With mrsModsInfo
                .Filter = "ϵͳ=" & arrTmp(1) & " And ���=" & strObj
                If .EOF Then Exit Sub 'û�и�ģ�飬���˳�
                If blnGrant Then
                    If !��Ȩ�� = 0 Then
                        strGrant = IIf(strFuncs = "" And intGrantType = 0, !Ĭ�Ϲ��� & "", strFuncs)
                        .Update Array("��Ȩ�ı�", "�ı�����", "��Ȩ��", "��Ȩ����"), Array(1, 1, 1, strGrant)
                    Else
                        '�Ѿ���Ȩ�ٴν���Ĭ����Ȩʱ��������Ȩ���ܣ��������赱ǰ����
                        strGrant = IIf(strFuncs = "" And intGrantType = 0, !��Ȩ���� & "", strFuncs)
                        blnChange = (strFuncs <> !��Ȩ���� & "") '������Ȩ�����Ƿ����仯
                        .Update Array("��Ȩ�ı�", "�ı�����", "��Ȩ��", "��Ȩ����"), Array(IIf(blnChange Or !��Ȩ�ı� <> 0, 1, 0), IIf(blnChange Or !�ı����� <> 0, 1, 0), 1, strGrant)
                    End If
                Else
                    If !��Ȩ�� = 1 Then 'ȡ����Ȩ
                        .Update Array("��Ȩ�ı�", "�ı�����", "��Ȩ��", "��Ȩ����"), Array(1, 1, 0, "")
                    End If
                End If
            End With
        Case "T"
            mrsTable.Filter = "����='" & strObj & "' And ϵͳ=" & arrTmp(1)
            mrsTable.Update "��Ȩ��", IIf(blnGrant, 1, 0)
        Case "F"
            mrsFunction.Filter = "������='" & strObj & "' And ϵͳ=" & arrTmp(1)
            mrsFunction.Update "��Ȩ��", IIf(blnGrant, 1, 0)
    End Select
    Call RefreshState
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "UpdateGrantState:" & err.Description, vbInformation, Me.Caption
End Sub

Private Function GetRelasTree(Optional ByVal blnLast As Boolean)
    Dim rsRelasTmp As ADODB.Recordset
    Dim rsRelasTree As ADODB.Recordset
    Dim blnDo As Boolean, blnGrant As Boolean
    Dim strPreTree As String
    Dim objNode As Node

    On Error GoTo errH
    If mrsModRelas Is Nothing Then Set mrsModRelas = ReadData(RDT_ModRelas)
    mrsModRelas.Filter = "": mrsModRelas.Sort = "ϵͳ , ģ��, ����, ���ϵͳ, ���ģ��,�����Ϣ"
    Set rsRelasTmp = CopyNewRec(mrsModRelas)
    Set rsRelasTree = CopyNewRec(mrsModRelas, True, "���ϵͳ ϵͳ,���ģ�� ģ��,�������,�����Ϣ,ȱʡֵ,����,���� �ϼ�����", Array("TreeName", adVarChar, 100, Empty, _
                                                     "TreeMain", adVarChar, 100, Empty, "TreeFunc", adVarChar, 100, Empty, "PreKey", adVarChar, 100, Empty, "MainPreKey", adVarChar, 100, Empty, _
                                                     "Key", adVarChar, 100, Empty, "MainKey", adVarChar, 100, Empty, "Level", adInteger, 1, 0, "���", adInteger, 1, 0, "��Ȩ��", adInteger, 1, 0, _
                                                     "��Ȩ����", adVarChar, 2000, Empty, "���й���", adVarChar, 2000, Empty, "Ĭ�Ϲ���", adVarChar, 2000, Empty))
    mrsModRelas.Filter = "": mrsModRelas.Sort = "ϵͳ , ģ��, ����, ���ϵͳ, ���ģ��,�����Ϣ"
    With mrsModRelas
        Do While Not .EOF
            blnDo = False
            mrsModsInfo.Filter = "ϵͳ=" & !���ϵͳ & " And ���=" & !���ģ��
            If Not mrsModsInfo.EOF Then
                mrsModsInfo.Filter = "ϵͳ=" & !ϵͳ & " And ���=" & !ģ��
                blnDo = Not mrsModsInfo.EOF And strPreTree <> !ϵͳ & "_" & !ģ�� & "_" & !����
            End If
            If blnDo Then
                strPreTree = !ϵͳ & "_" & !ģ�� & "_" & !����
                blnGrant = mrsModsInfo!��Ȩ�� = 1
                If blnGrant Then '�ж������Ƿ����
                    If !���� & "" <> "" Then blnGrant = InStr("," & mrsModsInfo!��Ȩ���� & ",", "," & !���� & ",") > 0
                End If
                Set objNode = tvwModRelas.Nodes.Add(, , strPreTree, "��(" & Format(!ģ�� & "", "000000") & ")" & !���� & "��" & !����, "����")
                objNode.Checked = blnGrant
                If FillRelasTreeRec(!ϵͳ & "_" & !ģ�� & "_" & !����, !ϵͳ & "_" & !ģ��, !���� & "", rsRelasTree, rsRelasTmp) Then
                    mrsModsInfo.Filter = "ϵͳ=" & !ϵͳ & " And ���=" & !ģ��
                    mrsModsInfo.Update "ģ�����", GetCurType(Val(mrsModsInfo!ģ����� & ""), True)
                End If
            End If
            .MoveNext
        Loop
    End With
    Set mrsRelasTree = rsRelasTree
'    '��¼��ͬ��������Ȩ��¼��ͬ��������
    Call SynchronizeState(Not blnLast)
    Exit Function
errH:
    MsgBox "GetRelasTree:" & err.Description, vbInformation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function FillRelasTreeRec(ByVal strTreeName As String, ByVal strTreeMain As String, ByVal strTreeFunc As String, ByRef rsTree As ADODB.Recordset, _
                                                        ByVal rsReals As ADODB.Recordset, Optional ByRef cllNodes As Collection, _
                                                        Optional ByVal strKey As String, Optional ByVal lngLevel As Long = 1, Optional ByRef lng��� As Long) As Boolean
'���ܣ�������ģ��
'������strKey=����ʽ�������⴦��
'                       ��ʽ1��ϵͳ_ģ��_����-���ظ�ģ�飨����Ϊ��ʱ����ģ�鹦�ܵ��������ģ��
'                       ��ʽ2��ϵͳ_ģ��_���ϵͳ_���ģ��-���ظ�ģ����Ӽ����ģ��
    Dim arrTmp As Variant
    Dim strFilter As String
    Dim strMainKey As String, strPreKey As String, strMainPreKey As String, strTMp As String
    Dim cllNodesKey As New Collection
    Dim Item As Variant
    Dim objNode As Node
    Dim intGant As Integer, strGrant As String, strPreGrant As String
    
    On Error GoTo errH
    If lngLevel = 5 Then Exit Function 'ֻ����4���ڵ�
    If lngLevel = 1 And cllNodes Is Nothing Then
        Set cllNodes = New Collection '�����жϽڵ��Ƿ����
        lng��� = 0
    End If
    arrTmp = Split(IIf(strKey = "", strTreeName, strKey), "_")
    If UBound(arrTmp) <= 2 Then '�����������ģ��
        strFilter = "ϵͳ=" & arrTmp(0) & " And ģ��=" & arrTmp(1) & IIf(arrTmp(2) <> "", " And ����='" & arrTmp(2) & "' And ����=0", " And ����=1")
        strMainPreKey = arrTmp(0) & "_" & arrTmp(1)
        strPreKey = ""
    Else '�����Ӽ����ģ��
        strFilter = "ϵͳ=" & arrTmp(2) & " And ģ��=" & arrTmp(3) & " And ����=1"
        strMainPreKey = arrTmp(2) & "_" & arrTmp(3)
        strPreKey = strKey
    End If
    '��ȡ��һ����Ȩ
    arrTmp = Split(strMainPreKey, "_")
    mrsModsInfo.Filter = "ϵͳ=" & arrTmp(0) & " And ���=" & arrTmp(1)
    strPreGrant = mrsModsInfo!��Ȩ���� & ""
    
    With rsReals
        .Filter = strFilter
        Do While Not .EOF
            mrsModsInfo.Filter = "ϵͳ=" & !���ϵͳ & " And ���=" & !���ģ��
            If Not mrsModsInfo.EOF Then '��ģ�������Ȩ
                strMainKey = !���ϵͳ & "_" & !���ģ��
                strTMp = !ϵͳ & "_" & !ģ�� & "_" & !���ϵͳ & "_" & !���ģ��
                On Error Resume Next
                cllNodesKey.Add strTMp
                cllNodes.Add "1", strTMp '�ж��Ƿ���ڸýڵ�
                If err.Number = 0 Then
                    On Error GoTo errH
                    lng��� = lng��� + 1
                    If lngLevel = 1 Then
                        Set objNode = tvwModRelas.Nodes.Add(strTreeName, 4, strTreeName & "M_" & strTMp, "��" & Format(!���ģ�� & "", "000000") & "��" & mrsModsInfo!����, IIf(!������� = 1, "Fixed", "Optional"))
                    Else
                        Set objNode = tvwModRelas.Nodes.Add(strTreeName & "M_" & strKey, 4, strTreeName & "M_" & strTMp, "��" & Format(!���ģ�� & "", "000000") & "��" & mrsModsInfo!����, IIf(!������� = 1, "Fixed", "Optional"))
                    End If
                    If lngLevel = 1 Then
                        objNode.Checked = mrsModsInfo!��Ȩ�� = 1 Or !������� = 1 And objNode.Parent.Checked
                    Else
                        objNode.Checked = (mrsModsInfo!��Ȩ�� = 1 Or !������� = 1) And objNode.Parent.Checked
                    End If
                    intGant = IIf(objNode.Checked, 1, 0)
                    If mrsModsInfo!��Ȩ�� = 0 And objNode.Checked Then
                        strGrant = GetGrantByRelasInfo(strPreGrant, mrsModsInfo!Ĭ�Ϲ��� & "", !�����Ϣ)
                    ElseIf mrsModsInfo!��Ȩ�� = 1 Then
                        strGrant = mrsModsInfo!��Ȩ���� & ""
                    End If
                    objNode.Tag = !������� & "^" & !�����Ϣ
                    rsTree.AddNew Array("TreeName", "TreeMain", "TreeFunc", "PreKey", "MainPreKey", "Key", "MainKey", "Level", "���", "ϵͳ", "ģ��", "�������", "�����Ϣ", "ȱʡֵ", "����", "�ϼ�����", "��Ȩ��", "��Ȩ����", "���й���", "Ĭ�Ϲ���"), _
                                    Array(strTreeName, strTreeMain, strTreeFunc, strPreKey, strMainPreKey, strTMp, strMainKey, lngLevel, lng���, !���ϵͳ, !���ģ��, !�������, !�����Ϣ, !ȱʡֵ, mrsModsInfo!����, !����, intGant, strGrant, mrsModsInfo!���й���, mrsModsInfo!Ĭ�Ϲ���)
                    mrsModsInfo.Update "ģ�����", GetCurType(Val(mrsModsInfo!ģ����� & ""))
                    If mrsModsInfo!��Ȩ�� = 0 And objNode.Checked Then
                        '������Ȩ
                        Call UpdateGrantState("M_" & strMainKey, True, strGrant)
                    End If
                    'λ��1
                Else
                    err.Clear: On Error GoTo errH
                End If
            End If
            .MoveNext
        Loop
    End With
    '�ݹ���ؽڵ㣬�öϴ��벻���ƶ�����λ��1������Ϊֱ�ӹ������ȼ���ԭ��
    For Each Item In cllNodesKey
        Call FillRelasTreeRec(strTreeName, strTreeMain, strTreeFunc, rsTree, rsReals, cllNodes, Item, lngLevel + 1, lng���)
    Next
    If lngLevel = 1 Then
        FillRelasTreeRec = cllNodes.Count <> 0
    End If
    Exit Function
errH:
    MsgBox "FillRelasTreeRec:" & err.Description, vbInformation, Me.Caption
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetCurType(ByVal intOldType As Integer, Optional ByVal blnHaveTree As Boolean) As Integer
'���ܣ���ȡһ��ģ���������
    Dim intReturn As Integer
    'ģ�����=0=���������ģ�鲻������
    '               1=���������ģ�鱻����
    '               2=�������ģ�鲻������
    '               3=�������ģ�鱻����
    Select Case intOldType
        Case 0
            intReturn = IIf(blnHaveTree, 2, 1)
        Case 1
            intReturn = IIf(blnHaveTree, 3, 1)
        Case 2
            intReturn = IIf(blnHaveTree, 2, 3)
        Case 3
            intReturn = 3
    End Select
    GetCurType = intReturn
End Function

Private Sub FillRelasModule(ByVal strTreeName As String)
'���ܣ�������ģ��
'������strTreeName=ϵͳ_ģ��_����-���ظ�ģ�飨����Ϊ��ʱ����ģ�鹦�ܵ��������ģ��
    Dim objNode As Node, strCaption As String
    Dim blnTreeGrant As Boolean
    Dim arrTmp As Variant
    Dim strTMp As String
    Dim blnUpdate As Boolean
    Dim blnNew As Boolean
    
    On Error GoTo errH
    If mrsModRelas Is Nothing Then Set mrsModRelas = ReadData(RDT_ModRelas)
    arrTmp = Split(strTreeName, "_")
    If UBound(arrTmp) <> 2 Then Exit Sub
    'ģ��û�����ģ�飬�����ж�
    If arrTmp(2) <> "" And tvwMenu(MT_����ģ��).Nodes.Count = 0 Then
        mstrCurRelas = strTreeName
        Exit Sub
    ElseIf arrTmp(2) <> "" Then '����սڵ���ɫ
        For Each objNode In tvwMenu(MT_����ģ��).Nodes
            objNode.BackColor = &HEFF0E0: objNode.Bold = False
        Next
    Else
        If Not mstrCurRelas Like arrTmp(0) & "_" & arrTmp(1) & "_*" Then
            blnNew = True
            Call ClearFace(CT_����ģ��)
        End If
    End If
    mstrCurRelas = strTreeName
    If mrsRelasTree Is Nothing Then
        mrsModRelas.Filter = "ϵͳ=" & arrTmp(0) & " And ģ��=" & arrTmp(1) & IIf(arrTmp(2) = "", "", " And ����='" & arrTmp(2) & "'")
        If mrsModRelas.RecordCount = 0 Then mstrCurRelas = strTreeName: Exit Sub
        Call GetRelasTree
    End If
    With mrsRelasTree
        .Filter = "TreeName='" & strTreeName & "'"
        .Sort = "Level,���"
        If blnNew Then
            If Not .EOF Then
                mrsModsInfo.Filter = "ϵͳ=" & arrTmp(0) & " And ���=" & arrTmp(1)
                If mrsModsInfo.EOF Then Exit Sub
                blnTreeGrant = mrsModsInfo!��Ȩ��
                If arrTmp(2) <> "" Then
                    blnTreeGrant = InStr("," & mrsModsInfo!��Ȩ���� & ",", "," & arrTmp(2) & ",")
                End If
            End If
            'ģ�����ģ����ع���ģ��
            Do While Not .EOF
                strCaption = "��" & Format(!ģ�� & "", "000000") & "��" & !����
                If !Level = 1 Then
                    Set objNode = tvwMenu(MT_����ģ��).Nodes.Add(, , "M_" & !Key, strCaption, IIf(!������� = 1, "Fixed", "Optional"))
                    objNode.Checked = !��Ȩ�� = 1 Or !������� = 1 And blnTreeGrant
                Else
                    Set objNode = tvwMenu(MT_����ģ��).Nodes.Add("M_" & !PreKey, tvwChild, "M_" & !Key, strCaption, IIf(!������� = 1, "Fixed", "Optional"))
                    objNode.Checked = (!��Ȩ�� = 1 Or !������� = 1) And objNode.Parent.Checked
                End If
                objNode.Tag = !������� & "^" & !�����Ϣ
                objNode.BackColor = &HEFF0E0
                If strTMp = "" Then strTMp = objNode.Key
                If objNode.Checked And tvwMenu(MT_����ģ��).Tag = "" Then tvwMenu(MT_����ģ��).Tag = objNode.Key
                '�����ܸ��µ���Ȩ��¼��
                If !��Ȩ�� = 0 And objNode.Checked Then
                    Call UpdateGrantState("M_" & !MainKey, True, !Ĭ�Ϲ���)
                    blnUpdate = True
                End If
                .MoveNext
            Loop
            
            If tvwMenu(MT_����ģ��).Tag = "" And strTMp <> "" Then
                tvwMenu(MT_����ģ��).Tag = strTMp
            End If
            '��ȡ��һ��ѡ��Ľڵ㣨��û��ѡ����ѡ�е�һ���ڵ㣩
            If tvwMenu(MT_����ģ��).Tag <> "" Then
                Set objNode = tvwMenu(MT_����ģ��).Nodes(tvwMenu(MT_����ģ��).Tag)
                Call SetNodeExpand(tvwMenu(MT_����ģ��), objNode.Key)
                tvwMenu(MT_����ģ��).Tag = "" '���Tag,�������¼��л��������ã����򲻻���ع���
                Call tvwMenu_NodeClick(MT_����ģ��, objNode)
            End If
        '    '��¼��ͬ��������Ȩ��¼��ͬ��������
            If blnUpdate Then Call SynchronizeState
        Else
            Do While Not .EOF
                tvwMenu(MT_����ģ��).Nodes("M_" & !Key).Bold = arrTmp(2) <> ""
                .MoveNext
            Loop
        End If
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "FillRelasModule:" & err.Description, vbInformation, Me.Caption
End Sub

Private Sub AdjustRelasTree()
'���ܣ��������Ȩʱʹ�ã�������ע����ģ����Ҫ��Ȩ
    Dim lngTimes As Long
    Dim blnDo As Boolean, strCurTree As String
    Dim arrTmp As Variant, blnTreeGrant As Boolean
    Dim objNode As Node
    Dim blnUpdate As Boolean
    
    On Error GoTo errH
    'ͬ����Ȩ״̬
    SynchronizeState (True)
    '�����ģ������ͬ�������οؼ�
    '�����Ȩ�����仯���߷�������ͬ��������ֹ��ͬ��
    lngTimes = 1: blnDo = True
    If mrsRelasTree Is Nothing Then Call GetRelasTree
    With mrsRelasTree
        Do While blnDo
            .Filter = ""
            .Sort = "TreeName,Level,���"
            strCurTree = ""
            Do While Not .EOF
                If strCurTree <> !TreeName & "" Then
                    strCurTree = !TreeName & ""
                    arrTmp = Split(strCurTree, "_")
                    mrsModsInfo.Filter = "ϵͳ=" & arrTmp(0) & " And ���=" & arrTmp(1)
                    blnTreeGrant = mrsModsInfo!��Ȩ�� = 1
                    If arrTmp(2) <> "" Then
                        blnTreeGrant = InStr("," & mrsModsInfo!��Ȩ���� & ",", "," & arrTmp(2) & ",")
                    End If
                    Set objNode = tvwModRelas.Nodes(strCurTree)
                    objNode.Checked = blnTreeGrant
                End If
                Set objNode = tvwModRelas.Nodes(strCurTree & "M_" & !Key)
                objNode.Checked = (!��Ȩ�� = 1 Or !������� = 1) And objNode.Parent.Checked
                If !��Ȩ�� = 0 And objNode.Checked Then
                    Call UpdateGrantState("M_" & !MainKey, True, !Ĭ�Ϲ���)
                End If
                .MoveNext
            Loop
            mrsModsInfo.Filter = "�ı�����=1"
            blnDo = Not mrsModsInfo.EOF And lngTimes < 3
            If Not mrsModsInfo.EOF Then
                'ͬ����Ȩ״̬
                SynchronizeState (True)
            End If
            lngTimes = lngTimes + 1
        Loop
    End With
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "AdjustRelasTree:" & err.Description, vbInformation, Me.Caption
End Sub

Private Function GetGrantByRelasInfo(ByVal strPreGrant As String, ByVal strGrtant As String, ByVal strRelasInfo As String, Optional ByVal blnGetDefault As Boolean) As String
'���ܣ����ݵ�ǰ��Ȩ�Լ������Ϣ��ȡӦ����Ȩ�Ĺ���
'������
'         strPreGrant=�ϼ���Ȩ
'         strGrtant=��ǰ��Ȩ����
'         strRelasInfo=��ǰ�ڵ����ϼ����������Ϣ
' ���أ�����֮�����Ȩ
    Dim arrRelasInfo As Variant, arrRelasTmp As Variant
    Dim i As Long, strReturn As String
    
    strReturn = strGrtant
    arrRelasInfo = Split(strRelasInfo, ";")
    For i = LBound(arrRelasInfo) To UBound(arrRelasInfo)
        arrRelasTmp = Split(arrRelasInfo(i), ",") '����,��ع���,�������,ȱʡֵ
        If arrRelasTmp(0) = "" And arrRelasTmp(1) = "" Then 'ģ����ز�������
        ElseIf arrRelasTmp(0) <> "" And arrRelasTmp(1) = "" Then '����ģ����ز�������
        ElseIf arrRelasTmp(0) = "" And arrRelasTmp(1) <> "" Then 'ģ�鹦�����
            '�̶�������û����Ȩ����Ҫ��Ȩ
            If arrRelasTmp(2) = 1 And InStr("," & strGrtant & ",", "," & arrRelasTmp(1) & ",") = 0 Then
                    strReturn = IIf(strReturn = "", arrRelasTmp(1), strReturn & "," & arrRelasTmp(1))
            End If
        Else '���ܹ��������Ҫ�����б�
            If InStr("," & strPreGrant & ",", "," & arrRelasTmp(0) & ",") > 0 Then
                '�̶�������û����Ȩ����Ҫ��Ȩ
                If arrRelasTmp(2) = 1 And InStr("," & strGrtant & ",", "," & arrRelasTmp(1) & ",") = 0 Then
                        strReturn = IIf(strReturn = "", arrRelasTmp(1), strReturn & "," & arrRelasTmp(1))
                End If
            End If
        End If
    Next
    GetGrantByRelasInfo = strReturn
End Function





