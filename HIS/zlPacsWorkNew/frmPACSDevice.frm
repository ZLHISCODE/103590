VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPACSDevice 
   AutoRedraw      =   -1  'True
   Caption         =   "Ӱ���豸Ŀ¼"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   Icon            =   "frmPACSDevice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   7305
      Left            =   75
      ScaleHeight     =   7275
      ScaleWidth      =   11175
      TabIndex        =   26
      Top             =   570
      Width           =   11205
      Begin MSComctlLib.ImageList imgKind 
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
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":058A
               Key             =   "Server"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":06E4
               Key             =   "Gate"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":0C7E
               Key             =   "Printer"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":1218
               Key             =   "Ӱ���豸"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPACSDevice.frx":1532
               Key             =   "Զ��Ŀ¼"
            EndProperty
         EndProperty
      End
      Begin VB.Frame FraInfor 
         Height          =   825
         Left            =   7455
         TabIndex        =   31
         Top             =   6195
         Width           =   3705
      End
      Begin VB.Frame FraDevice 
         Height          =   6165
         Left            =   7440
         TabIndex        =   27
         Top             =   0
         Width           =   3705
         Begin VB.TextBox txtSDPassword 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   37
            Top             =   3660
            Width           =   2060
         End
         Begin VB.TextBox txtSDUser 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   35
            Top             =   3315
            Width           =   2060
         End
         Begin VB.TextBox txtShareDir 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   33
            Top             =   2970
            Width           =   2060
         End
         Begin VB.ComboBox Cbosort 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmPACSDevice.frx":3CE4
            Left            =   1470
            List            =   "frmPACSDevice.frx":3CE6
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   2060
         End
         Begin VB.TextBox TxtName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1140
            Width           =   2060
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "ȡ��(&C)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2475
            TabIndex        =   32
            Top             =   5655
            Width           =   1000
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "���Ӳ���(&T)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   225
            TabIndex        =   24
            Top             =   5655
            Width           =   1140
         End
         Begin VB.ComboBox cboType 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmPACSDevice.frx":3CE8
            Left            =   1470
            List            =   "frmPACSDevice.frx":3CFB
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   810
            Width           =   2060
         End
         Begin VB.TextBox txtDevAdress 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   15
            TabIndex        =   7
            Top             =   1470
            Width           =   2060
         End
         Begin VB.TextBox txtDevPort 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   5
            TabIndex        =   11
            Top             =   4830
            Width           =   2060
         End
         Begin VB.TextBox txtDevNO 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   3
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   150
            Width           =   2060
         End
         Begin VB.TextBox txtDevLocalAE 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   17
            Top             =   4125
            Width           =   2060
         End
         Begin VB.TextBox txtDevAE 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   19
            ToolTipText     =   "�ȷ��ȷ�"
            Top             =   4470
            Width           =   2060
         End
         Begin VB.TextBox txtFtpPath 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   100
            TabIndex        =   9
            Top             =   1935
            Width           =   2060
         End
         Begin VB.TextBox txtUser 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   13
            Top             =   2280
            Width           =   2060
         End
         Begin VB.TextBox txtPassWord 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1470
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   2625
            Width           =   2060
         End
         Begin VB.CommandButton cmdPath 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   240
            Left            =   3225
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   5205
            Width           =   300
         End
         Begin VB.CommandButton CmdDevSave 
            Caption         =   "����(&S)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   1425
            TabIndex        =   23
            Top             =   5655
            Width           =   1000
         End
         Begin VB.TextBox txtDirPath 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1470
            MaxLength       =   100
            TabIndex        =   21
            ToolTipText     =   "FtpĿ¼�ڷ������ϵı���·��"
            Top             =   5175
            Width           =   2060
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000002&
            BorderStyle     =   3  'Dot
            X1              =   200
            X2              =   3500
            Y1              =   4040
            Y2              =   4040
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000002&
            BorderStyle     =   3  'Dot
            X1              =   200
            X2              =   3500
            Y1              =   1850
            Y2              =   1850
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����Ŀ¼����"
            Height          =   180
            Left            =   300
            TabIndex        =   38
            ToolTipText     =   "���洢�豸���������ӹ���FTPĿ¼�����롣"
            Top             =   3720
            Width           =   1080
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����Ŀ¼�û���"
            Height          =   180
            Left            =   60
            TabIndex        =   36
            ToolTipText     =   "���洢�豸���������ӹ���FTPĿ¼���û�����"
            Top             =   3375
            Width           =   1260
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����Ŀ¼"
            Height          =   180
            Left            =   600
            TabIndex        =   34
            ToolTipText     =   "���洢�豸������""FTPĿ¼""��ֻ������Ŀ¼���ƣ�����ʹ�ù���Ŀ¼��ʽ����ͼ��"
            Top             =   3015
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ӱ�����(&I)"
            Height          =   180
            Left            =   300
            TabIndex        =   0
            ToolTipText     =   "�����豸��Ӧ��Ӱ�����"
            Top             =   540
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&M)"
            Height          =   180
            Left            =   675
            TabIndex        =   4
            ToolTipText     =   "�����豸�����ơ�"
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblRoom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&T)"
            Height          =   180
            Left            =   675
            TabIndex        =   2
            ToolTipText     =   "�����豸��ְ�����͡�"
            Top             =   870
            Width           =   630
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "IP��ַ(&A)"
            Height          =   180
            Left            =   495
            TabIndex        =   6
            ToolTipText     =   "�����豸������IP��ַ��"
            Top             =   1530
            Width           =   810
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�˿�(&P)"
            Height          =   180
            Left            =   675
            TabIndex        =   10
            ToolTipText     =   $"frmPACSDevice.frx":3D31
            Top             =   4860
            Width           =   630
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "�豸��  "
            Height          =   255
            Left            =   330
            TabIndex        =   29
            ToolTipText     =   "�豸��Ψһ��ʶ��ֻ����"
            Top             =   173
            Width           =   975
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����AE(&L)"
            Height          =   180
            Left            =   495
            TabIndex        =   16
            ToolTipText     =   $"frmPACSDevice.frx":3DB2
            Top             =   4185
            Width           =   810
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�豸AE(&D)"
            Height          =   180
            Left            =   495
            TabIndex        =   18
            ToolTipText     =   $"frmPACSDevice.frx":3E77
            Top             =   4530
            Width           =   810
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FtpĿ¼(&F)"
            Height          =   180
            Left            =   405
            TabIndex        =   8
            ToolTipText     =   "���洢�豸���������Ӱ���FTPĿ¼���ơ�"
            Top             =   1995
            Width           =   900
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FTP�û���(&N)"
            Height          =   180
            Left            =   225
            TabIndex        =   12
            ToolTipText     =   "���洢�豸����������FTPĿ¼���û�����"
            Top             =   2310
            Width           =   1080
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "FTP����(&W)"
            Height          =   180
            Left            =   405
            TabIndex        =   14
            ToolTipText     =   "���洢�豸����������FTPĿ¼�����롣"
            Top             =   2655
            Width           =   900
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����·��(&L)"
            Height          =   180
            Left            =   315
            TabIndex        =   20
            ToolTipText     =   "����Զ��Ŀ¼����·����"
            Top             =   5235
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   6855
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgKind"
         SmallIcons      =   "imgKind"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   7875
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPACSDevice.frx":3F2D
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12965
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   255
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPACSDevice.frx":47C1
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPACSDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '**********************************API����*****************************************

Private mstrPrivs As String

'***********************************************************************************
Private blnBeginchange As Boolean   '��ʼ��������������޸�

Private Sub InitSubWindow()
Dim Pane1 As Pane
    With dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(0, 0, 0, DockTopOf, Nothing)
    Pane1.Title = "�豸�б�"
    Pane1.Handle = PicList.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable

End Sub
Private Sub InitCbosort()
Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select ����,���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Ӱ�������")
    Cbosort.Clear
        Cbosort.AddItem ""
    Do Until rsTemp.EOF
        Cbosort.AddItem rsTemp!���� & "-" & rsTemp!����
        rsTemp.MoveNext
    Loop
End Sub
Private Sub InitlvwItem()
    lvwItem.ColumnHeaders.Clear
    With lvwItem.ColumnHeaders
        .Clear
        .Add , "_�豸��", "�豸��", 800
        .Add , "_Ӱ�����", "Ӱ�����", 900
        .Add , "_����", "����", 1500
        .Add , "_����", "����", 900
        .Add , "_IP��ַ", "IP��ַ", 1500
        .Add , "_�˿ں�", "�˿ں�", 800
        .Add , "_FtpĿ¼", "FtpĿ¼", 900
        .Add , "_FTP�û���", "FTP�û���", 1200
        .Add , "_����Ŀ¼", "����Ŀ¼", 1200
        .Add , "_����Ŀ¼�û���", "����Ŀ¼�û���", 1600
        .Add , "_����·��", "����·��", 900
        .Add , "_����AE", "����AE", 1200
        .Add , "_�豸AE", "�豸AE", 1200
        .Add , "_״̬", "״̬", 800
        .Add , "_����Ŀ¼����", "����Ŀ¼����", 0
    End With
    With lvwItem
        .SortKey = .ColumnHeaders("_�豸��").Index - 1
        .SortOrder = lvwAscending
    End With
    lvwItem.ListItems.Add , , , , 1
    lvwItem.ListItems.Clear
    Call FillData 'д����
End Sub
Private Sub FillData()
Dim strCurrKey As String, objItem As ListItem, rsTemp As New ADODB.Recordset
    If Not lvwItem.SelectedItem Is Nothing Then strCurrKey = lvwItem.SelectedItem.Key
    gstrSQL = "Select A.�豸��,B.���� Ӱ�����,A.�豸��,Decode(Nvl(A.����,1),1,'�洢�豸',2,'��������',3,'��Ƭ��ӡ',4,'Ӱ���豸',5,'Զ��Ŀ¼') As �豸����," & _
        "Nvl(A.����,1) As ����,A.IP��ַ,A.�˿ں�,A.FtpĿ¼,A.FTP�û���,A.FTP����,A.����AE,A.�豸AE,A.����Ŀ¼,A.״̬, " & _
        "A.����Ŀ¼�û���,A.����Ŀ¼����,A.����Ŀ¼" & _
        " From Ӱ���豸Ŀ¼ A,Ӱ������� B WHERE A.Ӱ�����=B.����(+) order by �豸��"

    err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����")
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !�豸��, !�豸��, Val(!����), Val(!����))
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_Ӱ�����").Index - 1) = Nvl(!Ӱ�����)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = !�豸��
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = !�豸����
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_IP��ַ").Index - 1) = Nvl(!IP��ַ)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_�˿ں�").Index - 1) = Nvl(!�˿ں�)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_FtpĿ¼").Index - 1) = Nvl(!ftpĿ¼)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_FTP�û���").Index - 1) = Nvl(!FTP�û���)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����Ŀ¼").Index - 1) = Nvl(!����Ŀ¼)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����Ŀ¼�û���").Index - 1) = Nvl(!����Ŀ¼�û���)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����·��").Index - 1) = Nvl(!����Ŀ¼)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����AE").Index - 1) = Nvl(!����AE)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_�豸AE").Index - 1) = Nvl(!�豸AE)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_״̬").Index - 1) = Decode(Nvl(!״̬), 1, "����", "��ͣ��")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����Ŀ¼����").Index - 1) = Nvl(!����Ŀ¼����)
            objItem.tag = Nvl(!FTP����)
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.Count > 0 Then
        err = 0: On Error Resume Next
        lvwItem.ListItems(strCurrKey).Selected = True
        Me.lvwItem.SelectedItem.EnsureVisible
        'Ĭ��ѡ�е�һ�м�¼
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBars()
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
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        '.SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False


'�˵�����
'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.Title = "�˵�"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)") '����
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)") '����
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��") '����
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With


'Begin----------------------�༭�˵�--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ���豸(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�����豸(&R)")
    End With
    
'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '�����˵�
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With


'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������", -1, False)
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "WEB�ϵ�����(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With


'----------------------�����------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, VK_F1, conMenu_Help_Help              '����-------------F1
        .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem   '����-------------CTRL+N
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify    '�޸�-------------CTRL+M
        .Add FCONTROL, vbKeyD, conMenu_Edit_Delete    'ɾ��-------------CTRL+D
        .Add 0, VK_F5, conMenu_View_Refresh           'ˢ��-------------F5
        .Add FCONTROL, vbKeyP, conMenu_File_Parameter '��������
        .Add 0, VK_F9, conMenu_Edit_Stop              'ͣ��-------------F9
        .Add 0, VK_F10, conMenu_Edit_Reuse            '����-------------F10
    End With


'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.Style = xtpButtonIconAndCaption '����
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.Style = xtpButtonIconAndCaption '����
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"):  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"):  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ͣ��"):  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����"): cbrControl.BeginGroup = True:  cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption '����
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.Style = xtpButtonIconAndCaption '����
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.Style = xtpButtonIconAndCaption  '����
    End With
End Sub

Private Sub cboType_Click()
    Call ControlEnabled(cboType.list(cboType.ListIndex))
End Sub
Private Sub ControlEnabled(ByVal Typekey As String)
'�ؼ�����״̬
        Select Case Typekey
        Case "�洢�豸"
            Cbosort.Enabled = False: Cbosort.ListIndex = -1
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = False
            txtDirPath.Enabled = False: txtDirPath.BackColor = &H80000000
            txtDevAE.Enabled = False: txtDevAE.BackColor = &H80000000
            txtDevLocalAE.Enabled = False: txtDevLocalAE.BackColor = &H80000000
            txtPassWord.Enabled = True: txtPassWord.BackColor = &H80000005
            txtUser.Enabled = True: txtUser.BackColor = &H80000005
            txtDevPort.Enabled = False: txtDevPort.BackColor = &H80000000
            txtFtpPath.Enabled = True: txtFtpPath.BackColor = &H80000005
            txtDevAdress.Enabled = True: txtDevAdress.BackColor = &H80000005
            txtShareDir.Enabled = True: txtShareDir.BackColor = &H80000005
            txtSDUser.Enabled = True: txtSDUser.BackColor = &H80000005
            txtSDPassword.Enabled = True: txtSDPassword.BackColor = &H80000005
        Case "��������", "��Ƭ��ӡ"
            Cbosort.Enabled = True
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = False
            txtDirPath.Enabled = False: txtDirPath.BackColor = &H80000000
            txtDevAE.Enabled = True: txtDevAE.BackColor = &H80000005
            txtDevLocalAE.Enabled = True: txtDevLocalAE.BackColor = &H80000005
            txtDevLocalAE.ToolTipText = ""
            txtPassWord.Enabled = False: txtPassWord.BackColor = &H80000000
            txtUser.Enabled = False: txtUser.BackColor = &H80000000
            txtDevPort.Enabled = True: txtDevPort.BackColor = &H80000005
            txtFtpPath.Enabled = False: txtFtpPath.BackColor = &H80000000
            txtDevAdress.Enabled = True: txtDevAdress.BackColor = &H80000005
            txtShareDir.Enabled = False: txtShareDir.BackColor = &H80000000
            txtSDUser.Enabled = False: txtSDUser.BackColor = &H80000000
            txtSDPassword.Enabled = False: txtSDPassword.BackColor = &H80000000
        Case "Ӱ���豸"
            Cbosort.Enabled = True
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = False
            txtDirPath.Enabled = False: txtDirPath.BackColor = &H80000000
            txtDevAE.Enabled = True: txtDevAE.BackColor = &H80000005
            txtDevLocalAE.Enabled = True: txtDevLocalAE.BackColor = &H80000005
            txtDevLocalAE.ToolTipText = "����Q/R��ѯ�ı��ط���AE"
            txtPassWord.Enabled = False: txtPassWord.BackColor = &H80000000
            txtUser.Enabled = False: txtUser.BackColor = &H80000000
            txtDevPort.Enabled = True: txtDevPort.BackColor = &H80000005
            txtFtpPath.Enabled = False: txtFtpPath.BackColor = &H80000000
            txtDevAdress.Enabled = True: txtDevAdress.BackColor = &H80000005
            txtShareDir.Enabled = False: txtShareDir.BackColor = &H80000000
            txtSDUser.Enabled = False: txtSDUser.BackColor = &H80000000
            txtSDPassword.Enabled = False: txtSDPassword.BackColor = &H80000000
        Case "Զ��Ŀ¼"
            Cbosort.Enabled = True: Cbosort.ListIndex = -1
            txtName.Enabled = True: txtName.BackColor = &H80000005
            cmdPath.Enabled = True
            txtDirPath.Enabled = True: txtDirPath.BackColor = &H80000005
            txtDevAE.Enabled = False: txtDevAE.BackColor = &H80000000
            txtDevLocalAE.Enabled = False: txtDevLocalAE.BackColor = &H80000000
            txtPassWord.Enabled = True: txtPassWord.BackColor = &H80000005
            txtUser.Enabled = True: txtUser.BackColor = &H80000005
            txtDevPort.Enabled = False: txtDevPort.BackColor = &H80000000
            txtFtpPath.Enabled = False: txtFtpPath.BackColor = &H80000000
            txtDevAdress.Enabled = False: txtDevAdress.BackColor = &H80000000
            txtShareDir.Enabled = False: txtShareDir.BackColor = &H80000000
            txtSDUser.Enabled = False: txtSDUser.BackColor = &H80000000
            txtSDPassword.Enabled = False: txtSDPassword.BackColor = &H80000000
    End Select
End Sub
Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
'---------------------------�ļ�----------------
        Case conMenu_File_Exit      '�˳�
            Unload Me
        Case conMenu_File_PrintSet, conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Call Menu_File_PrintMore(control.ID)
        Case conMenu_Edit_NewItem '����
            Call Menu_Edit_NewItem
        Case conMenu_Edit_Modify '�޸�
            Call Menu_Edit_Modify
        Case conMenu_Edit_Delete 'ɾ��
            Call Menu_Edit_Delete
        Case conMenu_Edit_Stop   'ͣ��
            Call Menu_Edit_Stop
        Case conMenu_Edit_Reuse  '����
            Call Menu_Edit_Reuse
        Case conMenu_File_Parameter '��������
            Call Menu_File_Parameter
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_Refresh 'ˢ��
            Call InitlvwItem
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
End Sub
Private Sub Menu_Edit_Delete()
On Error GoTo errHand
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If MsgBoxD(Me, "��Ľ���" & Me.lvwItem.SelectedItem.SubItems(2) & "����Ӱ���豸Ŀ¼��ɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "zl_Ӱ���豸Ŀ¼_Delete('" & Mid(Me.lvwItem.SelectedItem.Key, 2) & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Call Me.lvwItem.ListItems.Remove(Me.lvwItem.SelectedItem.Key)
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog

End Sub
Private Sub Menu_Edit_Stop()
'ͣ���豸
    On Error GoTo errHand
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    With Me.lvwItem.SelectedItem
        If MsgBoxD(Me, "��Ľ���" & .SubItems(2) & "�� ͣ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_Ӱ���豸Ŀ¼_Update('" & .Text & "','" & .SubItems(lvwItem.ColumnHeaders("_����").Index - 1) & _
                    "'," & Decode(.SubItems(lvwItem.ColumnHeaders("_����").Index - 1), "�洢�豸", 1, "��������", 2, "��Ƭ��ӡ", 3, "Ӱ���豸", 4, "Զ��Ŀ¼", 5, 6) & ",'" & .SubItems(lvwItem.ColumnHeaders("_IP��ַ").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_FtpĿ¼").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_�˿ں�").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_FTP�û���").Index - 1) & "','" & Trim(.tag) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����Ŀ¼").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_����Ŀ¼�û���").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����Ŀ¼����").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����AE").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_�豸AE").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����·��").Index - 1) & "', 0)"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    Call InitlvwItem
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub
Private Sub Menu_Edit_Reuse()
'�����豸
On Error GoTo errHand
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    
    With Me.lvwItem.SelectedItem
        If MsgBoxD(Me, "��Ľ���" & .SubItems(2) & "��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_Ӱ���豸Ŀ¼_Update('" & .Text & "','" & .SubItems(lvwItem.ColumnHeaders("_����").Index - 1) & _
                    "'," & Decode(.SubItems(lvwItem.ColumnHeaders("_����").Index - 1), "�洢�豸", 1, "��������", 2, "��Ƭ��ӡ", 3, "Ӱ���豸", 4, "Զ��Ŀ¼", 5, 6) & ",'" & .SubItems(lvwItem.ColumnHeaders("_IP��ַ").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_FtpĿ¼").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_�˿ں�").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_FTP�û���").Index - 1) & "','" & Trim(.tag) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����Ŀ¼").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_����Ŀ¼�û���").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����Ŀ¼����").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����AE").Index - 1) & "','" & .SubItems(lvwItem.ColumnHeaders("_�豸AE").Index - 1) & _
                    "','" & .SubItems(lvwItem.ColumnHeaders("_����·��").Index - 1) & "', 1)"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    
    Call InitlvwItem
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub
Private Sub Menu_File_Parameter()
    Call frmPacsSrvSet.ShowMe(Mid(lvwItem.SelectedItem.Key, 2), _
                lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����").Index - 1), _
                lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_IP��ַ").Index - 1), Me)
End Sub
Private Sub Menu_Edit_NewItem()
    blnBeginchange = True
    txtDevNO = ""
    cboType.ListIndex = -1
    Cbosort.ListIndex = -1
    txtName = ""
    txtDirPath = ""
    txtDevAE = ""
    txtDevLocalAE = ""
    txtPassWord = ""
    txtUser = ""
    txtDevPort = ""
    txtFtpPath = ""
    txtDevAdress = ""

    CmdDevSave.Enabled = True
    cmdCancel.Enabled = True
    cboType.Enabled = True
    Cbosort.Enabled = True
    txtName.Enabled = True
    txtDevNO = GetNewNo
    If Cbosort.ListCount <= 0 Then Call InitCbosort
    txtName.SetFocus
End Sub
Private Sub Menu_Edit_Modify()
    blnBeginchange = True
    CmdDevSave.Enabled = True
    cmdCancel.Enabled = True
    Call ControlEnabled(lvwItem.SelectedItem.SubItems(2))
End Sub
Private Sub Menu_File_PrintMore(ByVal lngType As Long)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    
Dim objPrint As New zlPrintLvw, bytType As Byte
    
    On Error Resume Next
    If lvwItem.ListItems.Count <= 0 Then Exit Sub
    
    objPrint.Title.Text = "�豸�б�"
    Set objPrint.Body.objData = lvwItem
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")

    Select Case lngType
        Case conMenu_File_PrintSet
            zlPrintSet
        Case conMenu_File_Preview
            zlPrintOrViewLvw objPrint, 2
        Case conMenu_File_Print
            bytType = zlPrintAsk(objPrint)
            If bytType <> 0 Then zlPrintOrViewLvw objPrint, bytType
        Case conMenu_File_Excel
            zlPrintOrViewLvw objPrint, 3
    End Select
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_Edit_NewItem '���� ûȨ�޺���ɾ�ġ�������һ���ܿ�ʼ������������������
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Or blnBeginchange Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_Edit_Modify '�޸�
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_Edit_Delete 'ɾ��
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_File_Parameter '��������
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����").Index - 1) <> "Ӱ���豸" Then control.Enabled = False: Exit Sub
            control.Enabled = True
        Case conMenu_Edit_Stop      'ͣ��
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_״̬").Index - 1) <> "����" Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_Edit_Reuse     '����
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Or blnBeginchange Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem Is Nothing Then control.Enabled = False: Exit Sub
            If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_״̬").Index - 1) <> "��ͣ��" Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_File_Excel '����EXCEL
            If Not CheckPopedom(mstrPrivs, "��ɾ��") Then control.Enabled = False
    End Select
End Sub

Private Sub cmdCancel_Click()
    If MsgBoxD(Me, "��ǰ�����δ����,ȷʵҪȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then txtName.SetFocus: Exit Sub
    blnBeginchange = False
    Call ClearControl
    If Not lvwItem.SelectedItem Is Nothing Then
        Call lvwItem_ItemClick(lvwItem.SelectedItem)
    End If
End Sub
Private Sub ClearControl()
    cmdCancel.Enabled = False
    CmdDevSave.Enabled = False
    cmdTest.Enabled = False
    cboType.Enabled = False
    cmdPath.Enabled = False
    txtName.Enabled = False
    Cbosort.Enabled = False
    txtDirPath.Enabled = False
    txtDevAE.Enabled = False
    txtDevLocalAE.Enabled = False
    txtPassWord.Enabled = False
    txtUser.Enabled = False
    txtDevPort.Enabled = False
    txtFtpPath.Enabled = False
    txtDevAdress.Enabled = False
    txtDevNO = ""
    cboType.ListIndex = -1
    txtName = ""
    txtDirPath = ""
    txtDevAE = ""
    txtDevLocalAE = ""
    txtPassWord = ""
    txtUser = ""
    txtDevPort = ""
    txtFtpPath = ""
    txtDevAdress = ""
End Sub
Private Sub CmdDevSave_Click()
    '���ڡ�Ӱ���豸���͡���Ƭ��ӡ�����ȼ�������Ƿ񳬹���Ȩ����
    If cboType.ListIndex = 2 Or cboType.ListIndex = 3 Then
        If funCanAddModality = False Then Exit Sub
    End If
    '����
    If Not DevSave Then Exit Sub
    '���пؼ���Ϊ������
    Call ClearControl
    blnBeginchange = False
    'ˢ������
    Call InitlvwItem
End Sub
Private Function ValidData() As Boolean
    Dim j As Integer
    
    On Error GoTo err
    If cboType.list(cboType.ListIndex) = "�洢�豸" Then
        If Trim(txtUser) = "" Or Trim(txtPassWord) = "" Then
            MsgBoxD Me, "����ָ���豸���û���������,����.", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Len(Trim(txtDevNO)) = 0 Then
        MsgBoxD Me, "�������豸�ţ�", vbInformation, gstrSysName
        txtDevNO.SetFocus
        Exit Function
    End If
    
    If Not (cboType.list(cboType.ListIndex) = "�洢�豸" Or cboType.list(cboType.ListIndex) = "Զ��Ŀ¼") Then
        If Cbosort.Text = "" Then
            MsgBoxD Me, "��ѡ��Ӱ�����", vbInformation, gstrSysName
            Cbosort.SetFocus: Exit Function
        End If
    End If
    
    If Len(Trim(txtName)) = 0 Then
        MsgBoxD Me, "�������豸����", vbInformation, gstrSysName
        txtName.SetFocus: Exit Function
    End If
    
    If Me.cboType.ListIndex <> cboType.ListCount - 1 Then
        If UBound(Split(Trim(txtDevAdress), ".")) <> 3 Then
            MsgBoxD Me, "IP��ʽ����ȷ�����飡", vbInformation, gstrSysName
            txtDevAdress.SetFocus: Exit Function
        Else
            For j = 0 To 3
                If Not IsNumeric(Split(Trim(txtDevAdress), ".")(j)) Then
                    MsgBoxD Me, "IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                Else
                    If Split(Trim(txtDevAdress), ".")(j) < 0 Or Split(Trim(txtDevAdress), ".")(j) >= 256 Then
                        MsgBoxD Me, "IP��ʽ����ȷ�����飡", vbInformation, gstrSysName: Exit Function
                    End If
                End If
            Next
        End If
    Else
        If Trim(txtDirPath) = "" Then
            MsgBoxD Me, "�����뱾��·����", vbInformation, gstrSysName: txtDirPath.SetFocus: Exit Function
        End If
    End If
        
    If InStr(Trim(txtFtpPath.Text), ":") > 0 Then
        MsgBoxD Me, "FTPĿ¼��ʽ����ȷ�����飡", vbInformation, gstrSysName
        txtFtpPath.SetFocus: Exit Function
    End If
    
    If cboType.ListIndex = 1 And (Len(Trim(txtDevPort)) = 0 Or Not IsNumeric(txtDevPort)) Then
        MsgBoxD Me, "��������ȷ�Ķ˿ںţ�", vbInformation, gstrSysName
        txtDevPort.SetFocus: Exit Function
    End If
    If LenB(StrConv(Trim(txtName), vbFromUnicode)) > txtName.MaxLength Then
        MsgBoxD Me, "�豸�����������" & txtName.MaxLength & "���ַ���" & CInt(txtName.MaxLength / 2) & "�����֣���", vbInformation, gstrSysName
        txtName.SetFocus: Exit Function
    End If
    ValidData = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function DevSave() As Boolean
Dim DevType As String
Dim objFtp As New clsFtp
Dim strEncryptionPassW As String

    On Error GoTo DBError
    If ValidData = False Then Exit Function
    If zlStr.NeedName(Cbosort.list(Cbosort.ListIndex)) <> "" Then
        DevType = Split(Cbosort.list(Cbosort.ListIndex), "-")(0)
    End If
    
    '����ftp����
    If Trim(txtPassWord.Text) <> "" Then
        strEncryptionPassW = objFtp.GetEncryptionPassW(Trim(txtPassWord.Text))
        strEncryptionPassW = Mid(strEncryptionPassW, 1, 1) & "��" & Mid(strEncryptionPassW, 2)
        strEncryptionPassW = "��" & strEncryptionPassW & "��"
        strEncryptionPassW = Replace(strEncryptionPassW, "'", "''")
    End If
    
    gstrSQL = "zl_Ӱ���豸Ŀ¼_Update('" & txtDevNO & "','" & Trim(txtName) & "'," & cboType.ListIndex + 1 & _
        ",'" & Trim(txtDevAdress) & "','" & Trim(txtFtpPath) & "','" & Trim(txtDevPort) & "','" & Trim(txtUser) & "','" & _
        strEncryptionPassW & "','" & Trim(txtShareDir) & "','" & Trim(txtSDUser) & "','" & Trim(txtSDPassword) & "','" & Trim(txtDevLocalAE) & "','" & Trim(txtDevAE) & "','" & Trim(txtDirPath) & "', 1,'" & DevType & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    DevSave = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdPath_Click()
    Dim strTmp As String
    '�õ�·��
    strTmp = BrowPath(Me.hWnd, "��ѡ��������ļ�Ŀ¼��")
    '�����µ�·��ʱ�ű���
    If strTmp <> "" And strTmp <> txtDirPath.Text Then
        txtDirPath.Text = strTmp
    End If
End Sub

Private Sub cmdTest_Click()
    If ValidData = False Then Exit Sub
    Me.MousePointer = vbHourglass: cmdTest.Enabled = False
    Select Case cboType.Text
        Case "�洢�豸"
            If Len(Dir(Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsFtpTools.exe")) > 0 Then
                Shell Mid(App.Path, 1, InStrRev(App.Path, "\")) & "zlPacsFtpTools.exe   " & txtUser.Text & "||" & txtPassWord & "||" & txtDevAdress.Text & "||" & txtFtpPath.Text, 1
            Else
                Call TestFTPDev
            End If
        Case "��������", "��Ƭ��ӡ", "Ӱ���豸"
            Call TestDev
        Case "Զ��Ŀ¼"
            Call TestPath
    End Select
    Me.MousePointer = vbDefault: cmdTest.Enabled = True
End Sub

Private Sub Form_Load()
    blnBeginchange = False
    mstrPrivs = gstrPrivs
    Me.Icon = imgKind.ListImages(4).Picture
    Call InitCommandBars '��ʼ���˵�
    Call InitSubWindow  '��ʼ���Ӵ���
    Call InitlvwItem '��ʼ�����
    
    Call RestoreWinState(Me, App.ProductName)
    
    gintDICOM�豸���� = getLicenseCount(LOGIN_TYPE_DICOM�豸)
    gint��Ƭ��ӡ������ = getLicenseCount(LOGIN_TYPE_��Ƭ��ӡ��)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    Call SaveWinState(Me, App.ProductName)
    Unload Me
End Sub
Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '���ܣ����ð�������
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub
Private Sub Menu_Help_Web_Home_click()
    zlHomePage hWnd
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub
Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub




Private Sub lvwItem_DblClick()
Dim cbrControl As CommandBarControl
    If blnBeginchange Then txtName.SetFocus: Exit Sub
    Set cbrControl = cbrMain.FindControl(xtpControlButton, conMenu_Edit_Modify)
    If Not cbrControl Is Nothing Then Call cbrMain_Execute(cbrControl)
End Sub
Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim objFtp As New clsFtp
    Dim strDecryptionPassW As String
    Dim i As Integer
'��ʾ����
    If blnBeginchange Then '��ʼ�޸Ļ�������
        If MsgBoxD(Me, "��ǰ�����δ���棬ȷʵҪ�����鿴��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtName.SetFocus
            Exit Sub
        End If
    End If
    blnBeginchange = False
    cmdCancel.Enabled = False
    CmdDevSave.Enabled = False
    cmdTest.Enabled = True
    If Cbosort.ListCount <= 0 Then Call InitCbosort
    txtDevNO = lvwItem.SelectedItem.Text
    txtName = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����").Index - 1)
    
    cboType.ListIndex = Decode(lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����").Index - 1), "�洢�豸", 0, "��������", 1, "��Ƭ��ӡ", 2, "Ӱ���豸", 3, "Զ��Ŀ¼", 4)
    
    If lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_Ӱ�����").Index - 1) = "" Then '�洢��Զ��Ŀ¼��ָ��Ӱ�����
        Cbosort.ListIndex = -1
    Else
        Call SeekIndex(Cbosort, lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_Ӱ�����").Index - 1))
    End If
    txtDirPath = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����·��").Index - 1)
    txtDevAE = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_�豸AE").Index - 1)
    txtDevLocalAE = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����AE").Index - 1)
    
    '����ftp����
    If Len(lvwItem.SelectedItem.tag) >= 3 Then
        If Mid(lvwItem.SelectedItem.tag, 1, 1) & Mid(lvwItem.SelectedItem.tag, 3, 1) & Mid(lvwItem.SelectedItem.tag, Len(lvwItem.SelectedItem.tag), 1) = "�����" Then
            strDecryptionPassW = Mid(lvwItem.SelectedItem.tag, 2)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, Len(strDecryptionPassW) - 1)
            strDecryptionPassW = Mid(strDecryptionPassW, 1, 1) & Mid(strDecryptionPassW, 3)
            strDecryptionPassW = objFtp.GetDecryptionPassW(strDecryptionPassW)
            
            txtPassWord = strDecryptionPassW
        Else
            txtPassWord = lvwItem.SelectedItem.tag
        End If
    Else
        txtPassWord = lvwItem.SelectedItem.tag
    End If
    
    txtUser = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_FTP�û���").Index - 1)
    txtDevPort = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_�˿ں�").Index - 1)
    txtFtpPath = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_FtpĿ¼").Index - 1)
    txtDevAdress = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_IP��ַ").Index - 1)
    txtShareDir = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����Ŀ¼").Index - 1)
    txtSDUser = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����Ŀ¼�û���").Index - 1)
    txtSDPassword = lvwItem.SelectedItem.SubItems(lvwItem.ColumnHeaders("_����Ŀ¼����").Index - 1)
'�൱���޸�
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    With lvwItem
        .Top = 0
        .Left = 0
        .Width = PicList.Width - FraDevice.Width
        .Height = PicList.Height - 370
    End With
    With FraDevice
        .Top = 0
        .Left = lvwItem.Width
    End With
    With FraInfor
        .Top = FraDevice.Height
        .Left = FraDevice.Left
        .Height = PicList.Height - FraDevice.Height - 390
    End With
End Sub
Private Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    On Error GoTo OpenFileError
    With udtBI
        '�����������
        .lngHwnd = lWindowHwnd
        '����ѡ�е�Ŀ¼
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "��ѡ����ʼ�������ļ��У�"
        Else
            .lpszTitle = sTitle
        End If
    End With
    '�����������
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '��ȡ·��
        SHGetPathFromIDList lpIDList, sPath
        '�ͷ��ڴ�
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
    Exit Function
OpenFileError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetNewNo() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo DBError
    strSql = "Select Nvl(Max(To_Char(�豸��,'000')),1) From Ӱ���豸Ŀ¼"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsTmp.EOF Then
        GetNewNo = "001"
    Else
        GetNewNo = Format(Val(rsTmp(0)) + 1, "000")
    End If
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub TestDev()
Dim objGlobal As New DicomGlobal
    On Error GoTo TestError
    If Trim(txtDevAdress.Text) = "" Then
        MsgBoxD Me, "������IP��ַ��", vbInformation, gstrSysName
        txtDevAdress.SetFocus: Exit Sub
    End If
    If Trim(txtDevPort.Text) = "" Or Not IsNumeric(txtDevPort.Text) Then
        MsgBoxD Me, "��������ȷ�Ķ˿ںţ�", vbInformation, gstrSysName
        txtDevPort.SetFocus: Exit Sub
    End If
    
    If Trim(txtDevAE.Text) = "" Then
        MsgBoxD Me, "��������ȷ���豸AE��", vbInformation, gstrSysName
        txtDevAE.SetFocus: Exit Sub
    End If
    
    If Trim(txtDevLocalAE.Text) = "" Then
        MsgBoxD Me, "��������ȷ���豸����AE��", vbInformation, gstrSysName
        txtDevLocalAE.SetFocus: Exit Sub
    End If
    
    With objGlobal
        If .Echo(txtDevAdress, CLng(txtDevPort), txtDevLocalAE, txtDevAE) <> 0 Then
            MsgBoxD Me, "�޷����ӵ�ָ���Ľ���������", vbInformation, gstrSysName
            txtDevAdress.SetFocus
        Else
            MsgBoxD Me, "���Ӳ��Գɹ���", vbInformation, gstrSysName
        End If
    End With
    Exit Sub
TestError:
    Me.MousePointer = vbDefault
    MsgBoxD Me, "�޷����ӵ�ָ�����豸��", vbInformation, gstrSysName
End Sub
Private Sub TestFTPDev()
Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP��
    strPath = Format(zlDatabase.Currentdate, "yyyymmddHHMMSS")
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Print #1, "�����ļ�"
    Close #1
    If FtpNet.FuncFtpConnect(txtDevAdress, txtUser, txtPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir("/", "FTP����" & strPath) > 0 Then
            MsgBoxD Me, "��ǰ�豸дĿ¼����ʧ��", vbInformation, gstrSysName
        Else
            FtpNet.FuncFtpDelDir "/", "FTP����" & strPath
            If CheckFtpDir(FtpNet, txtFtpPath) Then
                If FtpNet.FuncFtpMkDir(txtFtpPath, "FTP����" & strPath) > 0 Then
                    MsgBoxD Me, "��ǰ�豸����Ŀ¼ʧ��", vbInformation, gstrSysName
                ElseIf FtpNet.FuncUploadFile(txtFtpPath, strTmpPath, "temp.txt") > 0 Then
                    MsgBoxD Me, "��ǰ�豸�ϴ��ļ�ʧ��", vbInformation, gstrSysName
                ElseIf FtpNet.FuncFtpGetFileSize(txtFtpPath, "temp.txt") <= 0 Then
                    MsgBoxD Me, "��ǰ�豸��ȡ�ļ���Сʧ�ܣ�" & IIf(GetSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", 1) <> 0, "����ȡ��ע������������FTP�ļ���С�Աȡ���", ""), vbInformation, gstrSysName
                ElseIf FtpNet.FuncDelFile(txtFtpPath, "temp.txt") > 0 Then
                    MsgBoxD Me, "��ǰ�豸ɾ���ļ�ʧ��", vbInformation, gstrSysName
                Else
                    FtpNet.FuncFtpDisConnect '�ȶϿ�����ɾ������Ȼɾ����
                    If FtpNet.FuncFtpConnect(txtDevAdress, txtUser, txtPassWord) <= 0 Then
                        MsgBoxD Me, "��ǰ�豸�������ӣ�", vbInformation, gstrSysName
                    ElseIf FtpNet.FuncFtpDelDir(txtFtpPath, "FTP����" & strPath) > 0 Then
                        MsgBoxD Me, "��ǰ�豸ɾ��Ŀ¼����ʧ��", vbInformation, gstrSysName
                    Else
                        MsgBoxD Me, "�������ӳɹ���", vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
    Else
        MsgBoxD Me, "��ǰ�豸�������ӣ�", vbInformation, gstrSysName
    End If
    FtpNet.FuncFtpDisConnect
    Kill strTmpPath
End Sub

Private Function CheckFtpDir(objFtp As clsFtp, strFtpDir As String) As Boolean
    
    CheckFtpDir = True
    If Len(Trim(Replace(Replace(strFtpDir, "/", ""), "\", ""))) <> 0 Then
        objFtp.FuncChangeDir ""
        If objFtp.FuncChangeDir(strFtpDir) <> 0 Then
            If MsgBox("��ǰFTPĿ¼�����ڣ��Ƿ񴴽�����ԣ�", vbYesNo, gstrSysName) = vbYes Then
                If objFtp.FuncFtpMkDir("", strFtpDir) <> 0 Then
                    CheckFtpDir = False
                    MsgBoxD Me, "��ǰFTPĿ¼���Ϸ���", vbInformation, gstrSysName
                    txtFtpPath.SetFocus
                End If
            Else
                CheckFtpDir = False
                MsgBoxD Me, "��ǰFTPĿ¼�����ڣ�", vbInformation, gstrSysName
                txtFtpPath.SetFocus
            End If
        End If
    End If
End Function

Private Sub TestPath()
Dim duTime As Double
    On Error GoTo TestError
    If Trim(txtDirPath.Text) = "" Then
        MsgBoxD Me, "��ѡ�������Ҫ���ʵ�Զ������", vbInformation, gstrSysName
        txtDirPath.SetFocus
        Exit Sub
    End If
    
    duTime = Timer
    Do Until CLng(Timer - duTime) >= 20
        Shell "net use " & txtDirPath & " " & txtPassWord & " /user:" & txtUser, vbHide
        If WriteTest(False) = True Then
            MsgBoxD Me, "���Ӳ��Գɹ���", vbInformation, gstrSysName
        Else
            MsgBoxD Me, "�޷����ӵ�ָ���Ľ���������", vbInformation, gstrSysName
        End If
        Exit Do
        DoEvents
    Loop
    Shell "net use " & txtDirPath & " /delete "
    Exit Sub
TestError:
    Me.MousePointer = vbDefault
    MsgBoxD Me, "�޷����ӵ�ָ�����豸��", vbInformation, gstrSysName
End Sub
Private Function WriteTest(ShowErrMsg As Boolean) As Boolean
    Dim strTmpPath As String
    On Error GoTo CopyError
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Close #1
    FileCopy strTmpPath, IIf(Len(txtDirPath) > 3, txtDirPath & "\", txtDirPath) & "temp.txt"
    Kill IIf(Len(txtDirPath) > 3, txtDirPath & "\", txtDirPath) & "temp.txt"
    Kill strTmpPath
    WriteTest = True
    Exit Function
CopyError:
    If ShowErrMsg = False Then Exit Function
    If err.Number = 75 Then
        MsgBoxD Me, "д�����ʧ��!��鿴[" & txtDirPath & "]�Ƿ���д��Ȩ��!", vbInformation, App.EXEName
    Else
        MsgBoxD Me, "������������", vbQuestion, App.EXEName
    End If
End Function

Private Function funCanAddModality() As Boolean
'���DICOM�豸�ͽ�Ƭ��ӡ�����������ж��Ƿ��������
'������
'����ֵ��   True--������ӣ�False--���������
    Dim i As Integer
    Dim str���� As String   '"Ӱ���豸"���ߡ���Ƭ��ӡ��
    Dim intSum As Integer
    
    On Error GoTo err
    str���� = cboType.list(cboType.ListIndex)
    intSum = 0
    For i = 1 To lvwItem.ListItems.Count
        If lvwItem.ListItems(i).SubItems(lvwItem.ColumnHeaders("_����").Index - 1) = str���� And _
            lvwItem.ListItems(i).SubItems(lvwItem.ColumnHeaders("_״̬").Index - 1) = "����" Then
            intSum = intSum + 1
        End If
    Next i
    '�������豸
    If txtDevNO = GetNewNo Then
        intSum = intSum + 1
    End If
    
    If str���� = "Ӱ���豸" Then
        If intSum <= gintDICOM�豸���� Or gintDICOM�豸���� = -1 Then
            funCanAddModality = True
            Exit Function
        End If
    ElseIf str���� = "��Ƭ��ӡ" Then
        If intSum <= gint��Ƭ��ӡ������ Or gint��Ƭ��ӡ������ = -1 Then
            funCanAddModality = True
            Exit Function
        End If
    End If
    funCanAddModality = False
    MsgBoxD Me, str���� & "�������������������" & _
        IIf(str���� = "Ӱ���豸", gintDICOM�豸����, gint��Ƭ��ӡ������) & _
        "�����޷���ӡ����������Ӧ����ϵ��", vbOKOnly, gstrSysName
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

