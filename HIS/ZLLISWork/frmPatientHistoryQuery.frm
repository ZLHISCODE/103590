VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPatientHistoryQuery 
   Caption         =   "������ʷ��¼��ѯ"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   7890
   Icon            =   "frmPatientHistoryQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleMode       =   0  'User
      ScaleWidth      =   1608.75
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2145
   End
   Begin MSComctlLib.ListView LivItem 
      Height          =   1815
      Left            =   90
      TabIndex        =   11
      Top             =   3120
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LivInfo 
      Height          =   1725
      Left            =   3780
      TabIndex        =   8
      Top             =   1890
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin C1Chart2D8.Chart2D ChartMain 
      Height          =   1695
      Left            =   5640
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1635
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   2884
      _ExtentY        =   2990
      _StockProps     =   0
      ControlProperties=   "frmPatientHistoryQuery.frx":020A
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   2730
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3195
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   45
   End
   Begin MSComctlLib.ListView LivPatient 
      Height          =   1635
      Left            =   30
      TabIndex        =   3
      Top             =   960
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   2940
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":078D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":09AD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":0BCD
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":0DED
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":100D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":122D
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":144D
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":166D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1889
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1AA9
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1CC9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   2910
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":1EE3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2103
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2323
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2543
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2763
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2983
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2BA3
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2DC3
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":2FDF
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":31FF
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientHistoryQuery.frx":341F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   7890
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "����"
      Child2          =   "CmbDepartment"
      MinWidth2       =   2505
      MinHeight2      =   300
      Width2          =   2685
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox CmbDepartment 
         Height          =   300
         Left            =   5295
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5055
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   635
      SimpleText      =   $"frmPatientHistoryQuery.frx":3639
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatientHistoryQuery.frx":3680
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8837
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
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   3645
      Left            =   3570
      TabIndex        =   6
      Top             =   900
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   6429
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ͼ��"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label LabItem 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ"
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label LabPatient 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      Height          =   180
      Left            =   30
      TabIndex        =   9
      Top             =   750
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "��������(&M)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnusplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort1 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "����(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "����(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmPatientHistoryQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MouseStartX As Single                       '�ƶ�ǰ����λ��X
Dim MouseStartY As Single                       '�ƶ�ǰ����λ��Y
Dim OutPatient As Boolean                       '=True���ﲡ��=FalseסԺ����
Dim StartDate As Date, EndDate As Date          '���ڹ��˵Ŀ�ʼ����ʱ��
Dim PatientInfo As String                       '������Ϣ(ͨ�����סԺ������)
Dim NowFocus As Integer                         '=1ѡ����LivPatient;=2ѡ����LivItem;=3ѡ����LivInfo

Private Sub CmbDepartment_Click()
    '���벿���µĲ���
    If Len(Me.CmbDepartment.Text) > 0 Then
        LoadPatientInfo Me.CmbDepartment.ItemData(Me.CmbDepartment.ListIndex)
    Else
        Me.LivPatient.ListItems.Clear
    End If
    '���벡�˵ļ�����Ŀ
    If Me.LivPatient.ListItems.Count > 0 Then
        Me.LivPatient.ListItems(1).Selected = True
        LoadItem Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
        If Me.LivItem.ListItems.Count > 0 Then
            '�������ݲ�����
            LoadInfo Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
        End If
    Else
        Me.LivItem.ListItems.Clear
        Me.LivInfo.ListItems.Clear
    End If
End Sub

Private Sub CoolBar1_Resize()
    Form_Resize
End Sub

Private Sub Form_Load()
    '��ʹ��
    Initialization
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'LabPatient
    Me.LabPatient.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    Me.LabPatient.Left = 0
    
    'LivPatient
    Me.LivPatient.Left = 0
    Me.LivPatient.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) + Me.LabPatient.Height
    Me.LivPatient.Width = Me.picSplit.Left
    Me.LivPatient.Height = Me.picSplit1.Top - Me.LabPatient.Height - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    
    'picSplit1
    Me.picSplit1.Left = 0
    Me.picSplit1.Width = Me.picSplit.Left
    
    'LabItem
    Me.LabItem.Top = Me.picSplit1.Top + Me.picSplit1.Height
    Me.LabItem.Left = 0
    
    'LivItem
    Me.LivItem.Top = Me.LabItem.Top + Me.LabItem.Height
    Me.LivItem.Left = 0
    Me.LivItem.Width = Me.LivPatient.Width
    Me.LivItem.Height = Me.ScaleHeight - Me.LabItem.Top - Me.LabItem.Height - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    'picSplit
    Me.picSplit.Top = IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
    Me.picSplit.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    'Tabstrip
    Me.TabStrip.Left = Me.picSplit.Left + Me.picSplit.Width
    Me.TabStrip.Top = Me.LabPatient.Top
    Me.TabStrip.Height = Me.ScaleHeight - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0) - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    Me.TabStrip.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
    
    'LivInfo
    Me.LivInfo.Top = Me.TabStrip.Top + 300
    Me.LivInfo.Left = Me.TabStrip.Left + 30
    Me.LivInfo.Height = Me.TabStrip.Height - 60 - 300
    Me.LivInfo.Width = Me.TabStrip.Width - 60
    
    'ChartMain
    Me.ChartMain.Top = Me.LivInfo.Top + 30
    Me.ChartMain.Left = Me.LivInfo.Left + 30
    Me.ChartMain.Width = Me.LivInfo.Width - 60
    Me.ChartMain.Height = Me.LivInfo.Height - 60
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�˳�ʱ����˽������
    SaveWinState Me, App.ProductName
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\����", "����", OutPatient
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\����", "��ʼ����", StartDate
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\����", "��������", EndDate
End Sub

Private Sub LivInfo_Click()
    NowFocus = 3
End Sub

Private Sub LivItem_Click()
    NowFocus = 2
End Sub

Private Sub LivItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '��������
    LoadInfo Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
End Sub

Private Sub LivItem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.stbThis.Panels(2).Text = "��ʾ:��ס<Ctrl>��������ѡ�ж��������Ŀ!"
End Sub

Private Sub LivPatient_Click()
    NowFocus = 1
End Sub

Private Sub LivPatient_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '������Ŀ
    LoadItem Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
    If Me.LivItem.ListItems.Count > 0 Then
        '�������ݲ�����
        LoadInfo Mid(Me.LivPatient.ListItems(Me.LivPatient.SelectedItem.Index).Key, 2)
    End If
End Sub
Private Sub mnuFileExcel_Click()
    '���Excel
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    '�˳�
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    'Ԥ��
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    subPrint 1
End Sub

Private Sub mnuFileSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuFileSetup_Click()
    '��������
    '����
    With frmPatientHistorySetup
        If OutPatient = True Then
            .OptInPatient.Value = 0
            .OptOutPatient.Value = 1
        Else
            .OptInPatient.Value = 1
            .OptOutPatient.Value = 0
        End If
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuHelp_Click()
    '��ʾ����
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpAbout_Click()
     '��ʾ����
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub mnuHelpWebHome_Click()
    '��ʾ��ҳ
    Call zlHomePage(Me.Hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '����Email
    Call zlMailTo(Me.Hwnd)
End Sub

Private Sub mnuViewFilter_Click()
    '����
    With frmPatientHistoryFilter
        .TxtPatient = PatientInfo
        .DTPBegin = StartDate
        .DTPEND = EndDate
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    '��ʾ�����ر�׼��ť
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    '��ʾ����������
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    
    For Each buttTemp In Toolbar1.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    
    Form_Resize
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        MouseStartX = x
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MoveTmp As Single
    '��ʱ���η��������
    On Error Resume Next
    If Button = 1 Then
        
        '�õ��ƶ����λ��
        MoveTmp = Me.picSplit.Left + x - MouseStartX
        
        '����������С���ʱ�˳�
        If MoveTmp <= 2000 Or Me.ScaleWidth - MoveTmp <= 2000 Then Exit Sub
        
        'picSplit
        Me.picSplit.Left = MoveTmp
        
        'picSplit1
        Me.picSplit1.Width = Me.picSplit.Left
        
        'Livpatient
        Me.LivPatient.Width = Me.picSplit.Left
        
        'LivItem
        Me.LivItem.Width = Me.picSplit.Left
        
        'TabStrip
        Me.TabStrip.Left = Me.picSplit.Left + Me.picSplit.Width
        Me.TabStrip.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
        Me.TabStrip.Refresh
        
        'LivInfo
        Me.LivInfo.Left = Me.TabStrip.Left + 30
        Me.LivInfo.Width = Me.TabStrip.Width - 60
        Me.LivInfo.Refresh
        
        'Chartmain
        Me.ChartMain.Left = Me.LivInfo.Left + 30
        Me.ChartMain.Width = Me.LivInfo.Width - 60
        Me.ChartMain.Refresh
    End If
End Sub



Private Sub picSplit1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        MouseStartY = y
    End If
End Sub

Private Sub picSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MoveTmp As Single
    '��ʱ���η��������
    On Error Resume Next
    If Button = 1 Then
        
        '�õ��ƶ����λ��
        MoveTmp = Me.picSplit1.Top + y - MouseStartY
        
        '����������С���ʱ�˳�
        If MoveTmp <= 2000 Or Me.ScaleHeight - MoveTmp <= 2000 Then Exit Sub
        
        'picSplit1
        Me.picSplit1.Top = MoveTmp
        
        'LivPatient
        Me.LivPatient.Height = Me.picSplit1.Top - Me.LabPatient.Height - IIf(Me.CoolBar1.Visible, Me.CoolBar1.Height, 0)
        
        'LabItem
        Me.LabItem.Top = Me.picSplit1.Top + Me.picSplit1.Height
        
        'LivItem
        Me.LivItem.Top = Me.LabItem.Top + Me.LabItem.Height
        Me.LivItem.Height = Me.ScaleHeight - Me.LabItem.Top - Me.LabItem.Height - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
        
    End If
End Sub

Private Sub TabStrip_Click()
    '��ʾ���ݻ�ͼ��
    If Me.TabStrip.SelectedItem.Index = 1 Then
        Me.ChartMain.Visible = False
        Me.LivInfo.Visible = True
    Else
        Me.LivInfo.Visible = False
        Me.ChartMain.Visible = True
    End If
End Sub

Sub LoadDepartmental(OutorIn As Boolean)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����               ���������סԺ����
    '����
    '    OutorIN        =True ��ʾ���� = False ��ʾסԺ
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    If OutorIn = True Then
        gstrSql = "select DISTINCT a.����id as id ,b.���� as �������� " & _
                  " from ���˹ҺŻ��� a , ���ű� b " & _
                  " Where a.����id = b.ID "
    Else
        gstrSql = "select DISTINCT a.��ǰ����id as id , b.���� as �������� " & _
                  " from ������Ϣ a , ���ű� b " & _
                  " Where a.��ǰ����id = b.ID " & _
                  " and a.��ǰ����id is not null"
    End If
    
    Me.CmbDepartment.Clear
    
    Me.MousePointer = 11
    
    OpenRecord rsTmp, gstrSql, Me.Caption
   
    Do Until rsTmp.EOF
        Me.CmbDepartment.AddItem rsTmp("��������")
        Me.CmbDepartment.ItemData(i) = rsTmp("ID")
        i = i + 1
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    Me.MousePointer = 1
End Sub
Sub Initialization()
    ''''''''''''''''''''''''''
    '����           ��ʹ��
    ''''''''''''''''''''''''''
    
    LoadColHead         'д���ͷ
    
    '�ָ�˽������
    RestoreWinState Me, App.ProductName
    
    StartDate = date
    StartDate = DateAdd("d", -DatePart("d", date) + 1, date)
    EndDate = DateAdd("d", -1, DateAdd("m", 1, StartDate))
    
    OutPatient = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\����", "����", "True")
    StartDate = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\����", "��ʼ����", StartDate)
    EndDate = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\����", "��������", EndDate)
    
    '���벿��
    LoadDepartmental OutPatient
    
    '���ñ�ע
    Me.ChartMain.Header.Text = "���˼�����ʷͼ��"
    Me.ChartMain.Header.Font.Size = 12
    Me.ChartMain.Header.Interior.ForegroundColor = vbBlue
    
    'X/Y���ע
    Me.ChartMain.ChartArea.Axes("X").Title.Text = "ʱ��"
    Me.ChartMain.ChartArea.Axes("Y").Title.Text = "���"
    
    NowFocus = 1
End Sub
Sub LoadPatientInfo(DepartmentID As Long)
    ''''''''''''''''''''''''''''''''''''
    '����              ���������µĲ���
    '����
    '    Department    ����ID
    ''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    
    Me.LivPatient.ListItems.Clear
    
    gstrSql = " Select DISTINCT d.����id, d.����, d.�Ա�, d.���� from ����걾��¼ a , ������ͨ��� b , ����ҽ����¼ c , ������Ϣ d " & _
              " Where a.ID = b.����걾id " & _
              " and a.������ = b.��¼���� " & _
              " and a.ҽ��id = c.id " & _
              " and c.����id = d.����id " & _
              " and d.��ǰ����ID = " & DepartmentID & _
              " and a.����ʱ�� Between [2] and [3]"
    
    If Len(Trim(PatientInfo)) > 0 Then
        gstrSql = gstrSql & " and (d.סԺ�� like [1] " & _
                                   " or d.����� like [1] " & _
                                   " or upper(d.����) like  upper([1]) )"
    End If
    
    Me.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, PatientInfo, CDate(StartDate), CDate(EndDate))
    
    
    Do Until rsTmp.EOF
        With Me.LivPatient
            Set ItmX = .ListItems.Add(, "A" & rsTmp("����ID"), rsTmp("����"))
            ItmX.SubItems(1) = rsTmp("�Ա�")
            ItmX.SubItems(2) = rsTmp("����")
        End With
        rsTmp.MoveNext
    Loop
    
    rsTmp.Close
    
    Me.MousePointer = 1
End Sub
Sub LoadItem(PatientID As Long)
    '''''''''''''''''''''''''''''''''''
    '����               ���������Ŀ
    '''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    
    Me.LivItem.ListItems.Clear
    
    gstrSql = " select distinct d.id , d.���� , d.������ , d.Ӣ���� " & _
              " from ����걾��¼ a , ������ͨ��� b , ����ҽ����¼ c , ����������Ŀ d " & _
              " Where a.ID = b.����걾id " & _
              " and a.������ = b.��¼���� " & _
              " and a.ҽ��id = c.id " & _
              " and b.������Ŀid = d.id " & _
              " and c.����id = " & PatientID & _
              " and a.����ʱ�� Between [2] and [3]"
              
    Me.MousePointer = 11
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, PatientID, CDate(StartDate), CDate(EndDate))
    
    Do Until rsTmp.EOF
        With Me.LivItem
            Set ItmX = .ListItems.Add(, "A" & rsTmp("ID"), zlCommFun.Nvl(rsTmp("����")))
            ItmX.SubItems(1) = zlCommFun.Nvl(rsTmp("������"))
            ItmX.SubItems(2) = zlCommFun.Nvl(rsTmp("Ӣ����"))
        End With
        rsTmp.MoveNext
    Loop
    
    rsTmp.Close
    Me.MousePointer = 1
End Sub
Sub LoadColHead()
    '''''''''''''''''''''''''''''''
    '����               ��ʹ����ͷ
    '''''''''''''''''''''''''''''''
    
    With Me.LivPatient
        .ColumnHeaders.Add , "A", "����", 1100
        .ColumnHeaders.Add , "B", "�Ա�", 700
        .ColumnHeaders.Add , "C", "����", 700
    End With
    
    With Me.LivItem
        .ColumnHeaders.Add , "A", "����", 1000
        .ColumnHeaders.Add , "B", "������", 800
        .ColumnHeaders.Add , "C", "Ӣ����", 800
        .ColumnHeaders.Add , "D", "��д", 800
    End With
    
    With Me.LivInfo
        .ColumnHeaders.Add , "A", "������Ŀ", 1100
        .ColumnHeaders.Add , "B", "����ʱ��", 2000
        .ColumnHeaders.Add , "C", "������", 1000
        .ColumnHeaders.Add , "D", "�����", 1000
        .ColumnHeaders.Add , "E", "���", 1100
    End With
End Sub
Sub LoadInfo(PatientID As Long)
    '''''''''''''''''''''''''''''''''''''''''
    '����               ���벡�˵ļ�����
    '����
    '    PatientID      ����ID
    '''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    Dim DateTmp As Date
    Dim XX As Variant, YY As Variant
    Dim i As Integer, j As Integer, N As Integer
    Dim NextID As Long
    Dim ItemID As Long
    
    Me.LivInfo.ListItems.Clear
    
    With Me.ChartMain.ChartGroups(1)
        '�����
        .Data.IsBatched = True
        .SeriesLabels.RemoveAll
        .PointLabels.RemoveAll
        .Data.NumSeries = 0
        .Data.IsBatched = False
    End With
    
    Me.ChartMain.ChartGroups(1).Data.IsBatched = True
    
    For i = 1 To Me.LivItem.ListItems.Count
        
        If Me.LivItem.ListItems(i).Selected = True Then
            
            ItemID = Mid(Me.LivItem.ListItems(i).Key, 2)
    
            gstrSql = " select a.id ,a.�걾���, a.����ʱ�� , a.������ , a.����� , b.������ " & _
                      " from ����걾��¼ a , ������ͨ��� b , ����ҽ����¼ c " & _
                      " Where a.ID = b.����걾id " & _
                      " and a.������ = b.��¼���� " & _
                      " and a.ҽ��id = c.id " & _
                      " and a.����� is not null " & _
                      " and c.����ID = [1] " & _
                      " and b.������ĿID = [2] " & _
                      " order by ����ʱ��"
              
            On Error GoTo Herr
                          
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, PatientID, ItemID)
    
            'д������
            Do Until rsTmp.EOF
                With Me.LivInfo
                    NextID = NextID + 1
                    Set ItmX = .ListItems.Add(, "A" & NextID, Me.LivItem.ListItems(i).SubItems(1))
                    ItmX.SubItems(1) = zlCommFun.Nvl(rsTmp("����ʱ��"))
                    ItmX.SubItems(2) = zlCommFun.Nvl(rsTmp("������"))
                    ItmX.SubItems(3) = zlCommFun.Nvl(rsTmp("�����"))
                    ItmX.SubItems(4) = zlCommFun.Nvl(rsTmp("������"))
                End With
                rsTmp.MoveNext
            Loop
            
            If rsTmp.RecordCount > 0 Then
                '�ƶ�����ʼλ��
                rsTmp.MoveFirst
            End If
            
            'û�м�¼ʱ�������˳�
            If rsTmp.EOF Then Exit Sub
            
            
            '����
            With Me.ChartMain.ChartGroups(1)
                
                .Data.Layout = oc2dDataGeneral  '�������÷�ʽΪÿ��Seriesӵ�и��Ե�X Points
                
                j = j + 1
                
                .Data.NumSeries = j             '�����м�����
                
                .Data.NumPoints(j) = rsTmp.RecordCount
                Me.ChartMain.ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateTimeLabels
                
                ReDim XX(rsTmp.RecordCount - 1) As Date
                ReDim YY(rsTmp.RecordCount - 1) As Double
                
                N = 0
                
                Do Until rsTmp.EOF
                    DateTmp = Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss")
                    XX(N) = DateTmp
                    YY(N) = Val(zlCommFun.Nvl(rsTmp("������")))
                    N = N + 1
                    rsTmp.MoveNext
                Loop
                
                .Data.CopyXVectorIn j, XX
                .Data.CopyYVectorIn j, YY
                
                'ͼ���Ա�
                .SeriesLabels.Add Me.LivItem.ListItems(i).SubItems(1)
                Me.ChartMain.Legend.Anchor = oc2dAnchorNorth            '�Ա�λ��
                Me.ChartMain.Legend.Orientation = oc2dOrientHorizontal  '�Ա귽��
                
                Select Case j
                    Case 1
                        .Styles(j).Symbol.Shape = oc2dShapeBox
                    Case 2
                        .Styles(j).Symbol.Shape = oc2dShapeCircle
                    Case 3
                        .Styles(j).Symbol.Shape = oc2dShapeCross
                    Case 4
                        .Styles(j).Symbol.Shape = oc2dShapeDiagonalCross
                    Case 5
                        .Styles(j).Symbol.Shape = oc2dShapeDiamond
                End Select

            End With
            rsTmp.Close
        End If
    Next
    Me.ChartMain.ChartGroups(1).Data.IsBatched = False
    Me.ChartMain.Refresh
    Exit Sub
Herr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Sub GetFilterStr(Patient As String, BegingDate As Date, OverDate As Date)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                   ��ȡ�����ִ�
    '����
    '    OutPatient         =True���ﲡ��;=FalseסԺ����
    '    Patient            ����
    '    StartDate          ��ʼ����
    '    EndDate            ��������
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PatientInfo = Patient
    StartDate = BegingDate
    EndDate = OverDate
    
    'ˢ��
    CmbDepartment_Click
End Sub

Public Sub GetFilterDate(InPatient As Boolean)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                   ��ȡ�����ִ�
    '����
    '    OutPatient         =True���ﲡ��;=FalseסԺ����
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If OutPatient <> InPatient Then
        '����סԺ������
        LoadDepartmental InPatient
    End If
    
    OutPatient = InPatient
    
    'ˢ��
    CmbDepartment_Click

End Sub

Private Sub subPrint(bytMode As Byte)
    '''''''''''''''''''''''''''''''''''''''''''
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '''''''''''''''''''''''''''''''''''''''''''
    Dim objPrint As New zlPrintLvw
    
    If gstrUserName = "" Then Call GetUserInfo
    
    Select Case NowFocus
        Case 1
            If LivPatient.SelectedItem Is Nothing Then Exit Sub
    
            If LivPatient.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivPatient
        Case 2
            If LivItem.SelectedItem Is Nothing Then Exit Sub
    
            If LivItem.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivItem
        Case 3
            If LivInfo.SelectedItem Is Nothing Then Exit Sub
    
            If LivInfo.ListItems.Count = 0 Then Exit Sub
            
            Set objPrint.Body.objData = Me.LivInfo
    End Select
    
    objPrint.Title.Text = "�ʿز�ѯ"
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    '�����°�ť
    Select Case Button.Key
        Case "Quit"
            '�˳�
            mnuFileExit_Click
        Case "Print"
            '��ӡ
            mnuFilePrint_Click
        Case "Preview"
            'Ԥ��
            mnuFilePreview_Click
        Case "Help"
            '����
            mnuHelp_Click
        Case "Filter"
            '����
            mnuViewFilter_Click
    End Select
End Sub
