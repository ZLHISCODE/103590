VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmVItemLists 
   BackColor       =   &H8000000C&
   Caption         =   "�������������"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9735
   Icon            =   "frmVItemLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2790
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5310
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   2520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Width           =   30
   End
   Begin VB.PictureBox picClass 
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "������Ŀ(&1)"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   5580
         Left            =   105
         TabIndex        =   5
         Tag             =   "1000"
         Top             =   405
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   9843
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1935
      Top             =   6825
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
            Picture         =   "frmVItemLists.frx":058A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":0B24
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":10BE
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   3210
      Left            =   2775
      TabIndex        =   1
      Top             =   1260
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6900
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmVItemLists.frx":1658
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12091
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
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9735
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ����ǰ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Class"
               Description     =   "����"
               Object.ToolTipText     =   "����ҩƷ����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "�����µ���Ŀ"
               Object.Tag             =   "����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ĵ�ǰ��Ŀ"
               Object.Tag             =   "�޸�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰ��Ŀ"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "���������Ŀ"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":1EEA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2104
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":231E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2538
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2B86
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2DA0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":2FC0
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":31E0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3400
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3620
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":383A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":3E88
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":40A2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemLists.frx":42C2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   8490
      Top             =   6135
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdInfo 
      Height          =   1695
      Left            =   2835
      TabIndex        =   8
      Top             =   5325
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "�����ߴ�"
      Height          =   180
      Left            =   8115
      TabIndex        =   9
      Top             =   5775
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
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
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "����(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "����(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "�޸�(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "ɾ��(&E)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "��Ŀ(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmVItemLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Public mstrPrivs As String       '�û����б�����ľ���Ȩ��

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String

Private Const conTabִ�п��� As Integer = 0
Private Const conTab�շѶ��� As Integer = 1
Private Const conTab����ָ�� As Integer = 2
Private Const conTab��鲿λ As Integer = 3
Private Const conTab�÷����� As Integer = 4
Private Const conTab������� As Integer = 5
Private Const conTab�䷽��� As Integer = 6
Private Const conTab���׷��� As Integer = 7
Private Const conTabӦ�òο� As Integer = 8

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    'װ���ݲ���������
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Val(tvwClass.Tag) <> Index Then
        Me.tvwClass.Tag = Index
        Call zlRefClasses
    End If
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
End Sub

Private Sub Form_Load()
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    On Error GoTo ErrHand
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "_������", "������", 1800
        .Add , "_����", "����", 1000
        .Add , "_Ӣ����", "Ӣ����", 1500
        .Add , "_����", "����", 800
        .Add , "_����", "����", 600
        .Add , "_С��", "С��", 600
        .Add , "_��λ", "��λ", 800
        .Add , "_����", "����", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("_����").Position = 1
        .SortKey = .ColumnHeaders("_����").Index - 1: .SortOrder = lvwAscending
    End With
    
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "1") = "1" Then
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picVBar.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picHBar.Top = CLng(strTemp)
        End If
    End If
    
    '��ֱ��ͨ���˵����е�Ȩ�޿���
    If InStr(1, mstrPrivs, "��ɾ��") = 0 Then
        Me.mnuClass.Visible = False
        Me.mnuEdit.Visible = False
        Me.tlbThis.Buttons("Class").Visible = False
        Me.tlbThis.Buttons("Split2").Visible = False
        Me.tlbThis.Buttons("Add").Visible = False
        Me.tlbThis.Buttons("Modify").Visible = False
        Me.tlbThis.Buttons("Delete").Visible = False
        Me.tlbThis.Buttons("Split3").Visible = False
    End If
    With Me.hgdInfo
        .Rows = 7: .Cols = 2: .FixedRows = 0: .FixedCols = 0
        .ColWidth(0) = 1000: .ColAlignment(0) = 6
        .ColWidth(1) = .Width - .ColWidth(0) - Me.SysInfo.ScrollBarSize: .ColAlignment(1) = 1
        .TextMatrix(0, 0) = "[�ٴ�����]"
        .TextMatrix(1, 0) = "[�Ա�����]"
        .TextMatrix(2, 0) = "[ ��ʾ�� ]"
        .TextMatrix(3, 0) = "[ ��ֵ�� ]"
'        .TextMatrix(4, 0) = "[ ��ʼֵ ]"
'        .TextMatrix(5, 0) = "[���ֱ���]"
        .TextMatrix(4, 0) = "[������Ŀ]"
    End With
    
    '�����Ѿ����õ�������������
    gstrSql = "select ����,���� from ������������ Where ����<>3 order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cmdKind.Count Then
                Load Me.cmdKind(.AbsolutePosition - 1)
            End If
            Me.cmdKind(.AbsolutePosition - 1).Caption = !���� & "(&" & .AbsolutePosition & ")"
            Me.cmdKind(.AbsolutePosition - 1).Left = 0
            Me.cmdKind(.AbsolutePosition - 1).ZOrder 0
            Me.cmdKind(.AbsolutePosition - 1).Visible = True
            .MoveNext
        Loop
    End With
    Call cmdKind_Click(0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Err = 0: On Error Resume Next
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 3000 Then .Top = 3000
        If .Top > Me.ScaleHeight - lngStatus - 1000 Then .Top = Me.ScaleHeight - lngStatus - 1000
    End With
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = lngTools
        .Height = Me.picHBar.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    With Me.hgdInfo
        .Left = Me.picVBar.Left + Me.picVBar.Width + 15
        .Top = Me.picHBar.Top + Me.picHBar.Height
        .Height = Me.ScaleHeight - lngStatus - .Top
        .Width = Me.ScaleWidth - .Left - 15
        .ColWidth(1) = .Width - .ColWidth(0) - Me.SysInfo.ScrollBarSize - 15
    End With
    Call zlGrdRowHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picHBar.Top)
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmVItemEdit.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Err = 0: On Error GoTo ErrHand
    '------------------------------------------------
    gstrSql = "select �ٴ�����,��ʾ��,�Ա���,��ֵ��,��ʼֵ,���ֱ���,��ֵ����,���� from ����������Ŀ where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Item.Key, 2))
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.hgdInfo.TextMatrix(0, 1) = IIf(IsNull(!�ٴ�����), "", !�ٴ�����)
            Select Case IIf(IsNull(!�Ա���), 0, !�Ա���)
            Case 0
                Me.hgdInfo.TextMatrix(1, 1) = "���Ա�����"
            Case 1
                Me.hgdInfo.TextMatrix(1, 1) = "������ʹ��"
            Case 2
                Me.hgdInfo.TextMatrix(1, 1) = "��Ů��ʹ��"
            End Select
            Select Case IIf(IsNull(!��ʾ��), 0, !��ʾ��)
            Case 0
                Me.hgdInfo.TextMatrix(2, 1) = "�ı���"
            Case 1
                Me.hgdInfo.TextMatrix(2, 1) = "���°�ť"
            Case 2
                Me.hgdInfo.TextMatrix(2, 1) = "����ѡ��"
            Case 3
                Me.hgdInfo.TextMatrix(2, 1) = "��ѡ��ť"
            Case 4
                Me.hgdInfo.TextMatrix(2, 1) = "��ѡ��ť"
            End Select
            Me.hgdInfo.TextMatrix(3, 1) = IIf(IsNull(!��ֵ��), "", !��ֵ��)
'            Me.hgdInfo.TextMatrix(4, 1) = IIf(IsNull(!��ʼֵ), "", !��ʼֵ)
'            Select Case IIf(IsNull(!���ֱ���), 0, !���ֱ���)
'            Case 0
'                Me.hgdInfo.TextMatrix(5, 1) = "��Ŀ��+��Ŀֵ+��λ"
'            Case 1
'                Me.hgdInfo.TextMatrix(5, 1) = "��Ŀֵ+��λ+��Ŀ��"
'            Case 2
'                Me.hgdInfo.TextMatrix(5, 1) = "��Ŀֵ+��λ"
'            End Select
'            strTemp = IIf(IsNull(!��ֵ����), "", !��ֵ����)
'            If Trim(strTemp) <> "" Then
'                Me.hgdInfo.TextMatrix(5, 1) = Me.hgdInfo.TextMatrix(5, 1) & "����ֵʱ����Ϊ��" & strTemp & "��"
'            End If
            Me.hgdInfo.TextMatrix(4, 1) = IIf(!���� = 0, "��", "��")
        Else
            Call zlClearDetail
        End If
    End With
    Call zlGrdRowHeight
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.mnuEdit.Visible Then Call PopupMenu(Me.mnuEdit, 2)
End Sub

Private Sub mnuClassAdd_Click()
    With frmVItemClass
        strTemp = Me.cmdKind(Val(Me.tvwClass.Tag)).Caption
        .lblKind.Tag = IIf(Val(Me.tvwClass.Tag) < 2, Val(Me.tvwClass.Tag), Val(Me.tvwClass.Tag) + 1) + 1
        If .lblKind.Tag = 5 Then .lblKind.Tag = 6
        
        .lblKind.Caption = Mid(strTemp, 1, Len(strTemp) - 4)
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "����"
        .Show 1, Me
    End With
    If Me.tvwClass.SelectedItem Is Nothing Then
        Call zlRefClasses
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuClassDel_Click()
    Err = 0: On Error GoTo ErrHand
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("���ɾ���÷��ࡰ" & Me.tvwClass.SelectedItem.Text & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "ZL_��������_DELETE(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Dim strParentKey As String
    If Me.tvwClass.SelectedItem.Next Is Nothing Then
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            Call zlRefClasses
        Else
            strParentKey = Me.tvwClass.SelectedItem.Parent.Key
            Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
            If Me.tvwClass.Nodes(strParentKey).Children = 0 Then
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Key, 2))
            Else
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Child.Key, 2))
            End If
        End If
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Next.Key, 2))
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()
    With frmVItemClass
        strTemp = Me.cmdKind(Val(Me.tvwClass.Tag)).Caption
        .lblKind.Tag = IIf(Val(Me.tvwClass.Tag) < 2, Val(Me.tvwClass.Tag), Val(Me.tvwClass.Tag) + 1) + 1
        If .lblKind.Tag = 5 Then .lblKind.Tag = 6
        .lblKind.Caption = Mid(strTemp, 1, Len(strTemp) - 4)
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(��)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
        .txtSymbol = Me.tvwClass.SelectedItem.Tag
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
End Sub

Private Sub mnuEditAdd_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾ��Ŀ��", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call frmVItemEdit.ShowMe(Me, 0, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
    Else
        Call frmVItemEdit.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    Call zlRefRecords
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHand
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If InStr(mnuEditDelete.Caption, "ɾ��") > 0 Then  'ɾ��
            If MsgBox("���ɾ����" & .SelectedItem.Text & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "ZL_������Ŀ_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Call .ListItems.Remove(.SelectedItem.Key)
            If .SelectedItem Is Nothing Then
                Call zlClearDetail
            Else
                Call lvwItems_ItemClick(.SelectedItem)
            End If
        Else                                    '���Ϊ�Ǳ�����
            Call EarMarkMustItem(Mid(.SelectedItem.Key, 2), False)
            Call zlRefRecords
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditModify_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾ��Ŀ��", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    If InStr(mnuEditModify.Caption, "�޸�") > 0 Then '�޸�
        Call frmVItemEdit.ShowMe(Me, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Else                                        '���Ϊ������
        Call EarMarkMustItem(Mid(lvwItems.SelectedItem.Key, 2), True)
    End If
    
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id����Ŀ=��Ŀid
    Dim lng����id As Long
    Dim lng��Ŀid As Long
    
    If Not Me.tvwClass.SelectedItem Is Nothing Then
        lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng��Ŀid = Mid(lvwItems.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIf(lng����id = 0, "", lng����id), _
        "��Ŀ=" & IIf(lng��Ŀid = 0, "", lng��Ŀid))
End Sub

Private Sub mnuViewFind_Click()
    With frmVItemFind
        .Show , Me
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picClass_Resize()
    Dim intCount As Integer
    Err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
            Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHBar.Top = Me.picHBar.Top + y
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + x
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        If Me.mnuClass.Visible Then Call PopupMenu(Me.mnuClass, 2)
    Case "Add"
        Call mnuEditAdd_Click
    Case "Modify"
        Call mnuEditModify_Click
    Case "Delete"
        Call mnuEditDelete_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.mnuClass.Visible Then Call PopupMenu(Me.mnuClass, 2)
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwItems.Tag = Node.Key Then Exit Sub
    Me.lvwItems.Tag = Node.Key
    Call zlRefRecords
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    '---------------------------------------------
    '��д���Ʒ�����Ŀ(�˴�ΪҩƷ����)�����ղ�ͬ���͵�������
    '---------------------------------------------
    Dim lngTmp As Long
    'Ȩ�޿���
    
    '������ʾ����
    If Val(Me.tvwClass.Tag) = 0 Then '�̶���Ŀ
        Me.mnuClass.Enabled = False: Me.mnuClassAdd.Enabled = False: Me.mnuClassMod.Enabled = False: Me.mnuClassDel.Enabled = False
        'Me.mnuEdit.Enabled = False: Me.mnuEditAdd.Enabled = False: Me.mnuEditModify.Enabled = False: Me.mnuEditDelete.Enabled = False
        '����Ŀ�˵���Ϊ������Ŀ/�Ǳ�����
        Me.mnuEditAdd.Visible = False: Me.mnuEditModify.Caption = "����(&M)": Me.mnuEditDelete.Caption = "��ѡ(&O)"
        Me.tlbThis.Buttons("Class").Enabled = False
        Me.tlbThis.Buttons("Add").Visible = False: Me.tlbThis.Buttons("Modify").Caption = "����": Me.tlbThis.Buttons("Delete").Caption = "��ѡ"
    Else                            '�ɸ�����Ŀ
        Me.mnuClass.Enabled = True: Me.mnuClassAdd.Enabled = True: Me.mnuClassMod.Enabled = True: Me.mnuClassDel.Enabled = True
        Me.mnuEditAdd.Visible = True: Me.mnuEditModify.Caption = "�޸�(&M)": Me.mnuEditDelete.Caption = "ɾ��(&D)"
        Me.tlbThis.Buttons("Class").Enabled = True
        Me.tlbThis.Buttons("Add").Visible = True: Me.tlbThis.Buttons("Modify").Caption = "�޸�": Me.tlbThis.Buttons("Delete").Caption = "ɾ��"
    End If
    
    Me.lvwItems.ListItems.Clear
    '��д����
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ������������" & _
            " Where ���� = [1] " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    lngTmp = 1 + IIf(Val(Me.tvwClass.Tag) < 2, Val(Me.tvwClass.Tag), Val(Me.tvwClass.Tag) + 1)
    lngTmp = IIf(lngTmp = 5, 6, lngTmp)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngTmp)
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Visible = True
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
        Else
            Me.tvwClass.Nodes(1).Selected = True
        End If
        Call zlRefRecords
    Else
        Call zlClearDetail
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    '---------------------------------------------
    '��д��Ŀ�б�
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select  I.ID,I.����,I.������,I.Ӣ����,I.����,I.����,I.С��,I.С��,I.��λ,I.����" & _
            " from ����������Ŀ I" & _
            " where I.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(Me.tvwClass.SelectedItem.Key, 2))
    
    With rsTemp
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !������)
            objItem.Icon = "item": objItem.SmallIcon = "item"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_Ӣ����").Index - 1) = IIf(IsNull(!Ӣ����), "", !Ӣ����)
            Select Case IIf(IsNull(!����), 0, !����)
            Case 0
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "��ֵ"
            Case 1
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "����"
            Case 2
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "����"
            Case 3
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = "�߼�"
            End Select
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_С��").Index - 1) = IIf(IsNull(!С��), "", !С��)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = IIf(!���� = 0, "��", "��")
            If !ID = lngItem Then
                objItem.Selected = True
            End If
            .MoveNext
        Loop
    
    End With
    If Me.lvwItems.ListItems.Count > 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then Me.lvwItems.ListItems(1).Selected = True
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        Err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.lvwItems.ListItems.Count & "����Ŀ"
    Else
        Call zlClearDetail
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlClearDetail()
    '---------------------------------------------
    '���������ϸ��Ϣ��ʾ����
    '---------------------------------------------
    With Me.hgdInfo
        .TextMatrix(0, 1) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(2, 1) = ""
        .TextMatrix(3, 1) = ""
        .TextMatrix(4, 1) = ""
        .TextMatrix(5, 1) = ""
        .TextMatrix(6, 1) = ""
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    strTemp = Me.cmdKind(Val(Me.tvwClass.Tag)).Caption
    objPrint.Title.Text = Mid(strTemp, 1, Len(strTemp) - 4) & "������Ŀ�嵥"
    
    objPrint.UnderAppItems.Add "���ࣺ" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '���ݵ������ݵ�������������и߶ȣ��Ա�֤���ݵ�������ʾ
    '---------------------------------------------
    Dim intRow As Integer, lngColWidth As Long
    With Me.hgdInfo
        For intRow = .FixedRows To .Rows - 1
            lngColWidth = .ColWidth(1)
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(intRow, 1)
            .RowHeight(intRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Public Sub zlLocateItem(lngClassId As Long, lngItemID As Long)
    '---------------------------------------------
    '��λ��ָ������Ŀ���ڲ���ʱʹ��
    '---------------------------------------------
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lngClassId)
    Me.tvwClass.Nodes("_" & lngClassId).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & lngItemID)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub
Private Sub EarMarkMustItem(ByVal lngItemID As Long, ByVal ItemMust As Boolean)
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSql = "Select ����id, ����, ������, Ӣ����, ����, ����, С��, ��λ, �ٴ�����, ��ʾ��, �Ա���, ��ֵ��, ��ʼֵ, ���ֱ���, ��ֵ����,��̬��" & vbNewLine & _
                "From ����������Ŀ" & vbNewLine & _
                "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        If .EOF Then Exit Sub
        gstrSql = Nvl(!����id, 0) & ",'" & !���� & "','" & !������ & "','" & !Ӣ���� & "'," & Nvl(!����, 0) & _
                "," & Nvl(!����, 0) & "," & Nvl(!С��, 0) & ",'" & !��λ & "','" & !�ٴ����� & "'," & Nvl(!��ʾ��, 0) & _
                "," & Nvl(!�Ա���, 0) & ",'" & !��ֵ�� & "','" & !��ʼֵ & "'," & Nvl(!���ֱ���, 1) & ",'" & !��ֵ���� & "'," & IIf(ItemMust, 1, 0) & "," & Nvl(!��̬��, 0)
    End With
    gstrSql = "ZL_������Ŀ_UPDATE(" & lngItemID & "," & gstrSql & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

