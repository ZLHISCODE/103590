VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDefTree 
   Caption         =   "��ѯĿ¼�滮"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmDefTree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvw 
      Height          =   1275
      Left            =   3330
      TabIndex        =   4
      Top             =   855
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2249
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��ʾ����"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ҳ������"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ҳ����"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�̶�ҳ��"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "��С"
         Object.Width           =   1058
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3120
      Left            =   45
      TabIndex        =   3
      Top             =   780
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   5503
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   1905
      Top             =   4740
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
            Picture         =   "frmDefTree.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":3194
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":4E9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1230
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":6BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":98FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":B604
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   6960
      Top             =   360
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
            Picture         =   "frmDefTree.frx":D30E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":D52E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":D74E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":D968
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":DB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":DDA2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":DFC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":E51C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":EA76
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":EC92
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":EEB2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7545
      Top             =   360
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
            Picture         =   "frmDefTree.frx":F0D2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":F2F2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":F512
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":F72C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":F94C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":FB66
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":FD86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":102E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":1083A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":10A56
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTree.frx":10C76
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8880
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ŀ¼"
               Key             =   "Ŀ¼"
               Object.ToolTipText     =   "Ŀ¼"
               Object.Tag             =   "Ŀ¼"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ҳ��"
               Key             =   "ҳ��"
               Object.ToolTipText     =   "ҳ��"
               Object.Tag             =   "ҳ��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��ʾ˳������"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��ʾ˳������"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "�鿴"
               Object.ToolTipText     =   "��ʾҳ��鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5805
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      SimpleText      =   $"frmDefTree.frx":10E96
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDefTree.frx":10EDD
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
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
   Begin VB.Image picX 
      Height          =   1530
      Left            =   2565
      MousePointer    =   9  'Size W E
      Top             =   930
      Width           =   210
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileUpdateFolder 
         Caption         =   "���²�ѯĿ¼(&U)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "����Ŀ¼(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditAddPage 
         Caption         =   "����ҳ��(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUp 
         Caption         =   "��ʾ˳������(&U)"
      End
      Begin VB.Menu mnuEditDown 
         Caption         =   "��ʾ˳������(&D)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
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
End
Attribute VB_Name = "frmDefTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFist As Boolean

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
    
    tvw.Nodes.Clear
    lvw.ListItems.Clear
    tvw.Nodes.Add , , "K0", "��ҳ", 1, 1
    tvw.Nodes(1).Selected = True
    tvw.Nodes(1).Expanded = True
    
    Call LoadTree("K0", 0)
            
    If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)
End Sub

Private Sub Form_Load()
    mblnFist = True
    
    RestoreWinState Me, App.ProductName
    
    Call mnuViewIcon_Click(lvw.View)
    
    picX.MousePointer = 9
    picX.Width = 45
            
    Call ReadRegister
    Call ModulePrivs
    
End Sub

Private Sub Form_Resize()
    '���ݴ���״̬,���������и��ؼ�����ʾλ��
    Dim sglCbrH As Single
    Dim sglStbH As Single
    
    On Error Resume Next
    sglCbrH = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sglStbH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    Call ResizeControl(tvw, 0, sglCbrH, picX.Left, Me.ScaleHeight - sglStbH - sglCbrH)
    Call ResizeControl(picX, picX.Left, lvw.Top, picX.Width, lvw.Height)
    Call ResizeControl(lvw, picX.Left + picX.Width, tvw.Top, Me.ScaleWidth - picX.Left - picX.Width, tvw.Height)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteRegister
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_DblClick()
    If mnuEdit.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call AdjustEnabled
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuEdit, 2
End Sub

Private Sub mnuEditAddPage_Click()
    If tvw.SelectedItem Is Nothing Then Exit Sub
    If tvw.SelectedItem.Image = 2 Then Exit Sub

    If frmDefTreePage.ShowTreePageBox(Me, 0, Val(tvw.SelectedItem.Tag)) Then
        
    End If
End Sub

Private Sub mnuEditDelete_Click()
    Dim strTmp As String
    Dim vIndex As Long
    Dim vTree As Boolean
    
    gstrSQL = ""
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If Val(lvw.SelectedItem.Tag) > 0 Then
        strTmp = "��ȷ��Ҫɾ����ѯҳ��[" & lvw.SelectedItem.Text & "]��"
    Else
        strTmp = "��ȷ��Ҫɾ����ѯĿ¼[" & lvw.SelectedItem.Text & "]�������Ĳ�ѯҳ����"
        vTree = True
    End If
    gstrSQL = "zl_��ѯҳ������_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
    
    If MsgBox(strTmp, vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errHand
    
    If gstrSQL <> "" Then Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    vIndex = lvw.SelectedItem.Index
    If vTree Then tvw.Nodes.Remove tvw.Nodes(lvw.SelectedItem.Key).Index
    lvw.ListItems.Remove lvw.SelectedItem.Index
    Call NextLvwPos(lvw, vIndex)
                    
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuEditDown_Click()
    '����ǰ����Ŀ������һ�У�ͬʱ�������ݿ�

    Dim svrAry(6) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    Dim svrIcon As Long
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    intPre = lvw.SelectedItem.Index + 1
    
    If intPre < lvw.ListItems.Count + 1 Then
        strSQL(0) = "zl_��ѯҳ������_AdjustOrder(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ",-1)"
        strSQL(1) = "zl_��ѯҳ������_AdjustOrder(" & Val(Mid(lvw.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_��ѯҳ������_AdjustOrder(-1," & Val(Mid(lvw.ListItems(intPre).Key, 2)) & ")"
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
        'gcnOracle.Execute strSQL(0), , adCmdStoredProc
        'gcnOracle.Execute strSQL(1), , adCmdStoredProc
        'gcnOracle.Execute strSQL(2), , adCmdStoredProc
        
        Call zlDatabase.ExecuteProcedure(strSQL(0), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(1), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(2), Me.Caption)
        
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        svrAry(0) = lvw.ListItems(intPre).Text
        svrAry(1) = lvw.ListItems(intPre).SubItems(1)
        svrAry(2) = lvw.ListItems(intPre).SubItems(2)
        svrAry(3) = lvw.ListItems(intPre).SubItems(3)
        svrAry(5) = lvw.ListItems(intPre).Tag
        svrIcon = lvw.ListItems(intPre).SmallIcon
        
        lvw.ListItems(intPre).Text = lvw.SelectedItem.Text
        lvw.ListItems(intPre).SubItems(1) = lvw.SelectedItem.SubItems(1)
        lvw.ListItems(intPre).SubItems(2) = lvw.SelectedItem.SubItems(2)
        lvw.ListItems(intPre).SubItems(3) = lvw.SelectedItem.SubItems(3)
        lvw.ListItems(intPre).Tag = lvw.SelectedItem.Tag
        lvw.ListItems(intPre).Icon = lvw.SelectedItem.Icon
        lvw.ListItems(intPre).SmallIcon = lvw.SelectedItem.SmallIcon
        
        lvw.SelectedItem.Text = svrAry(0)
        lvw.SelectedItem.SubItems(1) = svrAry(1)
        lvw.SelectedItem.SubItems(2) = svrAry(2)
        lvw.SelectedItem.SubItems(3) = svrAry(3)
        lvw.SelectedItem.Tag = svrAry(5)
        lvw.SelectedItem.Icon = svrIcon
        lvw.SelectedItem.SmallIcon = svrIcon
        
        lvw.ListItems(intPre).Selected = True
        Call mnuViewRefresh_Click
        Call AdjustEnabled
    End If
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuEditModify_Click()
    If tvw.SelectedItem Is Nothing Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If Val(lvw.SelectedItem.Tag) > 0 Then
        Call frmDefTreePage.ShowTreePageBox(Me, Val(Mid(lvw.SelectedItem.Key, 2)), Val(tvw.SelectedItem.Tag))
    Else
        Call frmDefTreeEdit.ShowTreeBox(Me, Val(Mid(lvw.SelectedItem.Key, 2)))
    End If
End Sub

Private Sub mnuEditNew_Click()
    If tvw.SelectedItem.Key <> "K0" Then Exit Sub
            
    If frmDefTreeEdit.ShowTreeBox(Me, 0) Then

    End If
End Sub

Private Sub mnuEditUp_Click()
    Dim svrAry(6) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    Dim svrIcon As Long
    
    intPre = lvw.SelectedItem.Index - 1
    
    If intPre > 0 Then
    
        strSQL(0) = "zl_��ѯҳ������_AdjustOrder(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ",-1)"
        strSQL(1) = "zl_��ѯҳ������_AdjustOrder(" & Val(Mid(lvw.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_��ѯҳ������_AdjustOrder(-1," & Val(Mid(lvw.ListItems(intPre).Key, 2)) & ")"
                    
        On Error GoTo errHand
        gcnOracle.BeginTrans
        'gcnOracle.Execute strSQL(0), , adCmdStoredProc
        'gcnOracle.Execute strSQL(1), , adCmdStoredProc
        'gcnOracle.Execute strSQL(2), , adCmdStoredProc
        
        Call zlDatabase.ExecuteProcedure(strSQL(0), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(1), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(2), Me.Caption)
        
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        svrAry(0) = lvw.ListItems(intPre).Text
        svrAry(1) = lvw.ListItems(intPre).SubItems(1)
        svrAry(2) = lvw.ListItems(intPre).SubItems(2)
        svrAry(3) = lvw.ListItems(intPre).SubItems(3)
        svrAry(5) = lvw.ListItems(intPre).Tag
        svrIcon = lvw.ListItems(intPre).SmallIcon
        
        lvw.ListItems(intPre).Text = lvw.SelectedItem.Text
        lvw.ListItems(intPre).SubItems(1) = lvw.SelectedItem.SubItems(1)
        lvw.ListItems(intPre).SubItems(2) = lvw.SelectedItem.SubItems(2)
        lvw.ListItems(intPre).SubItems(3) = lvw.SelectedItem.SubItems(3)
        lvw.ListItems(intPre).Tag = lvw.SelectedItem.Tag
        lvw.ListItems(intPre).Icon = lvw.SelectedItem.Icon
        lvw.ListItems(intPre).SmallIcon = lvw.SelectedItem.SmallIcon
        
        lvw.SelectedItem.Text = svrAry(0)
        lvw.SelectedItem.SubItems(1) = svrAry(1)
        lvw.SelectedItem.SubItems(2) = svrAry(2)
        lvw.SelectedItem.SubItems(3) = svrAry(3)
        lvw.SelectedItem.Tag = svrAry(5)
        lvw.SelectedItem.Icon = svrIcon
        lvw.SelectedItem.SmallIcon = svrIcon
        
        lvw.ListItems(intPre).Selected = True
        
        Call mnuViewRefresh_Click
        Call AdjustEnabled
    End If
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuFileExcel_Click()
    Call PrintObject(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFilePreView_Click()
    Call PrintObject(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintObject(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileUpdateFolder_Click()
    Call gfrmMain.FrameDefault.RefreshFolder
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuPop1_Click(Index As Integer)
    Select Case Index
    Case 0
        Call mnuEditAddPage_Click
    Case 1
        Call mnuEditModify_Click
    Case 2
        Call mnuEditDelete_Click
    Case 3
    Case 4
        Call mnuViewIcon_Click(0)
    Case 5
        Call mnuViewIcon_Click(1)
    Case 6
        Call mnuViewIcon_Click(2)
    Case 7
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub mnuPop2_Click(Index As Integer)
    Select Case Index
    Case 0
        Call mnuEditNew_Click
    Case 1
        Call mnuEditModify_Click
    Case 2
        Call mnuEditDelete_Click
    Case 3
    Case 4
        Call mnuEditUp_Click
    Case 5
        Call mnuEditDown_Click
    End Select
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Dim svrKey As String
    Dim svrKey2 As String
    
    svrKey = SaveLvwItem(tvw)
    svrKey2 = SaveLvwItem(lvw)
    
    tvw.Nodes.Clear
    lvw.ListItems.Clear
    tvw.Nodes.Add , , "K0", "��ҳ", 1, 1
    Call LoadTree("K0", 0)
    
    On Error Resume Next
    If Not (tvw.Nodes(svrKey) Is Nothing) Then
        tvw.Nodes(svrKey).EnsureVisible
        tvw.Nodes(svrKey).Selected = True
    End If
    On Error GoTo 0
    
    If tvw.SelectedItem Is Nothing And tvw.Nodes.Count > 0 Then tvw.Nodes(1).Selected = True
    
    Call LoadStatus
    If Not (tvw.SelectedItem Is Nothing) Then
        tvw.Nodes(tvw.SelectedItem.Key).Expanded = True
        Call tvw_NodeClick(tvw.SelectedItem)
        Call RestoreLvwItem(lvw, svrKey2)
    End If

End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub


Private Sub mnuViewToolText_Click()
    Dim i As Long
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub



Private Sub picX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picX.Left = picX.Left + X
    If picX.Left < 1500 Then picX.Left = 1500
    If Me.Width - picX.Left - picX.Width < 1500 Then picX.Left = Me.Width - picX.Width - 1500
    
    Form_Resize
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePreView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "Ŀ¼"
        Call mnuEditNew_Click
    Case "ҳ��"
        Call mnuEditAddPage_Click
    Case "�޸�"
        Call mnuEditModify_Click
    Case "ɾ��"
        Call mnuEditDelete_Click
    Case "����"
        Call mnuEditUp_Click
    Case "����"
        Call mnuEditDown_Click
    Case "�鿴"
        If lvw.View < 3 Then
            Call mnuViewIcon_Click(lvw.View + 1)
        Else
            Call mnuViewIcon_Click(0)
        End If
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub


Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
    Case "��ͼ��"
        Call mnuViewIcon_Click(0)
    Case "Сͼ��"
        Call mnuViewIcon_Click(1)
    Case "�б�"
        Call mnuViewIcon_Click(2)
    Case "��ϸ����"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub LoadTree(ByVal strUpKey As String, ByVal lngUpTag As Long)
    Dim rs As New ADODB.Recordset
    Dim imgIndex As Long
    Dim nod As Node
    
    On Error GoTo errHand
    If lngUpTag = 0 Then
        gstrSQL = "Select ���,�����,����,ҳ��,ҳ��ͼ��,����,��С,����,��ɫ From ��ѯҳ������ where (ҳ�� is null or ҳ��=0) and (����� is null or �����=0) order by ���"
    Else
        gstrSQL = "Select ���,�����,����,ҳ��,ҳ��ͼ��,����,��С,����,��ɫ From ��ѯҳ������ where (ҳ�� is null or ҳ��=0) and �����=[1] order by ���"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUpTag)
    If rs.BOF = False Then
        While Not rs.EOF
            imgIndex = 2
            If IIf(IsNull(rs!ҳ��), 0, rs!ҳ��) = 0 Then imgIndex = 1
            Set nod = tvw.Nodes.Add(strUpKey, tvwChild, "K" & rs!���, IIf(IsNull(rs!����), "", rs!����), imgIndex, imgIndex)
            nod.Tag = rs!���
            If imgIndex = 1 Then Call LoadTree("K" & nod.Tag, Val(nod.Tag))
            rs.MoveNext
        Wend
    End If
    CloseRecord rs
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
End Sub



Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim svrKey As String
    
    svrKey = SaveLvwItem(lvw)
    
    Call LoadPageList(Val(Mid(Node.Key, 2)))
    
    Call RestoreLvwItem(lvw, svrKey)
    Call LoadStatus
    Call AdjustEnabled
    
End Sub

Private Sub LoadPageList(ByVal Key As Long)
    Dim Itmx As ListItem
    
    On Error GoTo errHand
    lvw.ListItems.Clear
    If Key = 0 Then
        gstrSQL = "select A.����,A.ҳ��,A.ҳ��ͼ��,A.����,A.��С,A.����,A.��ɫ,A.���,A.�����,B.ҳ������,B.�̶�ҳ��,B.ҳ���� from ��ѯҳ������ A,��ѯҳ��Ŀ¼ B where A.ҳ��=B.ҳ�����(+) and (A.����� is null or A.�����=[1]) order by A.���"
    Else
        gstrSQL = "select A.����,A.ҳ��,A.ҳ��ͼ��,A.����,A.��С,A.����,A.��ɫ,A.���,A.�����,B.ҳ������,B.�̶�ҳ��,B.ҳ���� from ��ѯҳ������ A,��ѯҳ��Ŀ¼ B where A.ҳ��=B.ҳ�����(+) and A.�����=[1] order by A.���"
    End If
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Key)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!���, IIf(IsNull(gRs!����), "", gRs!����), 1, 1)
            Itmx.SubItems(1) = IIf(IsNull(gRs!ҳ������), "", gRs!ҳ������)
            
            Itmx.Tag = IIf(IsNull(gRs!ҳ��), 0, gRs!ҳ��)
            If Val(Itmx.Tag) > 0 Then
                Itmx.SubItems(2) = "��׼"
                Itmx.SubItems(3) = IIf(IsNull(gRs!�̶�ҳ��), "", IIf(gRs!�̶�ҳ�� = 1, "��", ""))
                If Itmx.SubItems(3) = "" Then
                    Itmx.SmallIcon = 2
                    Itmx.Icon = 2
                Else
                    Itmx.SmallIcon = 3
                    Itmx.Icon = 3
                End If
            End If
            
            Itmx.SubItems(4) = IIf(IsNull(gRs!����), "����", gRs!����)
            Select Case IIf(IsNull(gRs!����), 1, gRs!����)
            Case 1
                Itmx.SubItems(5) = "����"
            Case 2
                Itmx.SubItems(5) = "б��"
            Case 3
                Itmx.SubItems(5) = "����"
            Case 4
                Itmx.SubItems(5) = "��б��"
            End Select
            Itmx.SubItems(6) = IIf(IsNull(gRs!��С), 12, gRs!��С)
            
            gRs.MoveNext
        Wend
    End If
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool, 2
End Sub

Private Sub PrintObject(ByVal intMode As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������
    '     intMode: 2��ʾԤ�� 1��ӡ 3�����EXCEL
    '���أ�
    '---------------------------------------------------
    
    Dim objPrint As New zlPrintLvw
    Dim objRow As New zlTabAppRow
        
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If UserInfo.���� = "" Then Call GetUserInfo
    
    
    objPrint.Title = "[" & tvw.SelectedItem.Text & "]�µ�ҳ���嵥"
    Set objPrint.Body.objData = lvw
    
    objPrint.BelowAppItems.Add "��ӡ��:" & UserInfo.����
    objPrint.BelowAppItems.Add "��ӡʱ��:" & Format(zlDatabase.Currentdate, "YYYY��MM��DD��")
    objPrint.Footer = "��[ҳ��]ҳ;;"

    If intMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, intMode
    End If

End Sub

Private Sub ModulePrivs()
'����:����ģ��Ȩ��,������������ػ���ʾ
'     Ȩ����:��ɾ��
        
'    If InStr(gstrPrivs, "��ɾ��") = 0 Then
'        mnuEdit.Visible = False
'
'        tbrThis.Buttons("ҳ��").Visible = False
'        tbrThis.Buttons("Ŀ¼").Visible = False
'        tbrThis.Buttons("�޸�").Visible = False
'        tbrThis.Buttons("ɾ��").Visible = False
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("Split_3").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'    End If
End Sub

Private Sub AdjustEnabled()
'����:�������ܲ˵���ť�Ŀ���״̬
    mnuFilePreView.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileExcel.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditNew.Enabled = True
    mnuEditAddPage.Enabled = True
    
    mnuEditUp.Enabled = True
    mnuEditDown.Enabled = True
        
    If lvw.ListItems.Count = 0 Then
        mnuFilePreView.Enabled = False
        mnuFilePrint.Enabled = False
        mnuFileExcel.Enabled = False
                
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        
        mnuEditUp.Enabled = False
        mnuEditDown.Enabled = False
    End If
    
    If Not (lvw.SelectedItem Is Nothing) Then
        If lvw.SelectedItem.Index - 1 <= 0 Then mnuEditUp.Enabled = False
        If lvw.SelectedItem.Index + 1 > lvw.ListItems.Count Then mnuEditDown.Enabled = False
    End If
    
    If tvw.SelectedItem.Text <> "��ҳ" Then mnuEditNew.Enabled = False
    
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePreView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("ҳ��").Enabled = mnuEditAddPage.Enabled
    tbrThis.Buttons("Ŀ¼").Enabled = mnuEditNew.Enabled
    tbrThis.Buttons("�޸�").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("ɾ��").Enabled = mnuEditDelete.Enabled
        
    tbrThis.Buttons("����").Enabled = mnuEditUp.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditDown.Enabled
End Sub

Private Sub ReadRegister()
'����:��ȡע�����Ϣ
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    picX.Left = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\����\" & Me.Name, "picXλ��", 2385)
End Sub

Private Sub WriteRegister()
'����:����Ϣд��ע���
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\����\" & Me.Name, "picXλ��", picX.Left
End Sub

Private Sub LoadStatus()
'����:��д״̬����Ϣ
    Dim vTmp As String
    
    If lvw.ListItems.Count > 0 Then
        vTmp = "[" & tvw.SelectedItem.Text & "]�¹���" & lvw.ListItems.Count & "��ҳ����ɣ�"
    Else
        If tvw.SelectedItem Is Nothing Then
            vTmp = "��ǰû���κ���Ϣ��"
        Else
            vTmp = "[" & tvw.SelectedItem.Text & "]�»�û��ҳ�棡"
        End If
    End If
    stbThis.Panels(2).Text = vTmp
End Sub

Public Sub AddLvwItem(ByVal lngKey As Long)
    
    mnuViewRefresh_Click
    
    Call RestoreLvwItem(lvw, "K" & lngKey)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

