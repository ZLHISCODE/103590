VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDeptManage 
   Caption         =   "���Ź���"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmDeptManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Tag             =   "�ɱ仯��"
   Begin VB.PictureBox picList 
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   840
      Width           =   2775
      Begin VB.PictureBox picSplit2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   255
         ScaleWidth      =   3720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3120
         Width           =   3720
         Begin VB.Label lbl�������� 
            Caption         =   "  �������Ҷ�Ӧ��ϵ"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   60
            Width           =   1575
         End
      End
      Begin MSComctlLib.TreeView tvwMain_S 
         Height          =   2865
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   5054
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvwDept 
         Height          =   1140
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   2011
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
   End
   Begin XtremeSuiteControls.TabControl tbcDetails 
      Height          =   855
      Left            =   4560
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
      _Version        =   589884
      _ExtentX        =   2143
      _ExtentY        =   1508
      _StockProps     =   64
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   3720
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   45
   End
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   5280
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3000
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   635
      SimpleText      =   $"frmDeptManage.frx":030A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDeptManage.frx":0351
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12726
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
   Begin MSComctlLib.ListView lvw��������_S 
      Height          =   1095
      Left            =   4470
      TabIndex        =   3
      Top             =   3420
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "��������"
         Object.Tag             =   "��������"
         Text            =   "��������"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "�������"
         Object.Tag             =   "�������"
         Text            =   "�������"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "˵��"
         Object.Tag             =   "˵��"
         Text            =   "˵��"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3900
      Top             =   1290
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
            Picture         =   "frmDeptManage.frx":0BE5
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1231
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":154D
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Ilscolor 
      Left            =   4050
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":186D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1A8D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1CAD
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":1ECD
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":20ED
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":230D
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":252D
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":274D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":2969
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":2B89
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":2DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":3343
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3990
      Top             =   2160
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
            Picture         =   "frmDeptManage.frx":355D
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":3BA9
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":3EC5
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   4500
      TabIndex        =   4
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Ilsrw 
      Left            =   4770
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":41E5
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4405
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4625
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4845
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4A65
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4C85
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":4EA5
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":50C5
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":52E1
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":5501
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":5721
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptManage.frx":593B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10095
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Ilsrw"
         HotImageList    =   "Ilscolor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Key             =   "New"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Start"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Object.ToolTipText     =   "ͣ��"
               Object.Tag             =   "ͣ��"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sdf"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "֧�ֹؼ���ģ������"
               Object.ToolTipText     =   "���Ҳ���,֧�ֹؼ���ģ������"
               Object.Tag             =   "����"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "View"
               Object.ToolTipText     =   "��Ա�鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  �б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "  ��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   7560
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   15
            Top             =   240
            Width           =   495
            Begin VB.Label lbl���� 
               Caption         =   "����"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   74
               Width           =   495
            End
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   8520
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "����"
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label lblFind 
            Caption         =   "����"
            Height          =   255
            Left            =   8400
            TabIndex        =   14
            Top             =   2520
            Width           =   615
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcDept 
      Height          =   870
      Left            =   6555
      TabIndex        =   17
      Top             =   5010
      Width           =   1230
      _Version        =   589884
      _ExtentX        =   2170
      _ExtentY        =   1535
      _StockProps     =   64
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
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&T)"
      End
      Begin VB.Menu mnuEditRecovery 
         Caption         =   "�ָ�(&R)"
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExtend 
         Caption         =   "��չ��Ϣά��(&E)"
      End
      Begin VB.Menu mnuEditExpand 
         Caption         =   "�ӳ��¼�����(&X)"
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
      Begin VB.Menu mnuFind 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuViewSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "��ʾ�����¼�(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowStop 
         Caption         =   "��ʾͣ�ò���(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewShowDel 
         Caption         =   "��ʾ��ɾ������(&Y)"
      End
      Begin VB.Menu mnuViewReflash 
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
         Caption         =   "Web�ϵ�����"
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
      Begin VB.Menu mnuShortMenu1 
         Caption         =   "�ָ�(&R)"
         Index           =   4
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
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�ָ�(&R)"
         Index           =   4
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
Attribute VB_Name = "frmDeptManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sngStartX As Single, sngStartY As Single    '�ƶ�ǰ����λ��
Dim mblnLoad As Boolean  '���ڻ�δ��ʱΪ��
Dim mblnItem As Boolean  'Ϊ���ʾ������ListViewĳһ����
Dim mintColumn As Integer '
Dim mstrKey As String
Private Const mstrLvw As String = "����,2000,0,1;����,800,0,2;����,1440,0,0;λ��,2000,0,0;����ʱ��,1440,0,0;����ʱ��,1440,0,0;�ϼ�����,2000,0,0"
Dim mblnҩ��  As Boolean
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
Private mint��ɾ�� As Integer        '����"��ɾ������"����ʱ��ɾ��������Ϊ��ɾ����0-��ɾ��;1-��ɾ��
Private mint���� As Integer         '1-�ٴ�����;2-����;3-�ٴ��Ҳ���
Private mlng�������� As Long
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mrsFind As ADODB.Recordset
Private mstrFindValue As String     '��¼��ѯ�ı����ֵ
Private mint���� As Integer         '��¼����λ�� 1��listview�ؼ��� 2��������
Private Const mint����� As Integer = 0   'ҳ����ʾ��ʽ ����ι�ϵ��ʾ
Private Const mint������ As Integer = 1   'ҳ����ʾ��ʽ   �����ʹ�ϵ��ʾ
Private Const mCON�������� As Integer = 0
Private Const mCON��չ��Ϣ As Integer = 1
Private mobjForm As frmDeptExtend
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�

Private Function CheckExistDepPres(ByVal lngDepID As Long) As Boolean
    '���ò������Ƿ������Ա
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select ��Աid From ������Ա " & _
        " Where ����id In (Select ID From ���ű� " & _
        " Where ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null " & _
        " Start With ID = [1] Connect By Prior ID = �ϼ�id) And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鲿����Ա", lngDepID)
    
    If rsTemp.RecordCount > 0 Then
        CheckExistDepPres = True
        Exit Function
    End If
End Function

Private Sub InitTabControl()
    Dim i As Integer
    '��ʼ��Tabcontrol�ؼ�
    With Me.tbcDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        .InsertItem(mint�����, "�������ʾ", picList.hwnd, 0).Tag = "�������ʾ"
        .InsertItem(mint������, "��������ʾ", picList.hwnd, 0).Tag = "��������ʾ"
        
        .Item(mint������).Selected = True
        .Item(mint�����).Selected = True
    End With
    
    With Me.tbcDept
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        
        Set mobjForm = New frmDeptExtend
        Call SetFormVisible(mobjForm.hwnd) '�����������С������

        .InsertItem(mCON��������, "��������", lvw��������_S.hwnd, 0).Tag = "��������"
        .InsertItem(mCON��չ��Ϣ, "��չ��Ϣ", mobjForm.hwnd, 0).Tag = "��չ��Ϣ"
        
        .Item(mCON��չ��Ϣ).Selected = True
        .Item(mCON��������).Selected = True
    End With
End Sub

Private Sub CheckHaveDelDept()
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select ID From ���ű� Where ���� = '-'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�����ɾ�����ŷ���")
    
    If Not rsTemp.EOF Then
        mint��ɾ�� = 1
    Else
        mint��ɾ�� = 0
    End If
End Sub

Private Function Check��������(ByVal lngDeptID As Long) As Boolean
    '�����������
    Dim rsData As ADODB.Recordset
    
    If mlng�������� = 0 Then
        Check�������� = True
        Exit Function
    End If
    
    '�����ǰ�����������õ���Һ��������
    If mlng�������� = lngDeptID Then
        MsgBox "�ò����ѱ�����ΪҽԺ����Һ�������ģ�����ɾ����ͣ�ã����ڻ������������д���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����ǰ���ŵ��¼��������õ���Һ��������
    gstrSQL = "Select Id,���� || '-' || ���� As ���� From ���ű� " & _
        " Where ID In (Select ID From ���ű� Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id) And ID = [2] "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��鲿��", lngDeptID, mlng��������)
    
    If rsData.RecordCount > 0 Then
        MsgBox "�ò��ŵ��¼�����(" & rsData!���� & ")�ѱ�����ΪҽԺ����Һ�������ģ�����ɾ����ͣ�ã����ڻ������������д���", vbInformation, gstrSysName
        Exit Function
    End If
    
    Check�������� = True
End Function

Private Sub Show�������Ҷ�Ӧ(ByVal str����ID As String)
    'ȡ�������Ҷ�Ӧ��ϵ
    Dim strCon As String
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset
    Dim sngTop, sngBottom As Single
    Dim nod As Node
        
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    tvwDept.Visible = False
    tvwMain_S.Top = 0
    tvwMain_S.Height = IIF(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top - picSplit2.Height, 0)
    tvwMain_S.Left = 0
    
    If str����ID = "" Or str����ID = "Root" Then Exit Sub
     
    gstrSQL = "Select �������� From ��������˵�� Where ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", Val(Mid(str����ID, 2)))
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    mint���� = 0
    Do While Not rsTmp.EOF
        If rsTmp("��������") = "�ٴ�" Then
            If mint���� = 2 Then
                mint���� = 3
            Else
                mint���� = 1
            End If
        End If
        If rsTmp("��������") = "����" Then
            If mint���� = 1 Then
                mint���� = 3
            Else
                mint���� = 2
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If mint���� = 0 Then Exit Sub
    
    If mint���� = 1 Then
        strCon = " Select Distinct ����id From �������Ҷ�Ӧ Where ����id = [1] "
    ElseIf mint���� = 2 Then
        strCon = " Select Distinct ����id From �������Ҷ�Ӧ Where ����id = [1] "
    ElseIf mint���� = 3 Then
        strCon = " Select ����id As ID From �������Ҷ�Ӧ Where ����id = [1] " & _
                " Union " & _
                " Select ����id As ID From �������Ҷ�Ӧ Where ����id = [1] "
    End If

    If mnuViewShowStop.Checked = False Then
        strTemp = " And (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null ) "
    End If

    gstrSQL = "Select Id, ����, ����, ����ʱ�� From ���ű� Where ID In (" & strCon & ") " & strTemp & " Order by ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�������Ҷ�Ӧ", Val(Mid(str����ID, 2)))
    
    If rsTmp.RecordCount = 0 Then
        tvwDept.Visible = False
        tvwMain_S.Top = 0
        tvwMain_S.Height = IIF(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top - picSplit2.Height, 0)
        tvwMain_S.Left = 0
        Exit Sub
    End If
    
    tvwDept.Visible = True
    tvwDept.Height = 2000
    
    tvwMain_S.Height = tvwMain_S.Height - tvwDept.Height - picSplit2.Height - 50
    tvwDept.Left = tvwMain_S.Left
    tvwDept.Width = picSplit2.Width - 50
    tvwDept.Top = tvwMain_S.Top + tvwMain_S.Height + 50 + picSplit2.Height

    tvwDept.Nodes.Clear
    tvwDept.Nodes.Add , , "Root", tvwMain_S.SelectedItem.Text, "Root", "Root"
    tvwDept.Nodes("Root").Expanded = True
    Do Until rsTmp.EOF
        If CDate(IIF(IsNull(rsTmp("����ʱ��")), CDate("3000/1/1"), rsTmp("����ʱ��"))) = CDate("3000/1/1") Then
            strTemp = "Dept"
        Else
            strTemp = "Dept_No"
        End If

        tvwDept.Nodes.Add "Root", tvwChild, "C" & rsTmp("id"), "��" & rsTmp("����") & "��" & rsTmp("����"), strTemp, strTemp
        
        rsTmp.MoveNext
    Loop
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call Form_Resize 'Ϊ��ʹCoolBar����Ӧ�߶�
        FillTree
    End If
    mblnLoad = False
End Sub
Private Sub Form_Load()
    mblnLoad = True
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    SetParent txtFind.hwnd, Toolbar1.hwnd
    SetParent picFind.hwnd, Toolbar1.hwnd
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    Call InitTabControl
    
    Call Ȩ�޿���
    Call CheckHaveDelDept
    '���������ɾ����ListView�������
    lvwMain.Tag = "�ɱ仯��"
    '-----------
    RestoreWinState Me, App.ProductName
    lvw��������_S.Visible = True
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", 0)) = 1)
    mnuViewShowStop.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", 0)) = 1)
    mnuViewShowDel.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾɾ��", 0)) = 1)
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    If lvwMain.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
    End If
    '����LvwMain��ʾ���ö�Ӧ�˵�
     mnuViewIcon_Click lvwMain.View
     lvw��������_S.View = lvwReport
     lbl��������.BackStyle = 0
     
     mlng�������� = Val(zlDatabase.GetPara("��������", glngSys, 0))
    Call InitSystemPara
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    '��ʼ������RIS�ӿ�
    If mblnPACSInterface Then
        Call IniRIS
    End If
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    txtFind.Left = Me.Width - txtFind.Width
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    sngTop = IIF(CoolBar1.Visible, CoolBar1.Top + CoolBar1.Height, 0)
    sngBottom = Me.ScaleHeight - IIF(stbThis.Visible, stbThis.Height, 0)
    
    With Me.tbcDetails
        .Move 0, sngTop, Me.Width / 4, sngBottom - sngTop
        picList.Height = .Height - 400
    End With
    With tvwMain_S
        tvwMain_S.Top = 0
        tvwMain_S.Height = picList.Height
        tvwMain_S.Width = tbcDetails.Width - 60
        tvwMain_S.Left = 0
    End With
    
    If glngSys = 100 Then
        If tvwDept.Visible = True Then
            tvwDept.Height = 2000
            tvwMain_S.Height = tvwMain_S.Height - tvwDept.Height - 50 - picSplit2.Height
            tvwDept.Left = tvwMain_S.Left
            tvwDept.Width = picSplit2.Width - 50
            tvwDept.Top = tvwMain_S.Top + tvwMain_S.Height + 50 + picSplit2.Height
        End If
    End If
    
    picSplit.Top = sngTop
    picSplit.Height = IIF(sngBottom - picSplit.Top > 0, sngBottom - picSplit.Top, 0)
    picSplit.Left = tbcDetails.Left + tbcDetails.Width
    
    tbcDept.Height = sngBottom / 3
    tbcDept.Left = picSplit.Left + picSplit.Width
    tbcDept.Top = sngBottom - tbcDept.Height
    
    If tbcDept.Top < tvwMain_S.Top + 2000 Then tbcDept.Top = tvwMain_S.Top + 2000
    
    picSplitH.Left = tbcDept.Left
    picSplitH.Top = tbcDept.Top - picSplitH.Height
    
    picSplit2.Left = tvwMain_S.Left
    picSplit2.Top = tvwMain_S.Top + tvwMain_S.Height
    picSplit2.Width = picList.Width
    lbl��������.Width = picList.Width
    tvwDept.Width = picList.Width
    
    lvwMain.Left = picSplit.Left + picSplit.Width
    lvwMain.Top = sngTop
    lvwMain.Height = picSplitH.Top - lvwMain.Top
    If Me.ScaleWidth - lvwMain.Left > 0 Then lvwMain.Width = Me.ScaleWidth - lvwMain.Left
    tbcDept.Width = lvwMain.Width
    picSplitH.Width = lvwMain.Width
    
    lvw��������_S.Move 0, 400, tbcDept.Width, tbcDept.Height - 400
    lvw��������_S.ColumnHeaders.Item(3).Width = lvw��������_S.Width - lvw��������_S.ColumnHeaders.Item(1).Width - lvw��������_S.ColumnHeaders.Item(2).Width
    lvwMain.ColumnHeaders.Item(2).Width = 1000
    picSplit2.Visible = tvwDept.Visible
    Me.Refresh
    lvwMain.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrKey = ""
    If Not mobjForm Is Nothing Then Set mobjForm = Nothing
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIF(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", IIF(mnuViewShowStop.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾɾ��", IIF(mnuViewShowDel.Checked, 1, 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lbl��������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And picSplit2.Top + Y > tvwMain_S.Top + 200 And picSplit2.Top + Y < stbThis.Top - 1500 Then
        picSplit2.Top = picSplit2.Top + Y
        tvwMain_S.Height = tvwMain_S.Height + Y
        tvwDept.Move 0, tvwDept.Top + Y, picSplit2.Width - 50, tvwDept.Height - Y
    End If
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwMain.SortOrder = IIF(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwMain_DblClick()
    If mblnItem = True And mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
End Sub

Private Sub lvwMain_GotFocus()
    mint���� = 1
    With lvwMain
        If .ListItems.Count = 0 Or .SelectedItem Is Nothing Then
            lvw��������_S.ListItems.Clear
            Call mobjForm.initVSf(0)
            Call SetMenu
        Else
            Call SetMenu
        End If
        stbThis.Panels(2).Text = "�����б��й���ʾ��" & .ListItems.Count & "�����š�"
    End With
End Sub

Public Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowAttribe Mid(Item.Key, 2)
    Call mobjForm.initVSf(Val(Mid(Item.Key, 2)))
    
    Item.Tag = GetClerk(Item.Key)
    
    mblnItem = True
    Call SetMenu
    stbThis.Panels(2).Text = "�����б��й���ʾ��" & lvwMain.ListItems.Count & "������" & IIF(Item.Tag = "0", "��", "���ò�������Ա" & Item.Tag & "����")
End Sub

Private Sub ShowAttribe(ByVal strKey As String)
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim str������� As String
    
    gstrSQL = "select A.��������,A.�������,B.˵�� from ��������˵�� A,�������ʷ��� B where A.��������=B.���� and A.����ID= [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strKey))
        
    lvw��������_S.ListItems.Clear
    Do Until rsTemp.EOF
        Select Case rsTemp("�������")
             Case 1
                str������� = "���ﲡ��"
             Case 2
                str������� = "סԺ����"
             Case 3
                str������� = "�����סԺ����"
             Case Else
                str������� = "�������ڲ���"
        End Select
        Set lst = lvw��������_S.ListItems.Add(, rsTemp("��������"), rsTemp("��������"))
        If mblnҩ�� = True Then
            lst.SubItems(1) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        Else
            lst.SubItems(1) = str�������
            lst.SubItems(2) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        End If
        
        rsTemp.MoveNext
    Loop
    rsTemp.Close
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mnuEditModify.Enabled And mnuEditModify.Visible Then mnuEditModify_Click
    End If
End Sub
 
Private Sub lvwMain_LostFocus()
    mint���� = 0
End Sub

Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button = 2 Then
        If tbcDetails.Selected.Index = mint����� Then
            mnuShortMenu2(4).Visible = mnuViewShowDel.Checked And InStr(mstrPrivs, ";��ɾ��;") > 0
        
            '��ɾ�����ŷ��಻�������
            If InStr(tvwMain_S.SelectedItem.Text, "��-") > 0 Then
                mnuShortMenu2(1).Enabled = False
                mnuShortMenu2(2).Enabled = False
                mnuShortMenu2(3).Enabled = False
                mnuShortMenu2(4).Enabled = lvwMain.ListItems.Count > 0
            Else
                mnuShortMenu2(1).Enabled = mnuEditNew.Enabled
                mnuShortMenu2(2).Enabled = mnuEditModify.Enabled
                mnuShortMenu2(3).Enabled = mnuEditDelete.Enabled
                mnuShortMenu2(4).Enabled = False
            End If
        Else    '������
            If lvwMain.ListItems.Count = 0 Then
                mnuShortMenu2(1).Enabled = False
                mnuShortMenu2(2).Enabled = False
                mnuShortMenu2(3).Enabled = False
                mnuShortMenu2(4).Enabled = False
            End If
            If lvwMain.ListItems.Count <> 0 Then
                If lvwMain.SelectedItem.Icon = "Dept_No" And InStr(1, lvwMain.SelectedItem.ListSubItems(1).Text, "-") = 0 Then 'ͣ��
                    mnuShortMenu2(1).Enabled = False
                    mnuShortMenu2(2).Enabled = False
                    mnuShortMenu2(3).Enabled = False
                    mnuShortMenu2(4).Enabled = False
                End If
                If lvwMain.SelectedItem.Icon = "Dept_No" And InStr(1, lvwMain.SelectedItem.ListSubItems(1).Text, "-") > 0 Then 'ɾ��
                    mnuShortMenu2(1).Enabled = False
                    mnuShortMenu2(2).Enabled = False
                    mnuShortMenu2(3).Enabled = False
                    mnuShortMenu2(4).Enabled = True
                End If
                If lvwMain.SelectedItem.Icon = "Dept" Then  '����
                    mnuShortMenu2(1).Enabled = True
                    mnuShortMenu2(2).Enabled = True
                    mnuShortMenu2(3).Enabled = True
                    mnuShortMenu2(4).Enabled = False
                End If
            Else
                mnuShortMenu2(1).Enabled = False
                mnuShortMenu2(2).Enabled = False
                mnuShortMenu2(3).Enabled = False
                mnuShortMenu2(4).Enabled = False
            End If
        End If
                        
        For i = 0 To 3
            mnuShortIcon(i).Checked = mnuViewIcon(i).Checked
        Next
        PopupMenu mnuShort2, vbPopupMenuRightButton
    End If
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim intIndex As Long
    Dim strTemp As String
        
    If ActiveControl Is tvwMain_S Then
        If tbcDetails.Selected.Index = mint����� Then
            strTemp = Val(Mid(tvwMain_S.SelectedItem.Key, 2))
        Else
            strTemp = Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
        End If
    
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "�ò��Ż��¼����Ż���δɾ���Ĳ�����Ա������ɾ���ò��š�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & tvwMain_S.SelectedItem.Text & "���Ĳ�����" & vbCrLf & "������¼����ţ�Ҳ��һ��ɾ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            If Check��������(strTemp) = False Then Exit Sub
            If Int(glngSys / 100) = 1 And mblnPACSInterface Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(10, RISBaseItemOper.Delete, strTemp) <> 1 Then
                        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISBasicDictTable)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISBasicDictTable)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            MousePointer = 11
            gstrSQL = "zl_���ű�_DELETE(" & strTemp & "," & mint��ɾ�� & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            MousePointer = 0
            strKey = tvwMain_S.SelectedItem.Key
            If Not tvwMain_S.SelectedItem.Next Is Nothing Then
                tvwMain_S.SelectedItem.Next.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            Else
                tvwMain_S.SelectedItem.Parent.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            End If
            tvwMain_S.Nodes.Remove strKey
            '����ֻ�и����һ�������򶼿��޸�
            Call SetMenu
        End If
    Else
        If tbcDetails.Selected.Index = mint����� Then
            strTemp = Val(Mid(lvwMain.SelectedItem.Key, 2))
        Else
            strTemp = Mid(lvwMain.SelectedItem.Key, 2)
        End If
        If CheckExistDepPres(strTemp) = True Then
            MsgBox "�ò��Ż��¼����Ż���δɾ���Ĳ�����Ա������ɾ���ò��š�", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("��ȷ��Ҫɾ������Ϊ��" & lvwMain.SelectedItem.Text & "���Ĳ�����" & vbCrLf & "������¼����ţ�Ҳ��һ�뱻ɾ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            If Check��������(strTemp) = False Then Exit Sub
            If Int(glngSys / 100) = 1 And mblnPACSInterface Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(10, RISBaseItemOper.Delete, strTemp) <> 1 Then
                        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISBasicDictTable)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISBasicDictTable)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            Me.MousePointer = 11
            gstrSQL = "zl_���ű�_DELETE(" & strTemp & "," & mint��ɾ�� & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Me.MousePointer = 0
            With lvwMain
                '��ɾ��TreeView�ж�Ӧ�ڵ�
'                tvwMain_S.Nodes.Remove .SelectedItem.Key
                '��ɾ��ListView�ж�Ӧ�ڵ�
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                    lvwMain_ItemClick .SelectedItem
                Else
                    Call lvwMain_GotFocus
                End If
            End With
        End If
    End If
    FillTree
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = 0
End Sub

Private Sub mnuEditExtend_Click()
    Dim strKey As String
    Dim strName As String
    
    On Error Resume Next
    
    If ActiveControl Is lvwMain Then
        If tbcDetails.Selected.Index = mint������ Then
            strKey = Mid(lvwMain.SelectedItem.Key, 2, Len(lvwMain.SelectedItem.Key) - 1)
        Else
            strKey = Mid(lvwMain.SelectedItem.Key, 2)
        End If
        strName = lvwMain.SelectedItem.Text
    Else
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            If tbcDetails.Selected.Index = mint������ Then
                strKey = Mid(.Key, InStr(1, .Key, "|") + 1)
            Else
                strKey = Mid(.Key, 2)
            End If
            strName = Mid(.Text, InStr(1, .Text, "��") + 1)
        End With
    End If
    
    Call frmDeptExtend.ShowMe(Me, strKey, strName, 0, 1)
    Call mobjForm.initVSf(Val(strKey), 0)
End Sub

Private Sub mnuEditModify_Click()
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim strKey As String
    Dim rsTemp As ADODB.Recordset
    Dim str�ϼ����� As String
    Dim strTemp As String
    
    On Error Resume Next
    If ActiveControl Is lvwMain Then
'        If tvwMain_S.SelectedItem.Key = "Root" Then
'            Exit Sub
'        End If
        
        If mnuViewShowAll.Checked = True Then
            '���ϼ�������
            If tbcDetails.Selected.Index = mint������ Then
                strKey = Mid(lvwMain.SelectedItem.Key, 2, Len(lvwMain.SelectedItem.Key) - 1)
            Else
                strKey = Mid(lvwMain.SelectedItem.Key, 2)
            End If
            gstrSQL = "select a.�ϼ�id,b.����,b.����  from ���ű� a,���ű� b where a.�ϼ�id=b.id(+)  and a.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ϼ�id", strKey)
            If Not rsTemp Is Nothing Then
                strTemp = rsTemp!�ϼ�id
            Else
                Exit Sub
            End If
            If tvwMain_S.SelectedItem.Key = "Root" Then
                str�ϼ����� = ""
            Else
                str�ϼ����� = tvwMain_S.SelectedItem.Key
            End If
            
            frmDeptSet.�༭���� mstrPrivs, strKey, 2, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), str�ϼ�����, strTemp
        Else
            If tvwMain_S.SelectedItem.Key = "Root" Then
                Call frmDeptSet.�༭����(mstrPrivs, Mid(lvwMain.SelectedItem.Key, 2), 2, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), "")
            Else
                frmDeptSet.�༭���� mstrPrivs, Mid(lvwMain.SelectedItem.Key, 2), 2, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), tvwMain_S.SelectedItem.Parent.Key
            End If
        End If
    Else
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            If tbcDetails.Selected.Index = mint������ Then
                strKey = Mid(.Key, InStr(1, .Key, "|") + 1)
                gstrSQL = "select �ϼ�id from ���ű�  where id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ϼ�id", strKey)
                If Not rsTemp Is Nothing Then
                    strTemp = rsTemp!�ϼ�id
                Else
                    Exit Sub
                End If
            Else
                strKey = Mid(.Key, 2)
                strTemp = Mid(.Parent.Key, 2)
            End If
            If tvwMain_S.SelectedItem.Key = "Root" Then
                str�ϼ����� = ""
            Else
                str�ϼ����� = Mid(.Key, 1, InStr(1, .Key, "|") - 1)
            End If
            
            Call frmDeptSet.�༭����(mstrPrivs, strKey, 2, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), str�ϼ�����, strTemp)
        End With
    End If
End Sub

Private Sub mnuEditNew_Click()
    Dim str���� As String
    Dim str���� As String
    Dim i As Integer
    Dim strKey As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim str�ϼ����� As String
        
    If tbcDetails.Selected.Index = mint������ Then
        If ActiveControl Is tvwMain_S Then
            strKey = Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
        Else
            strKey = Mid(lvwMain.SelectedItem.Key, 2)
        End If
        
        gstrSQL = "Select c.����, a.Id, a.�ϼ�id" & _
                   " From ���ű� A, ��������˵�� B, �������ʷ��� C" & _
                   " Where b.�������� = c.���� And a.Id = b.����id and a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ϼ�id", strKey)
        If Not rsTemp Is Nothing Then
            strTemp = rsTemp!�ϼ�id
            str�ϼ����� = rsTemp!���� & "|" & rsTemp!ID
        End If
    Else
        strTemp = Mid(tvwMain_S.SelectedItem.Key, 2)
    End If

    Call frmDeptSet.�༭����(mstrPrivs, "", 1, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), str�ϼ�����, strTemp)
    Call FillTree
End Sub

Private Sub mnuEditRecovery_Click()
    Dim strKey As String
    Dim str�ϼ����� As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    '�ָ���ɾ������
    If ActiveControl Is tvwMain_S Then
        With tvwMain_S.SelectedItem
            If tbcDetails.Selected.Index = mint������ Then
                strKey = Mid(.Key, InStr(1, .Key, "|") + 1)
                gstrSQL = "select �ϼ�id from ���ű�  where id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ϼ�id", strKey)
                If Not rsTemp Is Nothing Then
                    strTemp = rsTemp!�ϼ�id
                Else
                    Exit Sub
                End If
            Else
                strKey = Mid(.Key, 2)
                strTemp = Mid(.Parent.Key, 2)
            End If
            If tvwMain_S.SelectedItem.Key = "Root" Then
                str�ϼ����� = ""
            Else
                If tbcDetails.Selected.Index = mint������ Then
                    str�ϼ����� = Mid(.Key, 1, InStr(1, .Key, "|") - 1)
                Else
                    str�ϼ����� = ""
                End If
            End If
        End With
        Call frmDeptSet.�༭����(mstrPrivs, strKey, 1, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), str�ϼ�����, strTemp)
    ElseIf ActiveControl Is lvwMain Then
        If tbcDetails.Selected.Index = mint������ Then
            strKey = Val(Mid(lvwMain.SelectedItem.Key, 2, Len(lvwMain.SelectedItem.Key) - 1))
        Else
            strKey = Val(Mid(lvwMain.SelectedItem.Key, 2))
        End If
        gstrSQL = "select a.�ϼ�id,b.����,b.����  from ���ű� a,���ű� b where a.�ϼ�id=b.id(+)  and a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ϼ�id", strKey)
        If Not rsTemp Is Nothing Then
            strTemp = rsTemp!�ϼ�id
        Else
            Exit Sub
        End If
        
        If tvwMain_S.SelectedItem.Key = "Root" Then
            str�ϼ����� = ""
        Else
            str�ϼ����� = tvwMain_S.SelectedItem.Key
        End If
        
        Call frmDeptSet.�༭����(mstrPrivs, strKey, 1, IIF(tbcDetails.Selected.Index = mint�����, 1, 2), str�ϼ�����, strTemp)
    End If
End Sub

Private Sub mnuEditStart_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim j As Integer
    Dim strTemp As String
    Dim str���� As String
            
    If ActiveControl Is tvwMain_S Then
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            If tbcDetails.Selected.Index = mint����� Then
                strKey = .Key
            Else
                strKey = "C" & Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
            End If
            
            gstrSQL = "zl_���ű�_reuse(" & Mid(strKey, 2) & ")"
        End With
    Else
        If tbcDetails.Selected.Index = mint����� Then
            strKey = lvwMain.SelectedItem.Key
        Else
            strKey = "C" & Mid(lvwMain.SelectedItem.Key, 2)
        End If
        gstrSQL = "zl_���ű�_reuse(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�ı�ͼ�����ɫ
    If ActiveControl Is tvwMain_S Then
        tvwMain_S.SelectedItem.Image = "Dept"
        tvwMain_S.SelectedItem.SelectedImage = "Dept"
    Else
        If tbcDetails.Selected.Index = mint����� Then
            tvwMain_S.Nodes(strKey).Image = "Dept"
            tvwMain_S.Nodes(strKey).SelectedImage = "Dept"
        Else
            str���� = "C" & Mid(tvwMain_S.SelectedItem.Key, 2, 1) & "|" & Mid(strKey, 2)
            tvwMain_S.Nodes(str����).Image = "Dept"
            tvwMain_S.Nodes(str����).SelectedImage = "Dept"
        End If
        With lvwMain.SelectedItem
            .Icon = "Dept"
            .SmallIcon = "Dept"
            .ForeColor = RGB(0, 0, 0)
            
            Dim i As Integer
            For i = 1 To lvwMain.ColumnHeaders.Count
                If i < lvwMain.ColumnHeaders.Count Then
                    .ListSubItems(i).ForeColor = RGB(0, 0, 0)
                End If
                '���³���ʱ��
                If lvwMain.ColumnHeaders(i).Text = "����ʱ��" Then
                    .SubItems(i - 1) = "3000-01-01"
                End If
            Next
        End With
    End If
    
    '�����ϼ�Ŀ¼
    If tbcDetails.Selected.Index = mint����� Then  'ֻ���ڰ������ʾ��ҳ���вŴ����ϼ�ͼ��
        If ActiveControl Is tvwMain_S Then
            j = Me.tvwMain_S.SelectedItem.Index
        Else
            j = Me.tvwMain_S.Nodes(lvwMain.SelectedItem.Key).Index
        End If
        
        While Me.tvwMain_S.Nodes(j).Parent.Image = "Dept_No"
            With tvwMain_S.Nodes(j)
                strKey = .Parent.Key
                gstrSQL = "zl_���ű�_reuse(" & Mid(.Parent.Key, 2) & ")"
                'ִ�����ù���
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                '����ͼ��
                .Parent.Image = "Dept"
                .Parent.SelectedImage = "Dept"
                j = .Parent.Index
            End With
        Wend
    End If
    
    '�ı�״̬���Ͳ˵�
    Call SetMenu
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    On Error GoTo ErrHandle
    Dim strKey As String
    Dim strTemp As String

    If ActiveControl Is tvwMain_S Then
        
        If tvwMain_S.SelectedItem.Key = "Root" Then Exit Sub
        If CheckStop = False Or tvwMain_S.SelectedItem.Tag <> "0" Then
            MsgBox "һ������ֻ��û���¼����Ż�������Աʱ���ܱ�ͣ�á�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If tbcDetails.Selected.Index = mint����� Then
            strKey = Val(Mid(tvwMain_S.SelectedItem.Key, 2))
        Else
            strKey = Mid(tvwMain_S.SelectedItem.Key, InStr(1, tvwMain_S.SelectedItem.Key, "|") + 1, Len(tvwMain_S.SelectedItem.Key) - InStr(1, tvwMain_S.SelectedItem.Key, "|"))
        End If
        
        '���ҵ�������
        If CheckBusiness(Val(strKey)) = False Then Exit Sub
        
        If Check��������(Val(strKey)) = False Then Exit Sub
        With tvwMain_S.SelectedItem
            If .Key = "Root" Then Exit Sub
            strKey = strKey
            gstrSQL = "zl_���ű�_stop(" & strKey & ")"
        End With
    Else
        If CheckStop = False Or lvwMain.SelectedItem.Tag <> "0" Then
            MsgBox "һ������ֻ��û���¼����Ż�������Աʱ���ܱ�ͣ�á�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '���ҵ�������
        If CheckBusiness(Val(Mid(lvwMain.SelectedItem.Key, 2))) = False Then Exit Sub
        
        If Check��������(Val(Mid(lvwMain.SelectedItem.Key, 2))) = False Then Exit Sub
        strKey = lvwMain.SelectedItem.Key
        gstrSQL = "zl_���ű�_stop(" & Mid(lvwMain.SelectedItem.Key, 2) & ")"
    End If
    'ִ�����ù���
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�ı�ͼ�����ɫ
    If mnuViewShowStop.Checked = True Then 'Ҫ��ʾͣ�ò���
        If ActiveControl Is tvwMain_S Then
            tvwMain_S.SelectedItem.Image = "Dept_No"
            tvwMain_S.SelectedItem.SelectedImage = "Dept_No"
        Else
            If tbcDetails.Selected.Index = mint����� Then
                tvwMain_S.Nodes(strKey).Image = "Dept_No"
                tvwMain_S.Nodes(strKey).SelectedImage = "Dept_No"
            Else
                tvwMain_S.Nodes(tvwMain_S.SelectedItem.Key).Image = "Dept_No"
                tvwMain_S.Nodes(tvwMain_S.SelectedItem.Key).SelectedImage = "Dept_No"
            End If
            With lvwMain.SelectedItem
                .Icon = "Dept_No"
                .SmallIcon = "Dept_No"
                .ForeColor = RGB(255, 0, 0)
                
                Dim i As Integer
                For i = 1 To lvwMain.ColumnHeaders.Count
                    If i < lvwMain.ColumnHeaders.Count Then
                        .ListSubItems(i).ForeColor = RGB(255, 0, 0)
                    End If
                    '���³���ʱ��
                    If lvwMain.ColumnHeaders(i).Text = "����ʱ��" Then
                        .SubItems(i - 1) = Format(Date, "yyyy-MM-dd")
                    End If
                Next
            End With
        End If
        Call SetMenu
    Else '����ʾͣ�ò���
        If ActiveControl Is tvwMain_S Then
            strKey = tvwMain_S.SelectedItem.Key
            If Not tvwMain_S.SelectedItem.Next Is Nothing Then
                tvwMain_S.SelectedItem.Next.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            Else
                tvwMain_S.SelectedItem.Parent.Selected = True
                tvwMain_S_NodeClick tvwMain_S.SelectedItem
            End If
            tvwMain_S.Nodes.Remove strKey
            '����ֻ�и����һ�������򶼿��޸�
            Call SetMenu
        Else
            With lvwMain
                tvwMain_S.Nodes.Remove .SelectedItem.Key
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    .ListItems(1).Selected = True
                    .ListItems(1).EnsureVisible
                    lvwMain_ItemClick .SelectedItem
                Else
                    Call lvwMain_GotFocus
                End If
            End With
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditExpand_Click()
    Dim strTemp As String
    Dim str������ As String
    Dim str���� As String
    Dim intNew As Integer 'Ŀǰ���

    On Error GoTo ErrHandle
    With tvwMain_S.SelectedItem
        If .Key = "Root" Then
            str������ = ""
            intNew = GetDownCodeLength("", "���ű�")
        Else
            str������ = Mid(.Text, 2, InStr(.Text, "��") - 2)
            intNew = GetDownCodeLength(Mid(.Key, 2), "���ű�")
        End If
        If intNew = 10 Then
            MsgBox "�����ټӳ����룬ĳһ���¼��Ѿ������˳��ȡ�", vbExclamation, gstrSysName
            Exit Sub
        End If
        str���� = Mid(.Child.Text, 2, InStr(.Child.Text, "��") - 2)
        intNew = frmCodingL.GetLength(Len(str����), 10 - (intNew - Len(str����)), .Text)
        If intNew = 0 Then Exit Sub
        strTemp = str������ & String(intNew - Len(str����), "0")
        If .Key = "Root" Then
            gstrSQL = "zl_���ű�_EXPAND('" & strTemp & "'," & Len(str������) + 1 & ",0)"
        Else
            gstrSQL = "zl_���ű�_EXPAND('" & strTemp & "'," & Len(str������) + 1 & "," & Mid(.Key, 2) & ")"
        End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        FillTree
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileParameter_Click()
    frmSetParameter.ShowMe Me
End Sub

Private Sub mnuFind_Click()
    frmPresFind.ShowOfType Me, 1, mnuViewShowStop.Checked, mnuViewShowDel.Checked, IIF(tbcDetails.Selected.Index = mint�����, 1, 2)
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id������=����id
    Dim lng����id As Long
    Dim lng����ID As Long
    
    If Not tvwMain_S.SelectedItem Is Nothing Then
        If tvwMain_S.SelectedItem.Key <> "Root" Then
            lng����id = Mid(tvwMain_S.SelectedItem.Key, 2)
        End If
    End If
    
    If Not lvwMain.SelectedItem Is Nothing Then
        lng����ID = Mid(lvwMain.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIF(lng����id = 0, "", lng����id), _
        "����=" & IIF(lng����ID = 0, "", lng����ID))
End Sub

'Private Sub mnuViewFind_Click()
'    frmDeptCharacter.��ʾ����
'End Sub

Private Sub mnuViewReflash_Click()
    FillTree
End Sub

Private Sub mnuViewSelect_Click()
    If zlControl.LvwSelectColumns(lvwMain, mstrLvw) = True Then
        '���б仯��Ҫ����ˢ��
        FillList tvwMain_S.SelectedItem.Key
    End If
End Sub

Private Sub mnuViewShowAll_Click()
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    FillList tvwMain_S.SelectedItem.Key
End Sub

Private Sub mnuViewShowDel_Click()
    mnuViewShowDel.Checked = Not mnuViewShowDel.Checked
    FillTree
End Sub

Private Sub mnuViewShowStop_Click()
    mnuViewShowStop.Checked = Not mnuViewShowStop.Checked
    FillTree
End Sub

'Private Sub opt���_Click()
'    Call FillTree
'End Sub

'Private Sub opt����_Click()
'    Call FillTree
'End Sub

Private Sub picsplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartX = X
    End If
End Sub

Private Sub picsplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplit.Left + X - sngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + picSplit.Width) > 500 Then
            picSplit.Left = sngTemp
            tbcDetails.Width = picSplit.Left - tvwMain_S.Left
            tvwMain_S.Width = tbcDetails.Width
            lvwMain.Left = picSplit.Left + picSplit.Width
            lvwMain.Width = Me.ScaleWidth - lvwMain.Left
            picSplit2.Width = tvwMain_S.Width
            
            If glngSys = 100 Then
                If tvwDept.Visible = True Then
                    tvwDept.Width = tvwMain_S.Width
                End If
            End If
            
            picSplitH.Left = lvwMain.Left
            tbcDept.Left = lvwMain.Left
            picSplitH.Width = lvwMain.Width
            tbcDept.Width = lvwMain.Width
            lvw��������_S.Left = 0
            lvw��������_S.Width = tbcDept.Width
        End If
        tvwMain_S.SetFocus
    End If
End Sub
'
Private Sub picSplit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And picSplit2.Top + Y > tvwMain_S.Top + 200 And picSplit2.Top + Y < stbThis.Top - 1500 Then
        picSplit2.Top = picSplit2.Top + Y
        tvwMain_S.Height = tvwMain_S.Height + Y
        tvwDept.Move 0, tvwDept.Top + Y, picSplit2.Width - 50, tvwDept.Height - Y
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        sngStartY = Y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    On Error Resume Next

    If Button = 1 Then
        sngTemp = picSplitH.Top + Y - sngStartY
        If sngTemp - lvwMain.Top > 2500 And IIF(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - (sngTemp + picSplitH.Height) > 1200 Then
            picSplitH.Top = sngTemp
            lvwMain.Height = picSplitH.Top - tvwMain_S.Top - 800
            tbcDept.Top = picSplitH.Top + picSplitH.Height
            tbcDept.Height = IIF(stbThis.Visible = True, stbThis.Top, Me.ScaleHeight) - tbcDept.Top
            lvw��������_S.Top = 400
            lvw��������_S.Height = tbcDept.Height - 400
        End If
        lvwMain.SetFocus
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufilePreview_Click()
    subPrint 2
End Sub


Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnufileset_Click()
    zlPrintSet
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Integer
    
    Call FillTree
    If InStr(tvwMain_S.SelectedItem.Text, "��������") > 0 And tbcDetails.Selected.Index = mint������ Then
        mnuEdit.Enabled = False
        mnuEditNew.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditStart.Enabled = False
        mnuEditStop.Enabled = False
        mnuEditRecovery.Enabled = False
        Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
        Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
        Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
        Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
        Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        Exit Sub
    End If
    If mblnLoad = False Then
        lvwMain.SetFocus
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuEditNew_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Start"
            mnuEditStart_Click
        Case "Stop"
            mnuEditStop_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnufilePreview_Click
        Case "Help"
            mnuhelptopic_Click
        Case "View"
            mnuViewIcon(lvwMain.View).Checked = False
            If lvwMain.View = 3 Then
                mnuViewIcon(0).Checked = True
                lvwMain.View = 0
            Else
                mnuViewIcon(lvwMain.View + 1).Checked = True
                lvwMain.View = lvwMain.View + 1
            End If
        Case "Find"
            mnuFind_Click
    End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(ButtonMenu.Index).Text, "  ", "��")
    lvwMain.View = ButtonMenu.Index - 1
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    CoolBar1.Visible = mnuViewToolButton.Checked
    CoolBar1.Bands("only").MinHeight = Toolbar1.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button

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

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
        Toolbar1.Buttons("View").ButtonMenus(i + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(i + 1).Text, "��", "  ")
    Next
    mnuViewIcon(Index).Checked = True
    Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text = Replace(Toolbar1.Buttons("View").ButtonMenus(Index + 1).Text, "  ", "��")
    lvwMain.View = Index
End Sub


Private Sub mnuShortMenu1_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
        Case 4
            mnuEditRecovery_Click
    End Select

End Sub

Private Sub mnuShortMenu2_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuEditNew_Click
        Case 2
            mnuEditModify_Click
        Case 3
            mnuEditDelete_Click
        Case 4
            mnuEditRecovery_Click
    End Select

End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuhelptopic_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

Private Sub tvwDept_NodeClick(ByVal Node As MSComctlLib.Node)
    '''''
End Sub


Private Sub tvwMain_S_GotFocus()
    mint���� = 2
    If tbcDetails.Selected.Index = mint����� Then
        stbThis.Panels(2).Text = "��������" & tvwMain_S.SelectedItem.Children & "��ֱ������" & IIF(Val(tvwMain_S.SelectedItem.Tag) = 0, "��", "����Ա" & tvwMain_S.SelectedItem.Tag & "����")
    Else
        If tvwMain_S.SelectedItem.Text = "��������" Then
            stbThis.Panels(2).Text = "��������" & tvwMain_S.SelectedItem.Children & "��ֱ������"
        ElseIf InStr(1, tvwMain_S.SelectedItem.Text, "��") = 0 Then
            stbThis.Panels(2).Text = "��������" & tvwMain_S.SelectedItem.Children & "��ֱ������" & IIF(Val(tvwMain_S.SelectedItem.Tag) = 0, "��", "����Ա" & tvwMain_S.SelectedItem.Tag & "����")
        ElseIf InStr(1, tvwMain_S.SelectedItem.Text, "��") > 0 Then
            stbThis.Panels(2).Text = "��������" & tvwMain_S.SelectedItem.Children & "��ֱ������" & IIF(Val(tvwMain_S.SelectedItem.Tag) = 0, "��", "����Ա" & tvwMain_S.SelectedItem.Tag & "����")
        End If
    End If
    Call SetMenu
End Sub

Private Sub tvwMain_S_LostFocus()
    mint���� = 0
End Sub

Private Sub tvwMain_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If tbcDetails.Selected.Index = mint����� Then
            If mnuShortMenu1(1).Visible = False Then Exit Sub
            
            mnuShortMenu1(4).Visible = mnuViewShowDel.Checked
            
            '��ɾ�����ŷ��಻�������
            If InStr(tvwMain_S.SelectedItem.Text, "��-") > 0 Then
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = Trim(tvwMain_S.SelectedItem.Text) <> "��-����ɾ������"
            Else
                mnuShortMenu1(1).Enabled = mnuEditNew.Enabled
                mnuShortMenu1(2).Enabled = mnuEditModify.Enabled
                mnuShortMenu1(3).Enabled = mnuEditDelete.Enabled
                mnuShortMenu1(4).Enabled = False
            End If
        Else    '������
            If InStr(1, tvwMain_S.SelectedItem.Text, "��") = 0 Then '����
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = False
            End If
            If InStr(1, tvwMain_S.SelectedItem.Text, "��-") Then    'ɾ��
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = True
            End If
            If InStr(1, tvwMain_S.SelectedItem.Text, "��") > 0 And InStr(1, tvwMain_S.SelectedItem.Text, "��-") = 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then 'ͣ��
                mnuShortMenu1(1).Enabled = False
                mnuShortMenu1(2).Enabled = False
                mnuShortMenu1(3).Enabled = False
                mnuShortMenu1(4).Enabled = False
            End If
            If InStr(1, tvwMain_S.SelectedItem.Text, "��") > 0 And InStr(1, tvwMain_S.SelectedItem.Text, "��-") = 0 And tvwMain_S.SelectedItem.Image = "Dept" Then  '����
                mnuShortMenu1(1).Enabled = True
                mnuShortMenu1(2).Enabled = True
                mnuShortMenu1(3).Enabled = True
                mnuShortMenu1(4).Enabled = False
            End If
        End If
        PopupMenu mnuShort1, vbPopupMenuRightButton
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strTemp As String
        
    If tbcDetails.Selected.Index = mint������ Then
        If InStr(1, tvwMain_S.SelectedItem.Text, "��������") > 0 Then
            mnuEdit.Enabled = False
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        
        If InStr(tvwMain_S.SelectedItem.Text, "��") = 0 Then    '����
            mnuEdit.Enabled = False
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "��-") > 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then 'ɾ��
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = True
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "��-") = 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then 'ͣ��
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = True
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "��") > 0 And tvwMain_S.SelectedItem.Image = "Dept" Then '����
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = True
            mnuEditModify.Enabled = True
            mnuEditDelete.Enabled = True
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = True
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        mnuEditExpand.Enabled = False
        strTemp = Mid(Node.Key, InStr(1, Node.Key, "|") + 1, Len(Node.Key) - InStr(1, Node.Key, "|"))
        strTemp = "C" & strTemp
    Else
        strTemp = Node.Key
        mnuEdit.Enabled = True
    End If
    
'    If mstrKey = strTemp Then Exit Sub
    
    mstrKey = strTemp
    
    Node.Tag = GetClerk(strTemp)
    
    FillList strTemp
    
    If InStr(tvwMain_S.SelectedItem.Text, "��") > 0 And tvwMain_S.SelectedItem.Image = "Dept" Then
        If tbcDetails.Selected.Index = mint������ Then
            Call ShowAttribe(Mid(Node.Key, 4))
            Call mobjForm.initVSf(Val(Mid(Node.Key, 4)))
        Else
            Call ShowAttribe(Mid(Node.Key, 2))
            Call mobjForm.initVSf(Val(Mid(Node.Key, 2)))
        End If
    End If
    
    If glngSys = 100 Then
        Show�������Ҷ�Ӧ (strTemp)
    End If
    picSplit2.Visible = tvwDept.Visible
    
    tvwMain_S_GotFocus
    
    If picSplit2.Visible = True Then
        picSplit2.Left = tvwMain_S.Left
        picSplit2.Top = tvwMain_S.Top + tvwMain_S.Height
        picSplit2.Width = tvwMain_S.Width
    End If
    If tvwDept.Visible = False Then
        tvwMain_S.Height = picList.Height
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "���ű�"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(Sys.Currentdate, "yyyy��MM��dd��")
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

Public Sub FillTree()
'����:װ�����в��ŵ�tvwMain_S
'����:
    Dim strTemp As String
    Dim strKey As String
    Dim rs���� As New ADODB.Recordset
    Dim rs���ʲ��� As ADODB.Recordset
    Dim str���� As String
    Dim i As Integer
    Dim rs�������� As ADODB.Recordset
    Dim str���� As String
    Dim nod As Node
    Dim strͼ�� As String
    Dim strɾ�� As String
    
    mstrKey = ""
    On Error GoTo ErrHandle
    rs����.CursorLocation = adUseClient
    rs����.CursorType = adOpenKeyset
    rs����.LockType = adLockReadOnly
    
    If tbcDetails.ItemCount = 0 Then Exit Sub
    If tbcDetails.Selected.Index = mint����� Then        '�������ʾ
        If Not tvwMain_S.SelectedItem Is Nothing Then
            strKey = tvwMain_S.SelectedItem.Key
        End If
    
        If mnuViewShowStop.Checked = False Then
            strTemp = " where (����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or ����ʱ�� is null ) "
        End If
       
        gstrSQL = "select id,�ϼ�id,���� ,����,to_char(����ʱ��,'YYYY-MM-DD') as ����ʱ��  from ���ű� " & strTemp & " " & _
                " start with �ϼ�id is null And ���� <> '-' " & _
                " connect by prior id =�ϼ�id"
        Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        tvwMain_S.Nodes.Clear
        tvwMain_S.Nodes.Add , , "Root", "���в���", "Root", "Root"
        tvwMain_S.Nodes("Root").Sorted = True
        
        Do Until rs����.EOF
            If CDate(IIF(IsNull(rs����("����ʱ��")), CDate("3000/1/1"), rs����("����ʱ��"))) = CDate("3000/1/1") Then
                strTemp = "Dept"
            Else
                strTemp = "Dept_No"
            End If
            
            If IsNull(rs����("�ϼ�id")) Then
                tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
            Else
                tvwMain_S.Nodes.Add "C" & rs����("�ϼ�id"), tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
            End If
            tvwMain_S.Nodes("C" & rs����("id")).Sorted = True
                    
            rs����.MoveNext
        Loop
        
        '"��ɾ������"����
        If mnuViewShowDel.Checked = True Then
            gstrSQL = " Select ID, �ϼ�id, ����, ����, To_Char(����ʱ��, 'YYYY-MM-DD') As ����ʱ�� " & _
                    " From ���ű� Start With ���� = '-' and �ϼ�id is null Connect By Prior ID = �ϼ�id"
                    
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    '        mint��ɾ�� = 0
            strTemp = "Dept_No"
            
            If Not rs����.EOF Then
                tvwMain_S.Nodes("Root").Sorted = False
    '            mint��ɾ�� = 1
            End If
            
            Do Until rs����.EOF
                If IsNull(rs����("�ϼ�id")) Then
                    tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
                Else
                    tvwMain_S.Nodes.Add "C" & rs����("�ϼ�id"), tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), strTemp, strTemp
                End If
                tvwMain_S.Nodes("C" & rs����("id")).Sorted = True
                
                '��ɾ���Ĳ����ú�ɫ���
                tvwMain_S.Nodes("C" & rs����("id")).ForeColor = &HFF&
                
                rs����.MoveNext
            Loop
        End If
    Else    '��������ʾ
        If Not tvwMain_S.SelectedItem Is Nothing Then
            strKey = tvwMain_S.SelectedItem.Key
        End If
        
        gstrSQL = "select distinct a.����,a.���� from �������ʷ��� a,��������˵�� c where a.����=c.��������"
        Set rs���ʲ��� = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʲ�ѯ����")
        
        tvwMain_S.Nodes.Clear
        tvwMain_S.Nodes.Add , , "Root", "��������", "Root", "Root"
        tvwMain_S.Nodes("Root").Sorted = True
        
        str���� = ""
        Do While Not rs���ʲ���.EOF
            tvwMain_S.Nodes.Add "Root", tvwChild, "C" & rs���ʲ���!����, rs���ʲ���!����, "Dept"
            str���� = str���� & rs���ʲ���!���� & "|"
            rs���ʲ���.MoveNext
        Loop
        
        If mnuViewShowStop.Checked = False Then
            strTemp = " and (a.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or a.����ʱ�� is null ) "
        End If
        
        strɾ�� = " and a.id not in(select id from ���ű� where ���� like '-%')"
        
        For i = 0 To UBound(Split(str����, "|"))
            gstrSQL = "Select a.id,a.�ϼ�id,a.����,a.���� as ���ű���,c.����,b.��������,a.����ʱ�� From ���ű� A, ��������˵�� B,�������ʷ��� c Where b.��������=c.���� " & strTemp _
                & " and A.ID=B.����ID and B.��������=[1]" & strɾ��
            str���� = Split(str����, "|")(i)
            Set rs�������� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
            Do While Not rs��������.EOF
                If CDate(IIF(IsNull(rs��������("����ʱ��")), CDate("3000/1/1"), rs��������("����ʱ��"))) = CDate("3000/1/1") Then
                    strͼ�� = "Dept"
                Else
                    strͼ�� = "Dept_No"
                End If
                If rs��������!���� Like "-*" Then
                    strͼ�� = "Dept_No"
                End If
            
                tvwMain_S.Nodes.Add "C" & rs��������!����, tvwChild, "C" & rs��������!���� & "|" & rs��������!ID, "��" & rs��������!���ű��� & "��" & rs��������!����, strͼ��
                rs��������.MoveNext
            Loop
        Next
        
        If mnuViewShowDel.Checked = True Then
            strɾ�� = "  and a.���� like '-%'"
            For i = 0 To UBound(Split(str����, "|"))
                gstrSQL = "Select a.id,a.�ϼ�id,a.����,a.���� as ���ű���,c.����,b.��������,a.����ʱ�� From ���ű� A, ��������˵�� B,�������ʷ��� c Where b.��������=c.���� " & _
                    " and A.ID=B.����ID and B.��������=[1]" & strɾ��
                str���� = Split(str����, "|")(i)
                Set rs�������� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
                Do While Not rs��������.EOF
                    If CDate(IIF(IsNull(rs��������("����ʱ��")), CDate("3000/1/1"), rs��������("����ʱ��"))) = CDate("3000/1/1") Then
                        strͼ�� = "Dept"
                    Else
                        strͼ�� = "Dept_No"
                    End If
                    If rs��������!���� Like "-*" Then
                        strͼ�� = "Dept_No"
                    End If
                
                    tvwMain_S.Nodes.Add "C" & rs��������!����, tvwChild, "C" & rs��������!���� & "|" & rs��������!ID, "��" & rs��������!���ű��� & "��" & rs��������!����, strͼ��
                    rs��������.MoveNext
                Loop
            Next
        End If
    End If
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes("Root")
        nod.Selected = True
        nod.Expanded = True
        tvwMain_S_NodeClick nod
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
        tvwMain_S_NodeClick nod
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FillList(ByVal str����ID As String)
'����:װ���Ӧ���ŵĲ��ŵ�lvwMain
'����:str����ID ���ŵı�ʶ

    Dim rs���� As New ADODB.Recordset
    Dim lst As ListItem
    Dim strKey As String
    Dim strͣ�� As String
    Dim strɾ�� As String
    
    If tbcDetails.Selected.Index = mint����� Then
        If Not lvwMain.SelectedItem Is Nothing Then
            '����ԭ�м�ֵ
            strKey = lvwMain.SelectedItem.Key
        End If
        
        rs����.CursorLocation = adUseClient
        
        If mnuViewShowStop.Checked = False And InStr(1, tvwMain_S.SelectedItem.Text, "��-") = 0 Then
            strͣ�� = " (A.����ʱ�� is null or A.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD')"
            If mnuViewShowDel.Checked = True Then
                strͣ�� = strͣ�� & " Or A.���� Like '-%'"
            End If
            strͣ�� = strͣ�� & ")"
        End If
        If mnuViewShowAll.Checked = True Then
            gstrSQL = "select A.*,B.���� as �ϼ����� from " & _
                "(select A.ID,A.�ϼ�ID,A.����,A.����,A.����,A.λ��,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ��,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ�� " & _
                " from ���ű� A " & IIF(strͣ�� = "", "", "where " & strͣ��) & " connect by prior A.id=A.�ϼ�id start with " & IIF(mnuViewShowDel.Checked = False, "���� <> '-' And ", "") & " " & IIF(str����ID = "Root", "A.�ϼ�ID is null ", "A.�ϼ�ID = [1]") & ") A,���ű� B where A.�ϼ�ID=B.ID(+)"
        Else
            gstrSQL = "select A.ID,A.�ϼ�ID,A.����,A.����,A.����,A.λ��,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ��,to_char(A.����ʱ��,'YYYY-MM-DD') as ����ʱ��,B.���� as �ϼ����� from ���ű� A,���ű� B where A.�ϼ�ID=B.ID(+) and " & IIF(strͣ�� = "", "", strͣ�� & " and ") & IIF(str����ID = "Root", "A.�ϼ�ID is null ", "A.�ϼ�ID = [1]")
            If mnuViewShowDel.Checked = False Then
                gstrSQL = gstrSQL & " And A.���� Not Like '-%'"
            End If
        End If
            
        Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(str����ID, 2)))
    Else
        If mnuViewShowDel.Checked = False Then
            gstrSQL = "Select a.Id, a.�ϼ�id, a.����, a.����, a.����,a.λ��, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��," & _
                              " c.���� as �ϼ�����" & _
                       " From ���ű� A, ��������˵�� B, ���ű� C" & _
                       " Where a.Id = b.����id And a.�ϼ�id = c.Id(+) And b.�������� = [1] and a.id not in(select id from ���ű� where ���� like '-%')"
        Else
            gstrSQL = "Select a.Id, a.�ϼ�id, a.����, a.����, a.����,a.λ��, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��," & _
                              " c.���� as �ϼ�����" & _
                       " From ���ű� A, ��������˵�� B, ���ű� C" & _
                       " Where a.Id = b.����id And a.�ϼ�id = c.Id(+) And b.�������� = [1] "
        End If
        If mnuViewShowStop.Checked = False Then
            strͣ�� = " and ((A.����ʱ�� is null or A.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD'))"
            If mnuViewShowDel.Checked = True Then
                strͣ�� = strͣ�� & " or a.���� like '-%'" & ")"
            Else
                strͣ�� = strͣ�� & ")"
            End If
            gstrSQL = gstrSQL & strͣ��
        End If
        Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�¼�", tvwMain_S.SelectedItem.Text)
    End If
    
    lvwMain.ListItems.Clear
    lvw��������_S.ListItems.Clear
    Call mobjForm.initVSf(0)
    
    Do Until rs����.EOF
        If CDate(IIF(IsNull(rs����("����ʱ��")), CDate("3000/1/1"), rs����("����ʱ��"))) = CDate("3000/1/1") And rs����("����") <> "-" Then
            strͣ�� = "Dept"
        Else
            strͣ�� = "Dept_No"
        End If
        Set lst = lvwMain.ListItems.Add(, "C" & rs����("ID"), rs����("����"), strͣ��, strͣ��)
        If strͣ�� = "Dept_No" Then lst.ForeColor = RGB(255, 0, 0)
        
        Dim lngCol  As Long
        Dim varValue As Variant
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwMain.ColumnHeaders.Count
            varValue = rs����(lvwMain.ColumnHeaders(lngCol).Text).value
            lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
            If strͣ�� = "Dept_No" Then lst.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rs����.MoveNext
    Loop
    If lvwMain.ListItems.Count > 0 Then
        Dim Item As ListItem
        On Error Resume Next
        Set Item = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            Set Item = lvwMain.ListItems(1)
            Item.Selected = True
            Item.EnsureVisible
        Else
            Item.Selected = True
            Item.EnsureVisible
        End If
        EnablePrint True
    Else
        EnablePrint False
    End If
End Sub

Private Sub SetMenu()
'����:�����޸ĺ�ɾ����ť����Чֵ
'����:blnEnabled ��Чֵ
'����:�������Ӱ�ť����Чֵ
    Dim blnEnabled As Boolean
    
    If tvwMain_S.SelectedItem.Image = "Dept_No" Then
        Toolbar1.Buttons("New").Enabled = False
        mnuEditNew.Enabled = False
    Else
        Toolbar1.Buttons("New").Enabled = True
        mnuEditNew.Enabled = True
    End If
    
    mnuEditRecovery.Enabled = False
    mnuEditRecovery.Visible = mnuViewShowDel.Checked
    
    '�Ƿ���ڵ�
    If tbcDetails.Selected.Index = mint����� And tvwMain_S Is ActiveControl And tvwMain_S.SelectedItem.Key = "Root" Or _
        Not (tvwMain_S Is ActiveControl) And lvwMain.ListItems.Count = 0 Then
        Toolbar1.Buttons("Modify").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        Toolbar1.Buttons("Start").Enabled = False
        Toolbar1.Buttons("Stop").Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditExtend.Enabled = False
        mnuEditStart.Enabled = False
        mnuEditStop.Enabled = False
        
        '��ɾ�����ŷ��಻�������
        If InStr(tvwMain_S.SelectedItem.Text, "��-") > 0 Then
            If mnuViewShowDel.Checked Then
                mnuEditRecovery.Enabled = True
                mnuEditNew.Enabled = False
                mnuEditModify.Enabled = False
                mnuEditExtend.Enabled = False
                mnuEditDelete.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
            Else
                mnuEdit.Enabled = False
            End If
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
        End If
        mnuEditExpand.Enabled = False
        Exit Sub
    ElseIf tbcDetails.Selected.Index = mint������ And tvwMain_S Is ActiveControl And tvwMain_S.SelectedItem.Key = "Root" Or _
        Not (tvwMain_S Is ActiveControl) And lvwMain.ListItems.Count = 0 Then
            mnuEdit.Enabled = False
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
    End If
    If tvwMain_S.SelectedItem Is Nothing Then
        mnuEditExpand.Enabled = False
    Else
        mnuEditExpand.Enabled = tvwMain_S.SelectedItem.Children <> 0
    End If
    
    If tvwMain_S Is ActiveControl Then
        blnEnabled = (tvwMain_S.SelectedItem.Image = "Dept")
    Else
        blnEnabled = (lvwMain.SelectedItem.Icon = "Dept")
    End If
    
    Toolbar1.Buttons("Modify").Enabled = blnEnabled
    Toolbar1.Buttons("Delete").Enabled = blnEnabled
    mnuEditDelete.Enabled = blnEnabled
    mnuEditModify.Enabled = blnEnabled
    mnuEditExtend.Enabled = blnEnabled
    Toolbar1.Buttons("Start").Enabled = Not blnEnabled
    Toolbar1.Buttons("Stop").Enabled = blnEnabled
    mnuEditStart.Enabled = Not blnEnabled
    mnuEditStop.Enabled = blnEnabled
                    
    '��ɾ�����ŷ��಻�������
    If InStr(tvwMain_S.SelectedItem.Text, "��-") > 0 Then
        If mnuViewShowDel.Checked Then
            If UCase(Me.ActiveControl.Name) = "TVWMAIN_S" Then
                mnuEditRecovery.Enabled = InStr(tvwMain_S.SelectedItem.Text, "��-����ɾ������") = 0
            Else
                mnuEditRecovery.Enabled = True
            End If
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
        Else
            mnuEdit.Enabled = False
        End If
        Toolbar1.Buttons("New").Enabled = False
        Toolbar1.Buttons("Modify").Enabled = False
        Toolbar1.Buttons("Delete").Enabled = False
        Toolbar1.Buttons("Start").Enabled = False
        Toolbar1.Buttons("Stop").Enabled = False
    End If
                    
    If tbcDetails.Selected.Index = mint������ And ActiveControl Is lvwMain Then
        With lvwMain
            If .ListItems.Count = 0 Then
                mnuEdit.Enabled = False
            Else
                If .SelectedItem.Icon = "Dept_No" And InStr(1, .SelectedItem.ListSubItems(1).Text, "-") > 0 Then 'ɾ��
                    mnuEdit.Enabled = True
                    mnuEditNew.Enabled = False
                    mnuEditModify.Enabled = False
                    mnuEditExtend.Enabled = False
                    mnuEditDelete.Enabled = False
                    mnuEditStart.Enabled = False
                    mnuEditStop.Enabled = False
                    mnuEditRecovery.Enabled = True
                    Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
                    Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
                    Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
                    Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
                    Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
                End If
                If .SelectedItem.Icon = "Dept_No" And InStr(1, .SelectedItem.ListSubItems(1).Text, "-") = 0 Then 'ͣ��
                    mnuEdit.Enabled = True
                    mnuEditNew.Enabled = False
                    mnuEditModify.Enabled = False
                    mnuEditExtend.Enabled = False
                    mnuEditDelete.Enabled = False
                    mnuEditStart.Enabled = True
                    mnuEditStop.Enabled = False
                    mnuEditRecovery.Enabled = False
                    Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
                    Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
                    Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
                    Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
                    Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
                End If
                If .SelectedItem.Icon = "Dept" Then  '����
                    mnuEdit.Enabled = True
                    mnuEditNew.Enabled = True
                    mnuEditModify.Enabled = True
                    mnuEditExtend.Enabled = True
                    mnuEditDelete.Enabled = True
                    mnuEditStart.Enabled = False
                    mnuEditStop.Enabled = True
                    mnuEditRecovery.Enabled = False
                    Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
                    Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
                    Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
                    Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
                    Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
                End If
            End If
        End With
    ElseIf tbcDetails.Selected.Index = mint������ And ActiveControl Is tvwMain_S Then
        
        If InStr(tvwMain_S.SelectedItem.Text, "��") = 0 Then    '����
            mnuEdit.Enabled = False
            Toolbar1.Buttons("New").Enabled = False
            Toolbar1.Buttons("Modify").Enabled = False
            Toolbar1.Buttons("Delete").Enabled = False
            Toolbar1.Buttons("Start").Enabled = False
            Toolbar1.Buttons("Stop").Enabled = False
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "��-") > 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then 'ɾ��
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = True
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "��-") = 0 And tvwMain_S.SelectedItem.Image = "Dept_No" Then 'ͣ��
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = True
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(tvwMain_S.SelectedItem.Text, "��") > 0 And tvwMain_S.SelectedItem.Image = "Dept" Then '����
            mnuEdit.Enabled = True
            mnuEditNew.Enabled = True
            mnuEditModify.Enabled = True
            mnuEditExtend.Enabled = True
            mnuEditDelete.Enabled = True
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = True
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        If InStr(1, tvwMain_S.SelectedItem.Text, "��������") > 0 Then
            mnuEdit.Enabled = False
            mnuEditNew.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditExtend.Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            mnuEditRecovery.Enabled = False
            Toolbar1.Buttons("New").Enabled = mnuEditNew.Enabled
            Toolbar1.Buttons("Modify").Enabled = mnuEditModify.Enabled
            Toolbar1.Buttons("Delete").Enabled = mnuEditDelete.Enabled
            Toolbar1.Buttons("Start").Enabled = mnuEditStart.Enabled
            Toolbar1.Buttons("Stop").Enabled = mnuEditStop.Enabled
        End If
        mnuEditExpand.Enabled = False
    End If
    EnablePrint lvwMain.ListItems.Count > 0
End Sub

Private Sub EnablePrint(ByVal blnEnabled As Boolean)
'����:���ô�ӡ��Ԥ����ť����Чֵ
'����:blnEnabled ��Чֵ
    Toolbar1.Buttons("Print").Enabled = blnEnabled
    Toolbar1.Buttons("Preview").Enabled = blnEnabled
    mnuFilePreview.Enabled = blnEnabled
    mnuFilePrint.Enabled = blnEnabled
    mnuFileExcel.Enabled = blnEnabled
End Sub

Private Sub Ȩ�޿���()
'����:�����е��û�Ȩ�޲���,��ʹһЩ�˵����ť���ɼ�

    If InStr(mstrPrivs, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Visible = False
        mnuShortMenu1(1).Visible = False
        mnuShortMenu2(1).Visible = False
        mnuShortMenu2(2).Visible = False
        mnuShortMenu2(3).Visible = False
        mnuShortMenu2(4).Visible = False
        mnuShortsplit1.Visible = -False
        Toolbar1.Buttons("Split").Visible = False
        Toolbar1.Buttons("New").Visible = False
        Toolbar1.Buttons("Modify").Visible = False
        Toolbar1.Buttons("Delete").Visible = False
        Toolbar1.Buttons("Split1").Visible = False
        Toolbar1.Buttons("Start").Visible = False
        Toolbar1.Buttons("Stop").Visible = False
    End If
    
    If InStr(mstrPrivs, ";��չ��Ϣά��;") = 0 Then
        mnuEditExtend.Visible = False
    End If
    
    mblnҩ�� = (glngSys \ 100 = 8)
    If mblnҩ�� = True Then
        '����ʾ�������
        lvw��������_S.ColumnHeaders.Remove 2
    End If
End Sub

Private Function GetClerk(ByVal strKey As String) As Long
    On Error GoTo errClerk
    Dim rsTemp As New ADODB.Recordset
    If strKey = "Root" Then Exit Function
    
    gstrSQL = "Select Count(b.Id) As ��Ա��" & vbNewLine & _
            "From ������Ա A, ��Ա�� B," & vbNewLine & _
            "     (Select ID" & vbNewLine & _
            "       From ���ű�" & vbNewLine & _
            "       Where ID = [1] Or ID In (Select ID From ���ű� Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id)) C" & vbNewLine & _
            "Where a.����id = c.Id And a.��Աid = b.Id And (b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') or ����ʱ�� is null)"


    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(strKey, 2)))
        
    GetClerk = rsTemp("��Ա��")
    Exit Function
errClerk:
    GetClerk = 0
End Function

Private Function CheckStop() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�޸���         ����
    '�޸�ʱ��       2004-10-18
    '����           ����¼������Ƿ�ȫ��Ϊͣ�ò���
    '����           =Trueȫ��Ϊͣ�ò���=Falseȫ����Ϊͣ�ò���
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    
    CheckStop = True
    
    If ActiveControl Is tvwMain_S Then
        With tvwMain_S.SelectedItem
            If .Children > 0 Then
                i = .Child.FirstSibling.Index
                If .Child.FirstSibling.Image = "Dept" Then
                    CheckStop = False
                    Exit Function
                End If
                While i <> .Child.LastSibling.Index
                    If Me.tvwMain_S.Nodes(i).Next.Image = "Dept" Then
                        CheckStop = False
                        Exit Function
                    End If
                    i = Me.tvwMain_S.Nodes(i).Next.Index
                Wend
            End If
        End With
    Else
        With tvwMain_S.Nodes(lvwMain.SelectedItem.Key)
            If .Children > 0 Then
                i = .Child.FirstSibling.Index
                If .Child.FirstSibling.Image = "Dept" Then
                    CheckStop = False
                    Exit Function
                End If
                While i <> .Child.LastSibling.Index
                    If Me.tvwMain_S.Nodes(i).Next.Image = "Dept" Then
                        CheckStop = False
                        Exit Function
                    End If
                    i = Me.tvwMain_S.Nodes(i).Next.Index
                Wend
            End If
        End With
    End If
    
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub LocateItem()
    Dim strTemp As String
    
    txtFind.SetFocus
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    If mrsFind.RecordCount = 0 Then
        MsgBox " û���ҵ�������������Ϣ��", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        MsgBox " �Ѿ���λ�������ҵ�����Ϣ������������������", vbInformation, gstrSysName
        txtFind.SetFocus
        Exit Sub
    End If
    
    With frmDeptManage.tvwMain_S
        If IsNull(mrsFind("�ϼ�ID")) Then
            If tbcDetails.Selected.Index = mint����� Then
                .Nodes("C" & mrsFind("ID")).Selected = True
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            Else
                strTemp = mrsFind!���� & "|" & mrsFind!ID
                .Nodes("C" & strTemp).Selected = True
                .SelectedItem.EnsureVisible
                frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            End If
        Else
            If tbcDetails.Selected.Index = mint������ Then
                strTemp = mrsFind!���� & "|" & mrsFind!ID
                .Nodes("C" & strTemp).Selected = True
                .Nodes("C" & strTemp).Expanded = True
            Else
                .Nodes("C" & mrsFind("�ϼ�ID")).Selected = True
                .Nodes("C" & mrsFind("�ϼ�ID")).Expanded = True
            End If
            
            .SelectedItem.EnsureVisible
            frmDeptManage.tvwMain_S_NodeClick .SelectedItem
            
            If tbcDetails.Selected.Index = mint����� Then
                frmDeptManage.lvwMain.ListItems("C" & mrsFind("ID")).Selected = True
                frmDeptManage.lvwMain.SelectedItem.EnsureVisible
                frmDeptManage.lvwMain_ItemClick frmDeptManage.lvwMain.SelectedItem
            End If
        End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    If KeyAscii = vbKeyReturn Then
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            strTemp = " and (a.����ʱ�� = to_date('3000-01-01','YYYY-MM-DD') or a.����ʱ�� is null ) "
            gstrSQL = "Select a.id,a.�ϼ�id,a.����,a.���� ,c.���� as ���� From ���ű� A, ��������˵�� B,�������ʷ��� c Where b.��������=c.���� " & _
                " and A.ID=B.����ID  and (a.���� like [1] or a.���� like [2] or a.���� like [3]) "
            
            If mnuViewShowStop.Checked = False Then
                gstrSQL = gstrSQL & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����", UCase(txtFind.Text) & "%", UCase(txtFind.Text) & "%", UCase(txtFind.Text) & "%")
            Call LocateItem
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                Call LocateItem
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call LocateItem
            End If
        End If
    End If
End Sub

Private Function CheckBusiness(ByVal lngDeptID As Long) As Boolean
'���ܣ����ҵ�������
'������
'  lngDeptID������ID
'���أ�True-ͨ����False��ͨ��

    Dim strSQL As String, strMess As String
    Dim rsTmp As ADODB.Recordset, rsBusiness As ADODB.Recordset
    Dim lngSys As Long
    Dim dblStock As Double
    
    On Error GoTo hErr
    strSQL = "Select Count(1) Rec, 400 ��� From zlSystems Where ��� = 400 And Nvl(�����, 0) = 100 Union all " & vbCr & _
             "Select Count(1) Rec, 600 ��� From zlSystems Where ��� = 600 And Nvl(�����, 0) = 100 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ǳ�׼ϵͳ��Ϣ")
    
    If rsTmp.EOF Then
        'û�й���װ����ϵͳ���Ͳ�������ص�ҵ����
        rsTmp.Close
        CheckBusiness = True
        Exit Function
    End If
    
    Do While rsTmp.EOF = False
        dblStock = 0
        lngSys = NVL(rsTmp!���, 0)
        Select Case lngSys
            Case 400    '����ϵͳ
                If NVL(rsTmp!Rec, 0) = 1 Then
                    strMess = "�ò������ʿ����ڣ����飡"
                    strSQL = "Select Sum(ʵ������) ʵ������ From ���ʿ�� Where �ⷿid = [1] "
                    Set rsBusiness = zlDatabase.OpenSQLRecord(strSQL, "������ʿ��", lngDeptID)
                    If rsBusiness.EOF = False Then
                        dblStock = NVL(rsBusiness!ʵ������, 0)
                    End If
                    rsBusiness.Close
                End If
            Case 600    '�豸ϵͳ
                If NVL(rsTmp!Rec, 0) = 1 Then
                    strMess = "�ò����豸�����ڣ����飡"
                    strSQL = "Select Sum(ʵ������) ʵ������ From �豸��� Where �ⷿid = [1] "
                    Set rsBusiness = zlDatabase.OpenSQLRecord(strSQL, "����豸���", lngDeptID)
                    If rsBusiness.EOF = False Then
                        dblStock = NVL(rsBusiness!ʵ������, 0)
                    End If
                    rsBusiness.Close
                End If
        End Select
        
        If dblStock <> 0 Then
            '���ڿ������
            rsTmp.Close
            MsgBox strMess, vbInformation, gstrSysName
            Exit Function
        End If
        
        rsTmp.MoveNext
    Loop
    
    CheckBusiness = True
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function
    
