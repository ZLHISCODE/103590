VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmӦ������� 
   Caption         =   "Ӧ�������"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9570
   FillColor       =   &H00404080&
   Icon            =   "frmӦ�������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicColor 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   7230
      ScaleHeight     =   540
      ScaleWidth      =   2265
      TabIndex        =   13
      Top             =   75
      Width           =   2295
      Begin VB.Label lblColor 
         BackColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   1215
         TabIndex        =   21
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   20
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF8080&
         Height          =   195
         Index           =   1
         Left            =   1215
         TabIndex        =   19
         Top             =   45
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00404080&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   3
         Left            =   1605
         TabIndex        =   17
         Top             =   307
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   375
         TabIndex        =   16
         Top             =   307
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�Ѹ���"
         Height          =   180
         Index           =   1
         Left            =   1605
         TabIndex        =   15
         Top             =   52
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ƻ�����"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   14
         Top             =   52
         Width           =   720
      End
   End
   Begin VB.PictureBox PicRang 
      BackColor       =   &H8000000C&
      Height          =   315
      Left            =   2835
      ScaleHeight     =   255
      ScaleWidth      =   6630
      TabIndex        =   11
      Top             =   765
      Width           =   6690
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         ForeColor       =   &H80000018&
         Height          =   180
         Left            =   75
         TabIndex        =   12
         Top             =   45
         Width           =   3330
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2445
      Left            =   2805
      TabIndex        =   10
      Top             =   1095
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1050
      Top             =   2970
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
            Picture         =   "frmӦ�������.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":0D22
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":117A
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":15CE
            Key             =   "No"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":1A26
            Key             =   "Write"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwList 
      Height          =   5355
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   2356
            Picture         =   "frmӦ�������.frx":1E7E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11792
            MinWidth        =   600
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6150
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":2712
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":2932
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":2B52
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":2D6E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":2F8E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":31AE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":33CA
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":35E6
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":3800
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":395A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":3B76
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":3D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":3FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":41CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":43E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":45FE
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":481E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":4A3E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":4C5A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":4E7A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":509A
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":52B6
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":54D2
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":56EC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":5846
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":5A66
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":5C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":5EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":60BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmӦ�������.frx":62D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   9570
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ƻ�"
               Key             =   "SplitDue"
               Description     =   "�ƻ�"
               Object.ToolTipText     =   "�ƶ�����ƻ�"
               Object.Tag             =   "�ƻ�"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "���ù�������"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "��λ"
               Key             =   "Find"
               Description     =   "��λ"
               Object.ToolTipText     =   "���ݶ�λ"
               Object.Tag             =   "��λ"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmӦ�������.frx":64EE
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSplit 
      Height          =   1800
      Left            =   2805
      TabIndex        =   6
      Top             =   3975
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   3175
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblTemp 
      Caption         =   "Ӧ�����ܶ"
      Height          =   165
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   5865
      Width           =   4440
   End
   Begin VB.Label lblTemp 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   2805
      TabIndex        =   8
      Top             =   5790
      Width           =   6750
   End
   Begin VB.Label lblHsc_s 
      Height          =   5355
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   750
      Width           =   60
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ����ƻ�"
      ForeColor       =   &H80000018&
      Height          =   180
      Index           =   0
      Left            =   2985
      TabIndex        =   7
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label lblVsc_s 
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2850
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   3645
      Width           =   6750
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
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "�ƻ�(&S)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&V)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditNO 
         Caption         =   "����(&D)"
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
         Begin VB.Menu mnuViewLine1 
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
      Begin VB.Menu mnuViewSp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSavePrint 
         Caption         =   "���̴�ӡ(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewVerifyPrint 
         Caption         =   "��˴�ӡ(&V)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "����(&J)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "���ݶ�λ(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
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
            Caption         =   "���ͷ���(&K)"
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)"
      End
   End
End
Attribute VB_Name = "frmӦ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mblnFirst  As Boolean

Private Enum HeadCol
    ID
    ��λID
    �շ�ID
    �������
    �ƻ����
    ��¼״̬
    �����־
    ���ݺ�
    ��Ӧ��
    Ʒ��
    ���
    ��λ
    ����
    ����
    �������
    ��Ʊ��
    ��Ʊ����
    ����
    ��Ʊ���
    �ɹ���
    �ɹ����
    ������
    ��������
    �����
    �������
    ��ǰ�ⷿ
    ��ǰ�ⷿ���
    ȫԺ���
    ҩ�ⵥλ
End Enum

Private mdtStartDate As Date    '��������
Private mdtEndDate As Date
Private mdtVerifyStartDate As Date  '�������
Private mdtVerifyEndDate As Date
Private mstrFind As String
Private mstr���� As String
Private msngDownX As Single, msngDownY As Single, mSelKey As String
Private mlngModule As Long
Private mcllFilter As Collection
Private mbln�����־ As Boolean

'by lesfeng 2010-1-7 �����Ż�
Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-08-18 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "��������"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "�������"
    mcllFilter.Add Array("", ""), "���ݺ�"
    mcllFilter.Add "", "�������"
    mcllFilter.Add "", "��Ӧ��id"
    mcllFilter.Add "", "�ⷿID"
    mcllFilter.Add "", "������"
    mcllFilter.Add "", "�����"
    mstrFind = ""
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call Ȩ�޿���
End Sub

Private Sub Form_Load()
    Dim strReg As String
    mstrPrivs = gstrPrivs
    mblnFirst = True
    mlngModule = glngModul
    mstr���� = "0000"
    '�ָ�����
    'by lesfeng 2010-1-7 �����Ż�
    Call InitFilter
    
    mnuViewSavePrint.Checked = IIf(Val(zldatabase.GetPara("���̴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1
    mnuViewVerifyPrint.Checked = IIf(Val(zldatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1
    mbln�����־ = Val(zldatabase.GetPara("�⹺�����Ҫ������Ǹ������ܽ��и������", glngSys, 0)) = 1
    
    mdtStartDate = Format(DateAdd("d", -7, zldatabase.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    mdtVerifyStartDate = "1901-01-01"
    mdtVerifyEndDate = "1901-01-01"
        
    'by lesfeng 2010-1-7 �����Ż�
'    mstrFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [1] And [2]"
    mstrFind = " And (A.�������� Between [1] And [2]) and ������� is null"
    lblRange = "��ѯ��Χ:" & Format(DateAdd("d", -7, zldatabase.Currentdate), "yyyy��MM��dd��") & "��" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    
    mcllFilter.Remove "��������"
    mcllFilter.Add Array(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00", Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), "��������"
       
    mSelKey = ""
    List��Ӧ��
    RestoreWinState Me, App.ProductName
   
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����id As Long
    Dim lngӦ����¼ID As Long
    Dim lng�շ�ID As Long
    Dim str��Ʊ�� As String
    Dim strNO As String
    Dim lng��λID As Long
    Dim lng��¼״̬ As Long
    Dim lng������� As Long
    If Not tvwList.SelectedItem Is Nothing Then
        lng����id = Val(Mid(tvwList.SelectedItem.Key, 2))
    End If
    
    lngӦ����¼ID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    lng�շ�ID = Val(mshList.TextMatrix(mshList.Row, HeadCol.�շ�ID))
    lng��λID = Val(mshList.TextMatrix(mshList.Row, HeadCol.��λID))
    str��Ʊ�� = Trim(mshList.TextMatrix(mshList.Row, HeadCol.��Ʊ��))
    strNO = Trim(mshList.TextMatrix(mshList.Row, HeadCol.���ݺ�))
    lng��¼״̬ = Val(mshList.TextMatrix(mshList.Row, HeadCol.��¼״̬))
    lng������� = Val(mshList.TextMatrix(mshList.Row, HeadCol.�������))
    
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "����=" & lng����id, "Ӧ����¼=" & lngӦ����¼ID, "��ⵥ��=" & lng�շ�ID, "��Ʊ��=" & str��Ʊ��, "NO=" & strNO, "��Ӧ��=" & lng��λID, "��¼״̬=" & lng��¼״̬, "�������=" & lng�������)
    
End Sub

Private Sub List��Ӧ��()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ع�Ӧ����������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rstTemp As New ADODB.Recordset
    Dim nodTemp As Node
    Dim strKey As String
    Dim strPrentKey As String
    
    If tvwList.SelectedItem Is Nothing Then
        strKey = "Root"
        strPrentKey = "Root"
    Else
        If tvwList.SelectedItem.Parent Is Nothing Then
            strKey = "Root"
            strPrentKey = "Root"
        Else
            strKey = tvwList.SelectedItem.Key
            strPrentKey = tvwList.SelectedItem.Parent.Key
        End If
    End If
    tvwList.Nodes.Clear
    tvwList.Nodes.Add , , "Root", "���й�Ӧ��", 1
    tvwList.Nodes("Root").Sorted = True
    Dim i As Long
    Dim str���� As String
    
    
    str���� = ""
    For i = 1 To Len(mstr����)
        If Mid(mstr����, i, 1) = 1 And Check���Ȩ��(mstrPrivs, i) Then
            str���� = str���� & " or substr(����," & i & ",1)=1"
        End If
    Next
    If str���� <> "" Then
        str���� = " And (" & Mid(str����, 4) & ") "
    End If
    
    Dim strȨ�� As String
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs, "")
    
    'by lesfeng 2010-1-7 �����Ż�
    gstrSQL = "" & _
        "   Select id,�ϼ�id,����,����,����,ĩ��" & _
        "   From ��Ӧ��" & _
        "       where (����ʱ�� is null or To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01') " & _
        "           and ( ĩ��<>1 or (ĩ��=1 " & zl_��ȡվ������() & "  " & _
                    str���� & strȨ�� & "))" & _
        "   start with �ϼ�id is null connect by prior id=�ϼ�id "
        
    Err = 0
    On Error GoTo ErrHand:
    zldatabase.OpenRecordset rstTemp, gstrSQL, Me.Caption
    With rstTemp
        While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set nodTemp = tvwList.Nodes.Add("Root", tvwChild, "K" & rstTemp("ID"), "��" & Nvl(!����) & "��" & Nvl(!����), IIf(Nvl(!ĩ��, 0) = 0, 5, 2))
            Else
                Set nodTemp = tvwList.Nodes.Add("K" & !�ϼ�ID, tvwChild, "K" & rstTemp("ID"), "��" & !���� & "��" & !����, IIf(!ĩ�� = 0, 5, 2))
            End If
            If strKey = "K" & Nvl(!ID) Then
                nodTemp.Selected = True
                nodTemp.Expanded = True
            End If
            nodTemp.Tag = Nvl(!����)
            nodTemp.Sorted = True
            rstTemp.MoveNext
        Wend
    End With
    If tvwList.SelectedItem Is Nothing Then
        Err = 0
        On Error Resume Next
        If strPrentKey <> "" Then
            tvwList.Nodes(strPrentKey).Selected = True
            tvwList.Nodes(strPrentKey).Expanded = True
        End If
        If Err <> 0 Then
            tvwList.Nodes("Root").Selected = True
            tvwList.Nodes("Root").Expanded = True
        End If
    End If
    Err.Clear
    tvwList_NodeClick tvwList.SelectedItem
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FullӦ����¼()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim str���� As String
    Dim lng�ϼ�id As Long
    
    str���� = ""
    For i = 1 To Len(mstr����)
        If Mid(mstr����, i, 1) = 1 Then
            str���� = str���� & " or substr(b.����," & i & ",1)=1"
        End If
    Next
    If str���� <> "" Then
        str���� = " And (" & Mid(str����, 4) & ") "
    End If
    mshList.Redraw = False
    Dim strȨ�� As String
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs, "a.", False)
    
    If tvwList.SelectedItem.Key = "Root" Then
        'IIf(mstr���� = "0000", "And (b.���� is null or b.����='" & mstr���� & "')", "       and b.����='" & mstr���� & "'")
        strSQL = "" & _
            "   Select  a.id,a.��λID as  ��λid,a.�շ�ID,nvl(�������,0) as �������,decode(a.�ƻ����,null,-1,0,-1,a.�ƻ����) as �ƻ����,nvl(a.��¼״̬,1) as ��¼״̬," & _
            IIf(mbln�����־, "decode(a.�����־,1,'����','') �����־,", "'' �����־,") & _
            "           a.no as ���ݺ�,'['||b.����||']'||b.���� as ��Ӧ��,a.Ʒ��,a.���,a.������λ,a.����,a.����,a.�������, a.��Ʊ��,to_char(a.��Ʊ����,'yyyy-mm-dd') as ��Ʊ����," & _
            "           a.����,a.��Ʊ���,a.�ɹ���,a.�ɹ����," & _
            "           a.������,to_char(a.��������,'yyyy-mm-dd hh24:mi:ss') as ��������," & _
            "           a.�����,to_char(a.�������,'yyyy-mm-dd hh24:mi:ss') as �������" & _
            IIf(mbln�����־, ", e.���� ��ǰ�ⷿ, c.ȫԺ���, d.��ǰ�ⷿ���, c.ҩ�ⵥλ ", "") & _
            "   From Ӧ����¼ a,��Ӧ�� b " & _
            IIf(mbln�����־, ",(Select a.ҩƷid, Round(a.ȫԺ��� / b.ҩ���װ, 5) ȫԺ���, b.ҩ�ⵥλ From (Select ҩƷid, Sum(ʵ������) ȫԺ��� From ҩƷ��� Group By ҩƷid) A, ҩƷ��� B Where a.ҩƷid = b.ҩƷid) C, " & _
                              " (Select a.�ⷿid, a.ҩƷid, Round(a.��ǰ�ⷿ��� / b.ҩ���װ, 5) ��ǰ�ⷿ��� From (Select �ⷿid, ҩƷid, Sum(ʵ������) ��ǰ�ⷿ��� From ҩƷ��� Group By �ⷿid, ҩƷid) A, ҩƷ��� B Where a.ҩƷid = b.ҩƷid) D, " & _
                              " ���ű� E ", "") & _
            "   Where a.��λID=b.id " & _
            IIf(mbln�����־, " and a.��Ŀid=c.ҩƷid(+) and a.��Ŀid=d.ҩƷid(+) and a.�ⷿid=d.�ⷿid(+) and a.�ⷿid=e.id(+) ", "") & _
            "     and not a.��¼���� in (-1,2) " & zl_��ȡվ������(True, "b") & "   " & _
            "" & str���� & strȨ�� & mstrFind & _
            "   Order By a.�������� desc,a.NO"
            lng�ϼ�id = 0
    Else
        strSQL = "" & _
            "   Select  a.id,a.��λID as  ��λid,a.�շ�ID,nvl(�������,0) as �������,decode(a.�ƻ����,null,-1,0,-1,a.�ƻ����) as �ƻ����,nvl(a.��¼״̬,1) as ��¼״̬," & _
            IIf(mbln�����־, "decode(a.�����־,1,'����','') �����־,", "'' �����־,") & _
            "           a.no as ���ݺ�,'['||b.����||']'||b.���� as ��Ӧ��,a.Ʒ��,a.���,a.������λ,a.����,a.����,a.�������, a.��Ʊ��,to_char(a.��Ʊ����,'yyyy-mm-dd') as ��Ʊ����," & _
            "           a.����,a.��Ʊ���,a.�ɹ���,a.�ɹ����," & _
            "           a.������,to_char(a.��������,'yyyy-mm-dd hh24:mi:ss') as ��������," & _
            "           a.�����,to_char(a.�������,'yyyy-mm-dd hh24:mi:ss') as �������" & _
            IIf(mbln�����־, ", e.���� ��ǰ�ⷿ, c.ȫԺ���, d.��ǰ�ⷿ���, c.ҩ�ⵥλ ", "") & _
            "   From Ӧ����¼ a,��Ӧ�� b " & _
            IIf(mbln�����־, ",(Select a.ҩƷid, Round(a.ȫԺ��� / b.ҩ���װ, 5) ȫԺ���, b.ҩ�ⵥλ From (Select ҩƷid, Sum(ʵ������) ȫԺ��� From ҩƷ��� Group By ҩƷid) A, ҩƷ��� B Where a.ҩƷid = b.ҩƷid) C, " & _
                              " (Select a.�ⷿid, a.ҩƷid, Round(a.��ǰ�ⷿ��� / b.ҩ���װ, 5) ��ǰ�ⷿ���, b.ҩ�ⵥλ From (Select �ⷿid, ҩƷid, Sum(ʵ������) ��ǰ�ⷿ��� From ҩƷ��� Group By �ⷿid, ҩƷid) A, ҩƷ��� B Where a.ҩƷid = b.ҩƷid) D, " & _
                              " ���ű� E", "") & _
            "   Where  a.��λID=b.id " & _
            IIf(mbln�����־, " and a.��Ŀid=c.ҩƷid(+) and a.��Ŀid=d.ҩƷid(+) and a.�ⷿid=d.�ⷿid(+) and a.�ⷿid=e.id(+) ", "") & _
            "     and not a.��¼���� in (-1,2)  And a.��λID in (select ID From ��Ӧ�� where  " & zl_��ȡվ������(False) & "  start with id= [12] connect by prior id=�ϼ�id )" & _
            "" & str���� & mstrFind & strȨ�� & _
            "   Order By a.�������� desc,a.NO"
            lng�ϼ�id = Val(Mid(tvwList.SelectedItem.Key, 2))
    End If
    
    'by lesfeng 2010-1-7 �����Ż�
    '��������: [1] [2]
    '�������: [3] [4]
    '���ݺ�:   [5] [6]
    '�������: [7]
    '��Ӧ��id: [8]
    '������: [9]
    '�����: [10]
    On Error GoTo errHandle
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mcllFilter("��������")(0)), CDate(mcllFilter("��������")(1)), _
        CDate(mcllFilter("�������")(0)), CDate(mcllFilter("�������")(1)), CStr(mcllFilter("���ݺ�")(0)), CStr(mcllFilter("���ݺ�")(1)), _
        CStr(mcllFilter("�������")), CLng(Val(mcllFilter("��Ӧ��id"))), CStr(mcllFilter("������")), CStr(mcllFilter("�����")), _
        Val(mcllFilter("�ⷿID")), lng�ϼ�id)

    mshList.Redraw = False
    If rsTemp.RecordCount > 0 Then
        Set mshList.Recordset = rsTemp
        mshList.Row = 1
        mshList.Col = HeadCol.���ݺ�
        mshList.ColSel = mshList.Cols - 1
    Else
        mshList.Clear
        mshList.Rows = 2
    End If
    stbThis.Panels(2).Text = "��ǰ����" & rsTemp.RecordCount & "�ŵ���"
    
    setGridColor
    formatMSH
    �ϼ�
    mshList.Redraw = True
    Full����ƻ�
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub �ϼ�()
    Dim dblSum As Double
    Dim lngRow As Long
    With mshList
        For lngRow = 1 To .Rows - 1
            dblSum = dblSum + Val(.TextMatrix(lngRow, HeadCol.��Ʊ���))
        Next
        '��ȡ��ǰδ�����
        
    End With
    lblTemp(2) = "Ӧ�����ܶ" & Format(dblSum, "###0.00;-###0.00;0;0") & "Ԫ"
    
End Sub

Private Sub SetRowColor(ByVal lngRow As Long, ByVal lngColor As Long, Optional blnList As Boolean = True)
    Dim intCol As Integer
    Dim objTmp As Object
    Set objTmp = IIf(blnList, mshList, mshSplit)
    With objTmp
        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next
    End With
End Sub

Private Sub setGridColor()
    Dim lngStatus As Long
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshList
        'If mrsDue.RecordCount = 0 Then Exit Sub
        'mrsDue.MoveFirst
        For intRow = 1 To .Rows - 1
            lngStatus = Val(.TextMatrix(intRow, HeadCol.�ƻ����))
            If lngStatus <> -1 Then    '�ƻ���Ų�Ϊ0˵���Ѿ��мƻ������˸���
                SetRowColor intRow, &H404080
            End If
            lngStatus = Val(.TextMatrix(intRow, HeadCol.�������))
            If lngStatus <> 0 Then '�˼�¼�Ѹ���
                SetRowColor intRow, &HFF8080
            Else
                lngStatus = Val(.TextMatrix(intRow, HeadCol.��¼״̬))
            
                If lngStatus Mod 3 = 0 Then             '��������¼
                    SetRowColor intRow, &H80000001
                ElseIf lngStatus Mod 3 = 2 Then         '������¼
                    SetRowColor intRow, &HFF
                End If
            End If
        Next
        .Col = 0
        .Row = 1
    End With
    
End Sub

Private Sub Full����ƻ�()
    Dim strSQL As String
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    
    On Error GoTo errHandle
    If lngID = 0 Then
        Set mshSplit.Recordset = Nothing
        mshSplit.Clear
        mshSplit.Rows = 2
    Else
        'by lesfeng 2010-1-7 �����Ż�
        strSQL = "Select �������,�ƻ����,�ƻ����,to_char(�ƻ�����,'yyyy-MM-dd'),�ƻ���,to_char(�ƶ�����,'yyyy-MM-dd') From Ӧ����¼ Where ID= [1] And ��¼����=-1 Order By �ƻ����"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        If rsTemp.RecordCount > 0 Then
            Set mshSplit.Recordset = rsTemp
        Else
            Set mshSplit.Recordset = Nothing
            mshSplit.Clear
            mshSplit.Rows = 2
        End If
    End If
    
    With mshSplit
        .Redraw = False
        .FormatString = "^�������|^�ƻ����|^�ƻ����|^�ƻ���������|^�ƻ���|^�ƶ��ƻ�����"
        .ColAlignment(0) = 7
        .ColWidth(0) = 0
        .ColWidth(1) = 1000: .ColAlignment(1) = 4
        .ColWidth(2) = 1100: .ColAlignment(2) = 7
        .ColWidth(3) = 1300: .ColAlignment(3) = 4
        .ColWidth(4) = 1100: .ColAlignment(4) = 1
        .ColWidth(5) = 1300: .ColAlignment(5) = 4
        setPlanGrdColor
        .Col = 1
        .ColSel = .Cols - 1
        .Redraw = True
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    Err = 0
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If
    
    If cbrThis.Bands(1).MinHeight <> tlbThis.Height Then cbrThis.Bands(1).MinHeight = tlbThis.Height
    cbrThis.Move 0, 0, Me.ScaleWidth
    
    If lblHsc_s.Left > Me.ScaleWidth - 2000 Then lblHsc_s.Left = Me.ScaleWidth - 2000
    
    lblHsc_s.Top = IIf(cbrThis.Visible, cbrThis.Height + 30, 0)
    lblHsc_s.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - lblHsc_s.Top - 15
    tvwList.Move 0, lblHsc_s.Top, lblHsc_s.Left, lblHsc_s.Height
    If lblVsc_s.Top > Me.ScaleHeight - 2000 Then lblVsc_s.Top = Me.ScaleHeight - 2000
    
    lblVsc_s.Left = lblHsc_s.Left + lblHsc_s.Width
    lblVsc_s.Width = Me.ScaleWidth - lblVsc_s.Left
    With PicRang
        .Top = lblHsc_s.Top
        .Width = lblVsc_s.Width
        .Left = lblVsc_s.Left
    End With
    With mshList
        .Left = lblVsc_s.Left
        .Top = lblHsc_s.Top + PicRang.Height + 50
        .Width = lblVsc_s.Width
        .Height = lblVsc_s.Top - .Top
    End With
    
    lblTemp(0).Move (lblVsc_s.Width - lblTemp(0).Width) / 2 + lblVsc_s.Left, (lblVsc_s.Height - lblTemp(0).Height) / 2 + lblVsc_s.Top + 1
    With mshSplit
        .Left = lblVsc_s.Left
        .Top = lblVsc_s.Top + lblVsc_s.Height
        .Width = lblVsc_s.Width
        .Height = tvwList.Top + tvwList.Height - .Top - lblTemp(1).Height
        
    End With
    lblTemp(1).Move lblVsc_s.Left, mshSplit.Top + mshSplit.Height, lblVsc_s.Width
    lblTemp(2).Move lblTemp(1).Left + 60, lblTemp(1).Top + (lblTemp(1).Height - lblTemp(2).Height) / 2, lblTemp(1).Width - 120
    With PicColor
        .Left = ScaleWidth - .Width - 100
        .Top = 80
    End With
    mnuViewToolButton.Checked = cbrThis.Visible
    mnuViewStatus.Checked = stbThis.Visible
    mnuViewToolText.Checked = tlbThis.Buttons(1).Caption <> ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuEditAdd_Click()
    Dim blnReturn As Boolean
    
    '����
    If tvwList.SelectedItem.Image = 2 Then
        Call frmӦ����༭.ShowCard(Me, 0, g����, mstrPrivs, Val(Mid(tvwList.SelectedItem.Key, 2)), , blnReturn)
    Else
        Call frmӦ����༭.ShowCard(Me, 0, g����, mstrPrivs, 0, , blnReturn)
    End If
    If blnReturn = False Then Exit Sub
    '���
    FullӦ����¼
    
End Sub

Private Sub mnuEditDelete_Click()
'ɾ��
    Dim strSQL As String
    Dim intRow As Integer
    Dim lngID As Long
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    If MsgBox("��ȷ��Ҫɾ����Ӧ����¼��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    strSQL = "ZL_Ӧ����¼_DELETE(" & lngID & ")"
    
    zldatabase.ExecuteProcedure strSQL, Me.Caption
    
    With mshList
        intRow = .Row
        If .Rows > 2 Then
            .RemoveItem intRow
        ElseIf .Rows = 2 Then
            .Rows = 3
            .RemoveItem intRow
            SetEnabled
        End If
        If intRow < .Rows - 1 Then
            .Row = intRow
        Else
            If .Rows = 2 Then
                .Row = 1
            Else
                .Row = intRow - 1
            End If
        End If
        .Col = 0
        .ColSel = .Cols - 1
    End With
    Full����ƻ�
End Sub

Private Sub mnuEditModify_Click()
    Dim blnReturn As Boolean
    '�޸�
    Dim lngID As Long
    Dim int��¼״̬ As Integer
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    Call frmӦ����༭.ShowCard(Me, lngID, g�޸�, mstrPrivs, , , blnReturn)
    If blnReturn = False Then Exit Sub
    
    FullӦ����¼
End Sub

Private Sub mnuEditNO_Click()
    Dim blnReturn  As Boolean
    Dim lngID As Long
    Dim int��¼״̬ As Integer
    Dim bytRec As RecBillStatus
    
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    int��¼״̬ = Val(mshList.TextMatrix(mshList.Row, HeadCol.��¼״̬))
    If lngID = 0 Then Exit Sub
'    Select Case int��¼״̬
'    Case 1
'        bytRec = ������¼
'    Case 2
'        bytRec = ������¼
'    Case Else
'        bytRec = ��������¼
'    End Select
    '���
    Call frmӦ����༭.ShowCard(Me, lngID, g�鿴, mstrPrivs, 0, int��¼״̬, blnReturn)
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditSplit_Click()
    '�ƻ�
    Dim lngID As Long
    If mnuEditSplit.Enabled = False Then Exit Sub
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    frm����ƻ�.�ƻ� Me, lngID
    Full����ƻ�
End Sub

Private Sub mnuEditStrike_Click()
    '����
    Dim blnReturn As Boolean
    Dim lngID As Long
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    
    Call frmӦ����༭.ShowCard(Me, lngID, gȡ��, mstrPrivs, 0, ������¼, blnReturn)
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
    
End Sub

Private Sub mnuEditVerify_Click()
    Dim blnReturn  As Boolean
    Dim lngID As Long
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID = 0 Then Exit Sub
    '���
    Call frmӦ����༭.ShowCard(Me, lngID, g���, mstrPrivs, 0, ������¼, blnReturn)
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuFileExcel_Click()
'�����Excel
    mshList.Redraw = False
    subPrint 3
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreView_Click()
'��ӡԤ��
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrint_Click()
'��ӡ
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub lblHsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc_s
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        Call Form_Resize
    End If
End Sub

Private Sub lblVsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
End Sub

Private Sub lblVsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblVsc_s
            If .Top + Y - msngDownY < 2000 Then Exit Sub
            If .Top + Y - msngDownY > ScaleHeight - 2000 Then Exit Sub
            .Top = .Top + Y - msngDownY
        End With
        Call Form_Resize
    End If
End Sub

Private Sub mnuViewFind_Click()
'��λ
End Sub

Private Sub mnuViewOpen_Click()
    Dim strCon As String
    Dim strFind As String
    Dim str���� As String
    Dim cllFilter As Collection
    
    str���� = mstr����
    'by lesfeng 2010-1-7 �����Ż�
    strFind = frmӦ�������.GetSearch(Me, mstrPrivs, mdtStartDate, mdtEndDate, mdtVerifyStartDate, mdtVerifyEndDate, mstr����, cllFilter)
    If strFind = "" Then Exit Sub
    mstrFind = strFind
    Set mcllFilter = cllFilter
    '��������
    '
    If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStartDate, "yyyy-mm-dd") = "1901-01-01" Then
    ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVerifyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
        strCon = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVerifyStartDate, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEndDate, "yyyy��MM��dd��")
    ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
        strCon = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
    ElseIf Format(mdtVerifyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
        strCon = "��ѯ��Χ:������� " & Format(mdtVerifyStartDate, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEndDate, "yyyy��MM��dd��")
    End If
    lblRange = strCon
    If str���� <> mstr���� Then
        '��������
        mSelKey = ""
        List��Ӧ��
    Else
        FullӦ����¼
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    Err = 0
    On Error Resume Next
    mSelKey = ""
    tvwList_NodeClick tvwList.SelectedItem
End Sub

Private Sub mnuViewSavePrint_Click()
    mnuViewSavePrint.Checked = Not mnuViewSavePrint.Checked
    Call zldatabase.SetPara("���̴�ӡ", IIf(mnuViewSavePrint.Checked, "1", "0"), glngSys, mlngModule)
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    PicColor.Visible = mnuViewToolButton.Checked And mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    PicColor.Visible = mnuViewToolButton.Checked And mnuViewToolButton.Checked
    For Each buttTemp In tlbThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub mnuViewVerifyPrint_Click()
    mnuViewVerifyPrint.Checked = Not mnuViewVerifyPrint.Checked
    Call zldatabase.SetPara("��˴�ӡ", IIf(mnuViewVerifyPrint.Checked, "1", "0"), glngSys, mlngModule)
    
End Sub

Private Sub mshList_Click()
    Dim strSQL As String
    Dim lngID As Long
    lngID = Val(mshList.TextMatrix(mshList.Row, HeadCol.ID))
    If lngID <> Val(mshSplit.Tag) Then
        Full����ƻ�
        lngID = lngID
    End If
End Sub

Private Sub mshList_DblClick()
    If mnuEditModify.Enabled And mnuEditModify.Visible Then
        mnuEditModify_Click
    Else
        mnuEditNO_Click
    End If
End Sub

Private Sub mshList_EnterCell()
    SetEnabled
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuEdit
End Sub

Private Sub mshList_RowColChange()
    mshList_Click
End Sub

Private Sub mshSplit_DblClick()
    mnuEditSplit_Click
End Sub

Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAdd_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "SplitDue"
            mnuEditSplit_Click
        Case "Find"
            mnuViewFind_Click
        Case "Search"
            Call mnuViewOpen_Click
        Case "Refresh"
            Call mnuViewRefresh_Click
        Case "Help"
            Call mnuHelpTitle_Click
        Case "Exit"
            Call mnuFileExit_Click
        Case "Verify"
            Call mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
    End Select
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    If mSelKey = Node.Key Then Exit Sub
    mSelKey = Node.Key
    FullӦ����¼
    SetEnabled
End Sub

Private Sub SetEnabled()
    Dim blnData As Boolean '��������
    Dim blnVerfiy As Boolean '�Ƿ����
    Dim blnCancel As Boolean  '��������
    Dim blnPayMoney As Boolean '�Ѿ�����
    Dim blnSys As Boolean       'ϵͳ��������
    Dim blnPlan As Boolean
    Dim blnSign As Boolean
    
    If mshList.Rows <= 1 Then
        blnData = False
        blnVerfiy = False
        blnPayMoney = False
        blnPlan = False
        blnSign = False
    Else
        With mshList
            blnData = Val(mshList.TextMatrix(1, HeadCol.ID)) <> 0
            blnVerfiy = Trim(.TextMatrix(.Row, HeadCol.�������)) <> ""
            blnCancel = Val(.TextMatrix(.Row, HeadCol.��¼״̬)) <> 1
            blnPlan = Val(.TextMatrix(.Row, HeadCol.��¼״̬)) = 1 Or Val(.TextMatrix(.Row, HeadCol.��¼״̬)) = 3
            blnPayMoney = Val(.TextMatrix(.Row, HeadCol.�������)) <> 0
            blnSys = Val(mshList.TextMatrix(.Row, HeadCol.�շ�ID)) <> 0
            blnSign = mshList.TextMatrix(.Row, HeadCol.�����־) = "����"
        End With
    End If
    
    mnuEditModify.Enabled = blnData And Not blnVerfiy And Not blnSys
    mnuEditDelete.Enabled = blnData And Not blnVerfiy And Not blnSys
    mnuEditStrike.Enabled = blnData And blnVerfiy And Not blnCancel And (Not blnPayMoney) And Not blnSys
    If mbln�����־ Then
        mnuEditVerify.Enabled = blnData And Not blnVerfiy And (Not blnSys Or blnSign)
    Else
        mnuEditVerify.Enabled = blnData And Not blnVerfiy And Not blnSys
    End If
    mnuEditSplit.Enabled = blnData And blnVerfiy And blnPlan And (Not blnPayMoney)
    mnuEditNO.Enabled = blnData
    
    tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tlbThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
    tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
    tlbThis.Buttons("SplitDue").Enabled = mnuEditSplit.Enabled
    
    mnuFilePreView.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    tlbThis.Buttons("Print").Enabled = blnData
    tlbThis.Buttons("PrintView").Enabled = blnData
    
End Sub


Private Sub formatMSH()
    Dim intCol As Integer
    With mshList
        .Cols = 29
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        .TextMatrix(0, HeadCol.ID) = "ID"
        .TextMatrix(0, HeadCol.�շ�ID) = "�շ�ID"
        .TextMatrix(0, HeadCol.�������) = "�������"
        .TextMatrix(0, HeadCol.�ƻ����) = "�ƻ����"
        .TextMatrix(0, HeadCol.��¼״̬) = "��¼״̬"
        
        .TextMatrix(0, HeadCol.�����־) = "�����־"
        
        .TextMatrix(0, HeadCol.���ݺ�) = "���ݺ�"
        .TextMatrix(0, HeadCol.�������) = "�������"
        .TextMatrix(0, HeadCol.��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, HeadCol.��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, HeadCol.��Ʊ���) = "��Ʊ���"
        .TextMatrix(0, HeadCol.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, HeadCol.Ʒ��) = "Ʒ��"
        .TextMatrix(0, HeadCol.���) = "���"
        .TextMatrix(0, HeadCol.��λ) = "��λ"
        .TextMatrix(0, HeadCol.����) = "����"
        .TextMatrix(0, HeadCol.����) = "����"
        .TextMatrix(0, HeadCol.����) = "����"
        .TextMatrix(0, HeadCol.�ɹ���) = "�ɹ���"
        .TextMatrix(0, HeadCol.�ɹ����) = "�ɹ����"
        
        .TextMatrix(0, HeadCol.������) = "������"
        .TextMatrix(0, HeadCol.��������) = "��������"
        .TextMatrix(0, HeadCol.�����) = "�����"
        .TextMatrix(0, HeadCol.�������) = "�������"
        
        .ColWidth(HeadCol.ID) = 0
        .ColWidth(HeadCol.�շ�ID) = 0
        .ColWidth(HeadCol.��λID) = 0
        .ColWidth(HeadCol.�������) = 0
        .ColWidth(HeadCol.�ƻ����) = 0
        .ColWidth(HeadCol.��¼״̬) = 0
    
        If mblnFirst = False Then
            SetEnabled
            Exit Sub
        End If
               
        .ColWidth(HeadCol.��Ӧ��) = 2000
        .ColWidth(HeadCol.���ݺ�) = 1400
        .ColWidth(HeadCol.�������) = 1400
        .ColWidth(HeadCol.��Ʊ��) = 1400
        .ColWidth(HeadCol.��Ʊ����) = 1400
        .ColWidth(HeadCol.��Ʊ���) = 1400
        .ColWidth(HeadCol.Ʒ��) = 2400
        .ColWidth(HeadCol.���) = 2000
        .ColWidth(HeadCol.��λ) = 800
        .ColWidth(HeadCol.����) = 1400
        .ColWidth(HeadCol.����) = 2000
        .ColWidth(HeadCol.����) = 1400
        .ColWidth(HeadCol.�ɹ���) = 1400
        .ColWidth(HeadCol.�ɹ����) = 1400
        .ColWidth(HeadCol.������) = 1000
        .ColWidth(HeadCol.��������) = 1600
        .ColWidth(HeadCol.�����) = 1000
        .ColWidth(HeadCol.�������) = 1600
        
        .ColAlignment(HeadCol.ID) = 1
        .ColAlignment(HeadCol.�շ�ID) = 1
        .ColAlignment(HeadCol.�������) = 1
        .ColAlignment(HeadCol.�ƻ����) = 1
        .ColAlignment(HeadCol.��¼״̬) = 1
        
        .ColAlignment(HeadCol.��Ӧ��) = 1
        .ColAlignment(HeadCol.������) = 4
        .ColAlignment(HeadCol.��������) = 4
        .ColAlignment(HeadCol.�����) = 4
        .ColAlignment(HeadCol.�������) = 4
        
        .ColAlignment(HeadCol.���ݺ�) = 4
        .ColAlignment(HeadCol.�������) = 1
        .ColAlignment(HeadCol.��Ʊ��) = 1
        .ColAlignment(HeadCol.��Ʊ����) = 4
        .ColAlignment(HeadCol.��Ʊ���) = 7
        .ColAlignment(HeadCol.Ʒ��) = 1
        .ColAlignment(HeadCol.���) = 1
        .ColAlignment(HeadCol.��λ) = 4
        .ColAlignment(HeadCol.����) = 1
        .ColAlignment(HeadCol.����) = 1
        .ColAlignment(HeadCol.����) = 7
        .ColAlignment(HeadCol.�ɹ���) = 7
        .ColAlignment(HeadCol.�ɹ����) = 7
        .ColAlignment(HeadCol.������) = 1
        
        If mbln�����־ Then
            .TextMatrix(0, HeadCol.��ǰ�ⷿ) = "��ǰ�ⷿ"
            .TextMatrix(0, HeadCol.��ǰ�ⷿ���) = "��ǰ�ⷿ���"
            .TextMatrix(0, HeadCol.ȫԺ���) = "ȫԺ���"
            .TextMatrix(0, HeadCol.ҩ�ⵥλ) = "ҩ�ⵥλ"
            .ColWidth(HeadCol.�����־) = 800
            .ColWidth(HeadCol.��ǰ�ⷿ) = 1400
            .ColWidth(HeadCol.��ǰ�ⷿ���) = 1400
            .ColWidth(HeadCol.ȫԺ���) = 1400
            .ColWidth(HeadCol.ҩ�ⵥλ) = 800
            .ColAlignment(HeadCol.��ǰ�ⷿ���) = 7
            .ColAlignment(HeadCol.ȫԺ���) = 7
        Else
            .ColWidth(HeadCol.�����־) = 0
            .ColWidth(HeadCol.��ǰ�ⷿ) = 0
            .ColWidth(HeadCol.��ǰ�ⷿ���) = 0
            .ColWidth(HeadCol.ȫԺ���) = 0
            .ColWidth(HeadCol.ҩ�ⵥλ) = 0
        End If
    End With
    SetEnabled
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    Set objPrint = New zlPrint1Grd
    
    If Me.ActiveControl Is mshSplit Then
        
        objRow.Add "Ӧ����λ:" & mshList.TextMatrix(mshList.Row, HeadCol.��Ӧ��)
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "Ӧ�����ݺ�:" & mshList.TextMatrix(mshList.Row, HeadCol.���ݺ�)
        objRow.Add "Ʒ��:" & mshList.TextMatrix(mshList.Row, HeadCol.Ʒ��)
        objRow.Add "���:" & mshList.TextMatrix(mshList.Row, HeadCol.���)
        objPrint.UnderAppRows.Add objRow
        
        objPrint.Title.Text = "����ƻ�����"
        Set objPrint.Body = mshSplit
    Else
        Dim strRange As String
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVerifyStartDate, "yyyy-mm-dd") = "1901-01-01" Then
            strRange = "������ڣ�" & Format(mdtVerifyStartDate, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtVerifyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            strRange = "�������ڣ�" & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������ڣ�" & Format(mdtVerifyStartDate, "yyyy��MM��dd��") & "��" & Format(mdtVerifyEndDate, "yyyy��MM��dd��")
        Else
            strRange = "�������ڣ�" & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
        End If
        objRow.Add strRange
        objPrint.Title.Text = "Ӧ��������"
        Set objPrint.Body = mshList
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡʱ�䣺" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub Ȩ�޿���()
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim blnDelete As Boolean
    Dim blnVerify As Boolean
    Dim blnCancel As Boolean
    Dim blnPlan As Boolean
    
    blnAdd = InStr(1, mstrPrivs, "�Ǽ�") <> 0
    blnModify = InStr(1, mstrPrivs, "�޸�") <> 0
    blnDelete = InStr(1, mstrPrivs, "ɾ��") <> 0
    blnVerify = InStr(1, mstrPrivs, "���") <> 0
    blnCancel = InStr(1, mstrPrivs, "����") <> 0
    blnPlan = InStr(1, mstrPrivs, "����ƻ�")
        
    mnuEditAdd.Visible = blnAdd
    mnuEditModify.Visible = blnModify
    mnuEditDelete.Visible = blnDelete
    mnuEditVerify.Visible = blnVerify
    mnuEditStrike.Visible = blnCancel
    mnuEditSplit.Visible = blnPlan
    
    tlbThis.Buttons("Add").Visible = blnAdd
    tlbThis.Buttons("Modify").Visible = blnModify
    tlbThis.Buttons("Delete").Visible = blnDelete
    
    tlbThis.Buttons("Verify").Visible = blnVerify
    tlbThis.Buttons("Strike").Visible = blnCancel
    
    tlbThis.Buttons("SplitDue").Visible = blnPlan
    
    
    If (Not blnAdd And Not blnModify And Not blnDelete) Or Not blnPlan Then
        tlbThis.Buttons("Split").Visible = False
        mnuEditLine1.Visible = False
    End If
    
    If (Not blnVerify And Not blnCancel) Or Not blnPlan Then
        mnuEditLine2.Visible = False
        tlbThis.Buttons("Split1").Visible = False
    End If
    
    If Not (blnAdd Or blnModify Or blnDelete Or blnPlan Or blnVerify Or blnVerify Or blnCancel) Then
        tlbThis.Buttons("Split2").Visible = False
        mnuEditLine3.Visible = False
    End If
    
End Sub

Private Sub setPlanGrdColor()
    Dim lngRow As Long
    
    With mshSplit
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                    '�Ѹ�
                    SetRowColor lngRow, &HFF8080, False
                Else
                    'δ��
                    
                End If
            Next
            
    End With
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub
