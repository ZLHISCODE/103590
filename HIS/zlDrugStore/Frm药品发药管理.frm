VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmҩƷ��ҩ���� 
   Caption         =   "ҩƷ������ҩ"
   ClientHeight    =   7560
   ClientLeft      =   3465
   ClientTop       =   1845
   ClientWidth     =   11400
   DrawMode        =   12  'Nop
   Icon            =   "FrmҩƷ��ҩ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11400
   Begin VB.CheckBox Chk��ʾ��ҩ�������� 
      Appearance      =   0  'Flat
      Caption         =   "��ʾ��ҩ��������"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7440
      TabIndex        =   45
      Top             =   720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�б� 
      Height          =   5415
      Left            =   30
      TabIndex        =   14
      Top             =   990
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicBackGroud 
      Height          =   5415
      Left            =   2280
      ScaleHeight     =   5355
      ScaleWidth      =   7245
      TabIndex        =   17
      Top             =   980
      Width           =   7305
      Begin VB.ComboBox cbo��ҩ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   48
         Text            =   "cbo��ҩ��"
         Top             =   4860
         Width           =   1215
      End
      Begin VB.PictureBox picRecipeColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   460
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   1095
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         Begin VB.Label lblRecipeType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ͨ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   47
            Top             =   105
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.CommandButton cmdAlley 
         Caption         =   "����ʷ/����״̬"
         Height          =   350
         Left            =   90
         TabIndex        =   37
         Top             =   630
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "��ҩ(&S)"
         Height          =   350
         Left            =   5910
         TabIndex        =   13
         ToolTipText     =   "�ȼ���F2"
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CheckBox Chkȫ�� 
         Appearance      =   0  'Flat
         Caption         =   "ȫ��(&A)"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6120
         TabIndex        =   30
         Top             =   4560
         Width           =   1005
      End
      Begin VB.TextBox Txt�շ�Ա 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4800
         TabIndex        =   31
         Top             =   4860
         Width           =   885
      End
      Begin VB.ComboBox Txt����ҽ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4860
         Width           =   1335
      End
      Begin ZL9BillEdit.BillEdit Bill������ϸ 
         Height          =   2655
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4683
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.ComboBox TxtNo 
         Height          =   315
         ItemData        =   "FrmҩƷ��ҩ����.frx":030A
         Left            =   4860
         List            =   "FrmҩƷ��ҩ����.frx":030C
         TabIndex        =   2
         Top             =   720
         Width           =   2325
      End
      Begin VB.PictureBox PicState 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6510
         ScaleHeight     =   375
         ScaleWidth      =   675
         TabIndex        =   21
         Top             =   90
         Visible         =   0   'False
         Width           =   675
         Begin VB.Label LblState 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   570
         TabIndex        =   4
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6390
         TabIndex        =   12
         Top             =   1080
         Width           =   795
      End
      Begin VB.TextBox TxtסԺ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4860
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3510
         TabIndex        =   8
         Top             =   1080
         Width           =   465
      End
      Begin VB.TextBox Txt�Ա� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2310
         TabIndex        =   6
         Top             =   1080
         Width           =   465
      End
      Begin VB.TextBox txtԭʼ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4440
         Width           =   525
      End
      Begin VB.TextBox txt��ҩ�巨 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4440
         Width           =   2955
      End
      Begin VB.Label Lbl�շ�Ա 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�շ�Ա"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4200
         TabIndex        =   32
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label Lbl��ҩ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2340
         TabIndex        =   18
         Top             =   4920
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4200
         TabIndex        =   1
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   120
         TabIndex        =   20
         Top             =   90
         Width           =   7140
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5970
         TabIndex        =   11
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label LblסԺ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4260
         TabIndex        =   9
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   7
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Lbl�Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1860
         TabIndex        =   5
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Lbl����ҽ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   4920
         Width           =   720
      End
      Begin VB.Label lbl��ҩ�巨 
         AutoSize        =   -1  'True
         Caption         =   "��ҩ�巨"
         Height          =   180
         Left            =   1980
         TabIndex        =   35
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lblԭʼ���� 
         AutoSize        =   -1  'True
         Caption         =   "ԭʼ����"
         Height          =   180
         Left            =   150
         TabIndex        =   34
         Top             =   4500
         Width           =   720
      End
   End
   Begin VB.Frame fraFind 
      Height          =   480
      Left            =   120
      TabIndex        =   40
      Top             =   6600
      Width           =   3975
      Begin VB.CommandButton cmdIC 
         Caption         =   "����"
         Height          =   300
         Left            =   3410
         TabIndex        =   43
         Top             =   135
         Width           =   495
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   690
         TabIndex        =   0
         Top             =   150
         Width           =   2325
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   3480
         Picture         =   "FrmҩƷ��ҩ����.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "������λ(F2)"
         Top             =   135
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgFilter 
         Height          =   240
         Left            =   3075
         Picture         =   "FrmҩƷ��ҩ����.frx":0458
         Top             =   150
         Width           =   240
      End
      Begin VB.Label lblFind 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨��"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   20
         TabIndex        =   42
         ToolTipText     =   "���˶�λ(F3)"
         Top             =   195
         Width           =   720
      End
   End
   Begin VB.PictureBox PicToolbar 
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   9510
      ScaleHeight     =   720
      ScaleWidth      =   1830
      TabIndex        =   38
      Top             =   15
      Width           =   1830
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������Ա"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   39
         Top             =   105
         Width           =   1500
      End
   End
   Begin VB.CheckBox Chk�嵥 
      Appearance      =   0  'Flat
      Caption         =   "��ʾ���й��̵���"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   720
      Width           =   1845
   End
   Begin VB.Timer TimeRefresh 
      Enabled         =   0   'False
      Left            =   5100
      Top             =   150
   End
   Begin VB.Timer TimePrintCancelBill 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5520
      Top             =   150
   End
   Begin VB.PictureBox PicCloseConsignment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   215
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   7440
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   6840
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar Cbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1164
      BandCount       =   1
      _CBWidth        =   11400
      _CBHeight       =   660
      _Version        =   "6.7.9782"
      Child1          =   "Tbar1"
      MinHeight1      =   600
      Width1          =   3000
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar1 
         Height          =   600
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "Cancel"
               Object.ToolTipText     =   "ȡ����ҩ"
               Object.Tag             =   "ȡ��"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Charge"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Stuff"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   7
            EndProperty
         EndProperty
         Begin VB.Timer TimePrint 
            Enabled         =   0   'False
            Left            =   4560
            Top             =   120
         End
         Begin MSComctlLib.ImageList imgPass 
            Left            =   8100
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   14
            ImageHeight     =   14
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmҩƷ��ҩ����.frx":6CAA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmҩƷ��ҩ����.frx":6F64
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmҩƷ��ҩ����.frx":721E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmҩƷ��ҩ����.frx":74D8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmҩƷ��ҩ����.frx":7792
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   7200
      Width           =   11400
      _ExtentX        =   20108
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
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsfPrint 
      Height          =   2985
      Left            =   390
      TabIndex        =   25
      Top             =   2550
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5265
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   345
      Left            =   0
      TabIndex        =   28
      Top             =   630
      Width           =   3950
      _ExtentX        =   6959
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "����ҩ(&1)"
      TabPicture(0)   =   "FrmҩƷ��ҩ����.frx":7A4C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "����ҩ(&2)"
      TabPicture(1)   =   "FrmҩƷ��ҩ����.frx":7A68
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "����ҩ(&3)"
      TabPicture(2)   =   "FrmҩƷ��ҩ����.frx":7A84
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "��ҩ(&4)"
      TabPicture(3)   =   "FrmҩƷ��ҩ����.frx":7AA0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin VB.Image img���� 
      Height          =   240
      Left            =   4320
      Picture         =   "FrmҩƷ��ҩ����.frx":7ABC
      ToolTipText     =   "ѡ����"
      Top             =   6720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgLeftRight_S 
      Height          =   5385
      Left            =   3720
      MousePointer    =   9  'Size W E
      Top             =   990
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu MnuFileSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu MnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu MnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu MnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileBillprint 
         Caption         =   "��ӡ��ҩ��(&B)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MnuFileRePrint 
         Caption         =   "��ӡ����ǩ(&D)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileReport 
         Caption         =   "��ӡ��ҩ�嵥(&W)"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "��ӡ��ҩ֪ͨ��(&R)"
      End
      Begin VB.Menu mnuFileLable 
         Caption         =   "��ӡҩƷ��ǩ(&L)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileBack 
         Caption         =   "��ӡ�˷ѵ���(T)"
      End
      Begin VB.Menu MnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFilePara 
         Caption         =   "��������(&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu MnuEditDosage 
         Caption         =   "��ҩģʽ(&D)"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuEditAbolish 
         Caption         =   "ȡ��ģʽ(&A)"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuEditConsignment 
         Caption         =   "��ҩģʽ(&C)"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditHandback 
         Caption         =   "��ҩģʽ(&H)"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
      Begin VB.Menu MnuEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditBatch 
         Caption         =   "������ҩ(&B)"
      End
      Begin VB.Menu MnuEditSendOther 
         Caption         =   "������ҩ���Ĵ���(&F)"
      End
      Begin VB.Menu MnuEditHandbackBatch 
         Caption         =   "������ҩ���Ĵ���(&T)"
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "��Ʊ�ݺŷ�ҩ(&I)"
      End
      Begin VB.Menu mnuEditBillRestore 
         Caption         =   "��Ʊ�ݺ���ҩ(&R)"
      End
      Begin VB.Menu mnuline9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlag 
         Caption         =   "ֹͣ��ҩ���(&S)"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "ȡ����ҩ(&Q)"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuCharge 
         Caption         =   "���ﻮ��(&M)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuStuff 
         Caption         =   "���ķ���(@W)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "�л���ҩ��(&E)"
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
      Begin VB.Menu MnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu MnuViewToolS 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu sdfsdfsd 
            Caption         =   "-"
         End
         Begin VB.Menu MnuViewToolT 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu MnuViewState 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&O)"
         Begin VB.Menu mnuViewFontSET 
            Caption         =   "С����(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSET 
            Caption         =   "������(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSET 
            Caption         =   "������(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu MnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLocate 
         Caption         =   "��λ��ʽ(&S)"
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "���￨(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "���ݺ�(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "�����(&3)"
            Index           =   2
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "����(&4)"
            Index           =   3
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "���֤(&5)"
            Index           =   4
         End
         Begin VB.Menu mnuViewLocateItem 
            Caption         =   "IC��(&6)"
            Index           =   5
         End
      End
      Begin VB.Menu mnuView4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu MnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuHelpWeb 
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu MnuHelpWebM 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu MnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuPass 
      Caption         =   "Pass"
      Visible         =   0   'False
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩ���ٴ���Ϣ�ο�(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩƷ˵����(&D)"
         Index           =   1
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "�й�ҩ��(&N)"
         Index           =   2
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "������ҩ����(&S)"
         Index           =   3
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "����ֵ(&T)"
         Index           =   4
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ר����Ϣ(&P)"
         Index           =   6
         Begin VB.Menu mnuPassSpec 
            Caption         =   "ҩ��-ҩ���໥����(&D)"
            Index           =   0
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "ҩ��-ʳ���໥����(&F)"
            Index           =   1
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����ע�������(&M)"
            Index           =   3
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����ע�������(&T)"
            Index           =   4
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����֢(&C)"
            Index           =   6
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "������(&S)"
            Index           =   7
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&G)"
            Index           =   9
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��ͯ��ҩ(&P)"
            Index           =   10
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&E)"
            Index           =   11
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&L)"
            Index           =   12
         End
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҽҩ��Ϣ����(&I)"
         Index           =   8
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩƷ�����Ϣ(&M)"
         Index           =   10
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "��ҩ;�������Ϣ(&R)"
         Index           =   11
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҽԺҩƷ��Ϣ(&F)"
         Index           =   12
      End
   End
End
Attribute VB_Name = "FrmҩƷ��ҩ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--ע�����ر���--
Private intFont As Integer                              '����
Private IntShowCol As Integer                           '�ڴ�����ϸ���Ƿ���ʾ����(0)

Private mintShowBill�շ� As Integer                     '��ʾ�շѴ�����Χ
Private mintShowBill���� As Integer                     '��ʾ���ʴ�����Χ
Private mstrShowBill As String                          '��ѯSQL�Ӵ�
Private mstrShowSendedBill As String                    '��ѯSQL�Ӵ����������ѷ�ҩ����

Private IntAutoPrint As Integer                         '��ҩ���ӡ������(1)
Private intУ����ҩ�� As Integer                        '��ҩʱ�Ƿ�У����ҩ��
Private intУ�鷢ҩ�� As Integer                        '��ҩʱ�Ƿ�У�鷢ҩ��
Private intҩƷ���� As Integer                          'ҩƷ������ʾ��ʽ��0-����������;1-������;2-������

Private intPrint As Integer                             '����ӡδ��ҩ����(0)
Private mbln���ʵ� As Boolean                           '��ӡ��ҩ��ʱ�Ƿ�������ʵ�
Private strPrintWindow As String                        '��ӡδ��ҩ����Ϊ3ʱ��Ч
Private mbln���￨ As Boolean                           '�Ƿ��Զ���λ�����￨
Private mlng�������� As Long                            '�Ƿ���ʾ��ҩ��������
Private mint��Ժ��ҩ As Integer
Private mint�Զ���ҩ As Integer                         '�Ƿ�ʹ���Զ���ҩ���ܣ�0-��ʹ�ã�1-ʹ��
Private mint�Զ���ҩʱ�� As Integer                     '������ʱ�޾���Ҫ��֤��ҩ�ˣ�Ĭ��Ϊʼ�ղ���֤��ҩ��
Private mint����ģʽ As Integer

'0-����ӡδ��ҩ����
'1-��ӡ����������δ��ҩ����
'2-��ӡ����������δ��ҩ����
'3-ѡ���ӡ(��ҩ����)
Private mlngRefresh As Long                             'ˢ�¼��(0)
Private mlngPrintInterval As Long                       '��ӡ��ҩ�����(0)
Private mIntPrintDelay As Integer                       '�ӳٴ�ӡ(60)
Private mIntPrintHandbackNO As Integer                  '��ӡ�˷ѵ��ݺ�(0)
Private mintPrintDrugLable  As Integer                  '��ӡҩƷ��ǩ
Private lngҩ��ID As Long                               'ҩ��(���ñ�������Ӧ��ҩ��)
Private Str��ҩ�� As String                             '������ҩ��
Private mstr�Զ���ҩ�� As String                        '�����Զ���ҩ������
Private Str���� As String                               '��ҩ����(���ñ�������Ӧ�ķ�ҩ����)
Private IntTimes As Integer                             '���ӳ�
Private intVerify As Integer                            '�Ƿ���ҪУ�鴦��
Private BlnEnterCell As Boolean                         '�Ƿ�������ENTERCELL()�¼�
Private str��� As String                               '���浱ǰ����ҩ������ϸ��ż�

Private mstrOracleMoneyForamt As String                 'ORACLE�н���ʽ
Private mstrVBMoneyForamt As String                     'VB�н���ʽ

'--ϵͳ����--
Private StrFindStyle As String                          'ƥ�䴮
Private IntCheckStock As Integer                        '�����
Private IntSendAfterDosage As Integer                   '�Ƿ���뾭����ҩ����(0)
Private mblnStarPass As Boolean                         '���ú�����ҩ(PASS)
Public Int����δ��˴�����ҩ As Integer                 'δ����Ƿ�����ҩ
Public mint����δ�շѴ�����ҩ As Integer                'δ�շ��Ƿ�����ҩ
Private blnҽ������ As Boolean                          '�Ƿ�����δ����ҽ����ҩ
Private int����λ�� As Integer                      '���ý���λ��
Private int��˻��۵� As Integer                        'ִ�к��Զ���˻��۵�
Private mint�Զ����� As Integer
Private mbln�����������۷��� As Boolean
Private mbln��ʾ��С��λ As Boolean

'--�������--
Private BlnStartUp As Boolean                           '�����ɹ�
Private BlnFirstStart As Boolean                        '��һ������
Private LngSendRow As Long                              '����
Private BlnInRefresh As Boolean                         '�Ƿ���ˢ��״̬
Private BlnInOper As Boolean                            '�Ƿ�����NO��
Private mstrFilter  As String                           '�����������
Private mrsBatchSend As ADODB.Recordset                 '����������ҩ
Private mblnFilterRefresh   As Boolean
Private mbln����ȡ����ҩ As Boolean                     '�Ƿ������δ��ҩƷ����ȡ����ҩ����
Private mstr����Ա As String
Private mstr��ҩ�� As String
Private mdate�ϴ�У��ʱ�� As Date
Private mblnIsFirst As Boolean                          'δУ��
Private mblnAuto As Boolean

Private mstrStartDate As String
Private mstrEndDate As String

Private mstrPrintRecipe As String                       '���ڷ�ҩ���ӡ����¼���ݺš��������ͣ����ݺ�1,��������1|���ݺ�2,��������2......

Private mblnDrop As Boolean                     '��KeyDown���ж������б��Ƿ񵯳�

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private mblnStateTimeRefresh As Boolean
Private mblnStateTimePrint As Boolean

'--������ʹ�ü�¼��--
Private RecPhysic As New ADODB.Recordset                'ҩƷ��¼
Private RecPart As New ADODB.Recordset                  '���ű�

Private mrsPASS As New ADODB.Recordset                  'PASS�����ݼ�

'--����--
Private BlnAllowClick As Boolean                        '����ִ��Click�¼�
Private strUnit As String                               '��λ����
Private str��λ�� As String                             '��λ��
Private mInt���� As Integer                             '��������  0-���ＰסԺ���е��� 8-���ﻮ�ۼ�������� 9-סԺ����
Private IntBillStyle As Integer                         '����
Private mstrNo As String                                'NO
Private mint�����־ As Integer                         '��ǰ���ݵ������־ 1-����;2-סԺ
Private mint��¼���� As Integer                         '��ǰ���ݵļ�¼���� 1-�շѼ�¼;2-���ʼ�¼
Private StrLastNo As String                             '�ϴ�ѡ��������NO
Private IntLastBill As Integer                          '�ϴ�ѡ�������ĵ���
Private strLastData As String                           '�ϴ�ѡ�������ĵ��ݵ����ƻ��������
Private mintLastSequence As Integer                     '������ϸ�б���ϴ�ѡ������
Private StrFind_1 As String                             'δ��ҩ�������Ҵ�
Private StrFind_2 As String                             '����ҩ�������Ҵ�
Private StrFind_3 As String                             'δ��ҩ�������Ҵ�
Private StrFind_4 As String                             '�ѷ�ҩ�������Ҵ�
Private StrDate As String                               '��ǰϵ������
Private strBill As String                               '��¼�����ѷ�ҩ�����ż���������
Private mblnAllBack As Boolean                          '�Ƿ�ȫ��
Private mblnCard As Boolean                             '�Ƿ�ˢ���￨
    
Private mblnIs��ҩ���� As Boolean                        '��ǰ�����Ƿ�Ϊ��ҩ����
Private mstr��������ʾ As String
Private mstr�۸�ʧЧ��ʾ As String
Private mbln��ʾ���� As Boolean

'PASS
Private mstr������λ As String
Private mlng����ID As Long
Private mlngPassPati As Long
Private mlng��ҳID As Long
Private mstr�Һŵ� As String

Private Const mlng��ɫ As Long = &HC000C0
'--����ʽ--
Private strOrder_1 As String                            'δ��ҩ��������
Private strOrder_2 As String                            '����ҩ��������
Private strOrder_3 As String                            'δ��ҩ��������
Private strOrder_4 As String                            '�ѷ�ҩ��������

'--���ز���--
Private mstrSourceDep As String                         '��Դ���Ҵ�
Public BlnSetParaSuccess As Boolean                     '���óɹ����
Public Intģʽ As Integer
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
Private strChargePrivs As String                        '���ﻮ��Ȩ�޴�
Private strStuffPrivs As String                         '���ķ��Ź���Ȩ�޴�
Private BlnRefresh As Boolean
Private mbln���������� As Boolean

Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private Enum ��������
    ��ɫ = 0
    �������� = 1
    ѡ�� = 2
    ��־ = 3
    ���� = 4
    ���� = 5
    �շ� = 6
    ��ҩ�� = 7
    NO = 8
    ���� = 9
    ��� = 10
    ���� = 11
    �ɲ��� = 12
    ˵�� = 13
    ���￨�� = 14
    ����� = 15
    ���֤ = 16
    IC�� = 17
    ����ID = 18

    '����ҩ����
    δ��� = 19
    ʵ�ս�� = 20
        
    '��ҩ����
    �����־ = 19
    ��¼���� = 20
        
    '��ҩ����ҩ����
    ��ҩ���� = 21
    ��ҩ���� = 21
End Enum

Private Enum ����
    ����� = 0
    ˳��� = 1       '���е����
    ҩƷ���� = 2
    ������ = 3
    Ӣ���� = 4
    ��� = 5
    ��� = 6
    ���� = 7
    Id = 8
    ҩƷID = 9
    ���� = 10
    ��λ = 11
    ���� = 12
    ���� = 13
    ���� = 14
    ��� = 15
    ���� = 16
    ���� = 17
    �÷� = 18
    Ƶ�� = 19
    ҽ������ = 20
    �ѱ� = 21
    ����� = 22
    ��λ = 23
    ������ = 24
    ׼���� = 25
    ׼������ = 26
    ׼����С = 27
    ��ҩ�� = 28
    ��ҩ���� = 29
    ��λ�� = 30
    ��ҩ��С = 31
    ��λС = 32
    ���� = 33
    ������ = 34
    ��Ч�� = 35
    �²��� = 36
    ��ע = 37
    ҽ��id = 38
    ʵ������ = 39
    ��װ = 40
    ���� = 41
End Enum

Private Type Type_SQLCondition
    date��ʼ���� As Date
    date�������� As Date
    str��ʼNO As String
    str����NO As String
    str���� As String
    str���￨ As String
    str��ʶ�� As String
    lng����ID As Long
    str������ As String
    str����� As String
    lngҩƷID As Long
    str��ǰNO As String
    str����� As String
    str���֤ As String
    strIC�� As String
    strҽ���� As String
End Type

Private SQLCondition As Type_SQLCondition

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object

Private Enum FindType
    ���￨ = 0
    ���ݺ� = 1
    ����� = 2
    ���� = 3
    ���֤ = 4
    IC�� = 5
End Enum

Private Const cstLocate As Integer = 0
Private Const cstFilter As Integer = 1

'�������ͣ���ͨ�����ơ������������һ������
Private Enum ��������
    ��ͨ = 0
    ���� = 1
    ���� = 2
    ���� = 3
    ��һ = 4
    ���� = 5
End Enum

'�����������ƣ���˳����;�ָ�
Private Const mconstrRecipeType = "��ͨ;����;����;����;��һ;����"

'Ĭ�ϴ�����ɫ����ͨ����ɫ���������ɫ�����ƣ�����ɫ��������һ������ɫ����������ɫ
Private Const mconlng��ͨ = &HFFFFFF
Private Const mconlng���� = &HC0FFC0
Private Const mconlng���� = &HC0FFFF
Private Const mconlng���� = &HFFFFFF
Private Const mconlng��һ = &HC0C0FF
Private Const mconlng���� = &HC0C0FF

'�û�����Ĵ�����ɫ����ע���ȡ���ַ�������;�ָ�
Private mstrUserRecipeColor As String

Private Function CheckBatchRecipe() As Boolean
    Dim n As Integer
    Dim rsTemp As ADODB.Recordset
    Dim BlnFirst As Boolean
    Dim lngRow As Long, lngҩƷID As Long, LngID As Long, lng���� As Long, lng���� As Long
    Dim blnBatchSend As Boolean
    Dim i As Integer
    
    On Error GoTo ErrHand
       
    '��鲡�˷������
    If Not CheckSendBillMoney(True) Then Exit Function
    
    For n = 1 To Msf�б�.Rows - 1
        If Val(Msf�б�.TextMatrix(n, ��������.��־)) = 1 Then
            Msf�б�.Row = n
            Call Msf�б�_EnterCell
            DoEvents
            
            '���ҩƷ�洢�ⷿ
            If CheckDrugStock = False Then Exit Function
            
            '����Ƿ�����
            If CheckBill(3, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Function
            
            '����Ƿ��շ�(��ҩ����)
            gstrSQL = " Select Decode(��ҩ��,Null,'','���ŷ�ҩ','',��ҩ��) ��ҩ��,���շ� From δ��ҩƷ��¼" & _
                     " Where No=[1] And (�ⷿID=[3] Or �ⷿID Is NULL) And ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex), lngҩ��ID)
            
            With rsTemp
                If IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
                    If IntSendAfterDosage = 0 Then
                        If IsNull(!��ҩ��) Then
                            MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If Trim(!��ҩ��) = "" Then
                            MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
                mstr��ҩ�� = NVL(!��ҩ��)
                
                If mint����δ�շѴ�����ҩ = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) = 8 Then
                    If !���շ� = 0 Then
                        MsgBox "�ô�����δ�շѣ�����ִ�з�ҩ������", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                If Int����δ��˴�����ҩ = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) <> 8 Then
                    If !���շ� = 0 Then
                        MsgBox "�ô�����δ��ˣ�����ִ�з�ҩ������", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                Call GetBillSequence
                If str��� = "" Then Exit Function
                If Not IsReceiptBalance(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), str���) Then Exit Function
                If Not IsOutPatient(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then Exit Function
                If Not CheckBillControl(tabShow.Tab + 1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), Msf�б�.TextMatrix(n, ��������.���)) Then Exit Function
                
                'У�鷢ҩ��
                If intУ�鷢ҩ�� = 1 And Not BlnFirst Then
                    mstr����Ա = zlDatabase.UserIdentify(Me, "У�鷢ҩ��", glngSys, 1341, "��ҩ")
                    BlnFirst = True
                Else
                    mstr����Ա = gstrUserName
                End If
                If mstr����Ա = "" Then Exit Function
                    
                If Not CheckSpec(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
                
                If mstr��������ʾ <> "" Then
                    If MsgBox("����Ϊ[" & TxtNo & "]" & "�Ĵ����к������¶�����ҩƷ��ȷ����ҩ��" & mstr��������ʾ, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                    mstr��������ʾ = ""
                End If
                
                If Not CheckStock(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
            End With
        End If
    Next
    
    CheckBatchRecipe = True
    Exit Function
ErrHand:
    CheckBatchRecipe = False
End Function

Private Function CheckDrugStock() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Integer
    
    For lngRow = 1 To Bill������ϸ.Rows - 2
        gstrSQL = "Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1] And �շ�ϸĿid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�洢�ⷿ", lngҩ��ID, Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID)))
        
        If rsTmp.EOF Then
            MsgBox Bill������ϸ.TextMatrix(lngRow, ����.ҩƷ����) & "δ���ô洢�ⷿ�����ܷ�ҩ��", vbInformation, gstrSysName
            Exit Function
        End If
    Next

    CheckDrugStock = True
End Function
Private Function CheckBillExist(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select ID From ҩƷ�շ���¼ " & _
             " Where ����=[1] And NO=[2] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鵥���Ƿ����", int����, strNo)
    CheckBillExist = Not rsTemp.EOF
End Function
Private Function CheckIsSended(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    '����Ƿ�����ҩ
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Count(Id) From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ��¼״̬ <> 1 And ������� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�����ȡ����ҩ", int����, strNo)
    
    CheckIsSended = (rsTemp.RecordCount > 0)
End Function

Private Function CheckRecipe() As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long, lngҩƷID As Long, LngID As Long, lng���� As Long, lng���� As Long
    
    On Error GoTo ErrHand
    
    '��鲡�˷������
    If Not CheckSendBillMoney(False) Then Exit Function
    
    '���ҩƷ�洢�ⷿ
    If CheckDrugStock = False Then Exit Function
    
    '����Ƿ�����
    If CheckBill(3, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Function
    '����Ƿ��շ�(��ҩ����)
    gstrSQL = " Select Decode(��ҩ��,Null,'','���ŷ�ҩ','',��ҩ��) ��ҩ��,���շ� From δ��ҩƷ��¼" & _
             " Where No=[1] And (�ⷿID=[3] Or �ⷿID Is NULL) And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex), lngҩ��ID)
    
    With rsTemp
        If IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
            If IntSendAfterDosage = 0 Then
                If IsNull(!��ҩ��) Then
                    MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                End If
                If Trim(!��ҩ��) = "" Then
                    MsgBox "�ô�����δ��ҩ������ִ�з�ҩ������", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        mstr��ҩ�� = NVL(!��ҩ��)
        
        If mint����δ�շѴ�����ҩ = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) = 8 Then
            If !���շ� = 0 Then
                MsgBox "�ô�����δ�շѣ�����ִ�з�ҩ������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If Int����δ��˴�����ҩ = 0 And Val(TxtNo.ItemData(TxtNo.ListIndex)) <> 8 Then
            If !���շ� = 0 Then
                MsgBox "�ô�����δ��ˣ�����ִ�з�ҩ������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If intVerify = 1 And Txt����ҽ��.ListIndex = 0 Then Txt����ҽ��.Enabled = True: Txt����ҽ��.SetFocus: Exit Function
        Call GetBillSequence
        If str��� = "" Then Exit Function
        If Not IsReceiptBalance(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), str���) Then Exit Function
        If Not IsOutPatient(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then Exit Function
        If Not CheckBillControl(tabShow.Tab + 1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), Msf�б�.TextMatrix(Msf�б�.Row, ��������.���)) Then Exit Function
        
        'У�鷢ҩ��
        If intУ�鷢ҩ�� = 1 Then
            mstr����Ա = zlDatabase.UserIdentify(Me, "У�鷢ҩ��", glngSys, 1341, "��ҩ")
        Else
            mstr����Ա = gstrUserName
        End If
        If mstr����Ա = "" Then Exit Function
        
        If Not CheckSpec(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
        
        If mstr��������ʾ <> "" Then
            If MsgBox("����Ϊ[" & TxtNo & "]" & "�Ĵ����к������¶�����ҩƷ��ȷ����ҩ��" & mstr��������ʾ, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("��ȷ������Ϊ[" & TxtNo & "]" & "�Ĵ�����ҩ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        End If
        
        If Not CheckStock(Mid(TxtNo.Text, 1, 8), TxtNo.ItemData(TxtNo.ListIndex)) Then Exit Function
    End With
    
    CheckRecipe = True
    Exit Function
ErrHand:
    CheckRecipe = False
End Function

Private Function CheckSendBillMoney(ByVal blnBatch As Boolean) As Boolean
    '��ҩ��飭��鲡�˷����������ݼ��ʱ�����������Ӧ����
    'blnBatch��True-������ҩ;False-��������ҩ
    '��Ҫ�㷨��
    '1��ϵͳ����"ִ�к��Զ����"��Чʱ�ż��
    '2��ֻ�Լ��ʻ��۵�
    '3��������ID���㵥�ݻ��ܽ��
    '4�����ݼ��ʱ�����������Ӧ����
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rs������� As ADODB.Recordset
    Dim strNo As String
    Dim lng����ID As Long
    Dim str����id As String
    Dim strFirstNo As String
    
    Dim cur������� As Currency
    
    Dim str������� As String
    Dim str��������� As String
    
    On Error GoTo errH
    
    'ϵͳ����"ִ�к��Զ����"��Чʱ�ż��
    If int��˻��۵� = 0 Then
        CheckSendBillMoney = True
        Exit Function
    End If
    
    If blnBatch Then
        With mrsBatchSend
            'ֻ�Լ��ʻ��۵��ż��
            .Filter = "����=9 And δ���=1"
            
            '������ID���㵥�ݻ��ܽ��
            .Sort = "����ID"
            
            If .RecordCount = 0 Then
                CheckSendBillMoney = True
                Exit Function
            End If
            
            .MoveFirst
            
            '���ݼ��ʱ�����������Ӧ����
            Do While Not .EOF
                If lng����ID <> Val(!����ID) Then
                    If lng����ID <> 0 Then
                        '�ж���סԺ�������ﲡ��
                        gstrSQL = "Select Distinct Decode(B.�����־, 1, '����', 4, '����', 'סԺ') As ��Դ, " & _
                            " B.����id,nvl(B.��ҳid,0) ��ҳid,Decode(B.�����־, 1, 0, 4, 0, B.���˲���id) ���˲���id, C.���� " & _
                            " From ҩƷ�շ���¼ A,���˷��ü�¼ B,������Ϣ C " & _
                            " Where A.����id=B.Id And b.����id = c.����id " & _
                            " And A.����=9 And A.no=[1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                        
                        'ȡ�������
                        gstrSQL = " Select Distinct b.����, b.���� " & _
                            " From ���˷��ü�¼ a, �շ���Ŀ��� b, ҩƷ�շ���¼ c " & _
                            " Where a.�շ���� = b.���� And a.Id = c.����id And c.���� = 9 And c.No In([1]) "
                        Set rs������� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                        
                        Do While Not rs�������.EOF
                            str������� = str������� & rs�������!����
                            str��������� = str��������� & "," & rs�������!����
                            rs�������.MoveNext
                        Loop
                                            
                        '���������
                        If Not FinishBillingWarn(rsTmp, cur�������, str�������, str���������) Then
                            CheckSendBillMoney = False
                            Exit Function
                        End If
                    End If
                    
                    strNo = !NO
                    cur������� = Val(!���)
                    strFirstNo = !NO
                    lng����ID = Val(!����ID)
                Else
                    strNo = strNo & "," & !NO
                    cur������� = cur������� + Val(!���)
                End If
                
                .MoveNext
                
                If .EOF Then
                    '�ж���סԺ�������ﲡ��
                    gstrSQL = "Select Distinct Decode(B.�����־, 1, '����', 4, '����', 'סԺ') As ��Դ, " & _
                        " B.����id,nvl(B.��ҳid,0) ��ҳid,Decode(B.�����־, 1, 0, 4, 0, B.���˲���id) ���˲���id, C.���� " & _
                        " From ҩƷ�շ���¼ A,���˷��ü�¼ B,������Ϣ C " & _
                        " Where A.����id=B.Id And b.����id = c.����id " & _
                        " And A.����=9 And A.no=[1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFirstNo)
                    
                    'ȡ�������
                    gstrSQL = " Select Distinct b.����, b.���� " & _
                        " From ���˷��ü�¼ a, �շ���Ŀ��� b, ҩƷ�շ���¼ c " & _
                        " Where a.�շ���� = b.���� And a.Id = c.����id And c.���� = 9 And c.No In([1]) "
                    Set rs������� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
                    
                    Do While Not rs�������.EOF
                        str������� = str������� & rs�������!����
                        str��������� = str��������� & "," & rs�������!����
                        rs�������.MoveNext
                    Loop
                                        
                    '���������
                    If Not FinishBillingWarn(rsTmp, cur�������, str�������, str���������) Then
                        CheckSendBillMoney = False
                        Exit Function
                    End If
                End If
            Loop
        End With
    Else
        If Val(TxtNo.ItemData(TxtNo.ListIndex)) <> 9 Or Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.δ���)) <> 1 Then
            CheckSendBillMoney = True
            Exit Function
        End If
        
        strNo = Mid(TxtNo.Text, 1, 8)
        
        cur������� = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.ʵ�ս��))
        
        '�ж���סԺ�������ﲡ��
        gstrSQL = "Select Distinct Decode(B.�����־, 1, '����', 4, '����', 'סԺ') As ��Դ, " & _
            " B.����id,nvl(B.��ҳid,0) ��ҳid,Decode(B.�����־, 1, 0, 4, 0, B.���˲���id) ���˲���id, C.���� " & _
            " From ҩƷ�շ���¼ A,���˷��ü�¼ B,������Ϣ C " & _
            " Where A.����id=B.Id And b.����id = c.����id " & _
            " And A.����=9 And A.no=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
        
        'ȡ�������
        gstrSQL = " Select Distinct b.����, b.���� " & _
            " From ���˷��ü�¼ a, �շ���Ŀ��� b, ҩƷ�շ���¼ c " & _
            " Where a.�շ���� = b.���� And a.Id = c.����id And c.���� = 9 And c.No In([1]) "
        Set rs������� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo)
        
        Do While Not rs�������.EOF
            str������� = str������� & rs�������!����
            str��������� = str��������� & "," & rs�������!����
            rs�������.MoveNext
        Loop
                            
        '���������
        If Not FinishBillingWarn(rsTmp, cur�������, str�������, str���������) Then
            CheckSendBillMoney = False
            Exit Function
        End If
    End If
    CheckSendBillMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub GetDosagePeople()
    Dim rsTemp As ADODB.Recordset
    '��ҩ��
    gstrSQL = " Select ����||'-'||���� As ���� From ��Ա��  Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in " & _
             " (Select Distinct ��ԱID From ��Ա����˵�� Where ��Ա����='ҩ����ҩ��' " & _
             " And ��ԱID IN (Select ��ԱID From ������Ա Where ����ID=[1]))" & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
    
    With rsTemp
        Me.cbo��ҩ��.Clear
        Do While Not .EOF
            cbo��ҩ��.AddItem !����
            .MoveNext
        Loop
        If cbo��ҩ��.ListCount = 0 Then
            cbo��ҩ��.Enabled = False
        End If
    End With
End Sub

Private Sub GetRecipeColor()
    mstrUserRecipeColor = zlDatabase.GetPara("������ɫ", glngSys, 1341)

    If mstrUserRecipeColor = "" Then
        Call GetDefaultRecipeColor
    End If
End Sub

Private Sub GetDefaultRecipeColor()
    mstrUserRecipeColor = CStr(mconlng��ͨ) & ";" & _
                    CStr(mconlng����) & ";" & _
                    CStr(mconlng����) & ";" & _
                    CStr(mconlng����) & ";" & _
                    CStr(mconlng��һ) & ";" & _
                    CStr(mconlng����)

End Sub
Private Function GetSumMoney(ByVal rsRecipt As ADODB.Recordset) As String
    Dim rsTemp As ADODB.Recordset
    Dim dblSum As Double
    
    Set rsTemp = rsRecipt.Clone
    
    With rsTemp
        .MoveFirst
        Do While Not .EOF
            dblSum = dblSum + Val(.Fields("���").Value)
            .MoveNext
        Loop
    End With
    
    GetSumMoney = FormatEx(dblSum, mintMoneyDigit)
End Function
Private Sub GetSysParms()
    Int����δ��˴�����ҩ = gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ
    mint����δ�շѴ�����ҩ = gtype_UserSysParms.P148_δ�շѴ�����ҩ
    
    mbln����ȡ����ҩ = (gtype_UserSysParms.P15_�����շ��뷢ҩ���� = 1 Or gtype_UserSysParms.P16_סԺ�����뷢ҩ���� = 1)
    
    blnҽ������ = (gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = 0)          'Ϊ���ʾ������ҩ
    
    '��ȡ���С��λ��
    int����λ�� = gtype_UserSysParms.P9_���ý���λ��
    
    '�жϻ��۵���ҩ���Ƿ��Զ����Ϊ���ʵ�
    int��˻��۵� = gtype_UserSysParms.P81_ִ�к��Զ���˻��۵�
    
    '���ʱ����������۷���
    mbln�����������۷��� = gtype_UserSysParms.P98_���ʱ����������۷��� <> 0

End Sub

Private Sub IniRecord()
    Set mrsBatchSend = New ADODB.Recordset
    With mrsBatchSend
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "δ���", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
Private Sub PrintRecipe()
    '��ӡ����
    Dim blnPrint As Boolean
    Dim arrRecipe
    Dim n As Integer
    Dim intNum As Integer
    Dim strRecipeNo As String
    Dim intBillType As Integer
    
    If mstrPrintRecipe = "" Then Exit Sub
    
    If IntAutoPrint < 2 Then
        blnPrint = IIf(IntAutoPrint = 1, True, False)
        If IntAutoPrint = 0 Then
            If MsgBox("��ӡ�ô���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnPrint = True
        End If
        
        If blnPrint Then
            arrRecipe = Split(mstrPrintRecipe, "|")
            intNum = UBound(arrRecipe)
            
            For n = 0 To intNum
                strRecipeNo = Split(arrRecipe(n), ",")(0)
                intBillType = Val(Split(arrRecipe(n), ",")(1))
            
                If Not BillHaveHerial(strRecipeNo, intBillType) Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                        "NO=" & strRecipeNo, _
                        "����=" & IIf(intBillType = 8, 1, 2), _
                        "ҩ��=" & lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
                        "ReportFormat=1", "PrintEmpty=0", 2)
                Else
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                        "NO=" & strRecipeNo, _
                        "����=" & IIf(intBillType = 8, 1, 2), _
                        "ReportFormat=1", "PrintEmpty=0", 2)
                End If
            Next
        End If
    End If
    
    Me.MnuFileRePrint.Caption = "�ش��ѷ�ҩ����-" & strRecipeNo & "(&D)"
    
    mstrPrintRecipe = ""
End Sub

Private Sub Select����()
    Dim rsTmp As ADODB.Recordset
    
    If cbo����.ListCount > 0 Then Exit Sub
    
    '����
    gstrSQL = " Select ����||'-'||���� ����,ID From ���ű� " & _
             " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ID in (Select ����ID From ��������˵�� Where ��������='����' And ������� IN(2,3))" & _
             " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order By ����||'-'||���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����")
    
    With cbo����
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            .ItemData(.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        If .ListIndex <> -1 Then
            .ListIndex = 0
        End If
    End With
End Sub

Private Function SendBatchRecipe() As Boolean
    Dim n As Integer
    Dim lngRow As Long, lngҩƷID As Long, LngID As Long, lng���� As Long, lng���� As Long
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim rsSendRecipeDetail As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rsSendRecipeDetail = New ADODB.Recordset
    With rsSendRecipeDetail
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    For n = 1 To Msf�б�.Rows - 1
        If Val(Msf�б�.TextMatrix(n, ��������.��־)) = 1 Then
            Msf�б�.Row = n
            Call Msf�б�_EnterCell
            DoEvents
            
            With rsSendRecipeByNo
                .AddNew
                !NO = Mid(TxtNo.Text, 1, 8)
                !���� = TxtNo.ItemData(TxtNo.ListIndex)
                !��ҩ�� = cbo��ҩ��.Text
                !������ = IIf(Txt����ҽ��.ListIndex = 0, "", Mid(Txt����ҽ��, InStr(1, Txt����ҽ��, "-") + 1))
                .Update
            End With
            
            With rsSendRecipeDetail
                For lngRow = 1 To Bill������ϸ.Rows - 2
                    .AddNew
                    !NO = Mid(TxtNo.Text, 1, 8)
                    !�շ�ID = Val(Bill������ϸ.TextMatrix(lngRow, ����.Id))
                    !ҩƷID = Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID))
                    !���� = Val(Bill������ϸ.TextMatrix(lngRow, ����.����))
                    .Update
                  Next
            End With
            
'            '�ȸ�������
'            For lngRow = 1 To Bill������ϸ.Rows - 2
'                LngID = Val(Bill������ϸ.TextMatrix(lngRow, ����.Id))
'                lngҩƷID = Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID))
'                lng���� = Val(Bill������ϸ.TextMatrix(lngRow, ����.����))
'                gstrSQL = "zl_ҩƷ�շ���¼_��������(" & LngID & "," & lngҩƷID & "," & lng���� & ")"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��������")
'            Next
'
'            If IntSendAfterDosage = 0 Then
'                '���뾭����ҩ���̣�����ҩ�˲���
'                gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
'                                  "','" & mstr����Ա & "'" & ",NULL," & IIf(Txt����ҽ��.ListIndex = 0, "NULL", _
'                                  "'" & Mid(Txt����ҽ��, InStr(1, Txt����ҽ��, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "', " & int����λ�� & "," & int��˻��۵� & ")"
'            Else
'                gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
'                                  "','" & mstr����Ա & "'" & ",'" & cbo��ҩ��.Text & "'," & IIf(Txt����ҽ��.ListIndex = 0, "NULL", _
'                                  "'" & Mid(Txt����ҽ��, InStr(1, Txt����ҽ��, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "'," & int����λ�� & "," & int��˻��۵� & ")"
'            End If
'            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
'
'            '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
'            If gblnҩƷʹ�õ���ǩ�� = True Then
'                If SaveSignatureRecored(EsignTache.send, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo.Text, 1, 8), lngҩ��ID) = False Then
'                    Exit Function
'                End If
'            End If
'
'            '��¼�ô����ż���������
'            strBill = Mid(TxtNo.Text, 1, 8) & "|" & TxtNo.ItemData(TxtNo.ListIndex)
'            mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & Mid(TxtNo.Text, 1, 8) & "," & TxtNo.ItemData(TxtNo.ListIndex)
        End If
    Next
    
    '�������������������ҩ
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For n = 1 To rsSendRecipeByNo.RecordCount
        rsSendRecipeDetail.Filter = "NO='" & rsSendRecipeByNo!NO & "'"
        rsSendRecipeDetail.MoveFirst
        For lngRow = 1 To rsSendRecipeDetail.RecordCount
            gstrSQL = "zl_ҩƷ�շ���¼_��������(" & rsSendRecipeDetail!�շ�ID & "," & rsSendRecipeDetail!ҩƷID & "," & rsSendRecipeDetail!���� & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��������")
            
            rsSendRecipeDetail.MoveNext
        Next
        
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
        '�ⷿID
        gstrSQL = gstrSQL & lngҩ��ID
        '����
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!����
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '��ҩ��(�����)
        gstrSQL = gstrSQL & ",'" & mstr����Ա & "'"
        '��ҩ��(���뾭����ҩ����ʱ������ҩ�˲���)
        gstrSQL = gstrSQL & "," & IIf(IntSendAfterDosage = 0, "Null", IIf(rsSendRecipeByNo!��ҩ�� = "", "NULL", "'" & rsSendRecipeByNo!��ҩ�� & "'")) & ""
        'У���ˣ�����ҽ����
        gstrSQL = gstrSQL & "," & IIf(rsSendRecipeByNo!������ = "", "NULL", "'" & rsSendRecipeByNo!������ & "'") & ""
        '��ҩ��ʽ
        gstrSQL = gstrSQL & ",1"
        '��ҩʱ��
        gstrSQL = gstrSQL & ",Null"
        '����Ա����
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '����Ա����
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '����λ��
        gstrSQL = gstrSQL & "," & int����λ��
        '�Զ���˼��˵�
        gstrSQL = gstrSQL & "," & int��˻��۵�
        gstrSQL = gstrSQL & ")"
       
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")

        '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
        If gblnҩƷʹ�õ���ǩ�� = True Then
            If SaveSignatureRecored(EsignTache.send, rsSendRecipeByNo!����, rsSendRecipeByNo!NO, lngҩ��ID) = False Then
                Exit Function
            End If
        End If

        '��¼�ô����ż���������
        strBill = rsSendRecipeByNo!NO & "|" & rsSendRecipeByNo!����
        mstrPrintRecipe = IIf(mstrPrintRecipe = "", "", mstrPrintRecipe & "|") & rsSendRecipeByNo!NO & "," & rsSendRecipeByNo!����
        
        rsSendRecipeByNo.MoveNext
    Next
    
    mstr����Ա = ""
    mstr��ҩ�� = ""
    Txt����ҽ��.Enabled = False
    Me.TxtNo.SetFocus
            
    SendBatchRecipe = True
    Exit Function
ErrHand:
    SendBatchRecipe = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SendRecipe() As Boolean
    Dim lngRow As Long, lngҩƷID As Long, LngID As Long, lng���� As Long, lng���� As Long
    
    On Error GoTo ErrHand
    
    '�ȸ�������
    For lngRow = 1 To Bill������ϸ.Rows - 2
        LngID = Val(Bill������ϸ.TextMatrix(lngRow, ����.Id))
        lngҩƷID = Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID))
        lng���� = Val(Bill������ϸ.TextMatrix(lngRow, ����.����))
        gstrSQL = "zl_ҩƷ�շ���¼_��������(" & LngID & "," & lngҩƷID & "," & lng���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��������")
    Next
    If IntSendAfterDosage = 0 Then
        '���뾭����ҩ���̣�����ҩ�˲���
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
                          "','" & mstr����Ա & "'" & ",NULL," & IIf(Txt����ҽ��.ListIndex = 0, "NULL", _
                          "'" & Mid(Txt����ҽ��, InStr(1, Txt����ҽ��, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "', " & int����λ�� & "," & int��˻��۵� & ")"
    Else
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & _
                          "','" & mstr����Ա & "'" & ",'" & cbo��ҩ��.Text & "'," & IIf(Txt����ҽ��.ListIndex = 0, "NULL", _
                          "'" & Mid(Txt����ҽ��, InStr(1, Txt����ҽ��, "-") + 1) & "'") & ",1,NULL,'" & gstrUserCode & "','" & gstrUserName & "'," & int����λ�� & "," & int��˻��۵� & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
    
    '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
    If gblnҩƷʹ�õ���ǩ�� = True Then
        If SaveSignatureRecored(EsignTache.send, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo.Text, 1, 8), lngҩ��ID) = False Then
            Exit Function
        End If
    End If
    
    '��¼�ô����ż���������
    strBill = Mid(TxtNo.Text, 1, 8) & "|" & TxtNo.ItemData(TxtNo.ListIndex)
    mstrPrintRecipe = Mid(TxtNo.Text, 1, 8) & "," & TxtNo.ItemData(TxtNo.ListIndex)
    
    mstr����Ա = ""
    mstr��ҩ�� = ""
    
    Txt����ҽ��.Enabled = False
    Me.TxtNo.SetFocus
    
    SendRecipe = True
    Exit Function
ErrHand:
    SendRecipe = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetBatchSendRecord()
    Dim n As Integer
    
    Call IniRecord
    With mrsBatchSend
        For n = 1 To Msf�б�.Rows - 1
            If Val(Msf�б�.TextMatrix(n, ��������.��־)) = 1 And Msf�б�.TextMatrix(n, ��������.NO) <> "" Then
                .AddNew
                !NO = Msf�б�.TextMatrix(n, ��������.NO)
                !���� = Msf�б�.TextMatrix(n, ��������.����)
                !����ID = Val(Msf�б�.TextMatrix(n, ��������.����ID))
                !δ��� = Val(Msf�б�.TextMatrix(n, ��������.δ���))
                !��� = Val(Msf�б�.TextMatrix(n, ��������.ʵ�ս��))
                .Update
            End If
        Next
    End With
End Sub
Private Sub SetCheckBox(Optional ByVal intRow As Integer = -1)
    'intRow = -1 ����������������־
    'intRow = 0  ���������ʱ����������������־
    'intRow > 0  ��ָ��������־
    
    Dim n As Integer
    Dim strFlagName As String
    Dim i As Integer
    
    With Msf�б�
        If .Rows <= 1 Then Exit Sub
         
        .Redraw = False
         
        .Col = ��������.ѡ��
         
        If intRow = -1 Then
            For n = 0 To .Rows - 1
                .Row = n
                If Val(.TextMatrix(n, ��������.��־)) = 0 Then
                    strFlagName = "checked"
                    .TextMatrix(n, ��������.��־) = 1
                Else
                    strFlagName = "unchecked"
                    .TextMatrix(n, ��������.��־) = 0
                End If
                Set .CellPicture = LoadResPicture(strFlagName, vbResBitmap)
            Next
        ElseIf intRow = 0 Then
            If Val(.TextMatrix(intRow, ��������.��־)) = 0 Then
                strFlagName = "checked"
            Else
                strFlagName = "unchecked"
            End If
            
            For n = 0 To .Rows - 1
                .Row = n
                .TextMatrix(n, ��������.��־) = Abs(Val(.TextMatrix(n, ��������.��־)) - 1)
                Set .CellPicture = LoadResPicture(strFlagName, vbResBitmap)
            Next
        Else
            .Row = intRow
            If Val(.TextMatrix(intRow, ��������.��־)) = 0 Then
                strFlagName = "checked"
            Else
                strFlagName = "unchecked"
            End If
            .TextMatrix(intRow, ��������.��־) = Abs(Val(.TextMatrix(intRow, ��������.��־)) - 1)
            Set .CellPicture = LoadResPicture(strFlagName, vbResBitmap)
        End If
        
        Call SetBatchSendRecord
        
        .Redraw = True
    End With
End Sub

Private Sub SetColHide()
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    
    '�����û��������ã���ֹ��ʾ������
    strSave = zlDatabase.GetPara("������", glngSys, 1341)
    If strSave = "" Then strSave = "0|ҩƷ����,0|������,0|Ӣ����,0|���,0|����,0|��λ,0|����,0|����,0|���,0|����,0|�÷�,0|����,0|Ƶ��,0|ҽ������,0|�ѱ�,0|�����,0|�ⷿ��λ,0|������,0|׼����,0|��ҩ��,0|��ע"
    arrColumn = Split(strSave, ",")
    intRows = UBound(arrColumn)
    mbln��ʾ���� = False
    With Bill������ϸ
        For intRow = 0 To intRows
            intCol = GetDetailCol(Split(arrColumn(intRow), "|")(1))
            If intCol > -1 Then
                If Split(arrColumn(intRow), "|")(1) = "ҩƷ����" Then
                    intҩƷ���� = Val(Split(arrColumn(intRow), "|")(0))
                Else
                    If Val(Split(arrColumn(intRow), "|")(0)) = 1 Then
                        .ColWidth(intCol) = 0
                    ElseIf .ColWidth(intCol) = 0 Then
                        Select Case Split(arrColumn(intRow), "|")(1)
                        Case "������"
                            .ColWidth(����.������) = 2000
                        Case "Ӣ����"
                            .ColWidth(����.Ӣ����) = 2000
                        Case "���"
                            .ColWidth(����.���) = 1500
                        Case "����"
                            .ColWidth(����.����) = 1500
                        Case "��λ"
                            .ColWidth(����.��λ) = IIf(mbln��ʾ��С��λ = True, 0, 500)
                        Case "����"
                            .ColWidth(����.����) = 1000
                        Case "����"
                            .ColWidth(����.����) = 1200
                        Case "���"
                            .ColWidth(����.���) = 1200
                        Case "����"
                            .ColWidth(����.����) = 1200
                        Case "�÷�"
                            .ColWidth(����.�÷�) = 1500
                        Case "Ƶ��"
                            .ColWidth(����.Ƶ��) = 1500
                        Case "��ע"
                            .ColWidth(����.��ע) = 1200
                        Case "�ѱ�"
                            .ColWidth(����.�ѱ�) = 1000
                        Case "�ⷿ��λ"
                            .ColWidth(����.��λ) = IIf(MnuEditHandback.Checked, 0, 1200)
                        Case "����"
                            mbln��ʾ���� = True
                            If mblnIs��ҩ���� Then
                                .ColWidth(����.����) = 1200
                            End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '�������ҩ״̬����Щ�б�����ʾ
        .ColWidth(����.������) = IIf(MnuEditHandback.Checked, 1200, 0)
        .ColWidth(����.׼����) = IIf(MnuEditHandback.Checked, 1200, 0)
        .ColWidth(����.��ҩ��) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = False, 1200, 0)
        .ColWidth(����.׼������) = 0
        .ColWidth(����.׼����С) = 0
        .ColWidth(����.��ҩ����) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 1500, 0)
        .ColWidth(����.��ҩ��С) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 1500, 0)
        .ColWidth(����.��λ��) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 500, 0)
        .ColWidth(����.��λС) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 500, 0)
    End With
End Sub
Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrOracleMoneyForamt = strOracleTmp
    mstrVBMoneyForamt = strVbTmp
    
End Sub


Private Function AdviceCheckWarn(ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ��ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0ʱ��Ҫ
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String
    Dim strSQL As String, i As Long, k As Long
    
    AdviceCheckWarn = -1
    
    On Error GoTo errH
    Screen.MousePointer = 11
        
    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If
    
    '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�������˳�
    strSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
        " From ҩƷ�շ���¼ A,���˷��ü�¼ B,����ҽ����¼ C " & _
        " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
        " And A.����=[2] And A.no=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo, IntBillStyle)
    
    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If
    
    mlng����ID = rsTmp!����ID
    mstr�Һŵ� = NVL(rsTmp!�Һŵ�)
    mlng��ҳID = rsTmp!��ҳid
    
    '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
    '-------------------------------------------------------------
    If mlng����ID <> mlngPassPati Then
        If mstr�Һŵ� <> "" Then               '���ﲡ��
            strSQL = "Select ����ID,Count(Distinct Trunc(�Ǽ�ʱ��)) as ������� From ���˹Һż�¼ Where ����ID=[1] Group by ����ID"
            strSQL = "Select D.�������,A.����,A.�Ա�,A.��������," & _
                " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
                " From ������Ϣ A,���˹Һż�¼ B,���ű� C,(" & strSQL & ") D,��Ա�� E" & _
                " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=D.����ID" & _
                " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlng����ID, rsTmp!�������, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), "")
        Else                                    'סԺ����
            strSQL = _
                " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
                " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
                " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
        
            Call PassSetPatientInfo(mlng����ID, mlng��ҳID, rsTmp!����, NVL(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
                rsTmp!������ & "/" & rsTmp!������, IIf(Not IsNull(rsTmp!ҽ����), NVL(rsTmp!ҽ����) & "/" & NVL(rsTmp!ҽ����), ""), _
                IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        End If
        mlngPassPati = mlng����ID
    End If
    
    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With Bill������ϸ
            'ȡҩƷ����
            strҩƷ = .TextMatrix(lngRow, ����.ҩƷ����)
            If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
            If InStr(strҩƷ, "(") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, "(") - 1)
            'ȡҩƷ��ҩ;��
            str�÷� = .TextMatrix(lngRow, ����.�÷�)
            
            '�����ѯҩƷ��Ϣ
            Call PassSetQueryDrug(.TextMatrix(lngRow, ����.ҩƷID), strҩƷ, mstr������λ, str�÷�)
                
            '���ò˵�����״̬
            Call SetPassMenuState
            
            AdviceCheckWarn = 1 '��ʾ���Ե����˵�
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    
    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub SetFilter(ByVal blnState As Boolean)
    Dim strFind As String
    
    strFind = Trim(txtFind.Text)
    
    If strFind = "" Then Exit Sub
    mstrFilter = ""
    
    Select Case lblFind.Tag
        Case FindType.���￨
            If blnState = False Then
                mstrFilter = mstrFilter & " And Upper(A.���￨��) = [6] "
            Else
                mstrFilter = mstrFilter & " And Upper(B.���￨��) = [6] "
            End If
            SQLCondition.str���￨ = strFind
        Case FindType.�����
            If blnState = False Then
                mstrFilter = mstrFilter & " And Upper(A.�����) = [14] "
            Else
                mstrFilter = mstrFilter & " And Upper(B.�����) = [14] "
            End If
            SQLCondition.str����� = strFind
        Case FindType.���ݺ�
            If IsNumeric(strFind) Then
                strFind = GetFullNO(strFind, 13)
                txtFind.Text = strFind
            End If
            strFind = UCase(strFind)
            mstrFilter = mstrFilter & " And A.NO Between [3] And [4] "
            SQLCondition.str��ʼNO = strFind
            SQLCondition.str����NO = strFind
        Case FindType.����
            If mblnCard = True Then
                If blnState = False Then
                    mstrFilter = mstrFilter & " And Upper(A.���￨��) = [6] "
                Else
                    mstrFilter = mstrFilter & " And Upper(B.���￨��) = [6] "
                End If
                SQLCondition.str���￨ = strFind
            Else
                If blnState = False Then
                    mstrFilter = mstrFilter & " And Upper(A.����) Like Upper([5]) "
                Else
                    mstrFilter = mstrFilter & " And Upper(B.����) Like Upper([5]) "
                End If
                SQLCondition.str���� = strFind & "%"
            End If
        Case FindType.���֤
            If blnState = False Then
                mstrFilter = mstrFilter & " And A.���֤�� = [15] "
            Else
                mstrFilter = mstrFilter & " And B.���֤�� = [15] "
            End If
            SQLCondition.str���֤ = strFind
        Case FindType.IC��
            If blnState = False Then
                mstrFilter = mstrFilter & " And A.IC���� = [16] "
            Else
                mstrFilter = mstrFilter & " And B.IC���� = [16] "
            End If
            SQLCondition.strIC�� = strFind
    End Select

    Call mnuViewRefresh_Click
'    mstrFilter = ""
End Sub
Private Sub SetPosition()
    If mbln���������� Then
        img����.Top = PicBackGroud.Top - img����.Height - 50
        img����.Left = PicBackGroud.Left
        
        If img����.BorderStyle = 1 Then
            cbo����.Visible = True
            cbo����.Top = img����.Top - 20
            cbo����.Left = img����.Left + img����.Width + 50
            
            Chk�嵥.Top = img����.Top + 20
            
            If Chk�嵥.Visible Then
                Chk�嵥.Left = cbo����.Left + cbo����.Width + 200
            End If
            
            If Chk��ʾ��ҩ��������.Visible Then
                Chk��ʾ��ҩ��������.Left = cbo����.Left + cbo����.Width + 200
            End If
            
            Call Select����
        Else
            cbo����.Visible = False
            
            Chk�嵥.Top = img����.Top + 20
            
            If Chk�嵥.Visible Then
                Chk�嵥.Left = img����.Left + img����.Width + 200
            End If
            
            If Chk��ʾ��ҩ��������.Visible Then
                Chk��ʾ��ҩ��������.Left = img����.Left + img����.Width + 200
            End If
        End If
    Else
        Chk�嵥.Top = PicBackGroud.Top - Chk�嵥.Height - 50
        Chk�嵥.Left = PicBackGroud.Left
    End If
    
    Chk��ʾ��ҩ��������.Top = Chk�嵥.Top
End Sub

Private Sub SetRecipeColor()
    '��Ǵ�����ɫ
    Dim lngRow As Integer
    
    Msf�б�.Redraw = False
'    Msf�б�.TextMatrix(0, ��������.��ɫ) = ""
    For lngRow = 1 To Msf�б�.Rows - 1
        Msf�б�.Row = lngRow
        Msf�б�.Col = ��������.��ɫ
        Msf�б�.CellBackColor = Split(mstrUserRecipeColor, ";")(Val(Msf�б�.TextMatrix(lngRow, ��������.��������)))
    Next
    Msf�б�.Redraw = True
End Sub

Private Sub SetTimerState(ByVal BlnSet As Boolean)
    '�رպ�����Timer�ؼ����е�������ʱ����
    'blnSet��True-������False-�ر�
    
    If BlnSet Then
        '����ʱ�ָ�ԭ����״̬
        TimeRefresh.Enabled = mblnStateTimeRefresh
        TimePrint.Enabled = mblnStateTimePrint
    Else
        '�ر�ʱ�ȼ�¼ԭ����״̬
        mblnStateTimeRefresh = TimeRefresh.Enabled
        mblnStateTimePrint = TimePrint.Enabled
        
        If mblnStateTimeRefresh Then TimeRefresh.Enabled = False
        If mblnStateTimePrint Then TimePrint.Enabled = False
    End If
End Sub

Private Function �ж��Ƿ���ҩ����(ByVal BillType As Integer, ByVal BillNo As String) As Boolean
    'ͨ��ҩƷid�ж��Ƿ�����ҩ
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim DblWidth As Double

    strSQL = "Select a.��� as ��� From �շ���ĿĿ¼ a ,ҩƷ�շ���¼ b Where b.ҩƷid=a.Id And b.����=[2] and b.No=[1] And (b.��¼״̬=1 Or Mod(b.��¼״̬,3)=0) and (b.�ⷿID+0=[3] OR b.�ⷿID IS NULL) " _
        & " union all " _
        & "Select a.��� as ��� From �շ���ĿĿ¼ a ,HҩƷ�շ���¼ b Where b.ҩƷid=a.Id And b.����=[2] and b.No=[1] And (b.��¼״̬=1 Or Mod(b.��¼״̬,3)=0) and (b.�ⷿID+0=[3] OR b.�ⷿID IS NULL) "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "[�ж��Ƿ���ҩ����]", BillNo, BillType, lngҩ��ID)
    
    mblnIs��ҩ���� = IIf(rs!��� = 7, True, False)
    rs.Close
    
    On Error Resume Next
    
    DblWidth = Me.ScaleWidth - (ImgLeftRight_S.Left + ImgLeftRight_S.Width)
    If mblnIs��ҩ���� Then
        With Bill������ϸ
            .Top = Txt����.Top + Txt����.Height + 50
            .Height = IIf(txtԭʼ����.Top - .Top - 50 < 0, .Height, txtԭʼ����.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
        End With
    Else
        With Bill������ϸ
            .Top = Txt����.Top + Txt����.Height + 50
            .Height = IIf(cbo��ҩ��.Top - .Top - 50 < 0, .Height, cbo��ҩ��.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
            If .ColWidth(����.����) <> 0 Then
                .ColWidth(����.����) = 0
            End If
        End With
    End If
    
    �ж��Ƿ���ҩ���� = mblnIs��ҩ����
    
End Function

Private Sub ��ҩ�����ر���(ByVal BillStyle As Integer, ByVal BillNo As String)
    '��ҩ������ʾԭʼ��������ҩ�巨
    Dim strSQL As String
    Dim rs As New ADODB.Recordset

    strSQL = "Select a.���,b.���� From ҩƷ�շ���¼ a ,���˷��ü�¼ b Where a.����id=b.Id " _
        & " And a.����=[2] And a.No=[1] " _
        & " And (a.��¼״̬=1 Or Mod(a.��¼״̬,3)=0) and (a.�ⷿID+0=[3] OR a.�ⷿID IS NULL) " _
        & " union all " _
        & " Select a.���,b.���� From HҩƷ�շ���¼ a ,H���˷��ü�¼ b Where a.����id=b.Id " _
        & " And a.����=[2] And a.No=[1] " _
        & " And (a.��¼״̬=1 Or Mod(a.��¼״̬,3)=0) and (a.�ⷿID+0=[3] OR a.�ⷿID IS NULL) "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "[��ҩ�����ر���]", BillNo, BillStyle, lngҩ��ID)
    
    txtԭʼ����.Text = CStr(IIf(IsNull(rs!����), 1, rs!����))
    txt��ҩ�巨.Text = IIf(IsNull(rs!���), "", rs!���)
    
    rs.Close
    
End Sub

Private Sub Bill������ϸ_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill������ϸ_cboClick(ListIndex As Long)
    With Bill������ϸ
        If Not .Active Then Exit Sub
        If .ListCount = 0 Then Exit Sub
        .TextMatrix(.Row, ����.����) = .CboText
        .TextMatrix(.Row, ����.����) = .ItemData(.ListIndex)
    End With
End Sub
Private Sub Bill������ϸ_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Call Bill������ϸ_cboClick(Bill������ϸ.ListIndex)
End Sub

Private Sub Bill������ϸ_EnterCell(Row As Long, Col As Long)
    Dim lng���� As Long, lngҩƷID As Long, Dbl���� As Double, blnAllow As Boolean
    Dim strNo As String, int���� As Integer, strUnit As String, str��װ As String
    Dim rs���� As New ADODB.Recordset
    
    If Not BlnEnterCell Then Exit Sub
    If TxtNo.ListIndex = -1 Or BlnRefresh = False Then Exit Sub
    
    '��鵥���Ƿ����
    If Not CheckBillExist(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)), Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)) Then
        MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    Call ShowStock
    With Bill������ϸ
        '���õ�ǰ�е���ɫ
        Call .SetRowColor(Row, &H8000000F, True)
        
        If .CboVisible Or .TxtVisible Then Exit Sub
        .ColData(.Col) = 0
        .Clear
        .Active = False
        .TxtVisible = False
        .CboVisible = False
        If .Row = .Rows - 1 Then
            If mblnAuto = False Then
                mintLastSequence = 0
            End If
            Exit Sub
        End If
        
        If Val(.TextMatrix(Row, ����.ҩƷID)) = 0 Then Exit Sub    'ҩƷIDΪ�գ����˳�
        
        mintLastSequence = Val(.TextMatrix(Row, ����.���))
        
        strUnit = GetUnit(lngҩ��ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
        Select Case strUnit
        Case "�ۼ۵�λ"
            str��װ = "1"
        Case "���ﵥλ"
            str��װ = "�����װ"
        Case "סԺ��λ"
            str��װ = "סԺ��װ"
        Case "ҩ�ⵥλ"
            str��װ = "ҩ���װ"
        End Select
        
        If (MnuEditDosage.Checked Or MnuEditConsignment.Checked) Then
            If Val(.TextMatrix(Row, ����.����)) = 0 Then Exit Sub    'ҩƷ����Ϊ�գ����˳�
            If Not (.Col = ����.����) Then Exit Sub
            lng���� = Val(.TextMatrix(Row, ����.����))
            lngҩƷID = Val(.TextMatrix(Row, ����.ҩƷID))
            Dbl���� = FormatEx(Val(.TextMatrix(Row, ����.����)), 5)
            strNo = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
            int���� = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����))

            '������ڷ�ҩ��¼�Ҳ�����ҩ���������޸�������Ϣ
            blnAllow = False

            gstrSQL = " Select count(*) Records From ҩƷ�շ���¼ " & _
                " Where (Mod(��¼״̬,3)=0 or ��¼״̬=1) And ����� Is Not NULL " & _
                " And NO=[1] And �ⷿID=[3] And ����=[2] " & _
                " And ҩƷID=[4] And Nvl(����,0)=[5]"
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, int����, lngҩ��ID, lngҩƷID, lng����)
            
            With rs����
                blnAllow = (!Records = 0)
            End With
            
            '��ȡ����������Ϣ
            gstrSQL = " SELECT B.�ϴ����� ����,B.����,ROUND(B.ʵ������/" & str��װ & ",2) ����" & _
                " FROM ҩƷ��� A,ҩƷ��� B,�շѼ�Ŀ C,�շ���ĿĿ¼ F" & _
                " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID AND B.�ⷿID = [1] AND B.ҩƷID=[2] AND A.ҩƷID = C.�շ�ϸĿID" & _
                " AND ((SYSDATE BETWEEN C.ִ������ AND C.��ֹ����) OR C.��ֹ���� IS NULL)" & _
                " AND NVL(����,0)<>0 AND NVL(ʵ������,0)<>0 AND ����=1" & _
                " AND ROUND(DECODE(F.�Ƿ���,NULL,C.�ּ�,0,C.�ּ�,B.ʵ�ʽ��/B.ʵ������),2)=" & _
                "     (SELECT ROUND(DECODE(F.�Ƿ���,NULL,C.�ּ�,0,C.�ּ�,B.ʵ�ʽ��/B.ʵ������),2) ����" & _
                "     FROM ҩƷ��� A,ҩƷ��� B,�շѼ�Ŀ C,�շ���ĿĿ¼ F" & _
                "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID AND B.�ⷿID = [1] AND B.ҩƷID=[2] AND A.ҩƷID = C.�շ�ϸĿID" & _
                "     AND ((SYSDATE BETWEEN C.ִ������ AND C.��ֹ����) OR C.��ֹ���� IS NULL)" & _
                "     AND NVL(����,0)<>0 AND NVL(ʵ������,0)<>0 AND ����=1 AND NVL(����,0)=[3])" & _
                " AND ROUND(B.ʵ������/" & str��װ & ",2)>=[4] AND (NVL(A.ҩ������,0)=0 OR (NVL(A.ҩ������,0)=1 AND (Ч�� IS NULL OR Ч��>TRUNC(SYSDATE))))" & _
                " ORDER BY B.����"
            Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID, lngҩƷID, lng����, Dbl����)
            
            With rs����
                Do While Not .EOF
                    If (!���� <> lng���� And blnAllow) Or !���� = lng���� Then
                        Bill������ϸ.AddItem IIf(IsNull(!����), "", !����) & "(" & !���� & ")"
                        Bill������ϸ.ItemData(Bill������ϸ.NewIndex) = !����
                    End If
                    .MoveNext
                Loop
            End With
        ElseIf MnuEditHandback.Checked Then
            If mbln��ʾ��С��λ = True Then
                .Tag = Val(.TextMatrix(.Row, ����.׼������)) * Val(.TextMatrix(.Row, ����.��װ)) + Val(.TextMatrix(.Row, ����.׼����С))
                If Not (.Col = ����.��ҩ���� Or .Col = ����.��ҩ��С) Then Exit Sub
            Else
                .Tag = Val(.TextMatrix(.Row, ����.׼����))
                If Not (.Col = ����.��ҩ��) Then Exit Sub
            End If
        Else
            Exit Sub
        End If
        
        If (MnuEditDosage.Checked Or MnuEditConsignment.Checked) Then
            .ColData(.Col) = IIf(.ListCount = 0, 0, 3)
            .Active = IIf((.ListCount > 0), True, False)
        ElseIf MnuEditHandback.Checked Then
            '����ô�����ת�������������
            If Not zlDatabase.NOMoved("ҩƷ�շ���¼", Mid(TxtNo.Text, 1, 8), "����=", TxtNo.ItemData(TxtNo.ListIndex)) Then
                .ColData(.Col) = 4
                .Active = CmdSend.Enabled
            End If
        End If
    End With
End Sub

Private Sub Bill������ϸ_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim blnUnValid As Boolean
    Dim dblCount As Double
    Dim dblSumCount As Double
    Dim rsTemp As New ADODB.Recordset
    
    With Bill������ϸ
        If KeyCode = vbKeyReturn Then
            If mbln��ʾ��С��λ = True Then
                If Not (.TxtVisible And (.Col = ����.��ҩ���� Or .Col = ����.��ҩ��С)) Then Exit Sub
            Else
                If Not (.TxtVisible And .Col = ����.��ҩ��) Then Exit Sub
            End If
            
            blnUnValid = False
            .Text = Trim(.Text)
            
            blnUnValid = (.Text = "")
            If Not blnUnValid Then blnUnValid = Not IsNumeric(.Text)
            If Not blnUnValid Then
                If mbln��ʾ��С��λ = True Then
                    If .Col = ����.��ҩ���� Then
                        dblSumCount = Val(.Text) * Val(.TextMatrix(.Row, ����.��װ)) + Val(.TextMatrix(.Row, ����.��ҩ��С))
                    Else
                        dblSumCount = Val(.TextMatrix(.Row, ����.��ҩ����)) * Val(.TextMatrix(.Row, ����.��װ)) + Val(.Text)
                    End If
                Else
                    dblSumCount = Val(.Text)
                End If
                blnUnValid = Not ((Abs(dblSumCount) <= Abs(.Tag)) And ((Val(dblSumCount) >= 0 And Val(.Tag) >= 0) Or (Val(dblSumCount) <= 0 And Val(.Tag) <= 0)))
            End If
            
            If blnUnValid Then
                If mbln��ʾ��С��λ = True Then
                    If .Col = ����.��ҩ���� Then
                        .Text = Val(.TextMatrix(.Row, ����.׼������))
                    Else
                        .Text = Val(.TextMatrix(.Row, ����.׼����С))
                    End If
                Else
                    .Text = Val(.Tag)
                End If
            End If
            
            '�ȼ���Ƿ���ҽ��������ҩƷ��¼
            '��������򲻹�
            '����ǣ����ϵͳ�����Ƿ�����δ����ҽ����ҩ�������������ҩ��Ϊ��
            '��������򲻹�
            dblCount = Val(FormatEx(.Text, 5))
            If dblCount <> 0 And blnҽ������ = False Then
                gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", Val(Bill������ϸ.TextMatrix(Bill������ϸ.Row, ����.Id)))
                
                If (rsTemp!���� Like "1*") Then       '����
                    gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ���˷��ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", Val(Bill������ϸ.TextMatrix(Bill������ϸ.Row, ����.Id)))
                    
                    If Not rsTemp.EOF Then
                        If (rsTemp!�����־ = 1 Or rsTemp!�����־ = 4) And rsTemp!ҽ����� <> 0 Then
                            gstrSQL = "Select decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ID=[1]"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rsTemp!ҽ�����))

                            If rsTemp!���� = 0 Then
                                dblCount = 0
                                MsgBox "�ñ�ҽ����δ���ϣ�������ҩ��", vbInformation, gstrSysName
                            End If
                        End If
                    End If
                End If
            End If
            
            .Text = FormatEx(dblCount, 5)
            
            If mbln��ʾ��С��λ = True Then
                If .Col = ����.��ҩ���� Then
                    .TextMatrix(.Row, ����.��ҩ����) = FormatEx(.Text, 5)
                Else
                    .TextMatrix(.Row, ����.��ҩ��С) = FormatEx(.Text, 5)
                End If
                .TextMatrix(.Row, ����.��ҩ��) = FormatEx(dblSumCount, 5) / Val(.TextMatrix(.Row, ����.��װ))
                
                If Val(.TextMatrix(.Row, ����.��ҩ��)) <> Val(.TextMatrix(.Row, ����.ʵ������)) / Val(.TextMatrix(.Row, ����.��װ)) Then
                    mblnAllBack = False
                End If
            Else
                .TextMatrix(.Row, ����.��ҩ��) = FormatEx(.Text, 5)
                
                If Val(.TextMatrix(.Row, ����.��ҩ��)) <> Val(.TextMatrix(.Row, ����.׼����)) Then
                    mblnAllBack = False
                End If
            End If
        End If
    End With
End Sub

Private Sub Bill������ϸ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    If Button = 2 Then
        With Bill������ϸ
            lngRow = .MouseRow
            If lngRow >= .MsfObj.FixedRows And lngRow < .Rows - 1 Then
                .Row = lngRow
            End If
        End With
    End If
    
End Sub

Private Sub Bill������ϸ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strҩƷ As String
    
    'Pass
    If Button = 2 And gblnPass And tabShow.Tab = 2 And Len(Bill������ϸ.TextMatrix(Bill������ϸ.Row, ����.ҽ��id)) > 0 Then
            With Bill������ϸ
            If .Rows > 1 And .Row < .Rows - 1 Then
                '���Pass״̬
                If AdviceCheckWarn(0, .Row) >= 0 Then PopupMenu mnuPass, 2
            End If
        End With
    End If
End Sub

Private Sub SetPassMenuState()
    '���ܣ�����Pass�˵�����״̬
    'Pass
    'һ���˵�
    'ҩ���ٴ���Ϣ�ο�
    mnuPassItem(0).Enabled = PassGetState("CPRRes") = 1
    'ҩƷ˵����
    mnuPassItem(1).Enabled = PassGetState("Directions") = 1
    '�й�ҩ��
    mnuPassItem(2).Enabled = PassGetState("Chp") = 1
    '������ҩ����
    mnuPassItem(3).Enabled = PassGetState("CPERes") = 1
    '����ֵ
    mnuPassItem(4).Enabled = PassGetState("CheckRes") = 1
    'ר����Ϣ
    'mnuPassItem(6).Enabled = PassGetState("") = 1
    'ҽҩ��Ϣ����
    mnuPassItem(8).Enabled = PassGetState("MEDInfo") = 1
    'ҩƷ�����Ϣ
    mnuPassItem(10).Enabled = PassGetState("MATCH-DRUG") = 1
    '��ҩ;�������Ϣ
    mnuPassItem(11).Enabled = PassGetState("MATCH-ROUTE") = 1
    'ҽԺҩƷ��Ϣ
    mnuPassItem(12).Enabled = PassGetState("HisDrugInfo") = 1
    
    '���˲˵�
    'ҩ��-ҩ���໥����
    mnuPassSpec(0).Enabled = PassGetState("DDIM") = 1
    'ҩ��-ʳ���໥ʹ��
    mnuPassSpec(1).Enabled = PassGetState("DFIM") = 1
    '����ע�����������
    mnuPassSpec(3).Enabled = PassGetState("MatchRes") = 1
    '����ע�����������
    mnuPassSpec(4).Enabled = PassGetState("TriessRes") = 1
    '����֢
    mnuPassSpec(6).Enabled = PassGetState("DDCM") = 1
    '������
    mnuPassSpec(7).Enabled = PassGetState("SIDE") = 1
    '��������ҩ
    mnuPassSpec(9).Enabled = PassGetState("GERI") = 1
    '��ͯ��ҩ
    mnuPassSpec(10).Enabled = PassGetState("PEDI") = 1
    '��������ҩ
    mnuPassSpec(11).Enabled = PassGetState("PREG") = 1
    '��������ҩ
    mnuPassSpec(12).Enabled = PassGetState("LACT") = 1
End Sub


Private Sub Cbar_Resize()
    Form_Resize
End Sub

Private Sub cbo����_Click()
    If cbo����.ListIndex = -1 Then Exit Sub
    
    If cbo����.ItemData(cbo����.ListIndex) <> Val(cbo����.Tag) Then
        cbo����.Tag = cbo����.ItemData(cbo����.ListIndex)
        Call mnuViewRefresh_Click
    End If
End Sub


Private Sub cbo��ҩ��_Click()
    '    Exit Sub
End Sub

Private Sub cbo��ҩ��_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo��ҩ��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub


Private Sub cbo��ҩ��_KeyPress(KeyAscii As Integer)
Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo��ҩ��.Text)
        If cbo��ҩ��.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo��ҩ��.List(cbo��ҩ��.ListIndex) Then Call zlControl.CboSetIndex(cbo��ҩ��.hWnd, -1)
        End If
        If strText = "" Then
            cbo��ҩ��.ListIndex = -1
        ElseIf cbo��ҩ��.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo��ҩ��.ListCount - 1
                If Mid(cbo��ҩ��.List(i), 1, InStr(1, cbo��ҩ��.List(i), "-") - 1) = strText _
                    Or Mid(cbo��ҩ��.List(i), InStr(1, cbo��ҩ��.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo��ҩ��.ListCount - 1
                    If UCase(cbo��ҩ��.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo��ҩ��.ListIndex = intIdx
            SendMessage cbo��ҩ��.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo��ҩ��_Click
            Exit Sub
        End If
        If cbo��ҩ��.ListIndex = -1 Then
            cbo��ҩ��.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo��ҩ��_Click
            ElseIf intIdx <> cbo��ҩ��.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo��ҩ��.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo��ҩ��_Click
            End If
        End If
    End If
End Sub


Private Sub cbo��ҩ��_LostFocus()
    Call cbo��ҩ��_Validate(True)
End Sub
Private Sub cbo��ҩ��_Validate(Cancel As Boolean)
    Dim n As Integer
    Dim blnFind As Boolean
    
    cbo��ҩ��.Text = Trim(cbo��ҩ��.Text)
    If InStr(cbo��ҩ��.Text, "-") > 0 Then
        cbo��ҩ��.Text = Mid(cbo��ҩ��.Text, InStr(cbo��ҩ��.Text, "-") + 1)
    End If
    If cbo��ҩ��.Text <> "" Then
        For n = 0 To cbo��ҩ��.ListCount - 1
            If cbo��ҩ��.Text = Mid(cbo��ҩ��.List(n), InStr(cbo��ҩ��.List(n), "-") + 1) Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            cbo��ҩ��.Text = ""
            Exit Sub
        End If
    End If
           
End Sub


Private Sub Chk�嵥_Click()
    Call mnuViewRefresh_Click
End Sub

Private Sub Chkȫ��_Click()
    Dim intRow As Integer
    Dim lng������ As Long
    Dim dblС���� As Double
    
    If Not Chkȫ��.Enabled Then Exit Sub
    With Bill������ϸ
        For intRow = 1 To .Rows - 2
            If mbln��ʾ��С��λ = True Then
                If Chkȫ��.Value = 1 Then
                    .TextMatrix(intRow, ����.��ҩ����) = .TextMatrix(intRow, ����.׼������)
                    .TextMatrix(intRow, ����.��ҩ��С) = .TextMatrix(intRow, ����.׼����С)
                    
                    .TextMatrix(intRow, ����.��ҩ��) = FormatEx(Val(.TextMatrix(intRow, ����.ʵ������)) / Val(.TextMatrix(intRow, ����.��װ)), mintNumberDigit)
                Else
                    .TextMatrix(intRow, ����.��ҩ��) = ""
                    .TextMatrix(intRow, ����.��ҩ����) = ""
                    .TextMatrix(intRow, ����.��ҩ��С) = ""
                End If
            Else
                .TextMatrix(intRow, ����.��ҩ��) = IIf(Chkȫ��.Value = 1, .TextMatrix(intRow, ����.׼����), "")
            End If
        Next
        mblnAllBack = (Chkȫ��.Value = 1)
    End With
End Sub
Private Sub Chk��ʾ��ҩ��������_Click()
    mlng�������� = Chk��ʾ��ҩ��������.Value
    Call mnuViewRefresh_Click
End Sub

Private Sub cmdAlley_Click()
    '���ܣ��Բ��˹���ʷ/����״̬���й���
    'Pass
    Call AdviceCheckWarn(21)
End Sub

Private Sub cmdFind_Click()
    Call Form_KeyDown(vbKeyF3, 0)
End Sub

Private Sub cmdIC_Click()
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If Not mobjICCard Is Nothing Then
        txtFind.Text = mobjICCard.Read_Card()
        If txtFind.Text <> "" Then Call txtFind_KeyPress(vbKeyReturn)
    End If
End Sub
Private Sub CmdSend_Click()
    Dim lngRow As Long, lngҩƷID As Long, LngID As Long, lng���� As Long, lng���� As Long
    Dim blnInput As Boolean, strShow As String, strReturn As String, str����Ա As String, strTmp As String
    Dim rsTemp As New ADODB.Recordset, blnInTrans As Boolean
    Dim intUnit As Integer
    Dim bln�Ƿ�����ҩ As Boolean
    Dim str��Ŵ� As String
    Dim n As Integer
    Dim BlnFirst As Boolean
    Dim strSignInfo As String
    
    blnInTrans = False
    On Error Resume Next
    
    mstr��������ʾ = ""
    
    err = 0
    
    If TxtNo.ListIndex = -1 Then   '��Ч����
        MsgBox "����ѡ�񴦷���", vbInformation, gstrSysName
        If TxtNo.Enabled Then TxtNo.SetFocus
        Exit Sub
    End If
    
    
    On Error GoTo ErrHand
    
    '��鵥���Ƿ����
    If Not CheckBillExist(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
        MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    strUnit = GetUnit(lngҩ��ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
    '--��ҩ����--wq
    If MnuEditDosage.Checked Then
        '������辭����ҩ���̣��������൱�ڷ�ҩ
        If Not IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
            GoTo SendBill
        End If
        
        '����Ƿ�����
        If CheckBill(1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Sub
        
        'У����ҩ�ˣ�������õ���ǩ����ʹ��
        If gblnҩƷʹ�õ���ǩ�� = False Then
            If intУ����ҩ�� = 1 Then
                str����Ա = zlDatabase.UserIdentify(Me, "У����ҩ��", glngSys, 1341, "��ҩ")
            Else
                str����Ա = Str��ҩ��
            End If
            If str����Ա = "" Then Exit Sub
        End If
        
        gcnOracle.BeginTrans
        blnInTrans = True
        
        '�ȸ�������
        For lngRow = 1 To Bill������ϸ.Rows - 2
            LngID = Val(Bill������ϸ.TextMatrix(lngRow, ����.Id))
            lngҩƷID = Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID))
            lng���� = Val(Bill������ϸ.TextMatrix(lngRow, ����.����))
            gstrSQL = "zl_ҩƷ�շ���¼_��������(" & LngID & "," & lngҩƷID & "," & lng���� & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��������")
        Next
        
        '��������ҩ��
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ��(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo, 1, 8) & "','" & IIf(gblnҩƷʹ�õ���ǩ�� = True, gstrUserName, IIf(intУ����ҩ�� = 1, str����Ա, IIf(Str��ҩ�� = "|��ǰ����Ա|", gstrUserName, str����Ա))) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-������ҩ��")
        
        '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
        If gblnҩƷʹ�õ���ǩ�� = True Then
            If SaveSignatureRecored(EsignTache.Dosage, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo, 1, 8), lngҩ��ID) = False Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gcnOracle.CommitTrans
        blnInTrans = False
    End If
    '--ȡ������--
    If MnuEditAbolish.Checked Then
        If Not IsDosage(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then
            MsgBox "���辭����ҩ���̣���˲�����ִ��ȡ����ҩ������", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����Ƿ�����
        If CheckBill(2, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Sub
        
        gcnOracle.BeginTrans
        blnInTrans = True
        
        '����������˵���ǩ������ȡ����ҩ�˵���ǩ��
        If gblnҩƷʹ�õ���ǩ�� = True Then
            If DelSignatureRecored(EsignTache.Dosage, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo, 1, 8), lngҩ��ID) = False Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ��(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo, 1, 8) & "',Null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-������ҩ��")
        
        gcnOracle.CommitTrans
        blnInTrans = False
    End If
    
    '--��ҩ����--
    If MnuEditConsignment.Checked Then
SendBill:
        '����״̬ʱΪ������ҩģʽ
        If imgFilter.BorderStyle = cstFilter Then
            '������鴦��
            If Not CheckBatchRecipe Then Exit Sub
            
            gcnOracle.BeginTrans
            blnInTrans = True
        
            '����������ҩ
            If Not SendBatchRecipe Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        Else
            '��鴦��
            If Not CheckRecipe Then Exit Sub
            
            gcnOracle.BeginTrans
            blnInTrans = True
            
            '������ҩ
            If Not SendRecipe Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gcnOracle.CommitTrans
        
        blnInTrans = False
        mblnFilterRefresh = True
        
        '��ӡ����
        Call PrintRecipe
    End If
    
    '--��ҩ����--
    If MnuEditHandback.Checked Then
        Dim str���� As String, sig��ҩ�� As Single, strSubSql As String
        '��ת�������ݲ��������
        If zlDatabase.NOMoved("ҩƷ�շ���¼", Mid(TxtNo.Text, 1, 8), "���� = ", TxtNo.ItemData(TxtNo.ListIndex)) Then
            MsgBox "�ô����ѱ�ת���������������ҩ������", vbInformation, gstrSysName
            Exit Sub
        End If
        '����Ƿ�����
        If CheckBill(4, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), True) <> 0 Then Exit Sub
        Call GetBillSequence
        If str��� = "" Then Exit Sub
        If Not IsReceiptBalance(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), str���) Then Exit Sub
        If Not IsOutPatient(mstrPrivs, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) Then Exit Sub
        If Not CheckBillControl(tabShow.Tab + 1, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), Msf�б�.TextMatrix(Msf�б�.Row, ��������.���)) Then Exit Sub
        '���汻עע����20020905 Modified by zyb
        'If ReadBillData(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) = False Then Exit Sub
        If MsgBox("��ȷ������Ϊ[" & TxtNo & "]" & "�Ĵ�����ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        For lngRow = 1 To Bill������ϸ.Rows - 2
            lng���� = Val(Bill������ϸ.TextMatrix(lngRow, ����.����))
            lng���� = Val(Bill������ϸ.TextMatrix(lngRow, ����.����))
            '���ԭ�������������ڷ���
            If lng���� = 0 And lng���� = 1 Then
                '������Ż�Ч��Ϊ�գ�����ȡ���û�����
                blnInput = IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.������)) = "", True, False)
                If blnInput Then
                    strShow = Txt����.Text & "|" & Txt����.Text & "|" & Msf�б�.TextMatrix(lngRow, ��������.����) & _
                    "|" & Bill������ϸ.TextMatrix(lngRow, ����.ҩƷ����) & "|" & Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID))
                    strReturn = Frm��ҩ����.ShowME(Me, strShow)
                    If strReturn = "" Then Exit Sub
                    '�������š�Ч�ڼ�����
                    Bill������ϸ.TextMatrix(lngRow, ����.������) = Split(strReturn, "|")(0)
                    Bill������ϸ.TextMatrix(lngRow, ����.��Ч��) = Split(strReturn, "|")(1)
                    Bill������ϸ.TextMatrix(lngRow, ����.�²���) = Split(strReturn, "|")(2)
                End If
            End If
        Next
        str���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        bln�Ƿ�����ҩ = False
        gcnOracle.BeginTrans
        blnInTrans = True
        For lngRow = 1 To Bill������ϸ.Rows - 2
            'modified.by.zyb ���ﵥλ��סԺ��λ��һ��ʱ����ҩδ���� 2003-01-10
            Select Case strUnit
            Case "�ۼ۵�λ"
                strSubSql = "*1"
            Case "���ﵥλ"
                strSubSql = "*Decode(�����װ,Null,1,0,1,�����װ)"
            Case "סԺ��λ"
                strSubSql = "*Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
            Case "ҩ�ⵥλ"
                strSubSql = "*Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
            End Select
            sig��ҩ�� = Val(Bill������ϸ.TextMatrix(lngRow, ����.��ҩ��))

            gstrSQL = " Select round(" & sig��ҩ�� & strSubSql & ",5) ���� From ҩƷ���" & _
                         " Where ҩƷID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Bill������ϸ.TextMatrix(lngRow, ����.ҩƷID)))
                         
            With rsTemp
                sig��ҩ�� = !����
            End With
            
            If mbln��ʾ��С��λ = True Then
                If (Val(Bill������ϸ.TextMatrix(lngRow, ����.��ҩ����)) = Val(Bill������ϸ.TextMatrix(lngRow, ����.׼������)) And _
                    Val(Bill������ϸ.TextMatrix(lngRow, ����.��ҩ��С)) = Val(Bill������ϸ.TextMatrix(lngRow, ����.׼����С))) Or _
                    (Val(Bill������ϸ.TextMatrix(lngRow, ����.��ҩ��)) = Val(Bill������ϸ.TextMatrix(lngRow, ����.׼������)) * Val(Bill������ϸ.TextMatrix(lngRow, ����.��װ)) + Val(Bill������ϸ.TextMatrix(lngRow, ����.׼����С))) Then
                    
                    sig��ҩ�� = Val(Bill������ϸ.TextMatrix(lngRow, ����.ʵ������))
                End If
            Else
                If Val(Bill������ϸ.TextMatrix(lngRow, ����.��ҩ��)) = Val(Bill������ϸ.TextMatrix(lngRow, ����.׼����)) Then
                    sig��ҩ�� = Val(Bill������ϸ.TextMatrix(lngRow, ����.ʵ������))
                End If
            End If
            
            If sig��ҩ�� <> 0 Then
                '���۸�
                If CheckPrice(Val(Bill������ϸ.TextMatrix(lngRow, ����.Id)), mstr�۸�ʧЧ��ʾ) = False Then
                    If MsgBox("ҩƷ[" & Bill������ϸ.TextMatrix(lngRow, ����.ҩƷ����) & "]" & mstr�۸�ʧЧ��ʾ, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & Val(Bill������ϸ.TextMatrix(lngRow, ����.Id)) & ",'" & gstrUserName & "'," & _
                            "to_date('" & str���� & "','yyyy-MM-dd hh24:mi:ss')," & _
                            IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.������)) = "", "NULL", "'" & Bill������ϸ.TextMatrix(lngRow, ����.������) & "'") & "," & _
                            "" & IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.��Ч��)) = "", "NULL", "to_date('" & Bill������ϸ.TextMatrix(lngRow, ����.��Ч��) & "','yyyy-MM-dd')") & "," & _
                            IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.�²���)) = "", "NULL", "'" & Trim(Bill������ϸ.TextMatrix(lngRow, ����.�²���)) & "'") & "," & _
                            sig��ҩ�� & ",NULL,NULL," & int����λ�� & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
                        bln�Ƿ�����ҩ = True
                    End If
                Else
                    gstrSQL = "zl_ҩƷ�շ���¼_������ҩ(" & Val(Bill������ϸ.TextMatrix(lngRow, ����.Id)) & ",'" & gstrUserName & "'," & _
                        "to_date('" & str���� & "','yyyy-MM-dd hh24:mi:ss')," & _
                        IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.������)) = "", "NULL", "'" & Bill������ϸ.TextMatrix(lngRow, ����.������) & "'") & "," & _
                        "" & IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.��Ч��)) = "", "NULL", "to_date('" & Bill������ϸ.TextMatrix(lngRow, ����.��Ч��) & "','yyyy-MM-dd')") & "," & _
                        IIf(Trim(Bill������ϸ.TextMatrix(lngRow, ����.�²���)) = "", "NULL", "'" & Trim(Bill������ϸ.TextMatrix(lngRow, ����.�²���)) & "'") & "," & _
                        sig��ҩ�� & ",NULL,NULL," & int����λ�� & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ҩƷ��ҩ")
                    bln�Ƿ�����ҩ = True
                End If
            End If
        Next
        
        '������ز����������Զ����ʣ����ҵ�ǰ�˷ѵ����Ǽ��ʵ�����ôִ������/סԺ����
        If mint�Զ����� = 1 And mint��¼���� = 2 And bln�Ƿ�����ҩ = True Then
            For lngRow = 1 To Bill������ϸ.Rows - 2
                If Val(Bill������ϸ.TextMatrix(lngRow, ����.��ҩ��)) <> 0 Then
                    str��Ŵ� = str��Ŵ� & IIf(str��Ŵ� = "", Bill������ϸ.TextMatrix(lngRow, ����.���), "," & Bill������ϸ.TextMatrix(lngRow, ����.���))
                End If
            Next
            If mint�����־ = 1 Or mint�����־ = 4 Then
                gstrSQL = "Zl_������ʼ�¼_Delete('" & mstrNo & "','" & str��Ŵ� & "','" & gstrUserCode & "','" & gstrUserName & "')"
            Else
                gstrSQL = "Zl_סԺ���ʼ�¼_Delete('" & mstrNo & "','" & str��Ŵ� & "','" & gstrUserCode & "','" & gstrUserName & "'," & mint��¼���� & ")"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-��ҩ����")
        End If
        
        gcnOracle.CommitTrans
        blnInTrans = False
        
        '��ӡ�˷�֪ͨ��
        Dim int���� As Integer, strNo As String
        Dim Str��ҩʱ�� As String, Int��װϵ�� As Integer
        
        If bln�Ƿ�����ҩ Then
            int���� = TxtNo.ItemData(TxtNo.ListIndex)
            strNo = Mid(TxtNo.Text, 1, 8)
            Str��ҩʱ�� = str����
            Int��װϵ�� = IIf(int���� = 8, 1, 2)
            
            If MsgBox("����Ҫ��ӡ��ҩ֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
                Me, "No=" & strNo, "����=" & int����, "��װϵ��=" & IIf(Int��װϵ�� = 1, "D.�����װ", "D.סԺ��װ"), "��ҩʱ��=" & Str��ҩʱ��, 2)
            End If
            
            '��ʾͣ��ҩƷ
            Call CheckStopMedi(int���� & "|" & strNo)
        Else
            MsgBox "����û����ҩ��"
        End If
    End If
    
    BlnInOper = False
    Call mnuViewRefresh_Click
    
    If txtFind.Text <> "" Then
        txtFind.SetFocus
        Call GetFocus(txtFind)
    End If
    Exit Sub
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
    ByVal cur���ս�� As Currency, ByVal Cur���ʽ�� As Currency, ByVal cur������� As Currency, _
    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
    intWarn As Integer) As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
'     str�շ����=��ǰҪ�������,���ڷ��౨��
'     str�������=�������,������ʾ
'     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
'����:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
'     intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
'     0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
    Dim bln�ѱ��� As Boolean, byt��־ As Byte
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim ArrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str���� As String, i As Long
    
    BillingWarn = 0
    
    '�����������:NULL��û������,0�������˵�
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    'ʾ����"-" �� ",ABC,567,DEF"
    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
    
    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
        If byt��־ = 2 Then
            If str�ѱ���� Like "-*" Then
                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
            Else
                ArrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(ArrTmp)
                    If InStr(ArrTmp(i), str�շ����) > 0 Then
                        byt�ѱ���ʽ = IIf(Right(ArrTmp(i), 1) = "��", 2, 1)
                        'Exit For 'ȡ��˵����סԺ����ģ��
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str������� <> "" Then str������� = """" & str������� & """����"
    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
    curʣ���� = curʣ���� + cur������� - Cur���ʽ��
    cur���ս�� = cur���ս�� + Cur���ʽ��
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If curʣ���� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & " ����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If curʣ���� < 0 Then
                        byt��ʽ = 2
                        If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str������� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf curʣ���� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If curʣ���� < 0 Then
                            byt��ʽ = 2
                            If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str������� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If curʣ���� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ս�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ս�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") = 0 Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox("ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gtype_UserSysParms.P9_���ý���λ��) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = "-"
            Else
                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
            End If
        End If
    End If
End Function

Private Function FinishBillingWarn(ByVal rsTmp As ADODB.Recordset, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
'���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
'������objRecord=����Ҫ���ִ�еĲ�����Ϣ��������
'      str���="CDE..."����������漰�����շ����
'      str�����="���,����,..."����Ӧ�������������ʾ
    Dim rsPati As ADODB.Recordset
    Dim rsWarn As ADODB.Recordset
    Dim strWarn As String, intWarn As Integer
    Dim strSQL As String, intR As Integer, i As Long
    Dim cur���� As Currency
    
    On Error GoTo errH
    
    If rsTmp!��Դ.Value = "סԺ" Then
        'סԺ���˱���
        strSQL = _
            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1]" & _
            " Union ALL" & _
            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
        strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
        
        strSQL = "Select zl_PatiWarnScheme(A.����ID,B.��ҳID) As ���ò���,C.ʣ���," & _
            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
            " From ������Ϣ A,������ҳ B,(" & strSQL & ") C" & _
            " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!����ID), Val(rsTmp!��ҳid))
    Else
        '���������ﱨ��
        strSQL = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����ID=[1]"
        strSQL = "Select zl_PatiWarnScheme(A.����ID) As ���ò���,A.������," & _
            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0)+Nvl(E.�ʻ����,0) as ʣ���" & _
            " From ������Ϣ A,(" & strSQL & ") B,ҽ�����˹����� D,ҽ�����˵��� E" & _
            " Where A.����ID=B.����ID(+) " & _
            " And A.����id = D.����id(+) And A.����=D.����(+) And D.����=E.����(+) And D.ҽ����=E.ҽ����(+) And D.��־(+)=1" & _
            " And A.����ID=[1]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!����ID))
    End If
    
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
    'ִ�б���:���ﲡ�˲���ID=0
    strSQL = "Select Nvl(��������,1) as ��������," & _
        " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
        " Where Nvl(����ID,0)=[1] And ���ò���=[2]"
    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!���˲���ID), CStr(NVL(rsPati!���ò���)))
    If Not rsWarn.EOF Then
        If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(Val(rsTmp!����ID))
        str����� = Mid(str�����, 2)
        For i = 1 To Len(str���)
            intR = BillingWarn(Me, mstrPrivs, rsWarn, rsTmp!����, NVL(rsPati!ʣ���, 0), cur����, cur���, NVL(rsPati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
            If InStr(",2,3,", intR) > 0 Then Exit Function
        Next
    End If
    
    FinishBillingWarn = True
    Exit Function
errH:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = NVL(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim BlnFirst As Boolean
    
    If KeyCode = vbKeyF2 Then
        If CmdSend.Enabled And CmdSend.Visible Then CmdSend_Click
    End If
    
    If KeyCode = vbKeyF3 Then
        If imgFilter.BorderStyle = cstLocate Then
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call txtFind_Validate(False)
                Call zlControl.TxtSelAll(txtFind)
                Call FindNextPati(txtFind.Tag <> txtFind.Text)
            End If
        Else
            Call SetFilter(MnuEditHandback.Checked)
        End If
    End If
    
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtFind.SetFocus
        End If
    End If
    
    'Ctrl+F4  ��IC��
    If KeyCode = vbKeyF4 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then
            If lblFind.Tag = FindType.IC�� Then
                Call cmdIC_Click
            End If
        End If
    End If
End Sub

Private Sub FindNextPati(ByVal BlnFirst As Boolean)
    Static intStar As Integer
    Dim n As Integer
    Dim strFind As String
    Dim blnDo As Boolean
    
    If BlnFirst Then intStar = 1
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    strFind = Trim(txtFind.Text)
    
    With Msf�б�
        If .Rows < 2 Then Exit Sub
        
        For n = intStar To .Rows - 1
            Select Case lblFind.Tag
                Case FindType.���￨
                    If Trim(.TextMatrix(n, ��������.���￨��)) = strFind Then blnDo = True
                Case FindType.�����
                    If Trim(.TextMatrix(n, ��������.�����)) = strFind Then blnDo = True
                Case FindType.���ݺ�
                    If Trim(.TextMatrix(n, ��������.NO)) = strFind Then blnDo = True
                Case FindType.����
                    If mblnCard = True Then
                        If Trim(.TextMatrix(n, ��������.���￨��)) = strFind Then blnDo = True
                    Else
                        If gbytCode = 1 Then
                            If Trim(.TextMatrix(n, ��������.����)) Like "*" & strFind & "*" Or mWBX(Trim(.TextMatrix(n, ��������.����)), 1) Like "*" & UCase(strFind) & "*" Then blnDo = True
                        Else
                            If Trim(.TextMatrix(n, ��������.����)) Like "*" & strFind & "*" Or mPinYin(Trim(.TextMatrix(n, ��������.����))) Like "*" & UCase(strFind) & "*" Then blnDo = True
                        End If
                    End If
                Case FindType.���֤
                    If Trim(.TextMatrix(n, ��������.���֤)) = strFind Then blnDo = True
                Case FindType.IC��
                    If Trim(.TextMatrix(n, ��������.IC��)) = strFind Then blnDo = True
            End Select
            
            If blnDo Then
                txtFind.Tag = txtFind.Text
                .Row = n
                Call Msf�б�_EnterCell
                .TopRow = n
                intStar = n + 1
                If intStar > .Rows - 1 Then intStar = .Rows - 1
                Exit Sub
            End If
        Next
    End With
    intStar = 1
    txtFind.SetFocus
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    
    zlDatabase.SetPara "��ʾ��������", img����.BorderStyle, glngSys, 1341
    
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & Me.Name, "���涨λ", imgFilter.BorderStyle)
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & Me.Name, "��ʾ��ҩ��������", Chk��ʾ��ҩ��������.Value)
    
    '��������
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "δ��ҩ��������", strOrder_1)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ҩ��������", strOrder_2)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "δ��ҩ��������", strOrder_3)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�ѷ�ҩ��������", strOrder_4)
    
    '��������ģʽ
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ģʽ", mint����ģʽ)
        
    If Not InDesign And glngOld > 0 Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    Call SaveFlexState(Bill������ϸ.MsfObj, Me.Name & "\" & tabShow.Tab)
    Call SaveFlexState(Msf�б�, Me.Name & "\" & tabShow.Tab)
    SaveWinState Me, App.ProductName
    
    'ж�����֤ˢ���ӿ�
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    
    'ж��IC��ˢ���ӿ�
    Set mobjICCard = Nothing
    
    'ж�ص���ǩ���ӿ�
    Set gobjESign = Nothing
End Sub





Private Sub imgFilter_Click()
    imgFilter.BorderStyle = Abs(imgFilter.BorderStyle - 1)
    If imgFilter.BorderStyle = cstFilter Then
        Msf�б�.ColWidth(��������.ѡ��) = IIf(MnuEditConsignment.Checked, 300, 0)
    Else
        Msf�б�.ColWidth(��������.ѡ��) = 0
    End If
    
    txtFind.Text = ""
    mstrFilter = ""
    Call mnuViewRefresh_Click
End Sub

Private Sub img����_Click()
    With img����
        .BorderStyle = Abs(.BorderStyle - 1)
    End With
    Call SetPosition
    Call mnuViewRefresh_Click
End Sub


Private Sub lblFind_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        PopupMenu mnuViewLocate, 2, fraFind.Left + lblFind.Left - 30, fraFind.Top + lblFind.Top + lblFind.Height + 30
    End If
End Sub


Private Sub mnuCancel_Click()
    Dim blnInTrans As Boolean
        
    On Error GoTo errHandle
    
    If mstrNo <> "" Or IntBillStyle <> 0 Then
        If CheckBill(5, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Sub
        
        gcnOracle.BeginTrans
        blnInTrans = True
        
        '����������˵���ǩ������ȡ����ҩ�˵���ǩ��
        If gblnҩƷʹ�õ���ǩ�� = True Then
            If DelSignatureRecored(EsignTache.send, Val(TxtNo.ItemData(TxtNo.ListIndex)), Mid(TxtNo, 1, 8), lngҩ��ID) = False Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
        End If
        
        gstrSQL = "Zl_ҩƷ�շ���¼_ȡ����ҩ(" & lngҩ��ID & "," & TxtNo.ItemData(TxtNo.ListIndex) & ",'" & Mid(TxtNo.Text, 1, 8) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-ȡ����ҩ")
        
        gcnOracle.CommitTrans
        blnInTrans = False
        
        mnuViewRefresh_Click
    End If
    Exit Sub
errHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuChange_Click()
    Dim strName As String
    
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    
    strName = zlDatabase.UserIdentify(Me, "У����ҩ��", glngSys, 1341, "��ҩ")
    
    TimeRefresh.Enabled = True
    TimePrint.Enabled = True
    
    If Trim(strName) = "" Then Exit Sub
    
    mstr�Զ���ҩ�� = strName
    
    mdate�ϴ�У��ʱ�� = zlDatabase.Currentdate

End Sub
Private Sub mnuCharge_Click()
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
            
    On Error Resume Next
    If gobjCharge Is Nothing Then
        Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
        If gobjCharge Is Nothing Then Exit Sub
    End If
    
    err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    blnOK = gobjCharge.Charge(Me, gcnOracle, glngSys, gstrDbUser, 1, 0)
    Call GlobalDeleteAtom(intAtom)
    
    '��ɻ���
    'ˢ��δ��ҩ����
    mnuViewRefresh_Click
End Sub

Private Sub MnuEditSendOther_Click()
    With FrmҩƷ������ҩ
        .In_���� = mInt����
        .In_��ҩ���� = Str����
        .In_ҩ��ID = lngҩ��ID
        .In_����� = IntCheckStock
        .In_У�鴦�� = intVerify
        .In_����δ��ҩ��ҩ = IntSendAfterDosage
        .IN_����δ��˷�ҩ = Int����δ��˴�����ҩ
        .IN_����δ�շѷ�ҩ = mint����δ�շѴ�����ҩ
        .In_Ȩ�� = mstrPrivs
        .Str��ҩ�� = IIf(Str��ҩ�� = "|��ǰ����Ա|", gstrUserName, Str��ҩ��)
        .In_����λ�� = int����λ��
        .IN_��˻��۵� = int��˻��۵�
        .In_������ҩ������ = True
        .Show 1, Me
    End With
    mnuViewRefresh_Click
End Sub

Private Sub mnuFileBack_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "ҩ��=" & lngҩ��ID)
End Sub

Private Sub mnuFileLable_Click()
    Dim int���� As Integer, strNo As String
    
    If Trim(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = "" Then Exit Sub
    
    int���� = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����))
    strNo = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
    
    '��鵥���Ƿ����
    If Not CheckBillExist(int����, strNo) Then
        MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    strUnit = GetUnit(lngҩ��ID, int����, strNo)

    If Not BillHaveHerial(strNo, int����) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
            "NO=" & strNo, "����=" & IIf(int���� = 8, 1, 2), "ҩ��=" & lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
            "NO=" & strNo, "ҩ��=" & lngҩ��ID, 2)
    End If
End Sub
Private Sub mnuStuff_Click()
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
    Dim lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    
    If Msf�б�.Rows = 1 Or Msf�б�.TextMatrix(Msf�б�.Rows - 1, ��������.NO) = "" Then
        mstrNo = ""
        lng����ID = 0
    End If
    
    If mstrNo <> "" Or IntBillStyle <> 0 Then
        gstrSQL = "Select Nvl(����id,0) ����ID From ���˷��ü�¼ Where Id=(Select ����id From ҩƷ�շ���¼ Where ����=[1] And No=[2] And Rownum=1)"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ����ID]", IntBillStyle, mstrNo)
        If Not rsTmp.EOF Then
            lng����ID = rsTmp!����ID
        End If
        rsTmp.Close
    End If
            
    On Error Resume Next
    If gobjStuff Is Nothing Then
        Set gobjStuff = CreateObject("zl9Stuff.clsStuff")
        If gobjStuff Is Nothing Then Exit Sub
    End If
    
    err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    Call gobjStuff.TransStuff(Me, gcnOracle, glngSys, gstrDbUser, lng����ID, mstrNo, lngҩ��ID, mstrStartDate, mstrEndDate)
    Call GlobalDeleteAtom(intAtom)

End Sub
Private Sub MnuEditAbolish_Click()
    '--��ʾ��ҩ--
    If MnuEditAbolish.Checked = True Then Exit Sub
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = True
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = False
    
    SetButtonState
End Sub

Private Sub MnuEditBill_Click()
    With Frm��Ʊ�ݺ�������ҩ
        .In_���� = mInt����
        .In_��ҩ���� = Str����
        .In_ҩ��ID = lngҩ��ID
        .In_����� = IntCheckStock
        .In_У�鴦�� = intVerify
        .In_����δ��ҩ��ҩ = IntSendAfterDosage
        .IN_����δ��˷�ҩ = Int����δ��˴�����ҩ
        .IN_����δ�շѷ�ҩ = mint����δ�շѴ�����ҩ
        .In_Ȩ�� = mstrPrivs
        .Str��ҩ�� = IIf(Str��ҩ�� = "|��ǰ����Ա|", gstrUserName, Str��ҩ��)
        .In_����λ�� = int����λ��
        .IN_��˻��۵� = int��˻��۵�
        .Show 1, Me
    End With
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuEditBillRestore_Click()
    frm��Ʊ�ݺ�������ҩ.In_Ȩ�� = mstrPrivs
    If Not frm��Ʊ�ݺ�������ҩ.ShowEditor(Me, lngҩ��ID, int����λ��) Then Exit Sub
    Call mnuViewRefresh_Click
End Sub

Private Sub MnuEditConsignment_Click()
    If MnuEditConsignment.Checked = True Then Exit Sub
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = True
    MnuEditHandback.Checked = False
    
    SetButtonState
End Sub

Private Sub MnuEditDosage_Click()
    '--��ʾ��ҩ--
    If MnuEditDosage.Checked = True Then Exit Sub
    MnuEditDosage.Checked = True
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = False
    
    SetButtonState
End Sub

Private Sub MnuEditHandback_Click()
    If MnuEditHandback.Checked = True Then Exit Sub
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = True
    
    SetButtonState
End Sub

Private Sub mnuEditHandbackBatch_Click()
    frm������ҩ.In_Ȩ�� = mstrPrivs
    If Not frm������ҩ.ShowEditor(Me, lngҩ��ID, True, int����λ��) Then Exit Sub
    Call mnuViewRefresh_Click
End Sub

Private Sub MnuFileBillprint_Click()
    Dim int���� As Integer, strNo As String
    
    If Trim(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = "" Then Exit Sub
    
    int���� = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����))
    strNo = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
    
    '��鵥���Ƿ����
    If Not CheckBillExist(int����, strNo) Then
        MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
        Call mnuViewRefresh_Click
        Exit Sub
    End If
    
    strUnit = GetUnit(lngҩ��ID, int����, strNo)

    If Not BillHaveHerial(strNo, int����) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strNo, "����=" & IIf(int���� = 8, 1, 2), "ҩ��=" & lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=2", "PrintEmpty=0", 1)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strNo, "����=" & IIf(int���� = 8, 1, 2), "ReportFormat=2", "PrintEmpty=0", 1)
    End If
End Sub
Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub MnuFilePara_Click()
    BlnSetParaSuccess = False
    BlnRefresh = False
    
    '�ر�Timer
    Call SetTimerState(False)
    
    With Frm��ҩ��������
        Set .RecPart = RecPart.Clone
        .mstrPrivs = mstrPrivs
        .Show 1, Me
    End With
    
    If Not BlnSetParaSuccess Then
        '�����ޱ仯ʱ
    
        '����Timer
        Call SetTimerState(True)
    Else
        '�����б仯ʱ
        Call ReadFromReg
        
        '����ʱ��ؼ�
        If mlngRefresh > 0 Then
            If mlngRefresh > 60 Then
                mlngRefresh = 60
            End If
            With TimeRefresh
                .Enabled = True
                .Interval = mlngRefresh * 1000
            End With
        Else
            TimeRefresh.Enabled = False
        End If
        
        If mlngPrintInterval > 0 Then
            If mlngPrintInterval > 60 Then
                mlngPrintInterval = 60
            End If
            With TimePrint
                .Enabled = True
                .Interval = mlngPrintInterval * 1000
            End With
        Else
            TimePrint.Enabled = False
        End If
        
        IntTimes = 0
        
        If mIntPrintHandbackNO <> 0 Then
            With TimePrintCancelBill
                .Enabled = False
                .Enabled = True
            End With
        Else
            TimePrintCancelBill.Enabled = False
        End If
        
        If CheckAnother = False Then Exit Sub
        Call SetFormat(2, True)
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFileReport_Click()
    Dim strҩ�� As String
    Dim rsPart As New ADODB.Recordset
    
    gstrSQL = "Select ���� From ���ű� Where ID=[1]"
    Set rsPart = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ��ǰҩ��������]", lngҩ��ID)
    
    strҩ�� = rsPart!����
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, "�ⷿ=" & strҩ�� & "|" & lngҩ��ID, "��װϵ��=" & IIf(mintUnit = mconint���ﵥλ, "D.�����װ", "D.סԺ��װ"))
End Sub
Private Sub MnuFileRePrint_Click()
    Dim strPrintNO As String, intBillType As Integer
    
    If Not MnuEditHandback.Checked Then
        If strBill = "" Then Exit Sub
        
        strPrintNO = Split(strBill, "|")(0)
        intBillType = Val(Split(strBill, "|")(1))
    Else
        If Trim(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = "" Then Exit Sub
        
        intBillType = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����))
        strPrintNO = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
    End If
    
    strUnit = GetUnit(lngҩ��ID, intBillType, strPrintNO)
    
    If Not BillHaveHerial(strPrintNO, intBillType) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
            "NO=" & strPrintNO, _
            "����=" & IIf(intBillType = 8, 1, 2), _
            "ҩ��=" & lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), _
            "ReportFormat=1", "PrintEmpty=0", 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
            "NO=" & strPrintNO, _
            "����=" & IIf(intBillType = 8, 1, 2), _
            "ReportFormat=1", "PrintEmpty=0", 2)
    End If
End Sub

Private Sub mnuFileRestore_Click()
    '��ӡ�˷�֪ͨ��
    Dim int���� As Integer, strNo As String
    Dim Str��ҩʱ�� As String, Int��װϵ�� As Integer
    If Trim(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = "" Then Exit Sub
    If Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) <> 3 Then Exit Sub
    
    int���� = Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)
    strNo = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
    Str��ҩʱ�� = Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)
    strUnit = GetUnit(lngҩ��ID, int����, strNo)
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_1", "ZL8_BILL_1341_1"), _
    Me, "No=" & strNo, "����=" & int����, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), "��ҩʱ��=" & Str��ҩʱ��, 2)
End Sub

Private Sub mnuFileset_Click()
    zlPrintSet
End Sub

Private Sub mnuFlag_Click()
    Dim frmFlag As New Frm���ٷ�ҩ������־
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    mnuViewRefresh_Click
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub MnuHelpWebM_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub MnuEditBatch_Click()
    With FrmҩƷ������ҩ
        .In_���� = mInt����
        .In_��ҩ���� = Str����
        .In_ҩ��ID = lngҩ��ID
        .In_����� = IntCheckStock
        .In_У�鴦�� = intVerify
        .In_����δ��ҩ��ҩ = IntSendAfterDosage
        .IN_����δ��˷�ҩ = Int����δ��˴�����ҩ
        .IN_����δ�շѷ�ҩ = mint����δ�շѴ�����ҩ
        .In_Ȩ�� = mstrPrivs
        .Str��ҩ�� = IIf(Str��ҩ�� = "|��ǰ����Ա|", gstrUserName, Str��ҩ��)
        .In_����λ�� = int����λ��
        .IN_��˻��۵� = int��˻��۵�
        .In_������ҩ������ = False
        .Show 1, Me
    End With
    mnuViewRefresh_Click
End Sub

Private Sub mnuPassItem_Click(Index As Integer)
    '���ܣ�ִ��PASS����
    'Pass
    Select Case Index
    Case 0 'ҩ���ٴ���Ϣ�ο�
        Call PassDoCommand(101)
    Case 1 'ҩƷ˵����
        Call PassDoCommand(102)
    Case 2 '�й�ҩ��
        Call PassDoCommand(107)
    Case 3 '������ҩ����
        Call PassDoCommand(103)
    Case 4 '����ֵ
        Call PassDoCommand(104)
    Case 8 'ҽҩ��Ϣ����
        Call PassDoCommand(106)
    Case 10 'ҩƷ�����Ϣ
        Call PassDoCommand(13)
    Case 11 '��ҩ;�������Ϣ
        Call PassDoCommand(14)
    Case 12 'ҽԺҩƷ��Ϣ
        Call PassDoCommand(105)
    End Select
End Sub
Private Sub mnuPassSpec_Click(Index As Integer)
    '���ܣ�ִ��ר��PASS����
    'Pass
    Select Case Index
    Case 0 'ҩ��-ҩ���໥����
        Call PassDoCommand(201)
    Case 1 'ҩ��-ʳ���໥ʹ��
        Call PassDoCommand(202)
    Case 3 '����ע�������
        Call PassDoCommand(203)
    Case 4 '����ע�������
        Call PassDoCommand(204)
    Case 6 '����֢
        Call PassDoCommand(205)
    Case 7 '������
        Call PassDoCommand(206)
    Case 9 '��������ҩ
        Call PassDoCommand(207)
    Case 10 '��ͯ��ҩ
        Call PassDoCommand(208)
    Case 11 '��������ҩ
        Call PassDoCommand(209)
    Case 12 '��������ҩ
        Call PassDoCommand(210)
    End Select
End Sub


Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ�����ҩƷ=ҩƷid��ҩ��=ҩ��id��NO=����NO����������=ҩƷ�շ���¼.���ݣ�����ID=����ID
    Dim lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    If Split(mnuReportItem(Index).Tag, ",")(1) = "ZL1_INSIDE_1341" Then
        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1341", "ZL8_INSIDE_1341"), Me)
    Else
        If mstrNo <> "" Or IntBillStyle <> 0 Then
            strSQL = "Select Nvl(����id, 0) ����id From ���˷��ü�¼ Where Id=(Select ����id From ҩƷ�շ���¼ Where ����=[1] And No=[2] And Rownum=1)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "[ȡ����ID]", IntBillStyle, mstrNo)
            If Not rsTmp.EOF Then
                lng����ID = rsTmp!����ID
            End If
            rsTmp.Close
        End If
        
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "ҩƷ=" & IIf(SQLCondition.lngҩƷID = 0, "", SQLCondition.lngҩƷID), _
            "ҩ��=" & IIf(lngҩ��ID = 0, "", lngҩ��ID), _
            "NO=" & mstrNo, _
            "��������=" & IIf(IntBillStyle = 0, "", IntBillStyle), _
            "����ID=" & IIf(lng����ID = 0, "", lng����ID))
    End If
End Sub



Private Sub MnuViewFind_Click()
    Dim strReturn As String, IntOper As Integer
    
    If MnuEditDosage.Checked Then
        IntOper = 1
    ElseIf MnuEditAbolish.Checked Then
        IntOper = 2
    ElseIf MnuEditConsignment.Checked Then
        IntOper = 3
    Else
        IntOper = 4
    End If
    
    With FrmҩƷ��ҩ����
        strReturn = .ShowME(Me, lngҩ��ID, mInt����, IntOper, mstrPrivs, mbln���￨, _
            SQLCondition.date��ʼ����, _
            SQLCondition.date��������, _
            SQLCondition.str��ʼNO, _
            SQLCondition.str����NO, _
            SQLCondition.str����, _
            SQLCondition.str���￨, _
            SQLCondition.str��ʶ��, _
            SQLCondition.lng����ID, _
            SQLCondition.str������, _
            SQLCondition.str�����, _
            SQLCondition.lngҩƷID, _
            SQLCondition.strҽ����, _
            mint��Ժ��ҩ)
        If strReturn = "" Then Exit Sub
    End With
    
    mstrStartDate = Format(SQLCondition.date��ʼ����, "yyyy-mm-dd hh:mm:ss")
    mstrEndDate = Format(SQLCondition.date��������, "yyyy-mm-dd hh:mm:ss")
    
    Select Case IntOper
    Case 1
        StrFind_1 = strReturn
    Case 2
        StrFind_2 = strReturn
    Case 3
        StrFind_3 = strReturn
    Case 4
        StrFind_4 = strReturn
    End Select
    
    If imgFilter.BorderStyle = cstFilter Then
        Call txtFind_KeyPress(13)
    Else
        Call mnuViewRefresh_Click
    End If
End Sub

Private Function DataRefresh() As Boolean
    Dim lngRow As Long, lngColor As Long     'ѭ�����ұ���
    Dim IntBillThis As Integer, StrNoThis As String
    Dim LngSelectRow As Long, intCol As Integer             '��ǰѡ����
    Dim strCond As String
    Dim strSendType As String
    Dim str�������� As String
    Dim strCon���� As String
    Dim blnҽ���� As Boolean
    Dim strSqlConҽ���� As String
    
    '--���ݵ�ǰ״̬ˢ������--
    On Error Resume Next
    err = 0
    DataRefresh = True
    If BlnInOper Then Exit Function
    If BlnInRefresh Then Exit Function

    Call zlCommFun.ShowFlash
    stbThis.Panels(2) = "����ˢ������,���Ժ�..."
    
    '����ؼ�ԭ����
    ClearCons
    BlnInRefresh = True
    DataRefresh = False
    Chkȫ��.Enabled = False
    
    If imgFilter.BorderStyle = cstFilter And Trim(txtFind.Text) = "" Then
        mstrFilter = " And 1 = 2 "
    End If
    
    strCon���� = ""
    If mbln���������� Then
        If img����.BorderStyle = 0 Then
            '����ʾ��������
            strCon���� = " And (D.�����־ <> 2 Or (D.�����־ = 2 And D.���˲���id <> D.��������id)) "
        End If
        If img����.BorderStyle = 1 And cbo����.ListIndex <> -1 Then
            'Ҫ��ʾ�������������Ҳ��˲������ڵ�ǰѡ��Ĳ���
            strCon���� = " And D.���˲���id = " & cbo����.ItemData(cbo����.ListIndex)
        End If
    End If
    
    If mInt���� = 0 Then
        strCond = " And A.���� In (8,9)" '���ＰסԺ���е���
    Else
        If mInt���� = 8 Then
            strCond = " And A.���� In (8,9) And A.��ҳID Is NULL " '���ﻮ�ۼ��������
        Else
            strCond = " And A.���� IN (8,9) And A.��ҳID Is Not NULL " 'סԺ����
        End If
    End If
    
    If mlng�������� = 0 Then
        str�������� = " And C.��¼״̬=1 "
    Else
        str�������� = " And MOD(C.��¼״̬,3)=1 "
    End If
    
    blnҽ���� = (SQLCondition.strҽ���� <> "")
    strSqlConҽ���� = IIf(blnҽ���� = True, " And B.ҽ����=[17] ", "")
    
    '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
    If mint��Ժ��ҩ = 0 Then
    ElseIf mint��Ժ��ҩ = 1 Then
        strSendType = " And Not Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    ElseIf mint��Ժ��ҩ = 2 Then
        strSendType = " And Ltrim(To_Char(Nvl(C.����,0),'00')) Like '_3'"
    End If
    
    Lbl��ҩ��.Caption = "��ҩ��"
    
    CmdSend.Visible = True
        
    '���еĲ�ѯ������һ���������ų��ѱ��Ϊ����ҩ�ļ�¼  by lyq 20050416
    If MnuEditDosage.Checked Then
        '��ȡ����
        gstrSQL = " Select '' As ��ɫ, ��������,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���,˵��,���￨��,�����,���֤��,IC����,����ID, ��¼״̬ As δ���,Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս�� " & _
                  " From (" & _
                  "     Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID, Decode(D.��¼״̬, 0, 1, 0) ��¼״̬ ,d.ʵ�ս��, A.�������� " & _
                  "     From " & _
                  "         (Select B.���￨��,B.�����,B.���֤��,B.IC����,B.סԺ��,A.���ȼ�,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����,A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, Nvl(A.��������, 0) �������� " & _
                  "         From δ��ҩƷ��¼ A,������Ϣ B" & _
                  "         Where A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & " ANd (A.�ⷿID=[13] " & IIf(Str���� = "", "", " And (A.��ҩ���� IN(" & Str���� & ") Or A.��ҩ���� Is NULL)") & " Or A.�ⷿID Is NULL)" & _
                  "         " & strCond & mstrShowBill & _
                  "         And A.��ҩ�� Is Null " & strSqlConҽ���� & " ) A,ҩƷ�շ���¼ C, ���˷��ü�¼ D" & _
                  "     Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & str�������� & strSendType & " And (C.�ⷿid=[13] Or C.�ⷿid Is null) " & IIf(mstrSourceDep = "", "", " And C.�Է�����id+0 in(" & mstrSourceDep & ") ") & _
                        IIf(StrFind_1 = "", " And A.�������� " & StrDate, StrFind_1) & mstrFilter & strCon���� & ") A" & _
                  "     GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.��¼״̬, A.��������"
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '��ҩ
        
        With Msf�б�
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "����" & RecPhysic.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(RecPhysic) & "Ԫ"
                If tabShow.Tab = 2 And mblnStarPass Then
                    cmdAlley.Visible = True
                End If
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
                If cmdAlley.Visible = True Then cmdAlley.Visible = False
            End If
            Call SetFormat(1, RecPhysic.EOF)
        End With
        
        CmdSend.Caption = "��ҩ(&V)"
        
        If mint�Զ���ҩ = 1 Then
            CmdSend.Visible = False
            MnuEditDosage.Visible = False
        Else
            MnuEditDosage.Visible = (IntSendAfterDosage = 0 And IsHavePrivs(mstrPrivs, "��ҩ"))
            CmdSend.Enabled = (RecPhysic.EOF <> True) And IsHavePrivs(mstrPrivs, "��ҩ")
        End If
    End If
    If MnuEditAbolish.Checked Then
        gstrSQL = " Select '' As ��ɫ, ��������,'' As ѡ��,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���,˵��,���￨��,�����,���֤��,IC����,����ID, ��¼״̬ As δ���,Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս�� " & _
                  " From (" & _
                  "     Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID, Decode(D.��¼״̬, 0, 1, 0) ��¼״̬ ,d.ʵ�ս��, A.�������� " & _
                  "     From " & _
                  "         (Select B.���￨��,B.�����,B.���֤��,B.IC����,B.סԺ��,A.���ȼ�,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����,A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, Nvl(A.��������, 0) �������� " & _
                  "         From δ��ҩƷ��¼ A,������Ϣ B" & _
                  "         Where A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & " ANd (A.�ⷿID=[13] " & IIf(Str���� = "", "", " And (A.��ҩ���� IN(" & Str���� & ") Or A.��ҩ���� Is NULL)") & " Or A.�ⷿID Is NULL)" & _
                  "         " & strCond & mstrShowBill & _
                  "         And A.��ҩ�� Is Not Null " & strSqlConҽ���� & ") A,ҩƷ�շ���¼ C, ���˷��ü�¼ D" & _
                  "     Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & str�������� & strSendType & " And (C.�ⷿid=[13] Or C.�ⷿid Is null) " & IIf(mstrSourceDep = "", "", " And C.�Է�����id+0 in(" & mstrSourceDep & ") ") & _
                        IIf(StrFind_2 = "", " And A.�������� " & StrDate, StrFind_2) & mstrFilter & strCon���� & ") A" & _
                  "     GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.��¼״̬, A.��������"
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function 'δ��ҩƷ��¼
        
        With Msf�б�
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "����" & RecPhysic.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(RecPhysic) & "Ԫ"
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
            End If
            Call SetFormat(1, RecPhysic.EOF)
        End With
        
        CmdSend.Caption = "ȡ����ҩ(&C)"
        CmdSend.Enabled = (RecPhysic.EOF <> True) And IsHavePrivs(mstrPrivs, "��ҩ")
    End If
    If MnuEditConsignment.Checked Then
        gstrSQL = " Select '' As ��ɫ, ��������,'' As ѡ��,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,to_Char(Sum(Round(���۽��," & mintMoneyDigit & ")),'" & mstrOracleMoneyForamt & "') AS ���,����,�ɲ���,˵��,���￨��,�����,���֤��,IC����,����ID, ��¼״̬ As δ���,Sum(Round(ʵ�ս��," & mintMoneyDigit & ")) ʵ�ս�� " & _
                  " From (" & _
                  "     Select A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID, Decode(D.��¼״̬, 0, 1, 0) ��¼״̬,d.ʵ�ս��, A.�������� " & _
                  "     From " & _
                  "         (Select B.���￨��,B.�����,B.���֤��,B.IC����,B.סԺ��,A.���ȼ�,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����,A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, Nvl(A.��������, 0) �������� " & _
                  "         From δ��ҩƷ��¼ A,������Ϣ B" & _
                  "         Where A.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & " ANd (A.�ⷿID=[13] " & IIf(Str���� = "", "", " And (A.��ҩ���� IN(" & Str���� & ") Or A.��ҩ���� Is NULL)") & " Or A.�ⷿID Is NULL)" & _
                  "         " & strCond & mstrShowBill & _
                        IIf(IntSendAfterDosage = 0, " And A.��ҩ�� Is Not Null", "") & strSqlConҽ���� & _
                  "     ) A,ҩƷ�շ���¼ C, ���˷��ü�¼ D" & _
                  "     Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & str�������� & strSendType & " And (C.�ⷿid=[13] Or C.�ⷿid Is null) " & IIf(mstrSourceDep = "", "", " And C.�Է�����id+0 in(" & mstrSourceDep & ") ") & _
                        IIf(StrFind_3 = "", " And A.�������� " & StrDate, StrFind_3) & mstrFilter & strCon���� & ") A" & _
                  "     GROUP BY A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.��¼״̬, A.��¼״̬, A.��������"
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '��ȡ����δ��ҩƷ��¼
        
        With Msf�б�
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "����" & RecPhysic.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(RecPhysic) & "Ԫ"
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
            End If
            Call SetFormat(1, RecPhysic.EOF)
            Call SetCheckBox(-1)
        End With
        
        CmdSend.Caption = "��ҩ(&S)"
        CmdSend.Enabled = (RecPhysic.EOF <> True) And IsHavePrivs(mstrPrivs, "��ҩ")
    End If
    If MnuEditHandback.Checked Then
        strCon���� = Replace(strCon����, "D.", "H.")
    
        Lbl��ҩ��.Caption = "��ҩ��"
        strCond = Replace(strCond, "A.��ҳID", "H.��ҳID")
        
        Dim strCond1 As String, strCond2 As String, strTemp As String
        Dim intRight As Integer, intLeft As Integer
        '��Ƕ�ײ�ѯ�У�û�����Ӳ��˷��ü�¼���������д��������ֶ�ʱ����ȥ���������������õ����˷��ü�¼��
        strCond1 = ""
        StrFind_4 = UCase(StrFind_4)
        strCond2 = StrFind_4
        intLeft = InStr(1, strCond2, " AND UPPER(H.����)")
        If intLeft <> 0 Then
            intRight = InStr(intLeft + 4, StrFind_4, " AND")
            strTemp = Mid(StrFind_4, 1, intLeft)
            If intRight <> 0 Then
                strCond1 = Mid(StrFind_4, intLeft, intRight - intLeft + 1)
                strCond2 = strTemp & Mid(StrFind_4, intRight)
            Else
                strCond1 = Mid(StrFind_4, intLeft)
                strCond2 = strTemp
            End If
        End If
        intLeft = InStr(1, strCond2, " AND UPPER(H.��ʶ��)")
        If intLeft <> 0 Then
            intRight = InStr(intLeft + 4, strCond2, " AND")
            strTemp = Mid(strCond2, 1, intLeft)
            If intRight <> 0 Then
                strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
                strCond2 = strTemp & Mid(strCond2, intRight)
            Else
                strCond1 = strCond1 & Mid(strCond2, intLeft)
                strCond2 = strTemp
            End If
        End If
        intLeft = InStr(1, strCond2, " AND UPPER(B.���￨��)")
        If intLeft <> 0 Then
            intRight = InStr(intLeft + 4, strCond2, " AND")
            strTemp = Mid(strCond2, 1, intLeft)
            If intRight <> 0 Then
                strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
                strCond2 = strTemp & Mid(strCond2, intRight)
            Else
                strCond1 = strCond1 & Mid(strCond2, intLeft)
                strCond2 = strTemp
            End If
        End If
        
        '��Ժ��ҩ:'0-����,1-������Ժ��ҩ,2-������Ժ��ҩ
        If mint��Ժ��ҩ = 0 Then
        ElseIf mint��Ժ��ҩ = 1 Then
            strSendType = " And Not Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3'"
        ElseIf mint��Ժ��ҩ = 2 Then
            strSendType = " And Ltrim(To_Char(Nvl(A.����,0),'00')) Like '_3'"
        End If
        
        
        '����κ�һ��ҩƷ�������������һ������ϸ�ֱ���������д��ڵ��������ˣ���ֱ��ͨ������UNION�󱸵ķ�ʽ���
        '���ڲ��˷��ü�¼������㣬��������Ҫ������ͨ�����˷��ü�¼������������Ӻ���Ч����ȫ��ɨ�裬��ˣ�ֻ��ͨ����������SQL UNION ������SQL�ķ�ʽ���
        If Chk�嵥.Value = 0 Then
            gstrSQL = " SELECT DISTINCT '' As ��ɫ, '' As ��������,'' As ѡ��,'0' As ��־,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����') ����,A.����,1 ���շ�,A.����� ��ҩ��," & _
                     "      A.NO,H.����,trim(to_char(sum(A.���۽��),'" & mstrOracleMoneyForamt & "')) AS ���,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS') ����,1 �ɲ���,' ' ˵��,B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,H.�����־, H.��¼���� " & _
                     " FROM " & _
                     "      (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                     "          NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬,A.��ҩ����," & _
                     "          A.���ۼ�,round(B.���۽��," & mintMoneyDigit & ") ���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID " & _
                     "      FROM" & _
                     "          (SELECT *" & _
                     "          FROM ҩƷ�շ���¼ A" & _
                     "          WHERE nvl(A.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                     "          AND A.�ⷿID+0=[13] " & strSendType & _
                     "      " & IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                     "          ) A," & _
                     "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����,SUM(A.���۽��) ���۽��" & _
                     "          FROM ҩƷ�շ���¼ A" & _
                     "          WHERE nvl(A.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL" & strSendType & _
                     "          AND A.�ⷿID+0=[13] " & IIf(mstrSourceDep = "", "", " And A.�Է�����id+0 in(" & mstrSourceDep & ") ") & _
                     "      " & IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                     "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
            gstrSQL = gstrSQL & _
                     "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0" & _
                     "     ) A,���˷��ü�¼ H,������Ϣ B" & _
                     " WHERE A.�ⷿID+0=[13] " & IIf(Str���� = "", "", " AND (A.��ҩ���� IN(" & Str���� & ") Or A.��ҩ���� Is NULL)") & _
                     " " & strCond & mstrShowSendedBill & strCond1 & mstrFilter & strCon���� & _
                     " AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0) AND A.����� IS NOT NULL AND A.����ID=H.ID AND A.ʵ������<>0 AND H.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & strSqlConҽ���� & _
                     " GROUP BY Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����'),A.����,1,A.�����,A.NO,H.����,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS'),B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID, H.�����־, H.��¼���� "
        Else
            gstrSQL = " SELECT DISTINCT '' As ��ɫ, '' As ��������,'' As ѡ��,'0' As ��־,Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����') ����,A.����,1 ���շ�,A.����� ��ҩ��," & _
                     "      A.NO,H.����,trim(to_char(sum(A.���۽��),'" & mstrOracleMoneyForamt & "')) AS ���,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS') ����,A.�ɲ���," & _
                     "      DECODE(A.��¼״̬,1,'��1�η�ҩ',DECODE(MOD(A.��¼״̬,3),0,'��1�η�ҩ',1,'��'||(FLOOR(A.��¼״̬/3)+1)||'�η�ҩ',2,'��'||(FLOOR(A.��¼״̬/3)+1)||'����ҩ')) ˵��,B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID,H.�����־, H.��¼���� " & _
                     " FROM " & _
                     "      (SELECT * FROM" & _
                     "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                     "              NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬,A.��ҩ����," & _
                     "              A.���ۼ� , round(A.���۽��," & mintMoneyDigit & ") ���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,1 �ɲ��� " & _
                     "          FROM" & _
                     "              (SELECT *" & _
                     "              FROM ҩƷ�շ���¼ A" & _
                     "              WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                     "              AND A.�ⷿID+0=[13] " & strSendType & _
                     "          " & IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                     "              ) A," & _
                     "              (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                     "              FROM ҩƷ�շ���¼ A" & _
                     "              WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and A.����� IS NOT NULL " & strSendType & _
                     "              AND A.�ⷿID+0=[13] " & IIf(mstrSourceDep = "", "", " And A.�Է�����id+0 in(" & mstrSourceDep & ") ") & _
                     "          " & IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                     "              GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
            gstrSQL = gstrSQL & _
                     "          WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.���)" & _
                     "          UNION" & _
                     "          SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                     "          NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬,A.��ҩ����," & _
                     "          A.���ۼ� , round(A.���۽��," & mintMoneyDigit & ") ���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
                     "          DECODE(��¼״̬,1,1,DECODE(MOD(��¼״̬,3),0,1,MOD(��¼״̬,3)+1)) �ɲ���" & _
                     "          FROM ҩƷ�շ���¼ A" & _
                     "          WHERE nvl(a.��ҩ��ʽ,-999)<>-1 and NOT (��¼״̬=1 OR MOD(��¼״̬,3)=0)" & IIf(mstrSourceDep = "", "", " And A.�Է�����id+0 in(" & mstrSourceDep & ") ") & strSendType & _
                     "          " & IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2)
            gstrSQL = gstrSQL & _
                     "     ) A,���˷��ü�¼ H,������Ϣ B" & _
                     " WHERE A.�ⷿID+0=[13] " & IIf(Str���� = "", "", " AND (A.��ҩ���� IN(" & Str���� & ") Or A.��ҩ���� Is NULL)") & _
                     " " & strCond & mstrShowSendedBill & strCond1 & mstrFilter & strCon���� & _
                     " AND A.����� IS NOT NULL AND A.����ID=H.ID AND H.����ID=B.����ID" & IIf(blnҽ���� = True, "", "(+)") & strSqlConҽ���� & _
                     " GROUP BY Decode(Nvl(h.��¼״̬, 0),  0,'(δ)','') || Decode(a.����, 8, '�շ�', 9, '����') ,A.����,1,A.�����," & _
                     "      A.NO,H.����,TO_CHAR(A.�������,'YYYY-MM-DD HH24:MI:SS'),A.�ɲ���," & _
                     "      DECODE(A.��¼״̬,1,'��1�η�ҩ',DECODE(MOD(A.��¼״̬,3),0,'��1�η�ҩ',1,'��'||(FLOOR(A.��¼״̬/3)+1)||'�η�ҩ',2,'��'||(FLOOR(A.��¼״̬/3)+1)||'����ҩ')),B.���￨��,B.�����,B.���֤��,B.IC����,B.����ID, H.�����־, H.��¼���� "
        End If
        
        Dim blnMoved As Boolean
        Dim str��ʼ���� As String, strSQL As String
       
        str��ʼ���� = Format(SQLCondition.date��ʼ����, "yyyy-mm-dd hh:mm:ss")
        
        '�жϴӿ�ʼ���ں��Ƿ����ת���Ĵ�������
        blnMoved = zlDatabase.DateMoved(str��ʼ����)
        
        '�����������ת��������Ҫͬʱ�Ӻ󱸱�����ȡ����
        If blnMoved Then
            strSQL = gstrSQL
            strSQL = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
            strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
            gstrSQL = gstrSQL & " UNION ALL " & strSQL
        End If
        
        If ReadData(gstrSQL) = False Then BlnInRefresh = False: Call zlCommFun.StopFlash: Exit Function '��ȡ����δ��ҩƷ��¼
        
        With Msf�б�
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                BlnRefresh = True
                stbThis.Panels(2) = "����" & RecPhysic.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(RecPhysic) & "Ԫ"
            Else
                .Clear
                .Rows = 2
                stbThis.Panels(2) = ""
            End If
            Call SetFormat(1, RecPhysic.EOF)
        End With
        
        '��ɫ
        Msf�б�.Redraw = False
        For lngRow = 1 To Msf�б�.Rows - 1
            Msf�б�.Row = lngRow
            lngColor = IIf(Val(Msf�б�.TextMatrix(lngRow, ��������.�ɲ���)) = 1, glng����, IIf(Val(Msf�б�.TextMatrix(lngRow, ��������.�ɲ���)) = 2, glng��ҩ, glng��ҩ))
            For intCol = ��������.ѡ�� To Msf�б�.Cols - 1
                Msf�б�.Col = intCol
                Msf�б�.CellForeColor = lngColor
            Next
        Next
        Msf�б�.Redraw = True
        
        CmdSend.Caption = "��ҩ(&R)"
        CmdSend.Enabled = (Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1) And IsHavePrivs(mstrPrivs, "��ҩ")
    End If
    
    Call SetFormat(2)
        
    '���ô�����ɫ
    Call SetRecipeColor
        
    '��λԭ��ѡ��Ĵ��������ʧ�ܣ���λ����һ��
    Msf�б�.Row = ReLocateRow
    Msf�б�_EnterCell
    '�󶨼�¼���������µ�����ťλ��
    ResizePicClose
    Call zlCommFun.StopFlash
    
    BlnInRefresh = False
    DataRefresh = True
End Function

'Modified By ���� 2003-12-10 ����������
Private Sub mnuViewFontSet_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSET(i).Checked = False
    Next
    Me.mnuViewFontSET(Index).Checked = True

    Select Case Index
    Case 0
        Me.Msf�б�.Font.Size = 9
        Bill������ϸ.Font.Size = 9
     Case 1
        Me.Msf�б�.Font.Size = 11
        Bill������ϸ.Font.Size = 11
    Case 2
        Me.Msf�б�.Font.Size = 15
        Bill������ϸ.Font.Size = 15
    End Select
    intFont = Index
    
    zlDatabase.SetPara "����", Index, glngSys, 1341
    
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewLocateItem_Click(Index As Integer)
    Dim strItem As String, i As Long
    
    For i = 0 To mnuViewLocateItem.UBound
        mnuViewLocateItem(i).Checked = i = Index
    Next
    strItem = Split(mnuViewLocateItem(Index).Caption, "(")(0)
    lblFind.Caption = strItem & "��"
    lblFind.Tag = Index
    mint����ģʽ = Index
    
    If Index <> FindType.IC�� Then
        cmdIC.Visible = False
        imgFilter.Left = fraFind.Width - imgFilter.Width - 80
        txtFind.Width = imgFilter.Left - txtFind.Left - 80
    End If
    
    txtFind.Text = "": txtFind.Tag = ""
    txtFind.PasswordChar = ""
    txtFind.MaxLength = 0
    
    Select Case Index
        Case FindType.���￨
            If gtype_UserSysParms.P12_���￨�Ƿ�������ʾ Then
                txtFind.PasswordChar = "*"
            End If
            txtFind.MaxLength = gtype_UserSysParms.P20_���￨�ų���
        Case FindType.IC��
            cmdIC.Visible = True
            cmdIC.Left = fraFind.Width - cmdIC.Width - 80
            imgFilter.Left = cmdIC.Left - imgFilter.Width - 80
            txtFind.Width = imgFilter.Left - txtFind.Left - 80
    End Select
        
    If Visible Then txtFind.SetFocus
End Sub

Private Sub mnuViewRefresh_Click()
    If Not BlnStartUp Then Exit Sub

    StrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    StrDate = " Between To_Date('" & StrDate & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_Date('" & StrDate & " 23:59:59','yyyy-MM-dd hh24:mi:ss') "
    
    'Modified by ZYB 2002-11-19 �����û�����
    Call SaveFlexState(Bill������ϸ.MsfObj, Me.Name & "\" & tabShow.Tab)
    Call SaveFlexState(Msf�б�, Me.Name & "\" & tabShow.Tab)
    '���¶�ȡ����
    DoEvents
    Call DataRefresh
    DoEvents
    
    '�ָ�����
    Call RestoreFlexState(Msf�б�, Me.Name & "\" & tabShow.Tab)
    Call RestoreFlexState(Bill������ϸ.MsfObj, Me.Name & "\" & tabShow.Tab)
    Bill������ϸ.ColWidth(����.�����) = IIf(Not mblnStarPass, 0, 240)
    Bill������ϸ.ColWidth(����.����) = IIf(mbln��ʾ���� And mblnIs��ҩ����, 1200, 0)
    Call SetColHide
    
    If imgFilter.BorderStyle = cstFilter Then
        Msf�б�.ColWidth(��������.ѡ��) = IIf(MnuEditConsignment.Checked, 300, 0)
    Else
        Msf�б�.ColWidth(��������.ѡ��) = 0
    End If
    
    mblnFilterRefresh = False
End Sub

Private Sub MnuViewState_Click()
    MnuViewState.Checked = MnuViewState.Checked Xor True
    stbThis.Visible = MnuViewState.Checked
    Form_Resize
End Sub

Private Sub MnuViewToolS_Click()
    MnuViewToolS.Checked = MnuViewToolS.Checked Xor True
    Cbar.Visible = MnuViewToolS.Checked
    MnuViewToolT.Enabled = MnuViewToolS.Checked
    
    Form_Resize
End Sub

Private Sub MnuViewToolT_Click()
    MnuViewToolT.Checked = MnuViewToolT.Checked Xor True
    If MnuViewToolT.Checked Then
        Tbar1.Buttons("Preview").Caption = "Ԥ��"
        Tbar1.Buttons("Print").Caption = "��ӡ"
        Tbar1.Buttons("Find").Caption = "����"
        Tbar1.Buttons("Help").Caption = "����"
        Tbar1.Buttons("Exit").Caption = "�˳�"
    Else
        Tbar1.Buttons("Preview").Caption = ""
        Tbar1.Buttons("Print").Caption = ""
        Tbar1.Buttons("Find").Caption = ""
        Tbar1.Buttons("Help").Caption = ""
        Tbar1.Buttons("Exit").Caption = ""
    End If
    
    Cbar.Bands(1).MinHeight = Tbar1.Height
End Sub

Private Sub Msf�б�_DblClick()
    Msf�б�_KeyDown vbKeyReturn, 0
End Sub

Private Sub Msf�б�_GotFocus()
    With Msf�б�
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf�б�_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = "" Then TxtNo.Clear: Exit Sub
    If KeyCode = vbKeyReturn Then TxtNo_Click
End Sub

Private Sub Msf�б�_LostFocus()
    With Msf�б�
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub SetFormat(ByVal IntStyle As Integer, Optional ByVal BlnSetHead As Boolean = True)
    Dim intCol As Integer
    Dim strSave As String
    Dim intRow As Integer, intRows As Integer
    Dim arrColumn
    '--���ø��б�ؼ��ĸ�ʽ--
    
    Select Case IntStyle
    Case 1
        With Msf�б�
            If BlnSetHead Then
                .Cols = IIf(MnuEditHandback.Checked, ��������.��ҩ����, ��������.��ҩ����)
                .TextMatrix(0, ��������.��ɫ) = "��ɫ"
                .TextMatrix(0, ��������.��������) = ""
                .TextMatrix(0, ��������.ѡ��) = ""
                .TextMatrix(0, ��������.��־) = "0"
                .TextMatrix(0, ��������.����) = "����"
                .TextMatrix(0, ��������.����) = "����"
                .TextMatrix(0, ��������.�շ�) = "�շ�"
                .TextMatrix(0, ��������.��ҩ��) = "��ҩ��"
                .TextMatrix(0, ��������.NO) = "NO"
                .TextMatrix(0, ��������.����) = "����"
                .TextMatrix(0, ��������.���) = "���"
                .TextMatrix(0, ��������.����) = "����"
                .TextMatrix(0, ��������.�ɲ���) = "�ɲ���"
                .TextMatrix(0, ��������.˵��) = "˵��"
                .TextMatrix(0, ��������.���￨��) = "���￨��"
                .TextMatrix(0, ��������.�����) = "�����"
                .TextMatrix(0, ��������.���֤) = "���֤��"
                .TextMatrix(0, ��������.IC��) = "IC����"
                .TextMatrix(0, ��������.����ID) = "����ID"
             End If
            .TextMatrix(0, ��������.ѡ��) = ""
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(��������.��ɫ) = 500
                .ColWidth(��������.��������) = 0
                .ColWidth(��������.ѡ��) = 300
                .ColWidth(��������.��־) = 0
                .ColWidth(��������.����) = 1000
                .ColWidth(��������.����) = 0
                .ColWidth(��������.�շ�) = 0
                .ColWidth(��������.��ҩ��) = 0
                .ColWidth(��������.NO) = 800
                .ColWidth(��������.����) = 800
                .ColWidth(��������.���) = 1200
                .ColWidth(��������.����) = 1500
                .ColWidth(��������.�ɲ���) = 0
                .ColWidth(��������.˵��) = 1500
                .ColWidth(��������.���￨��) = 1000
                .ColWidth(��������.�����) = 1000
                .ColWidth(��������.���֤) = 1600
                .ColWidth(��������.IC��) = 1600
                .ColWidth(��������.����ID) = 0
                .Row = 1
            End If
            
            If imgFilter.BorderStyle = cstFilter Then
                .ColWidth(��������.ѡ��) = IIf(MnuEditConsignment.Checked, 300, 0)
            Else
                .ColWidth(��������.ѡ��) = 0
            End If
            
            .ColAlignment(��������.ѡ��) = 4
            .ColAlignment(��������.���) = 7
            .ColAlignment(��������.���￨��) = 7
            .ColAlignment(��������.�����) = 7
            .ColAlignment(��������.���֤) = 7
            .ColAlignment(��������.IC��) = 7
            Call RestoreFlexState(Msf�б�, Me.Name & "\" & tabShow.Tab)
            .ColWidth(��������.����) = 0
            If MnuEditHandback.Checked Then
                .ColWidth(��������.��ɫ) = 0
                .ColWidth(��������.�����־) = 0
                .ColWidth(��������.��¼����) = 0
            Else
                .ColWidth(��������.��ɫ) = 500
                .ColWidth(��������.δ���) = 0
                .ColWidth(��������.ʵ�ս��) = 0
            End If
        End With
    Case 2
        With Bill������ϸ
            .Active = False
            .Rows = 2
            .Cols = ����.����
            
            .TextMatrix(0, ����.�����) = "��"
            .TextMatrix(0, ����.˳���) = "���"
            .TextMatrix(0, ����.ҩƷ����) = "ҩƷ����"
            .TextMatrix(0, ����.������) = "������"
            .TextMatrix(0, ����.Ӣ����) = "Ӣ����"
            .TextMatrix(0, ����.���) = "���"
            .TextMatrix(0, ����.���) = "���"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.Id) = "ID"
            .TextMatrix(0, ����.ҩƷID) = "ҩƷID"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.��λ) = "��λ"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.���) = "���"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.�÷�) = "�÷�"
            .TextMatrix(0, ����.Ƶ��) = "Ƶ��"
            .TextMatrix(0, ����.ҽ������) = "ҽ������"
            .TextMatrix(0, ����.������) = "������"
            .TextMatrix(0, ����.׼����) = "׼����"
            .TextMatrix(0, ����.׼������) = "׼������"
            .TextMatrix(0, ����.׼����С) = "׼����С"
            .TextMatrix(0, ����.��ҩ��) = "��ҩ��"
            .TextMatrix(0, ����.��ҩ����) = "��ҩ��(���װ)"
            .TextMatrix(0, ����.��λ��) = "��λ"
            .TextMatrix(0, ����.��ҩ��С) = "��ҩ��(С��װ)"
            .TextMatrix(0, ����.��λС) = "��λ"
            .TextMatrix(0, ����.�����) = "�����"
            .TextMatrix(0, ����.��λ) = "�ⷿ��λ"
            .TextMatrix(0, ����.����) = "����"
            .TextMatrix(0, ����.������) = "������"
            .TextMatrix(0, ����.��Ч��) = "��Ч��"
            .TextMatrix(0, ����.�²���) = "�²���"
            .TextMatrix(0, ����.��ע) = "��ע"
            .TextMatrix(0, ����.ҽ��id) = "ҽ��ID"
            .TextMatrix(0, ����.ʵ������) = "ʵ������"
            .TextMatrix(0, ����.�ѱ�) = "�ѱ�"
            .TextMatrix(0, ����.��װ) = "��װ"
            
            .ColWidth(����.�����) = IIf(Not mblnStarPass, 0, 240)
            .ColWidth(����.˳���) = 450
            .ColWidth(����.ҩƷ����) = 2500
            .ColWidth(����.������) = 2000
            .ColWidth(����.Ӣ����) = 2000
            .ColWidth(����.���) = 0
            .ColWidth(����.���) = 1500
            .ColWidth(����.����) = 1500
            .ColWidth(����.Id) = 0
            .ColWidth(����.ҩƷID) = 0
            .ColWidth(����.����) = 0
            .ColWidth(����.��λ) = IIf(mbln��ʾ��С��λ = True, 0, 500)
            .ColWidth(����.����) = 1000
            .ColWidth(����.����) = IIf(IntShowCol = 1, 800, 0)
            .ColWidth(����.����) = 1200
            .ColWidth(����.���) = 1200
            .ColWidth(����.����) = 1200
            .ColWidth(����.����) = 1200
            .ColWidth(����.�÷�) = 1500
            .ColWidth(����.Ƶ��) = 1500
            .ColWidth(����.ҽ������) = IIf(MnuEditHandback.Checked, 0, 1500)
            .ColWidth(����.�����) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(����.��λ) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(����.������) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(����.׼����) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(����.׼������) = 0
            .ColWidth(����.׼����С) = 0
            .ColWidth(����.��ҩ��) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = False, 1200, 0)
            .ColWidth(����.��ҩ����) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 1500, 0)
            .ColWidth(����.��ҩ��С) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 1500, 0)
            .ColWidth(����.��λ��) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 500, 0)
            .ColWidth(����.��λС) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 500, 0)
            
            .ColWidth(����.����) = 0
            .ColWidth(����.������) = 0
            .ColWidth(����.��Ч��) = 0
            .ColWidth(����.�²���) = 0
            .ColWidth(����.��ע) = 1200
            .ColWidth(����.ҽ��id) = 0
            .ColWidth(����.ʵ������) = 0
            .ColWidth(����.�ѱ�) = 1000
            .ColWidth(����.��װ) = 0
            
            .ColAlignment(0) = 1
            .ColAlignment(2) = 1
            .ColAlignment(����.ҩƷ����) = 1
            .ColAlignment(����.������) = 1
            .ColAlignment(����.Ӣ����) = 1
            .ColAlignment(����.���) = 1
            .ColAlignment(����.����) = 1
            .ColAlignment(����.��λ) = 1
            .ColAlignment(����.�÷�) = 1
            .ColAlignment(����.��ע) = 1
            
            'Modified by ZYB 2002-11-19 �ָ��û�����
            Call RestoreFlexState(.MsfObj, Me.Name & "\" & tabShow.Tab)
            .ColWidth(����.�����) = IIf(Not mblnStarPass, 0, 240)
            .ColWidth(����.����) = IIf(IntShowCol = 1, 800, 0)
            .ColWidth(����.ҽ������) = IIf(MnuEditHandback.Checked, 0, 1500)
            .ColWidth(����.�����) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(����.��λ) = IIf(MnuEditHandback.Checked, 0, 1200)
            .ColWidth(����.������) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(����.׼����) = IIf(MnuEditHandback.Checked, 1200, 0)
            .ColWidth(����.��ҩ��) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = False, 1200, 0)
            .ColWidth(����.��ҩ����) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 1500, 0)
            .ColWidth(����.��ҩ��С) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 1500, 0)
            .ColWidth(����.��λ��) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 500, 0)
            .ColWidth(����.��λС) = IIf(MnuEditHandback.Checked And mbln��ʾ��С��λ = True, 500, 0)
            .ColWidth(����.˳���) = 450
            .ColWidth(����.���) = 0
            .ColWidth(����.ʵ������) = 0
            .ColWidth(����.��װ) = 0
            .ColWidth(����.��λ) = IIf(mbln��ʾ��С��λ = True, 0, 500)
            .ColWidth(����.׼������) = 0
            .ColWidth(����.׼����С) = 0
        End With
    End Select
    
    Call SetColHide
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    
    Form_Resize
    BlnFirstStart = True
    
    If Me.Tag = "" Then
        Call tabShow_Click(tabShow.Tab)
        Me.Tag = "Refresh"
    End If
End Sub

Private Sub Form_Load()
    BlnEnterCell = False
    BlnStartUp = False
    
    lblUserName.Caption = gstrUserName
    lblUserName.Left = 0
    PicToolbar.Width = lblUserName.Width + 10
    PicToolbar.Height = Tbar1.Height
    lblUserName.Top = (PicToolbar.Height - lblUserName.Height) / 2 + 20
            
    cmdIC.Visible = False
          
    fraFind.Width = IIf(IntSendAfterDosage = 1, 3000, 3950)
    lblFind.Left = 90
    txtFind.Left = lblFind.Left + lblFind.Width + 80
    imgFilter.Left = fraFind.Width - imgFilter.Width - 80
    txtFind.Width = imgFilter.Left - txtFind.Left - 80
    
    '��ʼ�����������С�߽�
    glngMinW = 9555
    glngMinH = 6675
    glngMaxW = Screen.Width
    glngMaxH = Screen.Height
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    mint��Ժ��ҩ = 0
    mblnIsFirst = True
    mdate�ϴ�У��ʱ�� = zlDatabase.Currentdate
    mstr�Զ���ҩ�� = ""
    
    strChargePrivs = GetPrivFunc(glngSys, 1120)
    strStuffPrivs = GetPrivFunc(glngSys, 1723)
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If

    If gstrUserName = "" Then
        MsgBox "��Ϊ��ǰ�û����ö�Ӧ�Ĳ���Ա����ʹ�ñ�ģ�飡", vbInformation, gstrSysName
        Exit Sub
    End If
   
    'ȡϵͳ����
    Call GetSysParms
    
    'ȡ���λ��
    mintMoneyDigit = GetDigit(0, 1, 4)
    '���ý���ʽ
    Call GetMoneyFormat
    
    Call TradeName
    'Ϊ���ؼ�װ��ͼ��
    If LoadInIcon = False Then Exit Sub
    '�������ݼ��
    If DependOnCheck = False Then Exit Sub
    '��ע�����ȡ���û�����
    Call ReadFromReg
    '����������
    If CheckAnother = False Then Exit Sub
    Call mnuViewFontSet_Click(intFont)
    Lbl����.Caption = GetUnitName & Lbl����.Caption
    
    Set mobjIDCard = New clsIDCard
    
    '����ǩ���ӿڿ���
    If gblnҩƷʹ�õ���ǩ�� = True Then
        On Error Resume Next
        gblnҩƷʹ�õ���ǩ�� = False
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        err.Clear: On Error GoTo 0
        If Not gobjESign Is Nothing Then
            If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                Set gobjESign = Nothing
                gblnҩƷʹ�õ���ǩ�� = False
            End If
        End If
        gblnҩƷʹ�õ���ǩ�� = True
    End If
    
    Call mnuViewLocateItem_Click(mint����ģʽ)
    
    '��ʼ�����
    StrLastNo = ""
    IntLastBill = 0
    strLastData = ""
    mintLastSequence = 1
    StrFindStyle = "%"
    LngSendRow = 0
    Intģʽ = 1
    BlnFirstStart = False
    BlnInOper = False
    BlnAllowClick = True
    BlnInRefresh = False
    
    StrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    mstrStartDate = StrDate & " 00:00:00"
    mstrEndDate = StrDate & " 23:59:59"
    
    StrDate = " Between To_Date('" & StrDate & " 00:00:00','yyyy-MM-dd hh24:mi:ss') And To_Date('" & StrDate & " 23:59:59','yyyy-MM-dd hh24:mi:ss') "
    
    StrFind_1 = "": StrFind_2 = "": StrFind_3 = "": StrFind_4 = ""
    
    MnuEditDosage.Checked = False
    MnuEditAbolish.Checked = False
    MnuEditConsignment.Checked = False
    MnuEditHandback.Checked = False
    
    Call SetFormat(2, True)
    If IntSendAfterDosage = 0 Then
        If IsHavePrivs(mstrPrivs, "��ҩ") Then
            MnuEditDosage_Click
        End If
    Else
        If IsHavePrivs(mstrPrivs, "��ҩ") Then
            MnuEditConsignment_Click
        ElseIf IsHavePrivs(mstrPrivs, "��ҩ") Then
            MnuEditHandback_Click
        Else
            MnuEditConsignment_Click
        End If
    End If
    
    If glngSys \ 100 = 1 Then
        Me.Caption = "ҩƷ������ҩ"
    Else
        Me.Caption = "ҩ�괦����ҩ"
        Me.Lbl����.Caption = "����"
        Me.Lbl����.Visible = False
        Me.Txt����.Visible = False
        Me.TxtסԺ��.Visible = False
        Me.LblסԺ��.Visible = False
    End If
    Call SetFormat(1, True)
    Call SetFormat(2, True)
    Call Ȩ�޿���
    
    Call mnuViewRefresh_Click
    Call RestoreWinState(Me, App.ProductName)
    
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs, "ZL1_INSIDE_1341")
    
    Call SetFormat(1, True)
    Call SetFormat(2, True)
    StrLastNo = ""
    Call Msf�б�_EnterCell
    
    MnuEditDosage.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")
    MnuEditAbolish.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")
    MnuEditBatch.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")
    mnuEditBillRestore.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")

    '����ʱ��ؼ�
    TimeRefresh.Enabled = False
    TimePrint.Enabled = False
    If mlngRefresh > 0 Then
        If mlngRefresh > 60 Then
            mlngRefresh = 60
        End If
        With TimeRefresh
            .Enabled = True
            .Interval = mlngRefresh * 1000
        End With
    End If
    
    If mlngPrintInterval > 0 Then
        If mlngPrintInterval > 60 Then
            mlngPrintInterval = 60
        End If
        With TimePrint
            .Enabled = True
            .Interval = mlngPrintInterval * 1000
        End With
    End If
    IntTimes = 0
    If mIntPrintHandbackNO <> 0 Then
        With TimePrintCancelBill
            .Enabled = False
            .Enabled = True
        End With
    Else
        TimePrintCancelBill.Enabled = False
    End If
    
    BlnStartUp = True
    BlnEnterCell = True
    
    ImgLeftRight_S.Left = IIf(IntSendAfterDosage = 1, 3100, 4500)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    Dim DblWidth As Double, DblHeight As Double
    
    If IntSendAfterDosage = 0 Then
        If Me.Width < 13000 Then
            Me.Width = 13000
        End If
    Else
        If Me.Width < 11500 Then
            Me.Width = 11500
        End If
    End If
    If Me.Height < 8250 Then
        Me.Height = 8250
    End If
    
    tabShow.Width = IIf(IntSendAfterDosage = 1, 2500, 3930)
    Msf�б�.ZOrder 0
    
    With ImgLeftRight_S
        If .Left < IIf(IntSendAfterDosage = 1, 3100, 4500) Then .Left = IIf(IntSendAfterDosage = 1, 3100, 4500)
        If .Left > IIf(IntSendAfterDosage = 1, 3950, 5500) Then .Left = IIf(IntSendAfterDosage = 1, 3950, 5500)
    End With
    
    PicToolbar.Top = Tbar1.Top
    PicToolbar.Left = Me.Width - PicToolbar.Width - 200
        
    With Cbar
        .Align = 1
        If BlnFirstStart = False Then
            'Set .Bands(1).Child = Tbar1
            .Bands(1).MinHeight = Tbar1.Height
        End If
    End With
    
    With fraFind
        .Top = IIf(Cbar.Visible, Cbar.Height, 0)
        .Left = 10
        
        .Width = ImgLeftRight_S.Left - .Left - 80
        
        If cmdIC.Visible = True Then
            cmdIC.Left = .Width - cmdIC.Width - 80
            imgFilter.Left = cmdIC.Left - imgFilter.Width - 80
            txtFind.Width = imgFilter.Left - txtFind.Left - 80
        Else
            imgFilter.Left = .Width - imgFilter.Width - 80
            txtFind.Width = imgFilter.Left - txtFind.Left - 80
        End If
    End With
    
    With ImgLeftRight_S
        .Top = fraFind.Top + fraFind.Height - 20
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        DblHeight = .Height
    End With
    
    With tabShow
        .Top = fraFind.Top + fraFind.Height + 50
        .Left = 0
    End With
    
    With Msf�б�
        .Top = ImgLeftRight_S.Top + tabShow.Height
        .Height = ImgLeftRight_S.Height - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = ImgLeftRight_S.Left
        .Left = 0
    End With
    With PicCloseConsignment
        .Top = Msf�б�.Top + 30
    End With
    
    '�������ݴ�С
    DblWidth = Me.ScaleWidth - (ImgLeftRight_S.Left + ImgLeftRight_S.Width)
    With PicBackGroud
        .Left = ImgLeftRight_S.Left + ImgLeftRight_S.Width
        .Top = ImgLeftRight_S.Top
        .Width = DblWidth
        .Height = ImgLeftRight_S.Height
        .ZOrder 0
    End With
    
    '���������л���ť��������ѡ����λ������
    Call SetPosition
    
    With Lbl����
        .Width = DblWidth
    End With
    With PicState
        .Left = DblWidth - .Width - 80
    End With
    With TxtNo
        .Left = DblWidth - .Width - 80
    End With
    With LblNo
        .Left = TxtNo.Left - 80 - .Width
    End With
    
    With Txt����
        .Left = DblWidth - .Width - 80
    End With
    With Lbl����
        .Left = Txt����.Left - .Width - 80
    End With
    
    With Txt����ҽ��
        .Top = DblHeight - .Height - 100
    End With
    With Lbl����ҽ��
        .Top = Txt����ҽ��.Top + 60
    End With
    
    With Lbl��ҩ��
        .Top = Lbl����ҽ��.Top
    End With
    With cbo��ҩ��
        .Top = Txt����ҽ��.Top
    End With
    
    With Txt�շ�Ա
        .Top = Txt����ҽ��.Top
    End With
    With Lbl�շ�Ա
        .Top = Lbl����ҽ��.Top
    End With
    
    With CmdSend
        .Left = DblWidth - .Width - 50
        .Top = Txt����ҽ��.Top - 25
    End With
    
    With Chkȫ��
        .Top = Lbl����ҽ��.Top
        .Left = CmdSend.Left - Chkȫ��.Width - 150
    End With
    
    With txtԭʼ����
        .Top = DblHeight - .Height - 450
    End With
    
    With lblԭʼ����
        .Top = txtԭʼ����.Top + 60
    End With
    
    With txt��ҩ�巨
        .Top = txtԭʼ����.Top
    End With
        
    With lbl��ҩ�巨
        .Top = lblԭʼ����.Top
    End With

    If mblnIs��ҩ���� Then
        With Bill������ϸ
            .Top = Txt����.Top + Txt����.Height + 50
            .Height = IIf(txtԭʼ����.Top - .Top - 50 < 0, .Height, txtԭʼ����.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
        End With
    Else
        With Bill������ϸ
            .Top = Txt����.Top + Txt����.Height + 50
            .Height = IIf(cbo��ҩ��.Top - .Top - 50 < 0, .Height, cbo��ҩ��.Top - .Top - 50)
            .Width = IIf(DblWidth - .Left - 80 < 0, .Width, DblWidth - .Left - 80)
        End With
    End If
    
    '��������ͷ�ϵĿؼ�
    With Txt����
        If glngSys \ 100 = 1 Then
            .Left = DblWidth / 2 - .Width / 2
        Else
            .Left = DblWidth - .Width - 100
        End If
    End With
    With Lbl����
        .Left = Txt����.Left - .Width - 50
    End With
    With Txt�Ա�
        If glngSys \ 100 = 1 Then
            .Left = DblWidth / 3 - .Width / 2
        Else
            .Left = DblWidth / 2 - .Width / 2
        End If
    End With
    With Lbl�Ա�
        .Left = Txt�Ա�.Left - .Width - 50
    End With
    With TxtסԺ��
        .Left = (Txt����.Left - (Txt����.Left + Txt����.Width) / 2) + TxtסԺ��.Width / 2
    End With
    With LblסԺ��
        .Left = TxtסԺ��.Left - .Width
    End With
    
    ResizePicClose

End Sub

Private Sub ImgLeftRight_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight_S
        .Move .Left + x
    End With
    Form_Resize
End Sub
Private Function LoadInIcon() As Boolean
    '--Ϊ���ؼ�װ��ͼ��--
    On Error Resume Next
    err = 0
    LoadInIcon = False
    
    '������
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BCHARGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSTUFF", vbResIcon)
    End With
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CDOSAGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CEXIT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CCHARGE", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSTUFF", vbResIcon)
    End With
    With Tbar1
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor
        
        .Buttons("Preview").Image = 1
        .Buttons("Print").Image = 2
        .Buttons("Cancel").Image = 3
        .Buttons("Find").Image = 7
        .Buttons("Help").Image = 8
        .Buttons("Exit").Image = 9
        .Buttons("Charge").Image = 10
        .Buttons("Stuff").Image = 11
    End With
    Cbar.Bands(1).MinHeight = Tbar1.Height
    
    RaisEffect PicCloseConsignment, 2
    
    If err <> 0 Then
        MsgBox "�����Դ�ļ���ʧ�����������������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function ReadData(ByVal StrQuery As String, Optional ByVal IntStyle As Integer = 0) As Boolean
    Dim strOrder As String
    '--��ȡ���ݣ������û���Ҫ���������--
    'IntStyle:0-δ�䴦��;1-���䴦��;2-δ������;3-�ѷ�����
    
    On Error Resume Next
    err = 0
    ReadData = False
    
    gstrSQL = StrQuery
    
    Set RecPhysic = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.date��ʼ����, _
            SQLCondition.date��������, _
            SQLCondition.str��ʼNO, _
            SQLCondition.str����NO, _
            SQLCondition.str����, _
            SQLCondition.str���￨, _
            SQLCondition.str��ʶ��, _
            SQLCondition.lng����ID, _
            SQLCondition.str������, _
            SQLCondition.str�����, _
            SQLCondition.lngҩƷID, _
            SQLCondition.str��ǰNO, _
            lngҩ��ID, _
            SQLCondition.str�����, _
            SQLCondition.str���֤, _
            SQLCondition.strIC��, _
            SQLCondition.strҽ����)
            
   With RecPhysic
        'Add By ZYB 2002-11-27
        'ȡ��ҳ���Ӧ������
        If MnuEditDosage.Checked Then
            strOrder = strOrder_1
        ElseIf MnuEditAbolish.Checked Then
            strOrder = strOrder_2
        ElseIf MnuEditConsignment.Checked Then
            strOrder = strOrder_3
        Else
            strOrder = strOrder_4
        End If
        
        If strOrder <> "" And RecPhysic.RecordCount <> 0 Then
            strOrder = GetOrder(strOrder)
            RecPhysic.Sort = strOrder
        End If
    End With
    
    If err <> 0 Then
        MsgBox "��ȡ" & IIf(IntStyle = 0, "δ��ҩ����", IIf(IntStyle = 1, "����ҩ����", IIf(IntStyle = 2, "δ��ҩ����", "�ѷ�ҩ����"))) & "ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    ReadData = True
End Function

Private Sub Msf�б�_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    Dim lngColor As Long
    Dim bln��ҩ As Boolean
    Dim rsTmp As ADODB.Recordset
            
    picRecipeColor.Visible = False
    lblRecipeType.Visible = False
                
    mnuFileRestore.Enabled = (Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 3)
    
    If RecPhysic.State = 1 Then
        If RecPhysic.RecordCount > 0 Then
            stbThis.Panels(2) = "����" & RecPhysic.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(RecPhysic) & "Ԫ"
        Else
            stbThis.Panels(2) = ""
        End If
    End If
    
    With Msf�б�
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If LngSendRow > 0 And LngSendRow < .Rows Then
            .Row = LngSendRow       '����ϴ�ѡ����
            lngColor = Val(Msf�б�.TextMatrix(LngSendRow, ��������.�ɲ���))
            lngColor = IIf(tabShow.Tab <> 3 Or lngColor = 0, &H80000008, IIf(lngColor = 1, glng����, IIf(lngColor = 2, glng��ҩ, glng��ҩ)))
            For intCol = ��������.ѡ�� To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = lngColor
            Next
            .Col = ��������.ѡ��
        End If
        
        LngSendRow = LngSelectRow
        .Row = LngSendRow       '���õ�ǰѡ����
        lngColor = Val(Msf�б�.TextMatrix(LngSendRow, ��������.�ɲ���))
        lngColor = IIf(tabShow.Tab <> 3 Or lngColor <= 1, glng����, IIf(lngColor = 2, glng��ҩ, glng��ҩ))
        For intCol = ��������.ѡ�� To .Cols - 1
                .Col = intCol
                .CellBackColor = &HC0C0C0
                .CellForeColor = lngColor
        Next
        .Col = ��������.ѡ��
        .Redraw = True:
        
        '��ȡ����
        If .TextMatrix(.Row, ��������.����) = "" Then TxtNo.Clear: Exit Sub
        
        BlnInOper = False
        With TxtNo
            If Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO) & "--" & Msf�б�.TextMatrix(Msf�б�.Row, ��������.����) <> .Text Then
                .Clear
                .AddItem Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO) & "--" & Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)
                .ItemData(.NewIndex) = Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)
                
                BlnAllowClick = False
                .ListIndex = 0
                BlnAllowClick = True
            End If
        End With
        StrLastNo = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
        IntLastBill = Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)
        strLastData = Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)
        mstrNo = Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO)
        IntBillStyle = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����))
        
        mnuCancel.Enabled = False
        Tbar1.Buttons("Cancel").Enabled = False
            
        '��ҩ״̬ʱȡ��ǰ���ݵļ�¼���ʺ������־�����ж��Ƿ����ȡ����ҩ
        If MnuEditHandback.Checked Then
            mint�����־ = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�����־))
            mint��¼���� = Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.��¼����))
            
            '����ȡ����ҩģʽ�Ƿ����
            If mbln����ȡ����ҩ Then
                If (((mint�����־ = 1 Or mint�����־ = 4) And gtype_UserSysParms.P15_�����շ��뷢ҩ���� = 1) Or _
                    (mint�����־ = 2 And gtype_UserSysParms.P16_סԺ�����뷢ҩ���� = 1)) And CheckIsSended(IntLastBill, StrLastNo) = False Then
                    mnuCancel.Enabled = True
                    Tbar1.Buttons("Cancel").Enabled = True
                End If
            End If
        End If
        
        '��鵥���Ƿ����
        If Not CheckBillExist(IntBillStyle, mstrNo) Then
            MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
            Call mnuViewRefresh_Click
            Exit Sub
        End If
        
        '����cmdAlley��ť״̬
        If tabShow.Tab = 2 And mblnStarPass Then
            '�ж���סԺ�������ﲡ�ˣ����û���ҵ���¼����ҽ�����Ͳ���ʾcmdAlley��ť
            gstrSQL = "Select distinct B.����id,nvl(B.��ҳid,0) ��ҳid,nvl(C.�Һŵ�,'') �Һŵ� " & _
                " From ҩƷ�շ���¼ A,���˷��ü�¼ B,����ҽ����¼ C " & _
                " Where A.����id=B.Id And b.ҽ�����=c.Id And nvl(B.ҽ�����,0)<>0 And C.������� IN('5','6','7')" & _
                " And A.����=[2] And A.no=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo, IntBillStyle)
            If rsTmp.RecordCount = 0 Then
                If cmdAlley.Visible Then cmdAlley.Visible = False
            Else
                If Not cmdAlley.Visible Then cmdAlley.Visible = True
            End If
        End If
        
        Call ReadBillData(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����), Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO))
        
        If tabShow.Tab = 0 Then
            bln��ҩ = IsDosage(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)), Msf�б�.TextMatrix(Msf�б�.Row, ��������.NO))
            If bln��ҩ Then
                CmdSend.Caption = "��ҩ(&V)"
            Else
                CmdSend.Caption = "��ҩ(&S)"
            End If
        End If
        
        '���ô�����ϸ���еı�ǩ��ɫ��˵��
        If tabShow.Tab = 3 Then
            picRecipeColor.Visible = False
            lblRecipeType.Visible = False
        Else
            picRecipeColor.Visible = True
            lblRecipeType.Visible = True
            picRecipeColor.BackColor = Val(Split(mstrUserRecipeColor, ";")(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.��������))))
            lblRecipeType.Caption = Split(mconstrRecipeType, ";")(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.��������)))
        End If
    End With

End Sub

Private Sub Msf�б�_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strColumn As String, strOrder As String
    Dim lngRow As Long, lngColor As Long, intCol As Integer
    Dim intMouseRow As Integer, intMouseCol As Integer
    
    'Add by ZYB 2002-11-27
    '���ӵ��������Ĺ���
    If Button <> 1 Then Exit Sub
    intMouseRow = Msf�б�.MouseRow
    intMouseCol = Msf�б�.MouseCol
    If intMouseRow = 0 Then
        'ȡ����
        Select Case intMouseCol
        Case ��������.ѡ��
            'ȫ���򹴻���ȫ������
            Call SetCheckBox(0)
            Exit Sub
        Case ��������.����
            strColumn = "����"
        Case ��������.NO
            strColumn = "NO"
        Case ��������.����
            strColumn = "����ID,����"
        Case ��������.���
            strColumn = "���"
        Case ��������.����
            strColumn = "����"
        Case ��������.˵��
            strColumn = "˵��"
        Case Else
            Exit Sub
        End Select
        
        'ȡ����
        If MnuEditDosage.Checked Then
            strOrder = strOrder_1
        ElseIf MnuEditAbolish.Checked Then
            strOrder = strOrder_2
        ElseIf MnuEditConsignment.Checked Then
            strOrder = strOrder_3
        Else
            strOrder = strOrder_4
        End If
        
        '���������ͬ����ı�����ʽ����������ʽ
        If strOrder Like "*" & strColumn & "*" Then
            strOrder = ExchangeOrder(strOrder)
        Else
            strOrder = strColumn & strAsc
        End If
        
        '��ȫ�ֱ�����ֵ
        If MnuEditDosage.Checked Then
            strOrder_1 = strOrder
        ElseIf MnuEditAbolish.Checked Then
            strOrder_2 = strOrder
        ElseIf MnuEditConsignment.Checked Then
            strOrder_3 = strOrder
        Else
            strOrder_4 = strOrder
        End If
        strOrder = GetOrder(strOrder)
        
        '�Լ�¼�������������°�
        If RecPhysic.RecordCount = 0 Then Exit Sub
        RecPhysic.Sort = strOrder
        With Msf�б�
            .Redraw = False
            If Not RecPhysic.EOF Then
                Set .DataSource = RecPhysic
                stbThis.Panels(2) = "����" & RecPhysic.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(RecPhysic) & "Ԫ"
            End If
            DoEvents
            Call SetFormat(1, RecPhysic.EOF)
            Call SetCheckBox(-1)
            DoEvents
        
            '��ɫ
            If MnuEditHandback.Checked Then
                .Redraw = False
                For lngRow = 1 To .Rows - 1
                    .Row = lngRow
                    lngColor = IIf(Val(.TextMatrix(lngRow, ��������.�ɲ���)) = 1, glng����, IIf(Val(.TextMatrix(lngRow, ��������.�ɲ���)) = 2, glng��ҩ, glng��ҩ))
                    For intCol = ��������.ѡ�� To .Cols - 1
                        .Col = intCol
                        .CellForeColor = lngColor
                    Next
                Next
                .Redraw = True
            End If
            .Redraw = True
            .Row = 1
        End With
        '���ô�����ɫ
        Call SetRecipeColor
        Call Msf�б�_EnterCell
    Else
        '���������ʱ
        If intMouseCol = ��������.ѡ�� Then
            Call SetCheckBox(intMouseRow)
            Exit Sub
        End If
    End If
End Sub

Private Sub PicCloseConsignment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    RaisEffect PicCloseConsignment, -2     '�°�
End Sub

Private Sub PicCloseConsignment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    RaisEffect PicCloseConsignment, 2      '��͹
    
    '��������ڱ��ؼ��ϣ����˳�
    If x < 0 Or x > PicCloseConsignment.Width Then Exit Sub
    If y < 0 Or y > PicCloseConsignment.Height Then Exit Sub
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    'Modified by ZYB 2002-11-19 �����û�����
    Call SaveFlexState(Bill������ϸ.MsfObj, Me.Name & "\" & PreviousTab)
    Call SaveFlexState(Msf�б�, Me.Name & "\" & PreviousTab)
    '�ָ�����
    Call RestoreFlexState(Msf�б�, Me.Name & "\" & tabShow.Tab)
    Call RestoreFlexState(Bill������ϸ.MsfObj, Me.Name & "\" & tabShow.Tab)
    Bill������ϸ.ColWidth(����.�����) = IIf(Not mblnStarPass, 0, 240)
    
    txtFind.Text = ""
    Call SetMenuState
    
    BlnInOper = False
    Call mnuViewRefresh_Click
End Sub
Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreView_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Find"
        MnuViewFind_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnufileexit_Click
    Case "Charge"
        mnuCharge_Click
    Case "Stuff"
        mnuStuff_Click
    Case "Cancel"
        mnuCancel_Click
    End Select
End Sub

Private Sub Tbar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu MnuViewTool, 2
End Sub

Private Sub TimePrint_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    If tabShow.Tab = 3 Then
        If Chkȫ��.Value = 0 Or mblnAllBack = False Then Exit Sub
    End If
    
    TimePrint.Enabled = False
    DoEvents
    '���ô�ӡ����
    Call AutoPrint
    DoEvents
    TimePrint.Enabled = True
    
    If mint�Զ���ҩ = 1 Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub txtFind_Change()
    If txtFind.Text = "" Then txtFind.Tag = ""
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtFind.Text = "" And Me.ActiveControl Is txtFind)
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Tag = "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
    txtFind.Tag = ""
    
    If Not mobjIDCard Is Nothing And txtFind.Text = "" Then
        mobjIDCard.SetEnabled (True)
    End If
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    mblnCard = False
    If imgFilter.BorderStyle = cstLocate Then
        If KeyAscii = 13 Then
             Call Form_KeyDown(vbKeyF3, 0)
             Exit Sub
        End If
             
        If lblFind.Tag = FindType.���� Then
            mblnCard = zlCommFun.InputIsCard(txtFind, KeyAscii, glngSys)
        ElseIf lblFind.Tag = FindType.���￨ Then
            mblnCard = (KeyAscii <> 8 And Len(txtFind.Text) = gtype_UserSysParms.P20_���￨�ų��� - 1 And txtFind.SelLength <> Len(txtFind.Text))
        End If
        
        If mblnCard Or KeyAscii = 13 Then
            If KeyAscii <> 13 Then
                txtFind.Text = txtFind.Text & Chr(KeyAscii)
                txtFind.SelStart = Len(txtFind.Text)
            End If
            KeyAscii = 0
            Call Form_KeyDown(vbKeyF3, 0)
        Else
            Select Case lblFind.Tag
                Case FindType.���￨
                    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                Case FindType.�����
                    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
                Case FindType.���ݺ�
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    If Not (txtFind.Text = "" Or txtFind.SelLength = Len(txtFind.Text)) _
                        And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
                Case FindType.����
                    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                Case FindType.���֤
                Case FindType.IC��
            End Select
        End If
    Else
        If KeyAscii = 13 Then
            Call SetFilter(MnuEditHandback.Checked)
            Call zlControl.TxtSelAll(txtFind)
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFind_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    
    txtFind.MaxLength = 0
    lblFind.Tag = FindType.���֤
    lblFind.Caption = "���֤��"
    txtFind.Text = strID
    Call txtFind_KeyPress(vbKeyReturn)

    DoEvents

    txtFind.Text = ""

End Sub
Private Sub txtFind_Validate(Cancel As Boolean)
    If Val(lblFind.Tag) = FindType.���ݺ� Then
        If IsNumeric(txtFind.Text) Then
            txtFind.Text = GetFullNO(txtFind.Text, 13)
        End If
    End If
End Sub
Private Sub TxtNo_Click()
    Dim LngLocate As Long, blnFind As Boolean
    
    On Error GoTo ErrHand
    If BlnAllowClick = False Then Exit Sub
    If TxtNo.ListIndex = -1 Then
        Exit Sub
    End If
    '--Ϊ��ʾ������׼��--
    ClearCons
    
    '--��ȡ���ݲ���ʾ--
    blnFind = False
    StrLastNo = Mid(TxtNo.Text, 1, 8)
    IntLastBill = TxtNo.ItemData(TxtNo.ListIndex)
    TxtNo.Tag = TxtNo.Text
    '��λ���
    With Msf�б�
        For LngLocate = 1 To .Rows - 1
            If Trim(.TextMatrix(LngLocate, ��������.����)) <> "" Then
                If .TextMatrix(LngLocate, ��������.����) = TxtNo.ItemData(TxtNo.ListIndex) And .TextMatrix(LngLocate, ��������.NO) = Mid(TxtNo.Text, 1, 8) Then
                    .Row = LngLocate
'                    StrLastNo = ""
                    Msf�б�_EnterCell
                    blnFind = True
                    Exit For
                End If
            End If
        Next
    End With
    If Not blnFind Then If Not ReadBillData(TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8), blnFind) Then Exit Sub
    BlnInOper = False
    
    If CmdSend.Enabled Then Me.CmdSend.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub TxtNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intYear  As Integer, strYear As String, strCond As String
    Dim bln������ As Boolean            '�ٱ����ǲ��˱�ʶ��
    Dim RecRecord As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(TxtNo) = "" Then Exit Sub
    '--���������λ,�򰴹������--
    '--��A��ͷ��ʾ���˵ı�ʶ�ţ����������Ϊ��NO��
    Me.TxtNo = UCase(LTrim(Me.TxtNo))
    bln������ = Not ((Mid(TxtNo, 1, 1) = "B") Or (Mid(TxtNo, 1, 1) = "+"))
    If Not bln������ Then
        If MnuEditDosage.Checked Then
'            StrFind_1 = GetDateSQL(StrFind_1) & _
'            " And Upper(DECODE(A.����,8,A.�����,A.סԺ��)) Like '" & Mid(txtNo.Text, 2) & "%'"
            StrFind_1 = GetDateSQL(StrFind_1) & _
            " And Upper(DECODE(A.����,8,A.�����,A.סԺ��)) Like [12] "
            SQLCondition.str��ǰNO = Mid(TxtNo.Text, 2) & "%"
        ElseIf MnuEditAbolish.Checked Then
            StrFind_2 = GetDateSQL(StrFind_2) & _
            " And Upper(DECODE(A.����,8,A.�����,A.סԺ��)) Like [12] "
            SQLCondition.str��ǰNO = Mid(TxtNo.Text, 2) & "%"
        ElseIf MnuEditConsignment.Checked Then
            StrFind_3 = GetDateSQL(StrFind_3) & _
            " And Upper(DECODE(A.����,8,A.�����,A.סԺ��)) Like [12] "
            SQLCondition.str��ǰNO = Mid(TxtNo.Text, 2) & "%"
        Else
            StrFind_4 = GetDateSQL(StrFind_4) & _
            " And Upper(H.��ʶ��) Like [12] "
            SQLCondition.str��ǰNO = Mid(TxtNo.Text, 2) & "%"
        End If
        Call DataRefresh
        Exit Sub
    End If
    
    TxtNo.Text = GetFullNO(TxtNo.Text, 13)
    
    If mInt���� = 0 Then
        strCond = " And ���� In (8,9)" '���ＰסԺ���е���
    Else
        If mInt���� = 8 Then
            strCond = " And ���� In (8,9) And ��ҳID Is NULL " '���ﻮ�ۼ��������
        Else
            strCond = " And ���� = 9 And ��ҳID Is Not NULL " 'סԺ����
        End If
    End If

    '--�����������¼,��������û�ѡ��(����������ϴ�NO����������ȡ)--
    With RecRecord
        If .State = 1 Then .Close
        If MnuEditHandback.Checked = False Then
            gstrSQL = "Select A.No,A.����,A.���� " & _
                " From δ��ҩƷ��¼ A" & _
                " Where (Nvl(A.�ⷿID,0)=0 Or A.�ⷿID+0=[13] )" & strCond & _
                " And A.No =[12] "
            SQLCondition.str��ǰNO = Mid(TxtNo, 1, 8)
        Else
            strCond = Replace(strCond, "����", "A.����")
            strCond = Replace(strCond, "��ҳID", "H.��ҳID")
            
            Dim strCond2 As String
            strCond2 = ת����ҩ��
            gstrSQL = " Select Distinct A.No,A.����,H.���� " & _
                     " From " & _
                     "     (SELECT A.ID,A.No,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                     "          DECODE(SIGN((A.ʵ������*NVL(A.����,1))-B.�ѷ�����),0,A.����,1) ����," & _
                     "          DECODE(SIGN((A.ʵ������*NVL(A.����,1))-B.�ѷ�����),0,A.ʵ������,B.�ѷ�����) ʵ������,A.��¼״̬," & _
                     "          A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.������,A.��������,A.��ҩ��,A.�Է�����ID,A.�ⷿID" & _
                     "      From" & _
                     "          (SELECT *" & _
                     "          From ҩƷ�շ���¼ A" & _
                     "          WHERE A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                     "          And A.�ⷿID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " And A.������� " & StrDate & "", strCond2) & _
                     "          ) A," & _
                     "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                     "          From ҩƷ�շ���¼ A" & _
                     "          Where A.����� Is Not Null" & _
                     "          And A.�ⷿID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " And A.������� " & StrDate & "", strCond2) & _
                     "          GROUP BY A.no,A.����,A.ҩƷID,A.���) B" & _
                     "      Where A.no = B.no And A.���� = B.���� And A.ҩƷID+0 = B.ҩƷID And A.��� = B.���" & _
                     "     ) A,���˷��ü�¼ H" & _
                     " Where A.�ⷿID+0=[13] " & strCond & _
                     " And A.No ='" & Mid(TxtNo, 1, 8) & "'" & _
                     " And A.����ID=H.ID And (Mod(A.��¼״̬,3)=0 Or A.��¼״̬=1) And A.ʵ������<>0 "
        
            'һ�Ŵ���������ͬʱ������������󱸱��У���ˣ���������Ƴ�����ֱ�ӴӺ󱸱�����ȡ������ԭSQL����
            'ҩƷ������ҩ��ͬʱ�Ե��� IN (8,9)�ĵ��ݣ���˲��ų�����8���߶�9���е����
            Dim blnMoved As Boolean
            Dim strSQL As String
            
            blnMoved = zlDatabase.NOMoved("ҩƷ�շ���¼", Mid(TxtNo, 1, 8), " ���� IN ", " (8,9)")
            
            '�����������ת��������Ҫͬʱ�Ӻ󱸱�����ȡ���ݣ����ܴ��ڲ�ͬ���͵ĵ��ݷֱ�������󱸱��У�
            If blnMoved Then
                strSQL = gstrSQL
                strSQL = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
                strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
                gstrSQL = gstrSQL & " UNION ALL " & strSQL
            End If
        End If
    End With

   Set RecRecord = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.date��ʼ����, _
            SQLCondition.date��������, _
            SQLCondition.str��ʼNO, _
            SQLCondition.str����NO, _
            SQLCondition.str����, _
            SQLCondition.str���￨, _
            SQLCondition.str��ʶ��, _
            SQLCondition.lng����ID, _
            SQLCondition.str������, _
            SQLCondition.str�����, _
            SQLCondition.lngҩƷID, _
            SQLCondition.str��ǰNO, _
            lngҩ��ID)

    With RecRecord
        TxtNo.Clear
        If .EOF Then
            MsgBox "δ�ҵ�ָ�����������������룡", vbInformation, gstrSysName
            Msf�б�_EnterCell
            Exit Sub
        End If
        Do While Not .EOF
            TxtNo.AddItem !NO & "--" & !����
            TxtNo.ItemData(TxtNo.NewIndex) = !����
            .MoveNext
        Loop
        
        If TxtNo.ListCount = 0 Then Exit Sub
        
        If MnuEditDosage.Checked Then
            '��ҩ
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "��ҩ") And mint�Զ���ҩ = 0
        ElseIf MnuEditAbolish.Checked Then
            'ȡ��
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")
        ElseIf MnuEditConsignment.Checked Then
            '��ҩ
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")
        Else
            CmdSend.Enabled = IsHavePrivs(mstrPrivs, "��ҩ")
        End If
        
        TxtNo.ListIndex = 0
        StrLastNo = Mid(TxtNo, 1, 8)
        IntLastBill = TxtNo.ItemData(TxtNo.ListIndex)
        
        If .RecordCount > 1 Then
            MsgBox "���ֶ�����ͬ���ŵĴ������ݣ���ѡ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
End Sub

Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, Optional ByVal blnExist As Boolean = True) As Boolean
    Dim IntStyle As Integer, intUnit As Integer
    Dim strSubSql As String
    Dim strName As String
    Dim blnMoved As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    Dim RecBill As New ADODB.Recordset
    '--��ȡ��������--
    'BillStyle-��������;BIllNO-���ݺ�
    '��λ��ʾ���ݷ����������������ﵥλ��סԺ��סԺ���סԺ��λ���������ۼ۵�λ��
    On Error Resume Next
    err = 0
    ReadBillData = False
  
    strUnit = GetUnit(lngҩ��ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "1"
    Case "���ﵥλ"
        strSubSql = "Decode(�����װ,Null,1,0,1,�����װ)"
    Case "סԺ��λ"
        strSubSql = "Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
    End Select
    Call Get��λ��
    
    '�õ�ҩƷ���ƴ�
    Select Case intҩƷ����
    Case 0  'ҩƷ����������
        strName = "'['||C.����||']'||" & IIf(mblnTradeName, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    Case 1  'ҩƷ����
        strName = "C.���� As Ʒ��,"
    Case 2  'ҩƷ����
        strName = IIf(mblnTradeName, "NVL(E.����,C.����)", "C.����") & " As Ʒ��,"
    End Select
    
    strName = strName & IIf(Not mblnTradeName, "NVL(E.����,'')", "Decode(E.����,Null,'',C.����)") & " As ������, "
    
    If MnuEditHandback.Checked = False Then
        'Modified By ���� 2003-12-10 ���������� ���ӿ����
        gstrSQL = " SELECT DISTINCT B.NO,H.���,T.���� ����,H.����,H.�Ա�,H.����,H.��ʶ�� סԺ��,H.����,H.������,B.ID," & _
            " B.ҩƷID,DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����," & _
            " NVL(B.����,0) ����,NVL(D.ҩ������,0) ����," & strName & _
            " DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & str��λ�� & ",K.ʵ������/" & strSubSql & " �����," & _
            " NVL(B.����,1) ����,B.����,B.�÷�,B.Ƶ��,B.������,B.��������,H.����Ա����," & IIf(MnuEditHandback.Checked = False, "B.��ҩ��", "B.�����") & " ��ҩ��,L.�ⷿ��λ,M.ҽ������,M.id ҽ��id,nvl(M.�����,-1) �����,I.���㵥λ,round(B.���۽��," & mintMoneyDigit & ") ���۽��,H.�ѱ�,P.�������, " & _
            " B.ʵ������*D.����ϵ��* Nvl(B.����, 1) ����,Decode(Sign(Nvl(J.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ���� " & _
            " FROM ҩƷ�շ���¼ B,ҩƷ��� D,ҩƷ���� P,�շ���ĿĿ¼ C,�շ���Ŀ���� E," & _
            " ���˷��ü�¼ H,����ҽ����¼ M,���ű� S,���ű� T,ҩƷ��� K,ҩƷ�����޶� L,������ĿĿ¼ I,������Ŀ���� Z ," & _
            " (Select �ⷿid, ҩƷid, Nvl(Sum(ʵ������), 0) ������� From ҩƷ��� Where ���� = 1 And �ⷿid = [13] Group By �ⷿid, ҩƷid) J" & _
            " WHERE D.ҩƷID=C.ID And D.ҩ��ID=P.ҩ��ID And H.ҽ�����=M.ID(+) AND C.ID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
            " And B.ҩƷID=L.ҩƷID(+) And Nvl(B.�ⷿID,[13])=L.�ⷿID(+)" & _
            " AND H.��������ID=T.ID(+) AND B.ҩƷID=D.ҩƷID AND MOD(B.��¼״̬,3)=1" & _
            " AND S.ID=NVL(B.�ⷿID,[13]) AND B.����ID=H.ID AND B.NO=[14] AND B.����=[15] AND NVL(B.�ⷿID,[13])+0=[13] AND LTRIM(RTRIM(NVL(B.ժҪ,'С��')))<>'�ܷ�'" & _
            " AND B.ҩƷID=K.ҩƷID(+) AND K.����(+)=1 AND NVL(B.�ⷿID,[13])=K.�ⷿID(+) AND NVL(B.����,0)=NVL(K.����(+),0) AND B.����� IS NULL And D.ҩ��id=I.id " & _
            " And Nvl(B.�ⷿid, [13]) + 0 = J.�ⷿid(+) And B.ҩƷid = J.ҩƷid(+) And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 "
     Else
        '������ʾ��������
        '�����ܴ���һ�Ŵ���ͬʱ������󱸱��ж�����
        blnMoved = zlDatabase.NOMoved("ҩƷ�շ���¼", BillNo, " ���� = ", BillStyle)
        gstrSQL = " SELECT DISTINCT B.NO,H.���,T.���� ����,H.����,H.�Ա�,H.����,H.��ʶ�� סԺ��,H.����,H.������,B.ID,B.ҩƷID," & _
                 " DECODE(B.����,NULL,'',B.����)||DECODE(B.����,NULL,'',0,'','('||B.����||')') ����," & _
                 " NVL(B.����,0) ����,NVL(D.ҩ������,0) ����," & strName & _
                 " DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & str��λ�� & "," & _
                 " NVL(B.����,1) ����," & _
                 " B.��������/" & strSubSql & " ��������," & _
                 " B.�ѷ�����/" & strSubSql & " ׼����,B.�ѷ����� ʵ������," & _
                 " B.����,B.�÷�,B.Ƶ��,B.������,B.��������,H.����Ա����," & IIf(MnuEditHandback.Checked = False, "B.��ҩ��", "B.�����") & " ��ҩ��,I.���㵥λ,round(B.���۽��," & mintMoneyDigit & " ) ���۽��,H.�ѱ�,P.�������, "
        If Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist Then    '�����������ǽ�ȥ
            Dim strCond2 As String
            strCond2 = ת����ҩ��
            gstrSQL = gstrSQL & " B.�ѷ�����*D.����ϵ�� ����,Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ���� FROM "
            gstrSQL = gstrSQL & "   (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��," & _
                     "          NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
                     "          A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.������,A.��������,A.�����,A.�������,A.�Է�����ID,A.�ⷿID" & _
                     "      FROM" & _
                     "          (SELECT *" & _
                     "          FROM ҩƷ�շ���¼ A" & _
                     "          WHERE A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                     "          AND A.�ⷿID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " AND A.������� " & StrDate & "", strCond2) & _
                     "          ) A," & _
                     "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                     "          FROM ҩƷ�շ���¼ A" & _
                     "          WHERE A.����� IS NOT NULL" & _
                     "          AND A.�ⷿID+0=[13] " & _
                     "      " & IIf(StrFind_4 = "", " AND A.������� " & StrDate & "", strCond2) & _
                     "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
                     "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� " & _
                     "      )"
        Else
            gstrSQL = gstrSQL & " B.ʵ������*D.����ϵ�� ����,Decode(Sign(Nvl(K.�������, 0) - Nvl(L.����, 0)), -1, 0, 1) �������,Z.���� As Ӣ���� FROM "
            gstrSQL = gstrSQL & "(Select 0 �ѷ�����,0 ��������,0 ׼������,A.* From ҩƷ�շ���¼ A)"
        End If
        gstrSQL = gstrSQL & _
                 "       B,ҩƷ��� D,ҩƷ���� P,�շ���ĿĿ¼ C,�շ���Ŀ���� E,���˷��ü�¼ H,���ű� S,���ű� T,������ĿĿ¼ I,������Ŀ���� Z , " & _
                 "(Select �ⷿid, ҩƷid, Nvl(Sum(ʵ������), 0) ������� From ҩƷ��� Where ���� = 1 And �ⷿid = [13] Group By �ⷿid, ҩƷid) K, ҩƷ�����޶� L " & _
                 " Where H.��������ID=T.ID(+) And B.ҩƷID=D.ҩƷID And D.ҩ��ID=P.ҩ��ID And C.ID=D.ҩƷID " & _
                 " And D.ҩƷID=E.�շ�ϸĿID(+) and E.����(+)=3 And D.ҩ��id = Z.������Ŀid(+) And Z.����(+) = 2 " & _
                 " And S.ID=B.�ⷿID And B.����ID=H.ID And B.NO=[14] And B.����=[15] And B.�ⷿID+0=[13]"
                 
        If IsDate(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) Then
                 gstrSQL = gstrSQL & " And B.�������=To_Date('" & Msf�б�.TextMatrix(Msf�б�.Row, ��������.����) & "','yyyy-MM-dd hh24:mi:ss')"
        End If
        gstrSQL = gstrSQL & " And B.����� Is Not Null And D.ҩ��id=I.id " & _
                            " And B.ҩƷid = L.ҩƷid(+) And Nvl(B.�ⷿid, 24) = L.�ⷿid(+) And" & _
                            " D.ҩ��id = I.ID And Nvl(B.�ⷿid, 24) + 0 = K.�ⷿid(+) And B.ҩƷid = K.ҩƷid(+) "
        
        '�������ת������ֱ�ӴӺ󱸱�����ȡ����
        If blnMoved Then
            gstrSQL = Replace(gstrSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
            gstrSQL = Replace(gstrSQL, "���˷��ü�¼", "H���˷��ü�¼")
        End If
    End If
    gstrSQL = gstrSQL & " Order by H.���,B.ҩƷID,Nvl(B.����,0)"
     
    Set RecBill = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.date��ʼ����, _
            SQLCondition.date��������, _
            SQLCondition.str��ʼNO, _
            SQLCondition.str����NO, _
            SQLCondition.str����, _
            SQLCondition.str���￨, _
            SQLCondition.str��ʶ��, _
            SQLCondition.lng����ID, _
            SQLCondition.str������, _
            SQLCondition.str�����, _
            SQLCondition.lngҩƷID, _
            SQLCondition.str��ǰNO, _
            lngҩ��ID, BillNo, BillStyle)
    
    If WriteDataToBill(RecBill, blnExist) = False Then Exit Function
    
    '������ҩ������һЩ���� by lyq 2005-04-27
    If �ж��Ƿ���ҩ����(BillStyle, BillNo) Then
        Call ��ҩ�����ر���(BillStyle, BillNo)
    End If
    
    IntStyle = IIf(MnuEditDosage.Checked, 1, IIf(MnuEditAbolish.Checked, 2, IIf(MnuEditConsignment.Checked, 3, 4)))
    'ֻ�ж��������ݣ�δ��ҩ��δ��ҩ�����ݣ������Ƴ������ǲ���������ģ����ھ������������Ҳ�д��жϣ��ж�Ҳ��û��������
    If Not blnMoved Then
        If CheckBill(IntStyle, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8)) <> 0 Then Exit Function
    End If
    
    '���ð�ť״̬
    Select Case tabShow.Tab
    Case 0
        CmdSend.Enabled = (Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist) And MnuEditDosage.Visible And mint�Զ���ҩ = 0
    Case 1
        CmdSend.Enabled = (Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist) And MnuEditAbolish.Visible
    Case 2
        CmdSend.Enabled = (Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist) And IsHavePrivs(mstrPrivs, "��ҩ")
    Case 3
        CmdSend.Enabled = (Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist) And IsHavePrivs(mstrPrivs, "��ҩ")
        Chkȫ��.Enabled = CmdSend.Enabled
        Chkȫ��.Value = IIf(CmdSend.Enabled, 1, 0)
        mblnAllBack = (Chkȫ��.Value = 1)
        If blnMoved Then Bill������ϸ.Active = False
    End Select

    If err <> 0 Then
        MsgBox "��ȡ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        CmdSend.Enabled = False
        Chkȫ��.Enabled = False
        Exit Function
    End If
    ReadBillData = True
End Function

Private Function WriteDataToBill(ByVal RecData As ADODB.Recordset, Optional ByVal blnExist As Boolean = True) As Boolean
    Dim dblMoney As Currency, IntLocate As Integer
    Dim str����Ա As String, str�ϼ� As String
    Dim dbl������ As Double
    Dim str������λ As String
    Dim lng������ As Long
    Dim dblС���� As Double
    
    '--�򵥾ݿؼ���д����--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    Call ClearCons
    
    mblnAuto = True
    dblMoney = 0
    Lbl��ҩ��.Caption = IIf(tabShow.Tab = 3, IIf(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) <> 3, "��ҩ��", "��ҩ��"), "��ҩ��")
    Lbl�շ�Ա.Caption = IIf(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = 8, "�շ�Ա", "����Ա")
    
    '�������ԭʼ��¼������ʾ�У���������׼������
    Bill������ϸ.ColWidth(����.������) = 0
    Bill������ϸ.ColWidth(����.׼����) = 0
    Bill������ϸ.ColWidth(����.��ҩ��) = 0
    If tabShow.Tab = 3 Then
        Bill������ϸ.ColWidth(����.������) = IIf(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist, 1000, 0)
        Bill������ϸ.ColWidth(����.׼����) = IIf(Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist, 1000, 0)
        Bill������ϸ.ColWidth(����.��ҩ��) = IIf((Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.�ɲ���)) = 1 Or Not blnExist) And mbln��ʾ��С��λ = False, 1000, 0)
    End If
    
    '��䵥������
    With RecData
        '����ͷ
        If Not .EOF Then
            Me.Txt���� = IIf(IsNull(!����), "", !����)
            If Val(Msf�б�.TextMatrix(Msf�б�.Row, ��������.����)) = 8 Then Me.Txt���� = ""
            Me.Txt����ҽ��.ListIndex = 0
            If (intVerify = 0) And IsHavePrivs(mstrPrivs, "ҽ����ѯ") Then
                str����Ա = IIf(IsNull(!������), "", !������)
            Else
                If MnuEditHandback.Checked And intVerify = 1 Then
                    str����Ա = IIf(IsNull(!������), "", !������)
                Else
                    str����Ա = ""
                End If
            End If
            If str����Ա <> "" Then
                '��λҽ��
                For IntLocate = 1 To Txt����ҽ��.ListCount
                    If Mid(Txt����ҽ��.List(IntLocate), InStr(1, Txt����ҽ��.List(IntLocate), "-") + 1) = str����Ա Then
                        Txt����ҽ��.ListIndex = IntLocate
                        Exit For
                    End If
                Next
            End If
            If glngSys \ 100 = 1 Then
                If IsHavePrivs(mstrPrivs, "ҽ����ѯ") Then
                    Me.Txt���� = IIf(IsNull(!����), "", !����)
                End If
            Else
                Me.Txt���� = IIf(IsNull(!����), "", !����)
            End If
            Me.Txt���� = IIf(IsNull(!����), "", !����)
            If IIf(IsNull(!��ҩ��), "", !��ҩ��) <> "" Then
                Me.cbo��ҩ�� = IIf(IsNull(!��ҩ��), "", !��ҩ��)
            End If
            Me.Txt�շ�Ա = IIf(IsNull(!����Ա����), "", !����Ա����)
            Me.Txt�Ա� = IIf(IsNull(!�Ա�), "", !�Ա�)
            Me.TxtסԺ�� = IIf(IsNull(!סԺ��), "", !סԺ��)
        End If
            
        Bill������ϸ.Rows = 1
        Bill������ϸ.Rows = 2
        Bill������ϸ.MsfObj.FixedRows = 1
        Bill������ϸ.MsfObj.Redraw = False
        
        Do While Not .EOF
            Bill������ϸ.MergeRow .AbsolutePosition, False
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.˳���) = .AbsolutePosition
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.ҩƷ����) = !Ʒ��
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.������) = IIf(IsNull(!������), "", !������)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.Ӣ����) = IIf(IsNull(!Ӣ����), "", !Ӣ����)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.���) = !���
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.���) = IIf(IsNull(!���), "", !���)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = IIf(IsNull(!����), "", !����)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.Id) = !Id
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.ҩƷID) = !ҩƷID
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = !����
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��λ) = IIf(IsNull(!��λ), "", !��λ)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = GetFormat(!����, mintPriceDigit)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = Format(!����, "#####0;-#####0; ;")
            
            If mbln��ʾ��С��λ = True Then
                '����С��װ��ʾ����
                lng������ = Int(!����)
                If !�ۼ۵�λ = !��λ Or lng������ = !���� Then
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                Else
                    dblС���� = (Val(!����) - lng������) * !��װ
                    If lng������ = 0 Then
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                    Else
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = lng������ & IIf(IsNull(!��λ), "", !��λ) & dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                    End If
                End If
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��װ) = Val(!��װ)
            Else
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = FormatEx(!����, mintNumberDigit)
            End If
            
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.���) = GetFormat(Val(!���۽��), mintMoneyDigit)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = !���� & !���㵥λ
            
            dbl������ = dbl������ + !����
            str������λ = !���㵥λ
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.Ƶ��) = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            mstr������λ = NVL(!���㵥λ)
            If Not IsNull(!����) Then
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = FormatEx(!����, 5) & NVL(!���㵥λ)
            End If
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�÷�) = NVL(!�÷�)
            If MnuEditHandback.Checked Then
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��װ) = Val(!��װ)
                If mbln��ʾ��С��λ = True Then
                    '����С��װ��ʾ�������ֱ�������������׼����������ҩ����
                    '����������׼����������ʾģʽΪ"���װ����+���װ��λ+С��װ����+�ۼ۵�λ"����ҩ����������ʾ����ֻ��ʾ��ֵ
                    lng������ = Int(!��������)
                    If !�ۼ۵�λ = !��λ Or lng������ = !�������� Then
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.������) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                    Else
                        dblС���� = (Val(!��������) - lng������) * !��װ
                        If lng������ = 0 Then
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.������) = dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        Else
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.������) = lng������ & IIf(IsNull(!��λ), "", !��λ) & dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        End If
                    End If
                    
                    lng������ = Int(!׼����)
                    If !�ۼ۵�λ = !��λ Or lng������ = !׼���� Then
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                    Else
                        dblС���� = (Val(!׼����) - lng������) * !��װ
                        If lng������ = 0 Then
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����) = dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        Else
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����) = lng������ & IIf(IsNull(!��λ), "", !��λ) & dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        End If
                    End If
                    
                    lng������ = Int(!׼����)
                    If !�ۼ۵�λ = !��λ Then
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����С) = FormatEx(lng������, mintNumberDigit)
                    ElseIf lng������ = !׼���� Then
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼������) = FormatEx(lng������, mintNumberDigit)
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����С) = FormatEx(0, mintNumberDigit)
                    Else
                        dblС���� = (Val(!׼����) - lng������) * !��װ
                        If lng������ = 0 Then
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����С) = FormatEx(dblС����, mintNumberDigit)
                        Else
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼������) = FormatEx(lng������, mintNumberDigit)
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����С) = FormatEx(dblС����, mintNumberDigit)
                        End If
                    End If
                    
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��ҩ��) = FormatEx(!׼����, mintNumberDigit)
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��ҩ����) = Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼������)
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��ҩ��С) = Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����С)
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��λ��) = IIf(IsNull(!��λ), "", !��λ)
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��λС) = IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                Else
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.������) = FormatEx(!��������, mintNumberDigit)
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.׼����) = FormatEx(!׼����, mintNumberDigit)
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��ҩ��) = FormatEx(!׼����, mintNumberDigit)
                End If
            
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.ʵ������) = !ʵ������
            Else
                If mbln��ʾ��С��λ = True Then
                    '����С��װ��ʾ����
                    lng������ = Int(!�����)
                    If !�ۼ۵�λ = !��λ Or lng������ = !����� Then
                        Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�����) = lng������ & IIf(IsNull(!��λ), "", !��λ)
                    Else
                        dblС���� = (Val(!�����) - lng������) * !��װ
                        If lng������ = 0 Then
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�����) = dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        Else
                            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�����) = lng������ & IIf(IsNull(!��λ), "", !��λ) & dblС���� & IIf(IsNull(!�ۼ۵�λ), "", !�ۼ۵�λ)
                        End If
                    End If
                Else
                    Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�����) = FormatEx(NVL(!�����, 0), mintNumberDigit)
                End If
            
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��λ) = NVL(!�ⷿ��λ)
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.ҽ������) = NVL(!ҽ������)
                Bill������ϸ.TextMatrix(.AbsolutePosition, ����.ҽ��id) = NVL(!ҽ��id)
                If !����� <> -1 Then
                    BlnEnterCell = False
                    Bill������ϸ.Row = .AbsolutePosition
                    Bill������ϸ.Col = 0
                    Set Bill������ϸ.MsfObj.CellPicture = imgPass.ListImages(Val(!�����) + 1).Picture
                    Bill������ϸ.MsfObj.CellPictureAlignment = 4
'                    Bill������ϸ.CellBackColor = &H8000000F
                    BlnEnterCell = True
                End If
            End If
            
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.����) = IIf(IsNull(!����), 0, !����)
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.������) = ""
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��Ч��) = ""
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�²���) = ""
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.��ע) = ""
            Bill������ϸ.TextMatrix(.AbsolutePosition, ����.�ѱ�) = IIf(IsNull(!�ѱ�), "", !�ѱ�)
            If MnuEditHandback.Checked Then
                dblMoney = dblMoney + IIf(Chk�嵥.Value = 1, Val(!���۽��), FormatEx(!׼���� / (!���� * !����) * Val(!���۽��), mintMoneyDigit))
            Else
                dblMoney = dblMoney + Val(!���۽��)
            End If
            
            '�Ե��ڿ�����޵�ҩƷ��ɫ
            Bill������ϸ.MsfObj.Redraw = False
            If !������� = 0 Then
'            If IsLowerLimit(lngҩ��ID, !ҩƷID) Then
                Call SetForeColor_ROW(.AbsolutePosition, mlng��ɫ)
            Else
                Call SetForeColor_ROW(.AbsolutePosition, vbBlack)
            End If
                        
            '����ҩƷ������ʾ
            If InStr(";����ҩ;����ҩ;����I��;����II��;", NVL(!�������)) > 0 And NVL(!�������) <> "" Then
                Bill������ϸ.Col = ����.ҩƷ����
                Bill������ϸ.Row = .AbsolutePosition
                Bill������ϸ.MsfObj.CellFontBold = True
            End If
                        
            If .AbsolutePosition >= Bill������ϸ.Rows - 1 Then Bill������ϸ.Rows = Bill������ϸ.Rows + 1
            .MoveNext
        Loop
        Bill������ϸ.MsfObj.Redraw = True
        'ȡ�����հ���
        '--If Bill������ϸ.Rows - 1 >= 2 Then Bill������ϸ.Rows = Bill������ϸ.Rows - 1
    End With
    
    '���հ�����ʾ���ϼ�
    str�ϼ� = zlCommFun.UppeMoney(dblMoney)
    With Bill������ϸ
        .TextMatrix(.Rows - 1, 1) = "���ϼƣ�" & Format(dblMoney, mstrVBMoneyForamt)
        .TextMatrix(.Rows - 1, 2) = "���ϼƣ�" & Format(dblMoney, mstrVBMoneyForamt)
        .TextMatrix(.Rows - 1, 3) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 4) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 5) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 6) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 7) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 8) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 9) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 10) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 11) = "��д��" & str�ϼ�
        .TextMatrix(.Rows - 1, 12) = "��д��" & str�ϼ�
        If mbln��ʾ���� And mblnIs��ҩ���� Then
            .TextMatrix(.Rows - 1, 13) = "��������" & dbl������ & str������λ
        End If
        .MergeCell (1)
        .MergeRow .Rows - 1, True
        .MsfObj.LeftCol = 0
    End With
    
    mblnAuto = False
    
    If err <> 0 Then
        MsgBox "��ʾ����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    Bill������ϸ.Row = ReLocateDetailRow
    
    WriteDataToBill = True
End Function

Private Function DependOnCheck() As Boolean
    Dim strSQL As String
    '�������ݼ��
    DependOnCheck = False
    
    With RecPart
        gstrSQL = " Select A.����||'-'||A.���� ҽ�� From ��Ա�� A,��Ա����˵�� B" & _
                 " Where (A.վ�� = '" & gstrNodeNo & "' Or A.վ�� is Null) And B.��Ա����='ҽ��' And A.ID=B.��ԱID" & _
                 " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
                 " Order by A.����"
        Call zlDatabase.OpenRecordset(RecPart, gstrSQL, "�������ݼ��")
        
        If .EOF Then
            MsgBox "���ʼ����Ա��ҽ����", vbInformation, gstrSysName
            Exit Function
        End If
        
        Me.Txt����ҽ��.Clear
        Txt����ҽ��.AddItem ""
        Do While Not .EOF
            Txt����ҽ��.AddItem !ҽ��
            .MoveNext
        Loop
        Txt����ҽ��.ListIndex = 0
    End With
    
    If IsHavePrivs(mstrPrivs, "����ҩ��") Then
        strSQL = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��')"
    Else
        strSQL = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��')"
    End If
    gstrSQL = " Select Distinct P.ID,P.���� From ���ű� P " & _
             " Where (P.վ�� = '" & gstrNodeNo & "' Or P.վ�� is Null) And P.ID In " & strSQL & _
             " And (P.����ʱ�� Is Null Or P.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set RecPart = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With RecPart
        If .EOF Then
            If IsHavePrivs(mstrPrivs, "����ҩ��") Then
                strSQL = "���ʼ��ҩ���������Ź���"
            Else
                strSQL = "�㲻��ҩ����Ա������ʹ�ñ�ģ�飡"
            End If
            MsgBox strSQL, vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
End Function

Private Function ReadFromReg()
    Dim RecRead As New ADODB.Recordset
    Dim strSub1 As String, strSub2 As String
    Dim strSave As String
    Dim arrColumn
    Dim int��ʾ�������� As Integer
    
    On Error Resume Next
    
    'ȡ������˽�в���
    strSave = zlDatabase.GetPara("������", glngSys, 1341)
    intFont = Val(zlDatabase.GetPara("����", glngSys, 1341))
    
    mintShowBill�շ� = Val(zlDatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, 1341))
    mintShowBill���� = Val(zlDatabase.GetPara("���ʴ�����ʾ��ʽ", glngSys, 1341))
    mbln���ʵ� = (Val(zlDatabase.GetPara("��ӡ�������ʵ�", glngSys, 1341)) = 1)
    mIntPrintHandbackNO = Val(zlDatabase.GetPara("��ӡ�˷ѵ��ݼ��", glngSys, 1341))
    mIntPrintDelay = Val(zlDatabase.GetPara("��ӡ�ӳ�", glngSys, 1341))
    int��ʾ�������� = Val(zlDatabase.GetPara("��ʾ��������", glngSys, 1341))
    mlngRefresh = Val(zlDatabase.GetPara("ˢ�¼��", glngSys, 1341))
    mlngPrintInterval = Val(zlDatabase.GetPara("��ӡ���", glngSys, 1341))
    intУ�鷢ҩ�� = Val(zlDatabase.GetPara("У�鷢ҩ��", glngSys, 1341))
    intУ����ҩ�� = Val(zlDatabase.GetPara("У����ҩ��", glngSys, 1341))
    mint�Զ����� = Val(zlDatabase.GetPara("�Զ�����", glngSys, 1341))
    mbln��ʾ��С��λ = (Val(zlDatabase.GetPara("��ʾ��С��λ", glngSys, 1341)) = 1)
    
    IntShowCol = Val(zlDatabase.GetPara("��ʾ����", glngSys, 1341))
    IntAutoPrint = Val(zlDatabase.GetPara("��ҩ���Զ���ӡ", glngSys, 1341))
    
    '�������������״̬��0-��λ;1-���ˡ�Ĭ���Ƕ�λ
    imgFilter.BorderStyle = Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & Me.Name, "���涨λ", cstLocate))
    
    '�������˿��أ�Ĭ����0-����ʾ
    img����.BorderStyle = int��ʾ��������
    
    '��ʾ��ҩ��������
    mlng�������� = GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & Me.Name, "��ʾ��ҩ��������", 1)
    Chk��ʾ��ҩ��������.Value = mlng��������
    
    '0-����ӡδ��ҩ����
    '1-��ӡ����������δ��ҩ����
    '2-��ӡ����������δ��ҩ����
    '3-ѡ���ӡ(��ҩ����)
    intPrint = Val(zlDatabase.GetPara("�����µ����Ƿ��ӡ", glngSys, 1341))
    mintPrintDrugLable = Val(zlDatabase.GetPara("��ӡҩƷ��ǩ", glngSys, 1341))
    lngҩ��ID = Val(zlDatabase.GetPara("��ҩҩ��", glngSys, 1341))
    Call GetDrugDigit(lngҩ��ID, Me.Caption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Str���� = zlDatabase.GetPara("��ҩ����", glngSys, 1341)
    Str��ҩ�� = zlDatabase.GetPara("��ҩ��", glngSys, 1341)
    strPrintWindow = zlDatabase.GetPara("��ӡָ����ҩ����", glngSys, 1341)
    mint�Զ���ҩ = Val(zlDatabase.GetPara("�Զ���ҩ", glngSys, 1341))
    mint�Զ���ҩʱ�� = Val(zlDatabase.GetPara("�Զ���ҩʱ��", glngSys, 1341))
    
    mstrSourceDep = zlDatabase.GetPara("��Դ����", glngSys, 1341)
    
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1]"
    Set RecRead = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
    
    With RecRead
        If Not .EOF Then
            IntCheckStock = !�����
        End If
        
        .Close
        IntSendAfterDosage = 1          '��ʾ����Ҫ������ҩ����
    End With

   gstrSQL = " Select Nvl(��ҩ,0) AS ��ҩ From ҩ����ҩ���� Where ҩ��ID=[1] Order by ����"
   Set RecRead = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
   
    'ֻҪ��һ���ʾ��Ҫ������ҩ���̵ģ����Ϊ��Ҫ��ҩ
    Do While Not RecRead.EOF
        If RecRead!��ҩ = 1 Then
            IntSendAfterDosage = 0
            Exit Do
        End If
        RecRead.MoveNext
    Loop
    
    If IntSendAfterDosage = 1 And mint�Զ���ҩ = 0 Then
        cbo��ҩ��.Enabled = True
        Call GetDosagePeople
    Else
        cbo��ҩ��.Text = ""
        cbo��ҩ��.Enabled = False
    End If
    
    '���β˵������߰�ť
    MnuEditDosage.Visible = (IntSendAfterDosage = 0 And IsHavePrivs(mstrPrivs, "��ҩ"))
    MnuEditAbolish.Visible = MnuEditDosage.Visible
    tabShow.TabVisible(0) = MnuEditDosage.Visible
    tabShow.TabVisible(1) = MnuEditDosage.Visible
    mnuChange.Visible = (IntSendAfterDosage = 0 And IsHavePrivs(mstrPrivs, "��ҩ") And mint�Զ���ҩ = 1)
    mnuLine10.Visible = mnuChange.Visible
        
    '������ʾ���Զ���ӡ������:ע��"δ��ҩƷ��¼"�ı���ΪA
    Select Case mintShowBill�շ�
        Case 0  '����ʾ����
            strSub1 = "A.����<>9 And A.����<>8"
            mstrShowSendedBill = "A.����<>9 And A.����<>8"
        Case 1  '��ʾδ�շ�
            strSub1 = "A.����<>9 And Nvl(A.���շ�,0)=0 And A.����=8"
            mstrShowSendedBill = "A.����<>9 And A.����=8"
        Case 2  '��ʾ���շ�
            strSub1 = "A.����<>9 And A.���շ�=1 And A.����=8"
            mstrShowSendedBill = "A.����<>9 And A.����=8"
        Case 3  '��ʾ���д���
            strSub1 = "A.����<>9 And A.����=8"
            mstrShowSendedBill = "A.����<>9 And A.����=8"
    End Select
    Select Case mintShowBill����
        Case 0  '����ʾ����
            strSub2 = "A.����<>8 And A.����<>9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.����<>8 And A.����<>9"
        Case 1  '��ʾδ���
            strSub2 = "A.����<>8 And Nvl(A.���շ�,0)=0 And A.����=9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.����<>8 And A.����=9"
        Case 2  '��ʾ�����
            strSub2 = "A.����<>8 And A.���շ�=1 And A.����=9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.����<>8 And A.����=9"
        Case 3  '��ʾ���д���
            strSub2 = "A.����<>8 And A.����=9"
            mstrShowSendedBill = mstrShowSendedBill & " Or " & "A.����<>8 And A.����=9"
    End Select
    mstrShowBill = " And A.���� IN(8,9) And (" & strSub1 & " Or " & strSub2 & ")"
    mstrShowSendedBill = " And A.���� IN(8,9) And (" & mstrShowSendedBill & ")"
    
    'ȡ��ҩƷ���Ƶĸ�ʽ��ʽ
    If strSave = "" Then strSave = "0|ҩƷ����,0|������,0|Ӣ����,0|���,0|����,0|��λ,0|����,0|����,0|���,0|����,0|�÷�,0|Ƶ��,0|����,0|�����,0|�ⷿ��λ,0|������,0|׼����,0|��ҩ��,0|��ע"
    arrColumn = Split(strSave, ",")
    intҩƷ���� = Val(Split(arrColumn(0), "|")(0))
    
    'ȡ������ɫ
    Call GetRecipeColor
    
    'ȡ����
    strOrder_1 = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "δ��ҩ��������", "")
    strOrder_2 = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ҩ��������", "")
    strOrder_3 = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "δ��ҩ��������", "")
    strOrder_4 = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�ѷ�ҩ��������", "")
    
    'ȡ����ģʽ
    mint����ģʽ = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "����ģʽ", "0"))
    If mint����ģʽ < 0 Or mint����ģʽ > 5 Then
        mint����ģʽ = 0
    End If
End Function

Private Function CheckAnother() As Boolean
    Dim BlnInҩ�� As Boolean, blnסԺ As Boolean, Bln���� As Boolean
    Dim BlnSetPeople As Boolean
    Dim RecTestPeople As New ADODB.Recordset
    Dim LngOldҩ��ID As Long, StrOld��ҩ�� As String
    
    CheckAnother = False
    
    If lngҩ��ID <> 0 Then
        With RecPart
            .MoveFirst
            .Find "ID=" & lngҩ��ID
            BlnInҩ�� = (RecPart.EOF <> True)
            
            If BlnInҩ�� Then   '˵���ò�������ҩ��
                'ȡ��λ
                blnסԺ = False

                gstrSQL = "Select nvl(�������,1) ������� From ��������˵�� Where ����ID+0=[1]"
                Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !������� = 2 Or !������� = 3 Then blnסԺ = True: Exit Do
                        .MoveNext
                    Loop
                    Bln���� = False
                    If blnסԺ Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !������� = 3 Then Bln���� = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If blnסԺ = False Then
                    mInt���� = 8
                Else
                    mInt���� = IIf(Bln����, 0, 9)
                End If
            End If
        End With
    End If
    
    '���ö�Ӧ��ҩ��
    If lngҩ��ID = 0 Or BlnInҩ�� = False Then
        '�����ô���
        With Frm��ҩ��������
            MsgBox IIf(Str��ҩ�� = "", "������ҩ������ҩ�ˣ�", "������ҩ����"), vbInformation, gstrSysName
            Set .RecPart = RecPart.Clone
            .strShow = IIf(Str��ҩ�� = "", "������ҩ������ҩ�ˣ�", "������ҩ����")
            .mstrPrivs = mstrPrivs
            .Show 1, Me
        End With
        Call ReadFromReg

        '��δ����ҩ�����˳�
        If lngҩ��ID = 0 Then Exit Function
        '���»�ȡ��ҩ����ʹ�õ�λ
        With RecPart
            .MoveFirst
            .Find "ID=" & lngҩ��ID
            BlnInҩ�� = (RecPart.EOF <> True)
            
            If BlnInҩ�� Then   '˵���ò�������ҩ��
                'ȡ��λ
                blnסԺ = False

                gstrSQL = "Select nvl(�������,1) ������� From ��������˵�� Where ����ID+0=[1]"
                Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID)
                
                With RecTestPeople
                    Do While Not .EOF
                        If !������� = 2 Or !������� = 3 Then blnסԺ = True: Exit Do
                        .MoveNext
                    Loop
                    Bln���� = False
                    If blnסԺ Then
                        If .RecordCount <> 0 Then .MoveFirst
                        Do While Not .EOF
                            If !������� = 3 Then Bln���� = True: Exit Do
                            .MoveNext
                        Loop
                    End If
                End With
                If blnסԺ = False Then
                    mInt���� = 8
                Else
                    mInt���� = IIf(Bln����, 0, 9)
                End If
            Else
                Exit Function    '��ҩ�����˳�
            End If
        End With
    End If
    
    If IntSendAfterDosage = 0 And Str��ҩ�� <> "|��ǰ����Ա|" Then
        LngOldҩ��ID = lngҩ��ID
        StrOld��ҩ�� = Str��ҩ��
        
        '������ҩ��
        BlnSetPeople = False
        If Str��ҩ�� = "" Then
            MsgBox "��������ҩ�ˣ�", vbInformation, gstrSysName
            With Frm��ҩ��������
                Set .RecPart = RecPart.Clone
                .strShow = "��������ҩ�ˣ�"
                .mstrPrivs = mstrPrivs
                .Show 1, Me
            End With
            Call ReadFromReg

            If Str��ҩ�� = "" Then
                MsgBox "������������ҩ�ˣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '�����ҩ�˷Ǳ�����,�������������
        gstrSQL = " Select Count(*) Records From ������Ա Where ��ԱID=" & _
                 " (Select Distinct ID From ��Ա�� Where ����=[2]) And " & _
                 " ����ID+0 =[1]"
        Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID, Str��ҩ��)
        
        With RecTestPeople
            If .EOF Then
                BlnSetPeople = True
            Else
                If IsNull(!Records) Then
                    BlnSetPeople = True
                Else
                    If !Records = 0 Then
                        BlnSetPeople = True
                    End If
                End If
            End If
        End With
        If BlnSetPeople Then
            MsgBox "��������ҩ�ˣ�ԭ��ҩ���Ѳ����ڱ�ҩ������", vbInformation, gstrSysName
            With Frm��ҩ��������
                Set .RecPart = RecPart.Clone
                .strShow = "��������ҩ�ˣ�ԭ��ҩ���Ѳ����ڱ�ҩ������"
                .mstrPrivs = mstrPrivs
                .Show 1, Me
            End With
            Call ReadFromReg
            If Str��ҩ�� = "" Then
                MsgBox "������������ҩ�ˣ�ԭ��ҩ���Ѳ����ڱ�ҩ����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                Exit Function
            End If
            If StrOld��ҩ�� = Str��ҩ�� And LngOldҩ��ID = lngҩ��ID Then Exit Function
        End If
    End If
    
    CheckAnother = True
End Function

Private Function CheckSpec(ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim strNote As String
    Dim rsTemp As New ADODB.Recordset
    '�Զ�����ҩƷ���м��
    gstrSQL = " SELECT Distinct " & _
        " '['||C.����||']'||" & IIf(mblnTradeName, "NVL(L.����,C.����)", "C.����") & " Ʒ��,X.�������" & _
        " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� L,ҩƷ���� X " & _
        " WHERE A.ҩƷID=B.ҩƷID AND B.ҩ��ID=X.ҩ��ID And B.ҩƷID=C.ID " & _
        " AND B.ҩƷID=L.�շ�ϸĿID(+) AND L.����(+)=3 AND L.����(+)=1 " & _
        " AND A.����� IS NULL AND MOD(A.��¼״̬,3)=1 AND NVL(A.ժҪ,'С��')<>'�ܷ�'" & _
        " AND A.NO=[1] AND A.����=[2] AND (A.�ⷿID+0=[3] OR A.�ⷿID IS NULL) " & _
        " And X.�������<>'��ͨҩ'" & _
        " Order by X.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�Զ�����ҩƷ���м��]", strNo, IntBillStyle, lngҩ��ID)
    
    If rsTemp.RecordCount = 0 Then
        CheckSpec = True
        Exit Function
    End If
    
    With rsTemp
        Do While Not .EOF
            strNote = strNote & vbCrLf & Space(4) & !������� & "-" & !Ʒ��
            .MoveNext
        Loop
    End With
'    If MsgBox("�Ƿ�����¶����顢������ҩƷ���з�ҩ��" & strNote, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    mstr��������ʾ = strNote
    CheckSpec = True
End Function
Private Function CheckStock(ByVal strNo As String, ByVal IntBillStyle As Integer) As Boolean
    Dim RecCheckStock As New ADODB.Recordset, RecBillData As New ADODB.Recordset
    Dim dblStock As Double, intCheck As Integer
    '--�����--
    '0-�����;1-���,��������;2-���,�����ֹ
    On Error Resume Next
    err = 0
    CheckStock = False
    intCheck = IntCheckStock
    
    '���м��
    If intCheck <> 0 Then
        gstrSQL = " SELECT A.ҩƷID,SUM(NVL(A.ʵ������,0)*NVL(A.����,1)) ����," & _
                " '['||C.����||']'||" & IIf(mblnTradeName, "NVL(L.����,C.����)", "C.����") & " Ʒ��,NVL(A.����,0) ����" & _
                " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� L,���˷��ü�¼ D " & _
                " WHERE A.ҩƷID=B.ҩƷID AND B.ҩƷID=C.ID" & _
                " AND B.ҩƷID=L.�շ�ϸĿID(+) AND L.����(+)=3 AND L.����(+)=1 " & _
                " AND A.����� IS NULL AND MOD(A.��¼״̬,3)=1 AND NVL(A.ժҪ,'С��')<>'�ܷ�'" & _
                " AND A.����ID=D.ID AND A.NO=[1] AND A.����=[2] AND (A.�ⷿID+0=[3] OR A.�ⷿID IS NULL) " & _
                " GROUP BY A.ҩƷID,'['||C.����||']'||" & IIf(mblnTradeName, "NVL(L.����,C.����)", "C.����") & ",����"
        Set RecBillData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngҩ��ID)
        
        With RecBillData
            Do While Not .EOF
                gstrSQL = " Select nvl(ʵ������,0) ����" & _
                         " From ҩƷ��� " & _
                         " Where �ⷿID+0=[1] And ҩƷID=[2] " & _
                         " And ����=1 And Nvl(����,0)=[3]"
                Set RecCheckStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩ��ID, CLng(RecBillData!ҩƷID), CLng(RecBillData!����))
                
                With RecCheckStock
                    If .EOF Then
                        dblStock = 0
                    Else
                        dblStock = !����
                    End If
                    
                    If dblStock < RecBillData!���� Then
                        Select Case intCheck
                        Case 1
                            If MsgBox(RecBillData!Ʒ�� & "�Ŀ�����������Ƿ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox RecBillData!Ʒ�� & "�Ŀ�������������ܼ�����ҩ��", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End With
                .MoveNext
            Loop
        End With
    End If
    
    If err <> 0 Then
        MsgBox "�����ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Exit Function
    End If
    CheckStock = True
End Function

Private Function ResizePicClose()
    Dim DblHeight As Double, DblWidth As Double
    Dim intCols As Integer
    '--�����رհ�ť��λ��--
    
    With Msf�б�
        DblHeight = .CellHeight * .Rows
        DblWidth = 0
        For intCols = 0 To .Cols - 1
            DblWidth = DblWidth + .ColWidth(intCols)
        Next
        
        If DblHeight > .Height - 180 Or (DblHeight > .Height - 420 And DblWidth > .Width - 70) Then
            With PicCloseConsignment
                .Left = Msf�б�.Width - .Width - 30 - 250
            End With
        Else
            With PicCloseConsignment
                .Left = Msf�б�.Width - .Width - 30
            End With
        End If
    End With
End Function

Private Sub ClearCons()
    Me.Txt���� = ""
    Me.Txt����ҽ��.ListIndex = 0
    Me.Txt���� = ""
    Me.Txt���� = ""
'    Me.cbo��ҩ�� = ""
    Me.Txt�Ա� = ""
    Me.TxtסԺ�� = ""
    Me.Txt�շ�Ա = ""
    Me.txtԭʼ���� = ""
    Me.txt��ҩ�巨 = ""
    
    Bill������ϸ.ClearBill
End Sub

Private Function SetMenuState()
    MnuEditDosage.Checked = IIf(tabShow.Tab = 0, True, False)
    MnuEditAbolish.Checked = IIf(tabShow.Tab = 1, True, False)
    MnuEditConsignment.Checked = IIf(tabShow.Tab = 2, True, False)
    MnuEditHandback.Checked = IIf(tabShow.Tab = 3, True, False)
    
    mnuCancel.Enabled = False
    Tbar1.Buttons("Cancel").Enabled = False
    
    'ȡ����ҩ��������ҩģʽ�������ڷ���ģʽ��
    Chk�嵥.Visible = (tabShow.Tab = 3)
    Chk��ʾ��ҩ��������.Visible = (tabShow.Tab = 0 Or tabShow.Tab = 1 Or tabShow.Tab = 2)
    
    Call SetPosition
    
End Function

Private Function SetButtonState()
    If MnuEditDosage.Checked Then
        tabShow.Tab = 0
    ElseIf MnuEditAbolish.Checked Then
        tabShow.Tab = 1
    ElseIf MnuEditConsignment.Checked Then
        tabShow.Tab = 2
    Else
        tabShow.Tab = 3
    End If
    
    BlnInOper = False
    Call mnuViewRefresh_Click
End Function

Private Sub TimeRefresh_Timer()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    '�������ڲ��ǵ�ǰ����ʱ�˳�
    If Screen.ActiveForm.hWnd <> Me.hWnd Then Exit Sub
    
    If tabShow.Tab = 3 Then
        If Chkȫ��.Value = 0 Or mblnAllBack = False Then Exit Sub
    End If
    
    TimeRefresh.Enabled = False
    DoEvents
    Call mnuViewRefresh_Click
    DoEvents
    TimeRefresh.Enabled = True
End Sub
Private Sub TimePrintCancelBill_Timer()
    Dim curDateBegin As Date
    Dim curDateEnd As Date
    
    '���ô�ӡ�˷ѵ�
    IntTimes = IntTimes + 1
    '�����������˳�
    If IntTimes < mIntPrintHandbackNO Then Exit Sub
    IntTimes = 0
    
    curDateEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
    curDateBegin = DateAdd("n", 0 - mIntPrintHandbackNO, curDateEnd)
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_8", Me, "��ʼʱ��=" & Format(curDateBegin, "yyyy-MM-dd hh:mm"), "����ʱ��=" & Format(curDateEnd, "yyyy-MM-dd hh:mm"), "ҩ��=" & lngҩ��ID, 2)
End Sub
Private Function AutoPrint()
'���ܣ��Զ���ӡ����
    Dim recAutoPrint As New ADODB.Recordset, strErr As String
    Dim datCurr As Date, strRefresh As String, strCond As String
    Dim strUnit As String
    Dim str����Ա As String
    Dim blnInTrans As Boolean
    Dim blnIgnore As Boolean
    Dim strName As String
    
    '���ݴ�ӡ�����������
    '0-����ӡδ��ҩ����
    '1-��ӡ����������δ��ҩ����
    '2-��ӡ����������δ��ҩ����
    '3-ѡ���ӡ(��ҩ����)
    If BlnInRefresh Then Exit Function
    
    If mblnIsFirst = False And mint�Զ���ҩ = 1 Then
        If mint�Զ���ҩʱ�� > 0 Then
            If DateDiff("s", mdate�ϴ�У��ʱ��, zlDatabase.Currentdate) > mint�Զ���ҩʱ�� * 60 Then
                strName = zlDatabase.UserIdentify(Me, "У����ҩ��", glngSys, 1341, "��ҩ")
               
                If Trim(strName) = "" Then Exit Function
                mstr�Զ���ҩ�� = strName
                
                mdate�ϴ�У��ʱ�� = zlDatabase.Currentdate
            End If
        End If
    End If
    
    Select Case intPrint
        Case 0
            If mintPrintDrugLable = 0 Then Exit Function
        Case 1
            If Not mbln���ʵ� Then strRefresh = " And ����=8"
        Case 2
            If mbln���ʵ� Then
                If Str���� <> "" Then
                    strRefresh = " And (����=8 And ��ҩ���� IN(" & Str���� & ") Or ����=9)"
                End If
            Else
                If Str���� <> "" Then
                    strRefresh = " And ����=8 And ��ҩ���� IN(" & Str���� & ")"
                Else
                    strRefresh = " And ����=8"
                End If
            End If
        Case 3
            If mbln���ʵ� Then
                If strPrintWindow <> "" Then
                    strRefresh = " And (����=8 And ��ҩ���� IN(" & strPrintWindow & ") Or ����=9)"
                End If
            Else
                If strPrintWindow <> "" Then
                    strRefresh = " And ����=8 And ��ҩ���� IN(" & strPrintWindow & ")"
                Else
                    strRefresh = " And ����=8"
                End If
            End If
    End Select
    
    If mInt���� = 0 Then
        strCond = " And A.���� In (8,9)" '���ＰסԺ���е���
    Else
        If mInt���� = 8 Then
            strCond = " And A.���� In (8,9) And A.��ҳID Is NULL " '���ﻮ�ۼ��������
        Else
            strCond = " And A.���� = 9 And A.��ҳID Is Not NULL " 'סԺ����
        End If
    End If
            
    On Error GoTo ErrHand
    BlnInRefresh = True
    
    gstrSQL = " Select NO,����,��������" & _
               " From δ��ҩƷ��¼ A " & _
               " Where �ⷿID+0=[13] " & strRefresh & IIf(mint�Զ���ҩ = 1, " And ��ҩ�� Is Null ", "") & _
               " " & IIf(StrFind_1 = "", " And �������� " & StrDate, StrFind_1) & _
               " And ��ӡ״̬ Not In (1,2) " & strCond & mstrShowBill & _
               " " & IIf(mstrSourceDep = "", "", " And A.�Է�����id+0 in(" & mstrSourceDep & ") ") & _
               " Order by ���ȼ�,����,No"
    
    Set recAutoPrint = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
         SQLCondition.date��ʼ����, _
         SQLCondition.date��������, _
         SQLCondition.str��ʼNO, _
         SQLCondition.str����NO, _
         SQLCondition.str����, _
         SQLCondition.str���￨, _
         SQLCondition.str��ʶ��, _
         SQLCondition.lng����ID, _
         SQLCondition.str������, _
         SQLCondition.str�����, _
         SQLCondition.lngҩƷID, _
         SQLCondition.str��ǰNO, _
         lngҩ��ID)

    datCurr = zlDatabase.Currentdate()
        
    With recAutoPrint
        Do While Not .EOF
            '��ӡ����
            If DateDiff("s", !��������, datCurr) > mIntPrintDelay Then
                If intPrint > 0 Then
                    If mint�Զ���ҩ = 1 Then
                        '�����Զ���ҩ���ڴ�ӡǰ���
                        blnIgnore = False
                        
                        '����Ƿ���Ҫ��ҩ
                        If Not IsDosage(Val(!����), !NO) Then
                            blnIgnore = True
                        End If
                        
                        '����Ƿ�����
                        If CheckBill(1, Val(!����), !NO) <> 0 Then
                            blnIgnore = True
                        End If
                        
                        If blnIgnore = False Then
                            gcnOracle.BeginTrans
                            blnInTrans = True
        
                            '��������ҩ��
                            str����Ա = IIf(mstr�Զ���ҩ�� <> "", mstr�Զ���ҩ��, IIf(Str��ҩ�� = "|��ǰ����Ա|", gstrUserName, Str��ҩ��))
                            
                            gstrSQL = "zl_ҩƷ�շ���¼_������ҩ��(" & lngҩ��ID & "," & Val(!����) & ",'" & !NO & "','" & str����Ա & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-������ҩ��")
                            
                            gstrSQL = " Update δ��ҩƷ��¼ Set ��ӡ״̬=1 Where ����=" & !���� & " And No='" & !NO & "' And �ⷿID=" & lngҩ��ID & " " & IIf(mstrSourceDep = "", "", " And �Է�����id+0 in(" & mstrSourceDep & ") ")
                            Call ExecuteProcedure(Me.Caption & "-���µ����Ѵ�ӡ", False)
                            
                            '����������˵���ǩ��������Ҫ����ҩ�˽��е���ǩ������
                            If gblnҩƷʹ�õ���ǩ�� = True Then
                                If SaveSignatureRecored(EsignTache.Dosage, Val(!����), !NO, lngҩ��ID) = False Then
                                    gcnOracle.RollbackTrans
                                    Exit Function
                                End If
                            End If
                            
                            gcnOracle.CommitTrans
                            blnInTrans = False
                            
                            mblnIsFirst = False
                        End If
                    Else
                        gstrSQL = " Update δ��ҩƷ��¼ Set ��ӡ״̬=1 Where ����=" & !���� & " And No='" & !NO & "' And �ⷿID=" & lngҩ��ID & " " & IIf(mstrSourceDep = "", "", " And �Է�����id+0 in(" & mstrSourceDep & ") ")
                        Call ExecuteProcedure(Me.Caption & "-���µ����Ѵ�ӡ", False)
                    End If

                    strUnit = GetUnit(lngҩ��ID, !����, !NO)
                    If Not BillHaveHerial(!NO, !����) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_3", Me, _
                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), "ReportFormat=2", "PrintEmpty=0", 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_4", Me, _
                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ReportFormat=2", "PrintEmpty=0", 2)
                    End If
                End If
                
                '��ӡҩƷ��ǩ
                If mintPrintDrugLable = 1 Then
                    If Not BillHaveHerial(!NO, !����) Then
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_6", Me, _
                            "NO=" & !NO, "����=" & IIf(!���� = 8, 1, 2), "ҩ��=" & lngҩ��ID, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "D.�����װ", "D.סԺ��װ"), 2)
                    Else
                        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1341_7", Me, _
                            "NO=" & !NO, "ҩ��=" & lngҩ��ID, 2)
                    End If
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    BlnInRefresh = False
    Exit Function
ErrHand:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal IntOper As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, Optional ByVal bln��ʾ As Boolean = False) As Integer
    Dim dblCount As Double
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim RecCheck As New ADODB.Recordset
    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    'IntOper:1-��ҩ;2-ȡ����ҩ;3-��ҩ;4-��ҩ;5-ȡ����ҩ
    '����:
    '0-�������
    '1-δ��ҩ
    '2-����ҩ
    '3-�ѷ�ҩ
    '4-��ɾ��
    '5-δ��ҩ
    
    '��������ȡ����ҩʱ�ļ��
    If IntOper = 5 Then
        gstrSQL = "Select ����� From ҩƷ�շ���¼ Where No=[1] And ����=[2] And �ⷿID+0=[3] And ��¼״̬=1 And ����� IS Not Null And Rownum=1 "
        Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngҩ��ID)
        If RecCheck.EOF Then
            CheckBill = 4
            MsgBox "δ�ҵ�ָ�����ݣ����ѱ���������Ա����,����������ֹ��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
     
    gstrSQL = " Select A.��ҩ��,A.����� From ҩƷ�շ���¼ A" & _
        " Where A.No=[1] And A.����=[2] " & _
        " " & IIf(IntOper <> 4, " And mod(A.��¼״̬,3)=1", "") & " And Rownum=1 " & _
        " And Nvl(Ltrim(Rtrim(A.ժҪ)),'С��')<>'�ܷ�' And (A.�ⷿID+0=[3] Or A.�ⷿID Is NULL)"
    
    If IntOper = 4 Then
        gstrSQL = gstrSQL & " And ����� IS Not Null"
    Else
        gstrSQL = gstrSQL & " And ����� IS Null"
    End If

    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lngҩ��ID)
    
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            If InStr(1, "123", IntOper) <> 0 Then CheckBill = 3: MsgBox "�ô����ѱ���������Ա��ҩ��" & IIf(IntOper = 1, "��ҩ", IIf(IntOper = 2, "ȡ����ҩ", IIf(IntOper = 3, "��ҩ", "��ҩ"))) & "������ֹ��", vbInformation, gstrSysName: Exit Function
        Else
            If InStr(1, "4", IntOper) <> 0 Then CheckBill = 5: MsgBox "�ô�����δ��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
            If Not IsNull(!��ҩ��) Then
                If InStr(1, "1", IntOper) <> 0 Then CheckBill = 2: MsgBox "�ô�������ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
            Else
                If InStr(1, "2", IntOper) <> 0 Then CheckBill = 1: MsgBox "�ô���δ��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
            End If
        End If
    End With
    
    '�������ҩ������Ƿ�����δ����ҽ����ҩ
    If blnҽ������ = False And bln��ʾ Then
        intRows = Bill������ϸ.Rows - 2
        For intRow = 1 To intRows
            dblCount = Val(Bill������ϸ.TextMatrix(intRow, ����.��ҩ��))
            If dblCount <> 0 Then
                gstrSQL = "select ���� From ҩƷ�շ���¼ Where ID=[1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ�������]", Val(Bill������ϸ.TextMatrix(intRow, ����.Id)))

                If (rsTemp!���� Like "1*") Then       '����
                    gstrSQL = "Select Nvl(ҽ�����,0) ҽ�����,Nvl(�����־,1) �����־ From ���˷��ü�¼ Where ID=(Select ����ID From ҩƷ�շ���¼ Where ID=[1])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[����Ƿ���ҽ��]", Val(Bill������ϸ.TextMatrix(intRow, ����.Id)))
                    
                    If Not rsTemp.EOF Then
                        If (rsTemp!�����־ = 1 Or rsTemp!�����־ = 4) And rsTemp!ҽ����� <> 0 Then
                            gstrSQL = "Select decode(ҽ��״̬,4,1,0) ���� From ����ҽ����¼ Where ID=[1]"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[�жϸ�ҽ���Ƿ�����]", CLng(rsTemp!ҽ�����))
                            
                            If rsTemp!���� = 0 Then
                                CheckBill = 1
                                MsgBox "��" & intRow & "�е�ҩƷ��¼��Ӧ��ҽ����δ���ϣ���������ҩ��", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    CheckBill = 0
End Function

Private Sub subPrint(ByVal bytMode As Byte)
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim PrintRec As New ADODB.Recordset
    Dim strCond As String, strTemp As String
    Dim strCond2 As String, strCond1 As String, intLeft As Integer, intRight As Integer
    
    If Msf�б�.Rows = 2 Then
        If Msf�б�.TextMatrix(1, ��������.NO) = "" Then Exit Sub
    End If
    
    If mInt���� = 0 Then
        strCond = " And A.���� In (8,9)" '���ＰסԺ���е���
    Else
        If mInt���� = 8 Then
            strCond = " And A.���� In (8,9) " '���ﻮ�ۼ��������
        Else
            strCond = " And A.���� = 9 " 'סԺ����
        End If
    End If
    
    '��Ƕ�ײ�ѯ�У�û�����Ӳ��˷��ü�¼���������д��������ֶ�ʱ����ȥ���������������õ����˷��ü�¼��
    strCond1 = ""
    StrFind_4 = UCase(StrFind_4)
    strCond2 = StrFind_4
    intLeft = InStr(1, StrFind_4, " AND UPPER(H.����)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, StrFind_4, " AND")
        strTemp = Mid(StrFind_4, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = Mid(StrFind_4, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(StrFind_4, intRight)
        Else
            strCond1 = Mid(StrFind_4, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(H.��ʶ��)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(B.���￨��)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    
    '���ݵ��������ñ������ݵĵ�λ
    Const str�ۼ� As String = "X.���㵥λ ��λ,ltrim(to_char(S.���ۼ�,'999990.00000')) ����,ltrim(to_char(S.ʵ������,'999990.00000')) ����,LTRIM(TO_CHAR(S.��������,'999990.00000')) ��������,LTRIM(TO_CHAR(S.�ѷ�����,'999990.00000')) ׼����,"
    Const str���� As String = "D.���ﵥλ ��λ,ltrim(to_char(S.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����,ltrim(to_char(S.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����,LTRIM(TO_CHAR(S.��������/DECODE(D.�����װ,NULL,1,0,1,D.�����װ),'999990.00000')) ��������,LTRIM(TO_CHAR(S.�ѷ�����/DECODE(D.�����װ,NULL,1,0,1,D.�����װ),'999990.00000')) ׼����,"
    Const strסԺ As String = "D.סԺ��λ ��λ,ltrim(to_char(S.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����,ltrim(to_char(S.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����,LTRIM(TO_CHAR(S.��������/DECODE(D.סԺ��װ,NULL,1,0,1,D.סԺ��װ),'999990.00000')) ��������,LTRIM(TO_CHAR(S.�ѷ�����/DECODE(D.סԺ��װ,NULL,1,0,1,D.סԺ��װ),'999990.00000')) ׼����,"
    Const strҩ�� As String = "D.ҩ�ⵥλ ��λ,ltrim(to_char(S.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����,ltrim(to_char(S.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����,LTRIM(TO_CHAR(S.��������/DECODE(D.ҩ���װ,NULL,1,0,1,D.ҩ���װ),'999990.00000')) ��������,LTRIM(TO_CHAR(S.�ѷ�����/DECODE(D.ҩ���װ,NULL,1,0,1,D.ҩ���װ),'999990.00000')) ׼����,"
    Dim str��λ��1 As String
    
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��λ��1 = str�ۼ�
    Case "���ﵥλ"
        str��λ��1 = str����
    Case "סԺ��λ"
        str��λ��1 = strסԺ
    Case "ҩ�ⵥλ"
        str��λ��1 = strҩ��
    End Select
    
    '��ʼ����ӡ����
    With PrintRec
        If .State = 1 Then .Close
        If MnuEditHandback.Checked Then
            '����ҩ�嵥
            strCond = Replace(strCond, "A.", "S.")
            strCond1 = Replace(strCond1, "H.", "C.")
            If Chk�嵥.Value = 0 Then
                '##################������ʾÿ�ʼ�¼�������˶���##################
                gstrSQL = " Select ����, NO, ����, ����, �Ա�, ����, סԺ��, ����, Ʒ��, ���, ��λ, ����, LTrim(To_Char(Sum(����), '999990.00000')) ����," & _
                         " LTrim(To_Char(Sum(��������), '999990.00000')) ��������, LTrim(To_Char(Sum(׼����), '999990.00000')) ׼����, " & _
                         " LTrim(To_Char(Sum(���), '999990.00')) ���, ��ҩʱ��, ��ҩ��" & _
                         " From (SELECT Distinct DECODE(S.����,8,'�շ�',9,'����') ����,S.NO,P.���� ����,C.����,C.�Ա�,C.����,C.��ʶ�� סԺ��,C.����,'['||x.����||']'||" & IIf(mblnTradeName, "NVL(A.����,X.����)", "X.����") & " Ʒ��," & _
                         " DECODE(x.���,NULL,x.����,DECODE(x.����,NULL,x.���,x.���||'|'||x.����)) ���," & str��λ��1 & _
                         " LTRIM(TO_CHAR(S.���۽��,'999990.00')) ���,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ��ҩʱ��,S.����� ��ҩ��,S.���" & _
                         " FROM " & _
                         "      (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
                         "          NVL(A.����,1) ����,A.ʵ������ ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
                         "          A.���ۼ�,A.���۽��,A.����,A.Ƶ��,A.�÷�,A.ժҪ,A.�����,A.�������,A.�Է�����ID,A.�ⷿID, A.���" & _
                         "      FROM" & _
                         "          (SELECT *" & _
                         "          FROM ҩƷ�շ���¼ A" & _
                         "          WHERE A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                         "          AND A.�ⷿID+0=[13] " & _
                                    IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                         "          ) A," & _
                         "          (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                         "          FROM ҩƷ�շ���¼ A" & _
                         "          WHERE A.����� IS NOT NULL" & _
                         "          AND A.�ⷿID+0=[13] " & _
                                    IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                         "          GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
                         "      WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0) S,"
                gstrSQL = gstrSQL & "" & _
                         "      ���˷��ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,������Ϣ B " & _
                         " WHERE S.ҩƷID=D.ҩƷID AND D.ҩƷID=X.ID And C.����ID=B.����ID(+) " & _
                         " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & _
                         " AND S.�Է�����ID+0=P.ID " & strCond & strCond1 & _
                         " AND S.����ID=C.ID  AND (S.��¼״̬=1 OR MOD(S.��¼״̬,3)=0)" & _
                         " AND S.����� IS NOT NULL AND S.�ⷿID+0=[13] AND S.ʵ������*S.����>S.��������) " & _
                         " Group By ����, NO, ����, ����, �Ա�, ����, סԺ��, ����, Ʒ��, ���, ��λ, ����, ��ҩʱ��, ��ҩ�� "
            Else
                '##################�嵥��ʾÿ�ʲ�������##################
                gstrSQL = " Select ����, NO, ����, ����, �Ա�, ����, סԺ��, ����, Ʒ��, ���, ��λ, ����, LTrim(To_Char(Sum(����), '999990.00000')) ����," & _
                         " LTrim(To_Char(Sum(��������), '999990.00000')) ��������, LTrim(To_Char(Sum(׼����), '999990.00000')) ׼����, " & _
                         " LTrim(To_Char(Sum(���), '999990.00')) ���, ��ҩʱ��, ��ҩ�� From " & _
                         " (SELECT DECODE(S.����,8,'�շ�',9,'����') ����,S.NO,P.���� ����,C.����,C.�Ա�,C.����,C.��ʶ�� סԺ��,C.����,'['||X.����||']'||" & IIf(mblnTradeName, "NVL(A.����,X.����)", "X.����") & " Ʒ��," & _
                         " DECODE(X.���,NULL,X.����,DECODE(X.����,NULL,X.���,X.���||'|'||X.����)) ���," & str��λ��1 & _
                         " LTRIM(TO_CHAR(S.���۽��,'999990.00')) ���,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ��ҩʱ��,S.����� ��ҩ��" & _
                         " FROM "
                gstrSQL = gstrSQL & _
                         "      (SELECT * FROM" & _
                         "          (SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
                         "              NVL(A.����,1) ����,A.ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
                         "              A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID,1 �ɲ���" & _
                         "          FROM" & _
                         "              (SELECT *" & _
                         "              FROM ҩƷ�շ���¼ A" & _
                         "              WHERE A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)" & _
                         "              AND A.�ⷿID+0=[13] " & _
                                        IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                         "              ) A," & _
                         "              (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
                         "              FROM ҩƷ�շ���¼ A" & _
                         "              WHERE A.����� IS NOT NULL " & _
                         "              AND A.�ⷿID+0=[13] " & _
                                        IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                         "              GROUP BY A.NO,A.����,A.ҩƷID,A.���) B"
                gstrSQL = gstrSQL & _
                         "          WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.���)" & _
                         "          UNION" & _
                         "          SELECT A.ID,A.NO,A.����,A.ҩƷID,A.����ID,A.����,A.����,A.Ч��,NVL(A.����,0)," & _
                         "          NVL(A.����,1) ����,A.ʵ������,0 ������,0 �ѷ�����,A.��¼״̬," & _
                         "          A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID," & _
                         "          DECODE(A.��¼״̬,1,1,DECODE(MOD(A.��¼״̬,3),0,1,MOD(A.��¼״̬,3)+1)) �ɲ���" & _
                         "          FROM ҩƷ�շ���¼ A" & _
                         "          WHERE A.����� IS NOT NULL AND NOT (��¼״̬=1 OR MOD(��¼״̬,3)=0)" & _
                         "          AND A.�ⷿID+0=[13] " & _
                                    IIf(strCond2 = "", " AND A.������� " & StrDate & "", strCond2) & _
                         "          ) S,"
                gstrSQL = gstrSQL & "" & _
                         "      ���˷��ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,������Ϣ B " & _
                         " WHERE S.ҩƷID=D.ҩƷID AND D.ҩƷID=X.ID AND S.�Է�����ID+0=P.ID " & _
                         " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 And C.����ID=B.����ID(+) " & _
                         " AND S.����ID=C.ID " & strCond & strCond1 & " AND S.����� IS NOT NULL)  " & _
                         " Group By ����, NO, ����, ����, �Ա�, ����, סԺ��, ����, Ʒ��, ���, ��λ, ����, ��ҩʱ��, ��ҩ�� "
            End If
        
            Dim blnMoved As Boolean
            Dim str��ʼ���� As String, strSQL As String
            
            str��ʼ���� = IIf(strCond2 = "", StrFind_4, strCond2)
            'ȡ��ʼ����:intRight���浥���ŵ���ʼλ��
            intRight = InStr(1, str��ʼ����, "'") + 1
            str��ʼ���� = Mid(str��ʼ����, intRight, 19)
            '�жϴӿ�ʼ���ں��Ƿ����ת���Ĵ�������
            blnMoved = zlDatabase.DateMoved(str��ʼ����)
            
            '�����������ת��������Ҫͬʱ�Ӻ󱸱�����ȡ����
            If blnMoved Then
                strSQL = gstrSQL
                strSQL = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
                strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
                gstrSQL = gstrSQL & " UNION ALL " & strSQL
            End If
            
            If Chk�嵥.Value = 0 Then
                gstrSQL = gstrSQL & " ORDER BY NO,����"
            Else
                gstrSQL = gstrSQL & " ORDER BY NO,����,��ҩʱ��"
            End If
        Else
            'δ��ҩ�嵥
            Const str�ۼ�1 As String = "C.���㵥λ ��λ,ltrim(to_char(B.���ۼ�,'999990.00000')) ����,ltrim(to_char(B.ʵ������,'999990.00000')) ����,"
            Const str����1 As String = "D.���ﵥλ ��λ,ltrim(to_char(B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����,"
            Const strסԺ1 As String = "D.סԺ��λ ��λ,ltrim(to_char(B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����,"
            Const strҩ��1 As String = "D.ҩ�ⵥλ ��λ,ltrim(to_char(B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����,"
            
            Select Case strUnit
            Case "�ۼ۵�λ"
                str��λ�� = str�ۼ�1
            Case "���ﵥλ"
                str��λ�� = str����1
            Case "סԺ��λ"
                str��λ�� = strסԺ1
            Case "ҩ�ⵥλ"
                str��λ�� = strҩ��1
            End Select
            gstrSQL = "Select ����, NO, ����, ����, �Ա�, ����, סԺ��, ����, Ʒ��, ���, ��λ, ����, LTrim(To_Char(Sum(����), '999990.00000')) ����," & _
                     " LTrim(To_Char(Sum(���), '999990.00')) ���, ������, ��������, ��ҩ�� From " & _
                     " (SELECT DECODE(A.����,8,'�շ�',9,'����') ����,A.NO," & _
                     " T.���� ����,H.����,H.�Ա�,H.����,H.��ʶ�� סԺ��,H.����," & _
                     " '['||c.����||']'||C.���� Ʒ��,DECODE(C.���,NULL,C.����,DECODE(C.����,NULL,C.���,C.���||'|'||C.����)) ���," & str��λ�� & _
                     " LTRIM(TO_CHAR(B.���۽��,'999990.00')) ���,B.������,B.��������,DECODE(B.��ҩ��,'���ŷ�ҩ','',NULL,'',B.��ҩ��) ��ҩ��" & _
                     " FROM ҩƷ�շ���¼ B,ҩƷ��� D,�շ���ĿĿ¼ C,���˷��ü�¼ H,���ű� S,���ű� T,δ��ҩƷ��¼ A" & _
                     " WHERE D.ҩƷID=C.ID AND A.�ⷿID+0=[13] " & IIf(Str���� = "", "", " AND (A.��ҩ���� IN(" & Str���� & ") Or A.��ҩ���� Is NULL)") & _
                     " " & IIf(StrFind_1 = "", " AND A.�������� " & StrDate, StrFind_1) & _
                     " " & strCond & mstrShowBill & _
                     " AND B.����� IS NULL AND LTRIM(RTRIM(NVL(B.ժҪ,'С��')))<>'�ܷ�' " & _
                     " AND H.��������ID=T.ID AND B.ҩƷID=D.ҩƷID AND MOD(B.��¼״̬,3)=1" & _
                     " AND S.ID=B.�ⷿID AND B.����ID=H.ID AND B.NO=A.NO AND B.����=A.���� AND B.�ⷿID+0=[13]) " & _
                     " Group By ����, NO, ����, ����, �Ա�, ����, סԺ��, ����, Ʒ��, ���, ��λ, ����, ������, ��������, ��ҩ�� " & _
                     " ORDER BY ����, NO"
        End If
    End With
    
    Set PrintRec = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        SQLCondition.date��ʼ����, _
        SQLCondition.date��������, _
        SQLCondition.str��ʼNO, _
        SQLCondition.str����NO, _
        SQLCondition.str����, _
        SQLCondition.str���￨, _
        SQLCondition.str��ʶ��, _
        SQLCondition.lng����ID, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        SQLCondition.lngҩƷID, _
        SQLCondition.str��ǰNO, _
        lngҩ��ID)
    
    With PrintRec
        If .EOF Then Exit Sub
        Set MsfPrint.DataSource = PrintRec
    End With
    
    With MsfPrint
        .FixedCols = 0
        For intLeft = 0 To .Cols - 1
            .ColAlignmentFixed(intLeft) = 4
        Next
        
        .ColWidth(0) = 500
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
        .ColWidth(4) = 500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 2500
        .ColWidth(9) = 500
        .ColWidth(10) = 600
        '����
        If MnuEditHandback.Checked Then
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(14) = 1000
            .ColWidth(15) = 1000
            .ColWidth(16) = 1000
            .ColWidth(17) = 1000
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
            .ColAlignment(14) = 7
            .ColAlignment(15) = 7
        Else
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(14) = 1000
            .ColWidth(15) = 1000
            .ColWidth(16) = 1000
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            .ColAlignment(13) = 7
        End If
    End With
    
    ObjAppRow.Add "��ӡ��:" & gstrUserName
    ObjAppRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = "ҩƷ������"
    Set objPrint.Body = MsfPrint
    
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
    '��ҩ����ҩ����ҩ����������
    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then
        MnuEditDosage.Visible = False
        MnuEditAbolish.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then
        mnuEditBill.Visible = False
        MnuEditBatch.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "��ҩ") Then
        mnuEditBillRestore.Visible = False
        If MnuEditBatch.Visible = False Then MnuEdit1.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "��ӡ������ҩ��ϸ") Then
        mnuFileRestore.Visible = False
        MnuFile2.Visible = MnuFileBillprint.Visible
    End If
    If Not IsHavePrivs(mstrPrivs, "��ӡ�ѷ�ҩ�嵥") Then
        mnuFileReport.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "������ҩ���Ĵ���") Then
        MnuEditHandbackBatch.Visible = False
    End If
    If Not IsHavePrivs(mstrPrivs, "������ҩ���Ĵ���") Then
        MnuEditSendOther.Visible = False
    End If
    intVerify = IIf(IsHavePrivs(mstrPrivs, "У�鴦��"), 1, 0)
    If Not IsHavePrivs(strChargePrivs, "����") Then
        mnuCharge.Visible = False
        Tbar1.Buttons("Charge").Visible = False
    End If
    If Trim(strStuffPrivs) = "" Then
        mnuStuff.Visible = False
        Tbar1.Buttons("Stuff").Visible = False
    End If
    If gblnPass And IsHavePrivs(mstrPrivs, "������ҩ���") Then
        mblnStarPass = True
    End If
    
    mbln���������� = IsHavePrivs(mstrPrivs, "����������")
    img����.Visible = mbln����������
    
    mnuCancel.Visible = mbln����ȡ����ҩ And IsHavePrivs(mstrPrivs, "ȡ����ҩ")
    Tbar1.Buttons("Cancel").Visible = mbln����ȡ����ҩ And IsHavePrivs(mstrPrivs, "ȡ����ҩ")
    
    Tbar1.Buttons(9).Visible = (mnuCharge.Visible Or mnuStuff.Visible Or mnuCancel.Visible)
End Sub

Private Sub Txt����ҽ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Txt����ҽ��.ListIndex <> 0 Then Call CmdSend_Click
End Sub

Private Sub Txt����ҽ��_KeyPress(KeyAscii As Integer)
    Dim IntMatchIdx As Integer
    
    With Txt����ҽ��
        IntMatchIdx = MatchIndex(.hWnd, KeyAscii, 1)
        If IntMatchIdx = -2 Then Exit Sub
        .ListIndex = IntMatchIdx
        If .ListIndex = -1 Then .ListIndex = 0
    End With
End Sub

Private Function ת����ҩ��() As String
    Dim strCond1 As String, strCond2 As String, strTemp As String
    Dim intRight As Integer, intLeft As Integer
    '��Ƕ�ײ�ѯ�У�û�����Ӳ��˷��ü�¼���������д��������ֶ�ʱ����ȥ���������������õ����˷��ü�¼��
    strCond1 = ""
    StrFind_4 = UCase(StrFind_4)
    strCond2 = StrFind_4
    intLeft = InStr(1, strCond2, " AND UPPER(H.����)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, StrFind_4, " AND")
        strTemp = Mid(StrFind_4, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = Mid(StrFind_4, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(StrFind_4, intRight)
        Else
            strCond1 = Mid(StrFind_4, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(H.��ʶ��)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    intLeft = InStr(1, strCond2, " AND UPPER(B.���￨��)")
    If intLeft <> 0 Then
        intRight = InStr(intLeft + 4, strCond2, " AND")
        strTemp = Mid(strCond2, 1, intLeft)
        If intRight <> 0 Then
            strCond1 = strCond1 & Mid(strCond2, intLeft, intRight - intLeft + 1)
            strCond2 = strTemp & Mid(strCond2, intRight)
        Else
            strCond1 = strCond1 & Mid(strCond2, intLeft)
            strCond2 = strTemp
        End If
    End If
    ת����ҩ�� = strCond2
End Function

'Modified By ���� 2003-12-10 ����������
Private Sub ShowStock()
    Dim intUnit As Integer
    Dim lngҩƷID As Long, lng���� As Long
    Dim str��λ As String, str��װ As String
    Dim rsStock As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
'    stbThis.Panels(2).Text = ""
    
    If TxtNo.ListIndex < 0 Then Exit Sub
    If Trim(TxtNo.Text) = "" Then Exit Sub
    
    strUnit = GetUnit(lngҩ��ID, TxtNo.ItemData(TxtNo.ListIndex), Mid(TxtNo.Text, 1, 8))
    lngҩƷID = Val(Bill������ϸ.TextMatrix(Bill������ϸ.Row, ����.ҩƷID))
    lng���� = Val(Bill������ϸ.TextMatrix(Bill������ϸ.Row, ����.����))
    
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��λ = "C.���㵥λ"
        str��װ = "/1"
    Case "���ﵥλ"
        str��λ = "B.���ﵥλ"
        str��װ = "/B.�����װ"
    Case "סԺ��λ"
        str��λ = "B.סԺ��λ"
        str��װ = "/B.סԺ��װ"
    Case "ҩ�ⵥλ"
        str��λ = "B.ҩ�ⵥλ"
        str��װ = "/B.ҩ���װ"
    End Select
    
    gstrSQL = " Select A.ʵ������" & str��װ & " ʵ������," & str��λ & " ��λ" & _
             " From ҩƷ��� A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
             " Where A.ҩƷID=B.ҩƷID And B.ҩƷID=C.ID And A.����=1 " & _
             " And A.ҩƷID=[2] And Nvl(A.����,0)=[3] And A.�ⷿID=[1]"
    Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡҩƷ���]", lngҩ��ID, lngҩƷID, lng����)
    
    If rsStock.EOF Then Exit Sub
    
    If Me.ActiveControl Is Bill������ϸ Then
        stbThis.Panels(2).Text = "��ǰ��棺" & FormatEx(rsStock!ʵ������, 5) & rsStock!��λ
    End If
End Sub

Private Sub Get��λ��()
    Const str�ۼ� As String = "C.���㵥λ As �ۼ۵�λ,C.���㵥λ As ��λ,1 As ��װ,ltrim(to_char(B.���ۼ�,'999990.00000')) ����,ltrim(to_char(B.ʵ������,'999990.00000')) ����"
    Const str���� As String = "C.���㵥λ As �ۼ۵�λ,D.���ﵥλ As ��λ,D.�����װ As ��װ,ltrim(to_char(B.���ۼ�*Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.�����װ,Null,1,0,1,D.�����װ),'999990.00000')) ����"
    Const strסԺ As String = "C.���㵥λ As �ۼ۵�λ,D.סԺ��λ As ��λ,D.סԺ��װ As ��װ,ltrim(to_char(B.���ۼ�*Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.סԺ��װ,Null,1,0,1,D.סԺ��װ),'999990.00000')) ����"
    Const strҩ�� As String = "C.���㵥λ As �ۼ۵�λ,D.ҩ�ⵥλ As ��λ,D.ҩ���װ As ��װ,ltrim(to_char(B.���ۼ�*Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����,ltrim(to_char(B.ʵ������/Decode(D.ҩ���װ,Null,1,0,1,D.ҩ���װ),'999990.00000')) ����"
    
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��λ�� = str�ۼ�
    Case "���ﵥλ"
        str��λ�� = str����
    Case "סԺ��λ"
        str��λ�� = strסԺ
    Case "ҩ�ⵥλ"
        str��λ�� = strҩ��
    End Select
End Sub
Private Function BillHaveHerial(ByVal strNo As String, ByVal int���� As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH

    gstrSQL = "Select NO From ���˷��ü�¼ Where NO=[1] And ��¼״̬ IN(0,1,3)" & _
        " And ��¼����=[3] And �շ����='7' And ִ�в���ID+0=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lngҩ��ID, IIf(int���� = 8, 1, 2))
    
    BillHaveHerial = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDateSQL(ByVal strInput As String) As String
    Dim lngStart As Long
    Dim blnDefault As Boolean
    '�ֽ�SQL������ԭ������������
    strInput = Trim(UCase(strInput))
    If strInput = "" Then
        blnDefault = True
    Else
        lngStart = InStr(1, strInput, " AND TO_DATE(")
        If lngStart <> 0 Then
            lngStart = InStr(lngStart + 4, strInput, " AND")
            If lngStart <> 0 Then
                strInput = Mid(strInput, 1, lngStart)
            End If
        Else
            blnDefault = True
        End If
    End If
    If blnDefault Then
        If MnuEditDosage.Checked Then
            GetDateSQL = " And A.�������� " & StrDate
        ElseIf MnuEditAbolish.Checked Then
            GetDateSQL = " And A.�������� " & StrDate
        ElseIf MnuEditConsignment.Checked Then
            GetDateSQL = " And A.�������� " & StrDate
        Else
            GetDateSQL = " And A.������� " & StrDate
        End If
    Else
        GetDateSQL = strInput
    End If
End Function

Private Function ReLocateRow() As Long
    Dim lngRow As Long, lngRows As Long
    On Error GoTo ErrHand
    
    '��λ�ϴ�ѡ��Ĵ�����ʧ�ܷ���1
    lngRows = Msf�б�.Rows - 1
    For lngRow = 1 To lngRows
        If Val(Msf�б�.TextMatrix(lngRow, ��������.����)) = IntLastBill And _
            Msf�б�.TextMatrix(lngRow, ��������.NO) = StrLastNo And _
            Msf�б�.TextMatrix(lngRow, ��������.����) = strLastData Then
            ReLocateRow = lngRow
            Exit Function
        End If
    Next
ErrHand:
    ReLocateRow = IIf(LngSendRow > Msf�б�.Rows - 1 Or LngSendRow = 0, 1, LngSendRow)
End Function

Private Function ReLocateDetailRow() As Long
    Dim lngRow As Long, lngRows As Long
    On Error GoTo ErrHand
    
    '��λ�ϴ�ѡ��Ĵ�����ϸ�б�ʧ�ܷ���1
    lngRows = Bill������ϸ.Rows - 1
    
    If mintLastSequence = 0 Then
        ReLocateDetailRow = lngRows
        Exit Function
    End If
    
    For lngRow = 1 To lngRows
        If Val(Bill������ϸ.TextMatrix(lngRow, ����.���)) = mintLastSequence Then
            ReLocateDetailRow = lngRow
            Exit Function
        End If
    Next
ErrHand:
    ReLocateDetailRow = 1
End Function
Private Function GetDetailCol(ByVal strText As String) As Integer
    Dim intCol As Integer, intCols As Integer
    intCols = Bill������ϸ.Cols - 1
    If strText = "����" Then strText = "����"
    For intCol = 0 To intCols
        If Trim(Bill������ϸ.TextMatrix(0, intCol)) = strText Then
            GetDetailCol = intCol
            Exit Function
        End If
    Next
    GetDetailCol = -1
End Function

Private Function IsDosage(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim int���� As Integer, int��ҩ As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '��鵱ǰ�����Ƿ���Ҫ������ҩ����
    
    If int���� = 0 Then Exit Function
    If strNo = "" Then Exit Function
    
    'ȡ��ǰ�����Ĳ�����Դ
    gstrSQL = " Select �����־ From ���˷��ü�¼ " & _
              " Where ID=(" & _
              "     Select ����ID From ҩƷ�շ���¼ " & _
              "     Where (Nvl(�ⷿID,0)=[3] Or Nvl(�ⷿID,0)=0) And ����=[2] And NO=[1] And Rownum<2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������Դ]", strNo, int����, lngҩ��ID)
    
    int���� = IIf(rsTemp!�����־ = 1 Or rsTemp!�����־ = 4, 1, 2)
    
    '���ݵ�ǰ�����ж��Ƿ���Ҫ��ҩ
    gstrSQL = "Select Nvl(��ҩ,0) AS ��ҩ From ҩ����ҩ���� Where ҩ��ID=[1] And Nvl(����,1)=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[���ݵ�ǰ�����ж��Ƿ���Ҫ��ҩ]", lngҩ��ID, int����)
        
    If rsTemp.RecordCount = 0 Then Exit Function
    
    IsDosage = (rsTemp!��ҩ = 1)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetForeColor_ROW(ByVal lngRow As Long, ByVal lngColor As Long)
    Dim i As Integer, j As Integer
    Dim intCol As Integer, intRow As Integer
    '����ĳ�е���ɫ
    With Bill������ϸ
        BlnEnterCell = False
        intCol = .Col
        intRow = .Row
        .Row = lngRow
        For i = 0 To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            .MsfObj.CellForeColor = lngColor
            .ColData(i) = j
        Next
        .Col = intCol
        .Row = intRow
        BlnEnterCell = True
    End With
End Sub

Private Sub GetBillSequence()
    Dim intRow As Integer, intRows As Integer
    Dim int��� As Integer
    '��ȡ��ǰ����ҩ������ҩ��������Ч���
    str��� = ""
    intRows = Bill������ϸ.Rows - 2
    
    If MnuEditHandback.Checked Then
        '��ҩ����Ϊ���ʾ����Ҫ�˵���ϸ����ͳ�Ƴ�������ϸ�����
        For intRow = 1 To intRows
            If Val(Bill������ϸ.TextMatrix(intRow, ����.��ҩ��)) <> 0 Then
                int��� = Val(Bill������ϸ.TextMatrix(intRow, ����.���))
                If InStr(1, str��� & ",", "," & int��� & ",") = 0 Then
                    str��� = str��� & "," & int���
                End If
            End If
        Next
    Else
        For intRow = 1 To intRows
            int��� = Val(Bill������ϸ.TextMatrix(intRow, ����.���))
            If InStr(1, str��� & ",", "," & int��� & ",") = 0 Then
                str��� = str��� & "," & int���
            End If
        Next
    End If
    If str��� <> "" Then str��� = Mid(str���, 2)
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

