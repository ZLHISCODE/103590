VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPACStation 
   AutoRedraw      =   -1  'True
   Caption         =   "Ӱ����վ"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmPacStation.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLR_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   3330
      MousePointer    =   9  'Size W E
      TabIndex        =   13
      Top             =   750
      Width           =   30
   End
   Begin VB.PictureBox picKind 
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   2955
      TabIndex        =   18
      Top             =   720
      Width           =   3015
      Begin MSComctlLib.ListView lvwPati 
         Height          =   3180
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   5609
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Դ"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "���ݺ�"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "����"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "״̬"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "���˱�ʶ��"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "�ѱ�"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "���ʱ��"
            Object.Width           =   2081
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ִ�м�"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "����ʶ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "�걾��λ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "�����"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "����"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "��"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "��ӡ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "����ʱ��"
            Object.Width           =   2081
         EndProperty
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "����ɵļ��(&3)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   660
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "�����еļ��(&2)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   330
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "��ִ�еļ��(&1)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   3600
      ScaleHeight     =   5535
      ScaleWidth      =   6735
      TabIndex        =   15
      Top             =   1200
      Width           =   6735
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   735
      Top             =   2190
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
            Picture         =   "frmPacStation.frx":066C
            Key             =   "δִ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":0C06
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":11A0
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":173A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":735C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7676
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7AD1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1244
      BandCount       =   2
      _CBWidth        =   10995
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinWidth1       =   4995
      MinHeight1      =   645
      NewRow1         =   0   'False
      Caption2        =   "ҽ������"
      Child2          =   "cboDept"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   3495
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8445
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   2460
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   645
         Left            =   165
         TabIndex        =   12
         Top             =   30
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   34
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Ԥ��"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "��ӡ"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "��¼"
               Key             =   "��¼"
               Description     =   "��¼"
               Object.ToolTipText     =   "��¼ִ�����"
               Object.Tag             =   "��¼"
               ImageKey        =   "��¼"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����ִ�����"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "���"
               Key             =   "���"
               Description     =   "���"
               Object.ToolTipText     =   "ȷ��ִ�����"
               Object.Tag             =   "���"
               ImageKey        =   "���"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Exec_"
               Description     =   "ִ��"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����������"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "�������"
               Object.Tag             =   "����"
               ImageKey        =   "����"
               Style           =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ķ�"
               Key             =   "�ķ�"
               Description     =   "����"
               Object.ToolTipText     =   "�޸ķ���"
               Object.Tag             =   "�ķ�"
               ImageKey        =   "�ķ�"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Description     =   "����"
               Object.ToolTipText     =   "ɾ������"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "ɾ��"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Money_"
               Description     =   "����"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�¿�"
               Key             =   "�¿�"
               Description     =   "ҽ��"
               Object.ToolTipText     =   "�¿�ҽ��"
               Object.Tag             =   "�¿�"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�޸�"
               Description     =   "ҽ��"
               Object.ToolTipText     =   "�޸�ҽ��"
               Object.Tag             =   "�޸�"
               ImageKey        =   "�޸�"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Description     =   "ҽ��"
               Object.ToolTipText     =   "ɾ��ҽ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "ɾ��"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "ҽ��"
               Object.ToolTipText     =   "����ҽ��"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Advice_"
               Description     =   "ҽ��"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "��дһ���µĲ����ļ�"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�����޸�"
               Description     =   "����"
               Object.ToolTipText     =   "�޸Ļ���Ĳ����ļ�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "�޸�"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ����"
               Description     =   "����"
               Object.ToolTipText     =   "ɾ����ǰ�����ļ�"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "ɾ��"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "File_"
               Description     =   "����"
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ƭ"
               Key             =   "��Ƭ"
               Description     =   "Ӱ��"
               Object.ToolTipText     =   "�ڹ�Ƭ����վ�д���ǰѡ���Ӱ������"
               Object.Tag             =   "��Ƭ"
               ImageKey        =   "��Ƭ"
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ɼ�"
               Key             =   "�ɼ�"
               Object.ToolTipText     =   "�ɼ���Ƶͼ��(B����θ����)"
               Object.Tag             =   "�ɼ�"
               ImageKey        =   "Capture"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ʾ"
               Key             =   "��ʾ"
               Description     =   "Ӱ��"
               Object.ToolTipText     =   "��ʾ��ǰ����Ӱ��"
               Object.Tag             =   "��ʾ"
               ImageKey        =   "ViewPic"
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "Ӱ��"
               Object.ToolTipText     =   "ѡ������Ӱ������"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "Ӱ��"
               Object.ToolTipText     =   "�������Ӱ�����е�ѡ���־"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "View_"
               Description     =   "Ӱ��"
               Style           =   3
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��д��鱨��"
               Object.Tag             =   "����"
               ImageKey        =   "Report"
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "���"
               Object.ToolTipText     =   "��˱���"
               Object.Tag             =   "���"
               ImageKey        =   "Auditing"
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "���ص�ǰ��鱨��"
               Object.Tag             =   "����"
               ImageKey        =   "Rollback"
            EndProperty
            BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_Rep"
               Style           =   3
            EndProperty
            BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����¼���˲�ѯ"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7575
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7BB3
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7DCD
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":7FE7
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":8201
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":841B
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":8B15
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":920F
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":9909
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":A003
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":A6FD
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":ADF7
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":B4F1
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":BBEB
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":C2E5
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":C9DF
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":D0D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":D61A
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":DB5B
            Key             =   "ViewPic"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E2D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E4EF
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E709
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":E923
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":EB3D
            Key             =   "Capture"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10847
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10A61
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10C7B
            Key             =   "Rollback"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   8175
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":10E95
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":110AF
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":112C9
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":114E3
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":116FD
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":11DF7
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":124F1
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":12BEB
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":132E5
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":139DF
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":140D9
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":147D3
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":14ECD
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":155C7
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":15CC1
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":163BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":16B35
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":172AF
            Key             =   "ViewPic"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":174C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":176E3
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":178FD
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":17B17
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":17D31
            Key             =   "Capture"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":19A3B
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":19C55
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacStation.frx":19E6F
            Key             =   "Rollback"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraState 
      Height          =   1290
      Left            =   75
      TabIndex        =   14
      Top             =   5580
      Width           =   3165
      Begin VB.CheckBox chkFilter 
         Caption         =   "����ǰ����ɸѡ"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Top             =   1005
         Width           =   1860
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   2
         ItemData        =   "frmPacStation.frx":1A089
         Left            =   1080
         List            =   "frmPacStation.frx":1A09C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   1
         ItemData        =   "frmPacStation.frx":1A0C2
         Left            =   1080
         List            =   "frmPacStation.frx":1A0CF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboState 
         Height          =   300
         Index           =   0
         ItemData        =   "frmPacStation.frx":1A0E7
         Left            =   1080
         List            =   "frmPacStation.frx":1A0F4
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSeek 
         Height          =   360
         Left            =   2400
         Picture         =   "frmPacStation.frx":1A10E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "���˶�λ(F3)"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txt��ʶ�� 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   225
         Width           =   1275
      End
      Begin VB.CheckBox chk״̬ 
         Caption         =   "������δ���ű����ļ��"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CheckBox chk״̬ 
         Caption         =   "�����Ѿ�ִ����ɵļ��"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "״̬(&U)"
         Height          =   200
         Left            =   240
         TabIndex        =   3
         Top             =   650
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��(&N)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   285
         Width           =   810
      End
   End
   Begin MSComctlLib.TabStrip TabFile 
      Height          =   330
      Left            =   3480
      TabIndex        =   16
      Top             =   780
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      TabWidthStyle   =   2
      TabFixedWidth   =   1939
      TabFixedHeight  =   441
      HotTracking     =   -1  'True
      ImageList       =   "iLsTree"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&A)"
            Key             =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ӱ��(&B)"
            Key             =   "Ӱ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ��(&C)"
            Key             =   "ҽ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&D)"
            Key             =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgbLoad 
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   7005
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6945
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPacStation.frx":1A258
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14314
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
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1920
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnufile 
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSetup 
         Caption         =   "��������(&S)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileRoom 
         Caption         =   "ִ�м�����(&R)"
      End
      Begin VB.Menu mnufileImageDevice 
         Caption         =   "Ӱ���豸����(&I)"
      End
      Begin VB.Menu mnufileSendImage 
         Caption         =   "����ͼ��(&T)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuExec 
      Caption         =   "���(&E)"
      Begin VB.Menu mnuExecFunc 
         Caption         =   "�������(&R)"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ȡ������(&Q)"
         Index           =   1
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ȡ������(&E)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "��ʼ���(&A)"
         Index           =   4
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "Ӱ��ɼ�(&V)"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "��Ƭɨ��(&S)"
         Index           =   6
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ȡ�����(&D)"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "����Ӱ��(&S)"
         Index           =   9
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ȡ������(&G)"
         Index           =   10
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "��ȡ�豸Ӱ��(&I)"
         Index           =   11
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ɾ��Ӱ��(&P)"
         Index           =   12
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ȷ�ϼ�����(&F)"
         Index           =   14
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "ȡ��������(&C)"
         Index           =   15
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "��д����(&W)"
         Index           =   17
      End
      Begin VB.Menu mnuExecFunc 
         Caption         =   "������(&U)"
         Index           =   18
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "����(&R)"
      Begin VB.Menu mnuImageView 
         Caption         =   "Ӱ����(&K)"
         Index           =   0
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "Ӱ��Ա�(&B)"
         Index           =   1
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "ѡ����������(&A)"
         Index           =   2
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "���ѡ���־(&M)"
         Index           =   3
      End
      Begin VB.Menu mnuImageView 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "��д����(&W)"
         Index           =   0
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "������д����(&R)"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "������(&C)"
         Index           =   3
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "���沵��(&H)"
         Index           =   4
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "�����ӡ(&P)"
         Index           =   6
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "����Ԥ��(&V)"
         Index           =   7
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "��Ƭ��ӡ(&R)"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepFunc 
         Caption         =   "�����ʽ(&F)"
         Index           =   9
      End
   End
   Begin VB.Menu mnuMoney 
      Caption         =   "����(&M)"
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "����������(&N)"
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "���丽�ӷ���(&A)"
         Index           =   2
         Begin VB.Menu mnuMoneyAdd 
            Caption         =   "�շѵ���(&1)"
            Index           =   0
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuMoneyAdd 
            Caption         =   "���ʵ���(&2)"
            Index           =   1
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuMoneyAdd 
            Caption         =   "��Ѻ��õǼ�(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "�޸ĸ��ӷ���(&M)"
         Index           =   3
      End
      Begin VB.Menu mnuMoneyFunc 
         Caption         =   "ɾ�����ӷ���(&D)"
         Index           =   4
      End
   End
   Begin VB.Menu mnuReq 
      Caption         =   "����(&S)"
      Begin VB.Menu mnuReqFunc 
         Caption         =   "�������뵥(&S)"
         Index           =   0
         Begin VB.Menu ReqList 
            Caption         =   "�޿��õ���"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "�޸����뵥(&G)"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "ɾ�����뵥(&R)"
         Index           =   2
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "��ӡ֪ͨ��(&P)"
         Index           =   4
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "���ı���(&V)"
         Index           =   6
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "Ԥ������(&Y)"
         Index           =   7
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "��ӡ����(&D)"
         Index           =   8
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuReqFunc 
         Caption         =   "Ӱ��Ա�(&B)"
         Index           =   10
      End
   End
   Begin VB.Menu mnuPFile 
      Caption         =   "����(&L)"
      Begin VB.Menu mnuPFileFunc 
         Caption         =   "��������(&A)"
         Index           =   0
         Begin VB.Menu FileList 
            Caption         =   "�޲����ļ�"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuPFileFunc 
         Caption         =   "�޸Ĳ���(&M)"
         Index           =   1
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuPFileFunc 
         Caption         =   "ɾ������(&D)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuAdvice 
      Caption         =   "ҽ��(&Y)"
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "�¿�ҽ��(&A)"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "�޸�ҽ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "ɾ��ҽ��(&D)"
         Index           =   2
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "ҽ��ֹͣ(&S)"
         Index           =   4
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "ҽ������(&R)"
         Index           =   5
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "����ҽ��(&S)"
         Index           =   6
      End
      Begin VB.Menu mnuAdviceFunc 
         Caption         =   "����ҽ��(&R)"
         Index           =   7
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "����(&T)"
      Begin VB.Menu mnuToolItemRef 
         Caption         =   "���Ʋο�(&I)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuToolDiagRef 
         Caption         =   "��ϲο�(&D)"
      End
      Begin VB.Menu mnuTool_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolReport 
         Caption         =   "��鹤������(&1)"
         Index           =   0
      End
      Begin VB.Menu mnuReport 
         Caption         =   "��������"
         Begin VB.Menu mnuReportItem 
            Caption         =   "��"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "����ѡ��(&D)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "������Ϣ(&I)"
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCharge 
         Caption         =   "ֻ��ʾ�Ѿ��շѵĲ���(&P)"
      End
      Begin VB.Menu mnuViewAdviceSelf 
         Caption         =   "ֻ��ʾ�����´��ҽ��(&O)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewFileSelf 
         Caption         =   "ֻ��ʾ������д�Ĳ���(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "��ʾ������ʷ����(&H)"
      End
      Begin VB.Menu mnuViewAdviceAppend 
         Caption         =   "��ʾ������ϸ(&D)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewPic 
         Caption         =   "��ʾ��ǰ����ͼ��(&V)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "���ݹ���(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "��λ��ʽ"
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "��ʶ��(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "���￨(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "����(&3)"
            Index           =   2
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "���ݺ�(&4)"
            Index           =   3
         End
         Begin VB.Menu mnuViewFindItem 
            Caption         =   "����(&5)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuView_3 
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
         Caption         =   "&WEB�ϵ�����"
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
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmPACStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mstrFilter As String
Private mstrPrivs As String
Private mlngPreDept As Long
Private mstrPrePati As String
Private TabIndex As Integer
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mstrRoom As String, blnIfOnlyShow As Boolean '��ǰִ�м�

Private mfrmActive As Object '��ǰ�����
Private WithEvents mfrmRepEdit As Form
Attribute mfrmRepEdit.VB_VarHelpID = -1
Private aForms(4) As Object
Private objImgCapture As Object
Private mBeforeDays As Integer 'Ĭ�ϲ�ѯ������
Private mDispImgs As Integer '����ͼ��ʾ��
Private mblnEmergencyPrint As Boolean   'True=�������ʱ�������ֻ���Ѵ�ӡ��ǣ�False=�������ʱ�����������
'������������
Private mdatFBegin As Date
Private mdatFEnd As Date
Private mDatType As Integer '1=�����ʱ�䡢2=������ʱ��
Private mstrFNO As String
Private mlngF����ID As Long
Private mstrF��Դ As String
Private mdblF��ʶ�� As Double
Private mstrF���￨ As String
Private mstrF���� As String
Private mdblFChkNO As Double
Private mblnViewImage As Boolean '����ʱ�Ƿ��Ƭ
Private mblnSample As Boolean '�ǼǺ��Ƿ�ֱ�Ӻ���
Private mstr�걾��λ As String  '���걾��λ
Private mstrPatiName As String '��ǰɸѡ����

Private Sub cboState_Click(Index As Integer)
    If Me.Tag <> "" Then Me.Tag = "": Exit Sub
    
    Call LoadPatiList
End Sub

Private Sub chkFilter_Click()
    If Me.lvwPati.SelectedItem Is Nothing Then
        Me.chkFilter.Value = 0
    End If
    If Me.chkFilter.Value = 1 Then mstrPatiName = Me.lvwPati.SelectedItem.SubItems(2)
    Call LoadPatiList
End Sub

Private Sub chk״̬_Click(Index As Integer)
    If Me.Tag <> "" Then Me.Tag = "": Exit Sub
    
    Call LoadPatiList
End Sub

Private Sub cmdSeek_Click()
    Call Form_KeyDown(vbKeyF3, 0)
End Sub

Private Sub cmdKind_Click(Index As Integer)
    'װ���ݲ���������
    If Val(lvwPati.Tag) <> Index Then
        Me.lvwPati.Tag = Index
        Call picKind_Resize
        Call LoadPatiList
        '��λ������ҵĲ��˼��
        If txt��ʶ��.Text <> "" Then Call SeekNextPati(True)
    End If
    If Me.lvwPati.Visible Then
        Me.lvwPati.SetFocus
    End If
    ShowCheck Index
End Sub

Private Sub ShowCheck(ByVal Index As Integer)
    Dim intCount As Integer
    On Error Resume Next
'    With chk״̬
'        For intCount = .LBound To .UBound
'            .Item(intCount).Visible = False
'        Next
'        .Item(Index).Visible = True
'    End With
    With cboState
        For intCount = .LBound To .UBound
            .Item(intCount).Visible = False
        Next
        .Item(Index).Visible = True
    End With
End Sub

Private Sub FileList_Click(Index As Integer)
    mfrmActive.zlMenuClick FileList(Index)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "Loading" Then
        Me.Tag = ""
        TabFile.Tabs(TabIndex).Selected = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnFirst As Boolean
    If KeyCode = vbKeyF3 Then
        If txt��ʶ��.Text = "" Then
            txt��ʶ��.SetFocus
        Else
            Call txt��ʶ��_Validate(False)
            Call zlControl.TxtSelAll(txt��ʶ��)
            Call SeekNextPati(txt��ʶ��.Tag <> txt��ʶ��.Text)
        End If
    ElseIf KeyCode = vbKeyF4 Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean)
    Dim intB As Integer, blnDo As Boolean
    Dim strItem As String, strFind As String, i As Long
    
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    intB = 1
    If Not blnFirst Then
        intB = lvwPati.SelectedItem.Index + 1
        If intB > lvwPati.ListItems.Count Then intB = 1
    End If
    Do While True
        For i = intB To lvwPati.ListItems.Count
            blnDo = False
            If txt��ʶ��.Text <> "" Then
                strItem = Split(Label1.Caption, "(")(0)
                With lvwPati.ListItems(i)
                    If strItem = "��ʶ��" Then
                        strFind = .SubItems(6)
                    ElseIf strItem = "���￨" Then
                        strFind = .ListSubItems(10).Tag
                    ElseIf strItem = "����" Then
                        strFind = .SubItems(2)
                        If strFind Like txt��ʶ��.Text & "*" Then blnDo = True
                        If zlCommFun.SpellCode(strFind) Like UCase(txt��ʶ��.Text) & "*" Then blnDo = True
                    ElseIf strItem = "���ݺ�" Then
                        strFind = .SubItems(1)
                    ElseIf strItem = "����" Then
                        strFind = .SubItems(14)
                    End If
                    If strFind = txt��ʶ��.Text Then blnDo = True
                End With
            End If
            If blnDo Then
                txt��ʶ��.Tag = txt��ʶ��.Text
                If lvwPati.SelectedItem.Key <> lvwPati.ListItems(i).Key Then
                    lvwPati.ListItems(i).Selected = True
                    Call lvwPati_ItemClick(lvwPati.SelectedItem)
                    lvwPati.SelectedItem.EnsureVisible
                End If
                Exit Sub
            End If
        Next
        If Not blnFirst And intB > 1 Then
            intB = 1
        Else
            Exit Sub
        End If
    Loop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Label1_Click()
    Me.PopupMenu Me.mnuViewFind
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If Not lvwPati.SelectedItem Is Nothing And mnuRepFunc(0).Visible Then mnuRepFunc_Click 0
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key = mstrPrePati Then Exit Sub
    mstrPrePati = Item.Key
    Call tabFile_Click
End Sub

Private Sub lvwPati_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If lvwPati.Tag = "2" And mnuRep.Enabled Then
            If InStr(mstrPrivs, "��д����") > 0 Or InStr(mstrPrivs, "�������") > 0 Then Me.PopupMenu mnuRep
        ElseIf mnuExec.Enabled Then
            If InStr(mstrPrivs, "Ӱ����") > 0 Then Me.PopupMenu mnuExec
        End If
    End If
End Sub

Private Sub mfrmRepEdit_Unload(Cancel As Integer)
    Dim objPacsCore As Object
    
    Call LoadPatiList
    Set mfrmRepEdit = Nothing
    '�رչ�Ƭվ����
    If mblnViewImage Then
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.Closefrom
    End If
End Sub

Private Sub mnufileImageDevice_Click()
    frmPACSImageDeviceSetup.Show vbModal, Me
End Sub

Private Sub mnufileSendImage_Click()
    frmPacsSendImage.ShowMe Me
End Sub

Private Sub mnuImageView_Click(Index As Integer)
    mfrmActive.zlMenuClick mnuImageView(Index)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    With lvwPati.SelectedItem
        lngҽ��ID = Val(Split(Mid(.Key, 2), "_")(0))
        lng���ͺ� = Val(Split(Mid(.Key, 2), "_")(1))
    End With
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
         "ҽ��ID=" & lngҽ��ID, "���ͺ�=" & lng���ͺ�)
End Sub

Private Sub mnuViewFindItem_Click(Index As Integer)
    Dim strItem As String, i As Long
    
    For i = 0 To mnuViewFindItem.UBound
        mnuViewFindItem(i).Checked = i = Index
    Next
    strItem = Split(mnuViewFindItem(Index).Caption, "(")(0)
    Label1.Caption = strItem & "(&D)"
    If strItem = "���￨" And gblnCardHide Then
        txt��ʶ��.PasswordChar = "*"
    Else
        txt��ʶ��.PasswordChar = ""
    End If
    txt��ʶ��.Text = "": txt��ʶ��.Tag = ""
    If Visible Then txt��ʶ��.SetFocus
End Sub

Private Sub mnuViewInfo_Click()
    Dim lng����id As Long
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    lng����id = Val(Split(lvwPati.SelectedItem.Tag, "_")(0))
    Call frmDegreeCard.ShowInfo(Me, lng����id)
End Sub

Private Sub mnuViewPic_Click()
'    mnuViewPic.Checked = Not mnuViewPic.Checked
    mfrmActive.zlMenuClick mnuViewPic
End Sub

Private Sub picKind_Resize()
    Dim intCount As Integer
    On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picKind.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picKind.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If intCount <= Val(lvwPati.Tag) Then
            Me.cmdKind(intCount).Top = Me.picKind.ScaleTop + 285 * intCount
            Me.lvwPati.Top = Me.picKind.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picKind.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.lvwPati.Left = Me.picKind.ScaleLeft + 15
    Me.lvwPati.Width = Me.picKind.ScaleWidth
    Me.lvwPati.Height = Me.picKind.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Function Getִ������(ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal str��� As String) As String
'���ܣ�����ָ����ҽ��ID,����ҽ�����ݹ���ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim bln��ҩ;�� As Boolean, strTmp As String
    
    On Error GoTo errH
    
    '��ȡҽ������
    If str��� <> "E" Or lng���ID <> 0 Then
        '�䷽�巨,�������������ҽ��,ֱ����ʾҽ������
        strSQL = "Select ҽ������ From ����ҽ����¼ Where ID= " & IIf(str��� = "E", "[1]", "[2]")
        
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng���ID, lngҽ��ID)
            
        If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!ҽ������)
    Else
        strSQL = "Select A.ID,A.���ID,A.�������,A.ҽ������,A.ִ��Ƶ��,A.ִ��ʱ�䷽��,B.����" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B" & _
            " Where Not (A.�������='E' And ���ID is Not NULL) And A.������ĿID=B.ID" & _
            " And (A.���ID= [1] Or A.ID= [1] )" & _
            " Order by A.���"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
        rsTmp.Filter = "���ID=" & lngҽ��ID
        If Not rsTmp.EOF Then bln��ҩ;�� = InStr(",5,6,", rsTmp!�������) > 0
        
        If Not bln��ҩ;�� Then
            'һ��������Ŀ����ҩ�÷�
            rsTmp.Filter = ""
            strSQL = "Select ҽ������ From ����ҽ����¼ Where ID=[1] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then strTmp = Nvl(rsTmp!ҽ������)
        Else
            '��ҩ;��
            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & "," & rsTmp!ҽ������
                rsTmp.MoveNext
            Next
            rsTmp.Filter = "ID=" & lngҽ��ID
            strTmp = rsTmp!���� & "," & rsTmp!ִ��Ƶ�� & "(" & rsTmp!ִ��ʱ�䷽�� & "):" & Mid(strTmp, 2)
        End If
    End If
    
    '��ȡ��������
    strSQL = "Select A.��������,C.���㵥λ" & _
        " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C" & _
        " Where A.ҽ��ID= [1] And A.���ͺ�= [2] " & _
        " And A.ҽ��ID=B.ID And B.������ĿID=C.ID"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    If IsNull(rsTmp!��������) Then
        Getִ������ = "��ִ������:" & strTmp
    Else
        Getִ������ = "����������:" & FormatEx(rsTmp!��������, 5) & " " & Nvl(rsTmp!���㵥λ) & ",ִ������:" & strTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuAdviceFunc_Click(Index As Integer)
    If mfrmActive Is Nothing Then Exit Sub
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    If Val(lvwPati.SelectedItem.ListSubItems(3).Tag) = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
        MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mfrmActive.zlMenuClick mnuAdviceFunc(Index)
End Sub

Private Sub mnuExecFunc_Click(Index As Integer)
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, rsTmp As New ADODB.Recordset, rsSel As ADODB.Recordset
    Dim iCurrItemIndex As Integer
    
    Dim strImageType As String, strCheckUID As String
    Dim iReturn As Integer, blnCancel As Boolean
'    Dim inte As New clsFtp
    Dim strImageDeviceNumber As String                              '�豸��
    Dim strFilter As String                     'ȡ������ʱ������ѡ������ַ���
    
    On Error GoTo DBError
    If Me.lvwPati.SelectedItem Is Nothing And Index > 0 Then Exit Sub
    If Not Me.lvwPati.SelectedItem Is Nothing Then
        With lvwPati.SelectedItem
            lngҽ��ID = Val(Split(Mid(.Key, 2), "_")(0))
            lng���ͺ� = Val(Split(Mid(.Key, 2), "_")(1))
        End With
    End If
    Select Case Index
        Case 0 '����ԤԼ
'             If RequestRegister(Me, Me.cboDept.ItemData(Me.cboDept.ListIndex)) Then
             If frmPACSReqEdit.ShowMe_Request(Me, Me.cboDept.ItemData(Me.cboDept.ListIndex), blnCheck:=mblnSample) Then
                If mblnSample Then
                    lvwPati.Tag = 1: picKind_Resize
                    Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
                    ShowCheck 1
                Else
                    lvwPati.Tag = 0: picKind_Resize
                    Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
                    ShowCheck 0
                End If
             End If
        Case 1 'ȡ������
            '����ִ�л���ִ�в�����ܾ�
            If Val(lvwPati.SelectedItem.ListSubItems(3).Tag) = 3 Then
                MsgBox "��������Ŀ��ǰ����ִ�У�����ȡ����", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(lvwPati.SelectedItem.ListSubItems(3).Tag) = 1 Then
                MsgBox "��������Ŀ��ǰ�Ѿ�ִ�У�����ȡ����", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & lngҽ��ID & "," & lng���ͺ� & ")"
            ExecuteProc strSQL, Me.Caption
            '�����µĵ�ǰ����
            iCurrItemIndex = lvwPati.SelectedItem.Index
            If iCurrItemIndex < lvwPati.ListItems.Count Then
                lvwPati.ListItems(iCurrItemIndex + 1).Selected = True
            ElseIf lvwPati.ListItems.Count > 1 Then
                lvwPati.ListItems(lvwPati.ListItems.Count - 1).Selected = True
            End If
            
            Call LoadPatiList
            If Not lvwPati.SelectedItem Is Nothing Then
                lvwPati.SelectedItem.EnsureVisible
            End If
        Case 4 '��ʼ���
            '�ж�ִ��״̬
            If InStr("1", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) > 0 Then
                MsgBox "�ü�鲻�������¿�ʼ��", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
'            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag = 3 And _
'                Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) > 1 Then
'                MsgBox "�ü�����ڽ��У����������¿�ʼ��", vbInformation, gstrSysName
'                Exit Sub
'            End If
            
            iReturn = frmPACSReg.ShowMe(Me, lngҽ��ID, lng���ͺ�)
            Select Case iReturn
                Case 1
                    lvwPati.Tag = 1: picKind_Resize
                    Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
                    ShowCheck 1
                Case 2
                    Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
            End Select
        Case 5 '�ɼ�
            '�ж�ִ��״̬
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strSQL = "Select Ӱ�����,���UID From Ӱ�����¼ Where ҽ��ID= [1] And ���ͺ�= [2] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
            If rsTmp.EOF Then
                MsgBox "�ü��δ������ʼ����ȡ�������¿�ʼ��", vbInformation, gstrSysName
                Exit Sub
            End If
            strImageType = Nvl(rsTmp(0)): strCheckUID = Nvl(rsTmp(1))
            
            With lvwPati.SelectedItem
                objImgCapture.ImageCapture mstrPrivs, lngҽ��ID, lng���ͺ�, Me, .SubItems(1), CInt(.ListSubItems(5).Tag), _
                 CLng(Split(.ListSubItems(8).Tag, "|")(0)), CLng(Split(.ListSubItems(8).Tag, "|")(1)), "", _
                 strImageType, strCheckUID
            End With
            
            
'            With lvwPati.SelectedItem
'                EditReport Me, .SubItems(1), CInt(.ListSubItems(5).Tag), _
'                    CLng(Split(.ListSubItems(8).Tag, "|")(0)), CLng(Split(.ListSubItems(8).Tag, "|")(1)), "", _
'                    Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 6, False, tmpObject, , _
'                    Not (InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
'                InStr("3,4,5,6", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0), , lngҽ��ID, Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1"
'                Set mfrmRepEdit = tmpObject
'            End With
            
            Call tabFile_Click
            Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
            On Error GoTo DBError
        Case 6 'ɨ��
            '�ж�ִ��״̬
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strSQL = "Select Ӱ�����,���UID From Ӱ�����¼ Where ҽ��ID= [1] And ���ͺ�= [2] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
            If rsTmp.EOF Then
                MsgBox "�ü��δ������ʼ����ȡ�������¿�ʼ��", vbInformation, gstrSysName
                Exit Sub
            End If
            strImageType = Nvl(rsTmp(0)): strCheckUID = Nvl(rsTmp(1))
            On Error Resume Next
            objImgCapture.ImageScan lngҽ��ID, lng���ͺ�, strImageType, strCheckUID
            
            Call tabFile_Click
            On Error GoTo DBError
        Case 7 'ȡ�����
            '�ж�ִ��״̬
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("ȡ�����μ�齫ɾ����Ӧ�ļ�����м���ͼ���Ƿ������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            'ɾ��Ӱ���ļ���Ŀ¼
            RemoveCheckImages lngҽ��ID, lng���ͺ�
            strSQL = "ZL_Ӱ����_CANCEL(" & lngҽ��ID & "," & lng���ͺ� & ")"
            ExecuteProc strSQL, Me.Caption
            
            lvwPati.Tag = 0: picKind_Resize
            Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
            ShowCheck 0
        Case 9 '����ͼ��
'            If InStr("0,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
'                Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) > 2 Then
'                MsgBox "��ǰ�������ɣ�", vbInformation, gstrSysName
'                Exit Sub
'            End If
            If InStr("0,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
                InStr("1,2", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0 Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
'            strSQL = "Select Count(*) From Ӱ�����¼ A,Ӱ�������� B" & _
'                " Where A.���UID=B.���UID And A.ҽ��ID=" & lngҽ��ID & " And ���ͺ�=" & lng���ͺ�
'            OpenRecord rsTmp, strSQL, Me.Caption
'            If rsTmp(0) > 0 Then
'                MsgBox "��ǰ�����Ŀ���м��ͼ�񣬲�����ȡ��������¼��", vbInformation, gstrSysName
'                Exit Sub
'            End If
            strSQL = "Select A.���UID As ID,Nvl(A.����豸,' ') As ����豸,Nvl(A.����,0) As ����,Nvl(A.��������,Sysdate) As ���ʱ��," & _
                "Nvl(A.����,' ') As ����,Nvl(A.Ӣ����,' ') As Ӣ����,Nvl(A.�Ա�,' ') As �Ա�,Nvl(A.����,' ') As ����," & _
                "Nvl(A.��������,Sysdate) As ��������," & _
                "Nvl(A.���,0) As ���,Nvl(A.����,0) As ����" & _
                " From Ӱ����ʱ��¼ a,����ҽ����¼ b,Ӱ�����¼ c,������Ϣ d" & _
                " Where c.ҽ��ID=b.ID And b.����ID=d.����ID" & _
                " And (a.����=c.���� Or a.����=c.ҽ��ID Or a.����=Decode(b.������Դ,2,d.סԺ��,d.�����))" & _
                " And c.ҽ��ID=" & lngҽ��ID & " And c.���ͺ�=" & lng���ͺ�
                
            Set rsTmp = OpenSQLRecord(strSQL, "�Զ���Ӧ��Ŀ", lngҽ��ID, lng���ͺ�)
'''            If rsTmp.State <> adStateClosed Then rsTmp.Close
'''            Set rsSel = zlDatabase.ShowSelect(Me, strSQL, 0, "���Ӱ��", blnNoneWin:=False, Cancel:=blnCancel, blnSearch:=True)
'''            If rsSel Is Nothing And Not blnCancel Then

            'û�з���ɸѡ�����ļ�¼����ѯȫ��
'                strSQL = "Select A.���UID As ID,Nvl(A.����豸,' ') As ����豸,Nvl(A.����,0) As ����,Nvl(A.��������,Sysdate) As ���ʱ��," & _
'                    "Nvl(A.����,' ') As ����,Nvl(A.Ӣ����,' ') As Ӣ����,Nvl(A.�Ա�,' ') As �Ա�,Nvl(A.����,' ') As ����," & _
'                    "Nvl(A.��������,Sysdate) As ��������," & _
'                    "Nvl(A.���,0) As ���,Nvl(A.����,0) As ����" & _
'                    " From Ӱ����ʱ��¼ a,����ҽ����¼ b,Ӱ������Ŀ c" & _
'                    " Where a.Ӱ�����=c.Ӱ����� And b.������Ŀid=c.������Ŀid And b.id= " & lngҽ��ID

                strSQL = "Select A.���UID As ID,Nvl(A.����,0) As ����,Nvl(A.����,' ') As ����,Nvl(A.����豸,' ') As ����豸,Nvl(A.��������,Sysdate) As ���ʱ��," & _
                    "Nvl(A.Ӣ����,' ') As Ӣ����,Nvl(A.�Ա�,' ') As �Ա�,Nvl(A.����,' ') As ����," & _
                    "Nvl(A.��������,Sysdate) As ��������," & _
                    "Nvl(A.���,0) As ���,Nvl(A.����,0) As ����" & _
                    " From Ӱ����ʱ��¼ a,Ӱ�����¼ b" & _
                    " Where a.Ӱ�����=b.Ӱ����� And b.ҽ��id= " & lngҽ��ID & " And b.���ͺ�=" & lng���ͺ� & " order by A.����"
                    
                Set rsSel = zlDatabase.ShowSelect(Me, strSQL, 0, "���Ӱ��", False, IIf(rsTmp.EOF, "", rsTmp!����), "", , , False, , _
                                                    picKind.Top + lvwPati.Top * 3 + lvwPati.SelectedItem.Height, , _
                                                    blnCancel, , True)
                                                    
'''                If rsTmp.State <> adStateClosed Then rsTmp.Close
'''                Set rsSel = zlDatabase.ShowSelect(Me, strSQL, 0, "���Ӱ��", blnNoneWin:=False, blnSearch:=True)
'''            End If

            If Not rsSel Is Nothing Then
                If MsgBox("�Ƿ�ȷ��ѡ���Ӱ���ǵ�ǰ���ģ�", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                '�ƶ�Ftp�ϵ�Ӱ���ļ�
                strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1] And ���ͺ�=[2]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
                If Not rsTmp.EOF Then
                    If Len(Trim(Nvl(rsTmp(0)))) > 0 Then Call MergeImageFiles(rsSel("ID"), rsTmp(0))
                End If
                
                strSQL = "ZL_Ӱ����_SET(" & lngҽ��ID & "," & lng���ͺ� & ",'" & _
                    rsSel("ID") & "')"
                ExecuteProc strSQL, Me.Caption
                
                lvwPati.Tag = 1: picKind_Resize
                Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
                ShowCheck 1
            End If
        Case 10   'ȡ������
            'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
            
            If InStr("0,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
                InStr("1,2", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0 Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '��ʾ����ѡ�񴰿�
            strSQL = "select 0 as ѡ��,B.����UID as ID ,B.���к�,B.��������,SUM(1) AS ͼ���� from Ӱ�����¼ A ," & _
                    "Ӱ�������� B, Ӱ����ͼ�� C Where a.���UID = B.���UID And B.����UID = C.����UID" & _
                    " And a.ҽ��ID = " & lngҽ��ID & " and A.���ͺ�= " & lng���ͺ� & " group by B.����UID,B.���к�,B.��������"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
            
            frmSelectMuli.ShowSelect Me, rsTmp, "ID,3000,0,1;���к�,800,0,1;��������,2000,0,1;ͼ����,800,0,1", 0, 0, 7000, 5000
            
            If frmSelectMuli.mblnOK = True Then
                strFilter = frmSelectMuli.strFilter
                rsTmp.Filter = strFilter
                '�����ѡ�����У�����ÿһ�����е�ȡ��
                While Not rsTmp.EOF
                    subCancelSeriesRelate lngҽ��ID, lng���ͺ�, rsTmp!ID
                    rsTmp.MoveNext
                Wend
                
                '����װ�ز��˼�¼
                lvwPati.Tag = 1: picKind_Resize
                Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
                ShowCheck 1
            End If
        Case 11  '���豸ֱ����ȡͼ��
            strImageDeviceNumber = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
            
            'û��Ĭ���豸ʱ����
            If strImageDeviceNumber = "" Then
                If MsgBox("û������Ĭ��Ӱ�����豸���Ƿ��������ã�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    frmPACSImageDeviceSetup.Show vbModal, Me
                    Exit Sub
                End If
            End If
            
            strSQL = "select �豸�� , �豸��, IP��ַ,�˿ں�,����AE,�豸AE from Ӱ���豸Ŀ¼ where �豸�� = [1] "
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Mid(strImageDeviceNumber, 2))
            
            '��Ĭ���豸��ɾ������������
            If rsTmp.EOF = True Then
                MsgBox "Ĭ���豸�ѱ�ɾ�������������ã�", vbInformation, gstrSysName
                frmPACSImageDeviceSetup.Show vbModal, Me
                Exit Sub
            End If
                
            frmPACSGetDeviceImage.ShowMe Me, rsTmp("IP��ַ"), rsTmp("�˿ں�"), rsTmp("�豸��"), Nvl(rsTmp("����AE")), Nvl(rsTmp("�豸AE")), lngҽ��ID
            Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
        
        Case 12 'ɾ�����ͼ��
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            strSQL = "select ���UID from Ӱ�����¼ where ҽ��ID = " & lngҽ��ID & " and  ���ͺ� = [1]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lng���ͺ�)
            If rsTmp.RecordCount = 0 Then Exit Sub
            If MsgBox("�Ƿ�ȷ��Ҫɾ���ü�������Ӱ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            'ɾ��Ӱ���ļ���Ŀ¼
            RemoveCheckImages lngҽ��ID, lng���ͺ�
            strSQL = "ZL_Ӱ����_PhotoDelete(" & lngҽ��ID & "," & lng���ͺ� & ")"
            ExecuteProc strSQL, Me.Caption
            
            Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
        Case 14 '������
            '�ж�ִ��״̬
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Or _
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "2" Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("ȷ�ϸ������������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            strSQL = "ZL_Ӱ����_STATE(" & lngҽ��ID & "," & lng���ͺ� & ",3)"
            ExecuteProc strSQL, Me.Caption
            
            lvwPati.Tag = 2: picKind_Resize
            Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
            ShowCheck 2
        Case 15 'ȡ��������
            '�ж�ִ��״̬
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Then
                MsgBox "��ǰδ���иü�飡", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) < 3 Then
                MsgBox "��ǰ���δ��ɣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("ȷ�ϼ������и�������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            strSQL = "ZL_Ӱ����_STATE(" & lngҽ��ID & "," & lng���ͺ� & ",2)"
            ExecuteProc strSQL, Me.Caption
            
            lvwPati.Tag = 1: picKind_Resize
            Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
            ShowCheck 1
        Case 17 '��д����
            mnuRepFunc_Click 0
        Case 18 '������
            mnuRepFunc_Click 3
    End Select
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RemoveCheckImages(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long)
    'ɾ��ָ��ҽ���ļ��Ӱ��
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    Dim Inte As New clsFtp
    Dim strDeviceNO As String
    On Error GoTo ProcError
    '��ɾ��ͼ��
    strSQL = "select a.IP��ַ, a.FTPĿ¼, a.�û���, a.����, a.ҽ��ID, a.���ͺ�, a.���UID, a.λ��, a.�������� ,a.�豸�� ,c.ͼ��UID" & _
             " from (select IP��ַ, FTPĿ¼, �û���, ����, ҽ��ID, ���ͺ�, ���UID, λ��һ as λ��, ��������, a.�豸�� " & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ��һ " & _
             "       Union All " & _
             "       select IP��ַ, FTPĿ¼, �û���, ����, ҽ��ID, ���ͺ�, ���UID, λ�ö� as λ��, ��������, a.�豸��" & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ�ö� " & _
             "       Union All " & _
             "       select IP��ַ, FTPĿ¼, �û���, ����, ҽ��ID, ���ͺ�, ���UID, λ���� as λ��, ��������, a.�豸�� " & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ���� " & _
             "       ) a , Ӱ�������� b , Ӱ����ͼ�� c " & _
             " Where a.���uid = B.���uid " & _
             " and b.����uid = c.����uid " & _
             " and a.ҽ��ID = [1] And ���ͺ� = [2] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    Do Until rsTmp.EOF
'        inte.strIPAddress = rsTmp("IP��ַ")
'        inte.strUser = IIf(IsNull(rsTmp("�û���")), "", rsTmp("�û���"))
'        inte.strPsw = IIf(IsNull(rsTmp("����")), "", rsTmp("����"))
        If strDeviceNO <> rsTmp("�豸��") Then
            strDeviceNO = rsTmp("�豸��")
            Inte.FuncFtpConnect rsTmp("IP��ַ"), rsTmp("�û���"), rsTmp("����")
        End If
        Inte.FuncDelFile IIf(IsNull(rsTmp("FTPĿ¼")), "", rsTmp("FTPĿ¼") & "/") & Format(rsTmp("��������"), "YYYYMMDD") & "/" & rsTmp("���UID"), rsTmp("ͼ��UID")
        rsTmp.MoveNext
    Loop
    strDeviceNO = ""
    Inte.FuncFtpDisConnect
    'ɾ��Ŀ¼
    strSQL = "select IP��ַ,FTPĿ¼,�û���,����,ҽ��ID,���ͺ�,���UID,�豸��,λ��,�������� from " & _
             "      (select IP��ַ,FTPĿ¼,�û���,����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ��һ as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      Where a.�豸�� = B.λ��һ " & _
             "      Union All " & _
             "      select IP��ַ,FTPĿ¼,�û���,����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ�ö� as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      Where a.�豸�� = B.λ�ö� " & _
             "      Union All " & _
             "      select IP��ַ,FTPĿ¼,�û���,����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ���� as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      where a.�豸�� = b.λ���� ) a " & _
             " Where a.ҽ��ID = [1] And ���ͺ� = [2] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    Do Until rsTmp.EOF
'        Inte.strIPAddress = rsTmp("IP��ַ")
'        Inte.strUser = IIf(IsNull(rsTmp("�û���")), "", rsTmp("�û���"))
'        Inte.strPsw = IIf(IsNull(rsTmp("����")), "", rsTmp("����"))
        If strDeviceNO <> rsTmp("�豸��") Then
            strDeviceNO = rsTmp("�豸��")
            Inte.FuncFtpConnect rsTmp("IP��ַ"), rsTmp("�û���"), rsTmp("����")
        End If
        Inte.FuncFtpDelDir IIf(IsNull(rsTmp("FTPĿ¼")), "", rsTmp("FTPĿ¼")), Format(rsTmp("��������"), "YYYYMMDD") & "/" & rsTmp("���UID")
        rsTmp.MoveNext
    Loop
    Inte.FuncFtpDisConnect
    Exit Sub
ProcError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuFileRoom_Click()
    Dim strTmp As String, blnTmp As Boolean
    
    On Error Resume Next
    strTmp = mstrRoom: blnTmp = blnIfOnlyShow
    If frmPACSRoom.ShowMe(Me, strTmp, blnTmp, cboDept.ItemData(cboDept.ListIndex)) Then
        mstrRoom = strTmp: blnIfOnlyShow = blnTmp
        Call LoadPatiList
    End If
End Sub

Private Sub mnuMoneyAdd_Click(Index As Integer)
    If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
        MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
        Exit Sub
    End If
    mfrmActive.zlMenuClick mnuMoneyAdd(Index)
End Sub

Private Sub mnuMoneyFunc_Click(Index As Integer)
    If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
        MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
        Exit Sub
    End If
    mfrmActive.zlMenuClick mnuMoneyFunc(Index)
End Sub

Private Sub mnuPFileFunc_Click(Index As Integer)
    mfrmActive.zlMenuClick mnuPFileFunc(Index)
End Sub

Private Sub mnuRepFunc_Click(Index As Integer)
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strRptName As String
    Dim tmpObject As Object
    Dim iMsgReturn As Integer
    Dim strAudiName As String '������
    Dim blnEmerge As Boolean '����
    Dim strIfAuditing As String, strIfRollback As String
    Dim strIfPrint As String
    
    On Error GoTo DBError
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    With lvwPati.SelectedItem
        lngҽ��ID = Val(Split(Mid(.Key, 2), "_")(0))
        lng���ͺ� = Val(Split(Mid(.Key, 2), "_")(1))
    End With
    
    Select Case Index
        Case 0 '��д����
            '�ж�ִ��״̬
'            If InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
'                InStr("3,4,5,6", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0 Then
'                MsgBox "���������ڲ�����д���棡", vbInformation, gstrSysName
'                Exit Sub
'            End If
            'ˢ�±�����¼
            strSQL = "Select ִ��״̬,Nvl(ִ�й���,0) As ִ�й��� From ����ҽ������ Where ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
            If Not rsTmp.EOF Then
                Me.lvwPati.SelectedItem.ListSubItems(3).Tag = rsTmp("ִ��״̬")
                Me.lvwPati.SelectedItem.ListSubItems(9).Tag = rsTmp("ִ�й���")
            End If
            If Not mfrmRepEdit Is Nothing Then
                Unload mfrmRepEdit
'                MsgBox "������д���档Ҫ�༭�������棬��رյ�ǰ�ı�����д���ڣ�", vbInformation, gstrSysName
'                Call ShowWindow(mfrmRepEdit.Hwnd, SW_RESTORE)
'                Call BringWindowToTop(mfrmRepEdit.Hwnd)
'                Exit Sub
            End If
            '�ж��Ƿ��������
            strIfAuditing = IIf((InStr(mstrPrivs, "�������") <> 0), "1", "0")
'            If InStr(mstrPrivs, "�������") = 0 And Len(Me.lvwPati.SelectedItem.SubItems(15)) = 0 Then
'                '�ж��Ƿ���
'                blnEmerge = False
'                strSQL = "Select ���� From ���˹Һż�¼ Where NO=[1]"
'                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Split(Me.lvwPati.SelectedItem.Tag, "_")(2))
'                If Not rsTmp.EOF Then blnEmerge = (Nvl(rsTmp(0), 0) = 1)
'
'                If Not blnEmerge Then strIfAuditing = "0"
'            End If
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Then strIfAuditing = "0"
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then strIfAuditing = "0"
            '�ж��Ƿ�������
            strIfRollback = IIf((InStr(mstrPrivs, "���沵��") <> 0), "1", "0")
'            If InStr(mstrPrivs, "���沵��") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.���� Then
'                strIfRollback = "0"
'            End If
            If InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Then strIfRollback = "0"
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then strIfRollback = "0"
            '�ж��Ƿ������ӡ
            strIfPrint = "0"
            If InStr(mstrPrivs, "�������") <> 0 Or Me.lvwPati.SelectedItem.SubItems(12) = "" Or Me.lvwPati.SelectedItem.SubItems(12) = UserInfo.���� Then
                strIfPrint = "1"
            End If
            
            With lvwPati.SelectedItem
                EditReport Me, .SubItems(1), CInt(.ListSubItems(5).Tag), _
                    CLng(Split(.ListSubItems(8).Tag, "|")(0)), CLng(Split(.ListSubItems(8).Tag, "|")(1)), "", _
                    Val(Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 6 Or InStr(mstrPrivs, "��д����") = 0, False, tmpObject, , _
                    Not (InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Or _
                    InStr("3,4,5,6", Me.lvwPati.SelectedItem.ListSubItems(9).Tag) = 0), True, lngҽ��ID, _
                    Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1", strIfAuditing & strIfRollback & strIfPrint
                Set mfrmRepEdit = tmpObject
            End With
            DoEvents
            '�򿪹�Ƭվ
            If mblnViewImage Then
                Me.TabFile.Tabs("Ӱ��").Selected = True
                mfrmActive.zlMenuClick Me.mnuImageView(2) 'ѡ����������
                mfrmActive.zlMenuClick Me.mnuImageView(0)
            End If
        Case 3 '�������
           
            If InStr(mstrPrivs, "�������") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.���� Then
                    MsgBox "�������ֻ������Լ���д�ı��棡", vbInformation, gstrSysName
                    Exit Sub
            End If
            
            blnEmerge = (Me.lvwPati.SelectedItem.SubItems(15) = "��")
            If InStr(mstrPrivs, "�������") = 0 And Len(Me.lvwPati.SelectedItem.SubItems(15)) = 0 Then
                'û�н�����־��,�����ж��Ƿ���
                strSQL = "Select ���� From ���˹Һż�¼ Where NO=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Split(Me.lvwPati.SelectedItem.Tag, "_")(2))
                If Not rsTmp.EOF Then blnEmerge = (Nvl(rsTmp(0), 0) = 1)
                
                If Not blnEmerge Then
                    MsgBox "��ֻ����˼����飡", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            
            
            If Me.lvwPati.SelectedItem.ListSubItems(3).Tag <> "3" Then
                MsgBox "�����鱨�治�������ˣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            'û�б��棬��ֱ��ָ��������
            If CLng(Split(lvwPati.SelectedItem.ListSubItems(8).Tag, "|")(1)) = 0 Then
                If Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���Խ��������", 0)) = 0 Then
                    iMsgReturn = MsgBox("��ȷ�ϼ�����Ƿ�Ϊ���ԣ�" & vbCrLf & "ѡ��ȡ���������ˡ�", vbYesNoCancel + vbQuestion + vbDefaultButton1, gstrSysName)
                    If iMsgReturn = vbCancel Then Exit Sub
                    iMsgReturn = IIf(iMsgReturn = vbYes, 1, 0)
                Else
                    iMsgReturn = 0
                End If
            Else
                iMsgReturn = -1
            End If
            If InStr(mstrPrivs, "�������") = 0 And mblnEmergencyPrint = True And blnEmerge = True Then
                '��������Ҵ���ɴ�ӡ
                strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID= [1] And ���ͺ�= [2] "
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
                If Not rsTmp.EOF Then
                    strSQL = "ZL_Ӱ�����¼_FilmState('" & rsTmp(0) & "',1)"
                    ExecuteProc strSQL, Me.Caption
                End If
                Call LoadPatiList
            Else
                '�����������
                Call ExeFinish(lngҽ��ID, lng���ͺ�, False, iMsgReturn)
            
                If lvwPati.Tag <> "2" Then
                    lvwPati.Tag = 2: picKind_Resize
                    Call LoadPatiList("_" & lngҽ��ID & "_" & lng���ͺ�)
                    ShowCheck 2
                Else
                    Call LoadPatiList
                End If
            End If
            
        Case 4 '���沵��
            
'            If InStr(mstrPrivs, "���沵��") = 0 And strAudiName <> UserInfo.���� Then
            If InStr(mstrPrivs, "���沵��") = 0 And _
                (Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.���� Or (Me.lvwPati.SelectedItem.SubItems(13) <> UserInfo.���� And Me.lvwPati.SelectedItem.SubItems(13) <> "")) Then
                MsgBox "��ֻ�ܲ����Լ��ı��棡", vbInformation, gstrSysName
                Exit Sub
            End If
            If InStr("1,3", Me.lvwPati.SelectedItem.ListSubItems(3).Tag) = 0 Then
                MsgBox "�����鱨�滹δ��д�����貵�أ�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1" Then
                MsgBox "��ǰ�����ת�뱸�ݣ�����ִ�б�������", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If MsgBox("ȷ��Ҫ���ظ����鱨����", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            If InStr(mstrPrivs, "���沵��") = 0 And mblnEmergencyPrint = True Then
                '�������Ȩ��,ͬʱʹ�ý�����˴�ӡ�ķ����������,����ʱ,ֻ���ش�ӡ״̬
                strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID= [1] And ���ͺ�= [2] "
                Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
                If Not rsTmp.EOF Then
                    strSQL = "ZL_Ӱ�����¼_FilmState('" & rsTmp(0) & "',0)"
                    ExecuteProc strSQL, Me.Caption
                End If
            Else
                If Me.lvwPati.SelectedItem.ListSubItems(9).Tag <> "6" Then
                    strSQL = "ZL_Ӱ����_STATE(" & lngҽ��ID & "," & lng���ͺ� & ",5)"
                    ExecuteProc strSQL, Me.Caption
                Else
                    Call ExeFinish(lngҽ��ID, lng���ͺ�, True)
                End If
            End If
            Call LoadPatiList
        Case 6 '�����ӡ
            If InStr(mstrPrivs, "�������") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.���� Then
                MsgBox "��ֻ�ܴ�ӡ�Լ���д�ı��棡", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
            PrintDiagReport lngҽ��ID, lng���ͺ�, Me, , Me.picBuffer
            Me.MousePointer = vbDefault
        Case 7 '����Ԥ��
            If InStr(mstrPrivs, "�������") = 0 And Me.lvwPati.SelectedItem.SubItems(12) <> UserInfo.���� Then
                MsgBox "��ֻ�ܴ�ӡ�Լ���д�ı��棡", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
            Me.lvwPati.Enabled = False
            PrintDiagReport lngҽ��ID, lng���ͺ�, Me, 1, Me.picBuffer
            Me.lvwPati.Enabled = True
            Me.MousePointer = vbDefault
        Case 9 '�����ʽ
            On Error Resume Next
            gintReportFormat = Val(InputBox("�����뱨���ʽ��ţ�", "���뱨���ʽ", 1))
            If gintReportFormat = 0 Then gintReportFormat = 1
            On Error GoTo 0
    End Select
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExeFinish(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal blnCancel As Boolean, Optional ByVal iCheckResult As Integer = -1)
'������iCHeckResult ���������
'       -1�����ԣ�������ǰ���
'       0���������
'       1���������
    Dim strSQL As String
    
    gcnOracle.BeginTrans
    On Error GoTo DBError
    If blnCancel Then
        strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngAdviceID & "," & lngSendNO & ")"
        ExecuteProc strSQL, Me.Caption
        strSQL = "ZL_Ӱ����_STATE(" & lngAdviceID & "," & lngSendNO & ",5)"
        ExecuteProc strSQL, Me.Caption
    Else
        If iCheckResult = -1 Then
            strSQL = "ZL_����ҽ��ִ��_Finish(" & lngAdviceID & "," & lngSendNO & ")"
        Else
            strSQL = "ZL_����ҽ��ִ��_Finish(" & lngAdviceID & "," & lngSendNO & "," & iCheckResult & ")"
        End If
        ExecuteProc strSQL, Me.Caption
        strSQL = "ZL_Ӱ����_STATE(" & lngAdviceID & "," & lngSendNO & ",6)"
        ExecuteProc strSQL, Me.Caption
    End If
    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "Ӱ��ҽ�����"
End Sub

Private Sub mnuReqFunc_Click(Index As Integer)
    mfrmActive.zlMenuClick mnuReqFunc(Index)
End Sub

Private Sub mnuToolReport_Click(Index As Integer)
    Select Case Index
        Case 0 'ҽ����������
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1206_1", Me, _
                "ִ�п���=" & NeedName(cboDept.Text) & "|" & cboDept.ItemData(cboDept.ListIndex))
    End Select
End Sub

Private Sub mnuViewAdviceAppend_Click()
'���ܣ���ʾ������ҽ�����ӱ��
'�ӿڣ�Function zlMenuClick(objMenu as Menu) as Boolean
    
    '���õ�ǰ���ܴ��ڽӿ�
    If Not mfrmActive Is Nothing Then
        Call mfrmActive.zlMenuClick(mnuViewAdviceAppend)
    End If
End Sub

Private Sub mnuViewAdviceSelf_Click()
    mnuViewAdviceSelf.Checked = Not mnuViewAdviceSelf.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewCharge_Click()
    mnuViewCharge.Checked = Not mnuViewCharge.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewFileSelf_Click()
    mnuViewFileSelf.Checked = Not mnuViewFileSelf.Checked
    If mnuViewFileSelf.Checked Then mnuViewHistory.Checked = False
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewFilter_Click()
    frmPACSFilter.mBeforeDays = mBeforeDays
    frmPACSFilter.Show 1, Me
    mBeforeDays = frmPACSFilter.mBeforeDays
    If frmPACSFilter.mblnOK Then
        '���ù��˱���
        With frmPACSFilter
            '����ʱ��
            mdatFBegin = Format(.dtpBegin.Value, "yyyy-MM-dd HH:mm:00")
            If Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
                mdatFEnd = CDate(0) '��ʾȡ��ǰʱ��
            Else
                mdatFEnd = Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm:59")
            End If
            mDatType = .FindType
            
            '���ݺ�
            If .txtNO.Text <> "" Then
                mstrFNO = .txtNO.Text
            Else
                mstrFNO = ""
            End If
            
            '���걾��λ
            If Trim(.cboPart.Text) <> "" Then
                mstr�걾��λ = .cboPart.Text
            End If
            
            '���˿���
            If .cboDept.ListIndex <> 0 Then
                mlngF����ID = .cboDept.ItemData(.cboDept.ListIndex)
            Else
                mlngF����ID = 0
            End If
            
            '������Դ
            If Not (.chk��Դ(0).Value = 1 And .chk��Դ(1).Value = 1) Then
                If .chk��Դ(0).Value = 1 Then
                    mstrF��Դ = "1,3,4"
                ElseIf .chk��Դ(1).Value = 1 Then
                    mstrF��Դ = "2,3"
                End If
            Else
                mstrF��Դ = ""
            End If
            
            '���˱�ʶ
            If .txt��ʶ��.Text <> "" Then
                mdblF��ʶ�� = Val(.txt��ʶ��.Text)
            Else
                mdblF��ʶ�� = 0
            End If
            If .txt���￨.Text <> "" Then
                mstrF���￨ = .txt���￨.Text
            Else
                mstrF���￨ = ""
            End If
            If .txt����.Text <> "" Then
                mstrF���� = .txt����.Text
            Else
                mstrF���� = ""
            End If
            If .txtChkNO.Text <> "" Then
                mdblFChkNO = Val(.txtChkNO.Text)
            Else
                mdblFChkNO = 0
            End If
        End With
        Call mnuViewRefresh_Click
        
        Me.chkFilter.Value = 0
    End If
End Sub

Private Sub mnuViewHistory_Click()
    mnuViewHistory.Checked = Not mnuViewHistory.Checked
    If mnuViewHistory.Checked Then mnuViewFileSelf.Checked = False
    Call mnuViewRefresh_Click
End Sub

Private Sub picFile_Resize()
'���ܣ��������Resize
    If Not mfrmActive Is Nothing Then
        SetWindowPos mfrmActive.Hwnd, 0, 0, 0, picFile.ScaleWidth / Screen.TwipsPerPixelX, picFile.ScaleHeight / Screen.TwipsPerPixelY, SWP_NOREPOSITION Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub ReqList_Click(Index As Integer)
    mfrmActive.zlMenuClick ReqList(Index)
End Sub

Private Sub tabFile_Click()
'���ܣ�����ѡ��ֱ������Ӧ����
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim lng����id As Long, int������Դ As Integer
    Dim lng��ҳID As Long, str�Һŵ� As String
    Dim int�Ʒ�״̬ As Integer, int��¼���� As Integer
    Dim iNum As Integer
    Dim lngPatientID As Long, strCheckID As String
    Dim i As Integer
    Dim strMsg As String
        
    '1.��ʾҽ������
'    lblAdvice.Caption = Getִ������(lng���ͺ�, lngҽ��ID, Val(Item.ListSubItems(1).Tag), Item.ListSubItems(2).Tag)
    On Error Resume Next
    If Not mfrmActive Is Nothing Then 'And TabIndex <> TabFile.SelectedItem.Index Then
        mfrmActive.Hide
        Set mfrmActive = Nothing
    End If
    TabIndex = TabFile.SelectedItem.Index
    
    Me.mnuPFile.Visible = False
    Me.mnuReq.Visible = False
    Me.mnuAdvice.Visible = False
    Me.mnuMoney.Visible = False
    
    Me.mnuViewAdviceSelf.Visible = False
    Me.mnuViewFileSelf.Visible = False
    Me.mnuViewHistory.Visible = False
    Me.mnuViewAdviceAppend.Visible = False
    Me.mnuViewPic.Visible = False
    
    Me.mnuImageView(0).Visible = False
    Me.mnuImageView(1).Visible = False
    Me.mnuImageView(2).Visible = False
    Me.mnuImageView(3).Visible = False
    
    Select Case TabFile.SelectedItem.Key
        Case "����"
            Me.mnuPFile.Visible = True
            Me.mnuReq.Visible = True
            Me.mnuViewFileSelf.Visible = True
            Me.mnuViewHistory.Visible = True
            If aForms(4) Is Nothing Then Set aForms(4) = New frmPACSRec
            Set mfrmActive = aForms(4)
            
            Set mfrmActive.mfrmParent = Me
            mfrmActive.mstrPrivs = mstrPrivs
        Case "ҽ��"
            Me.mnuAdvice.Visible = True
            Me.mnuViewAdviceSelf.Visible = True
            Me.mnuViewAdviceAppend.Visible = True
            
            If Me.lvwPati.SelectedItem Is Nothing Then
                If aForms(1) Is Nothing Then Set aForms(1) = InDoctorAdvice
                Set mfrmActive = aForms(1)
            ElseIf lvwPati.SelectedItem.Text = "����" Then
                If aForms(2) Is Nothing Then Set aForms(2) = OutDoctorAdvice
                Set mfrmActive = aForms(2)
            Else
                If aForms(1) Is Nothing Then Set aForms(1) = InDoctorAdvice
                Set mfrmActive = aForms(1)
            End If
            
            Set mfrmActive.mfrmParent = Me
            mfrmActive.mstrPrivs = mstrPrivs
        Case "����"
            Me.mnuMoney.Visible = True
            If aForms(0) Is Nothing Then Set aForms(0) = New frmPACSReq
            Set mfrmActive = aForms(0)
        Case "Ӱ��"
            Me.mnuImageView(0).Visible = True
            Me.mnuImageView(1).Visible = True
            Me.mnuImageView(2).Visible = True
            Me.mnuImageView(3).Visible = True
            Me.mnuViewPic.Visible = True
            If aForms(3) Is Nothing Then Set aForms(3) = New frmPACSImg
            Set mfrmActive = aForms(3)
    End Select
    
    '����������
    For iNum = 1 To Me.tbrMain.Buttons.Count
        If Len(Me.tbrMain.Buttons(iNum).Description) > 0 And _
            Me.tbrMain.Buttons(iNum).Description <> TabFile.SelectedItem.Key Then
            Me.tbrMain.Buttons(iNum).Visible = False
        Else
            Me.tbrMain.Buttons(iNum).Visible = True
        End If
    Next
    '������Ȩ����������Ȩ��(Visible),����Ȩ��(Enabled)���Ӵ����д���
    Call SetFuncPrivs
    If mfrmActive Is Nothing Then Exit Sub
    
    SetWindowLong mfrmActive.Hwnd, GWL_STYLE, WS_CHILD
    mfrmActive.Show , Me
    SetParent mfrmActive.Hwnd, picFile.Hwnd
    mfrmActive.ZOrder 0
    
    picFile_Resize
    
    If Me.lvwPati.SelectedItem Is Nothing Then
        lngҽ��ID = 0
        lng���ͺ� = 0
        lng����id = 0
        lng��ҳID = 0
        str�Һŵ� = ""
        int������Դ = 2
    Else
        With lvwPati.SelectedItem
            lngҽ��ID = Val(Split(Mid(.Key, 2), "_")(0))
            lng���ͺ� = Val(Split(Mid(.Key, 2), "_")(1))
            lng����id = Val(Split(.Tag, "_")(0))
            lng��ҳID = Val(Split(.Tag, "_")(1))
            str�Һŵ� = Split(.Tag, "_")(2)
            int������Դ = IIf(.Text = "����", 1, 2)
        End With
    End If
    '�˵�������������
    ShowMenu
    If Me.Visible Then
        Select Case TabFile.SelectedItem.Key
            Case "����"
                ShowAddFileMenu int������Դ '��ʾ�����˵�
                
                Me.MousePointer = vbHourglass
                mfrmActive.zlRefresh lng����id, IIf(int������Դ = 1, str�Һŵ�, lng��ҳID), lngҽ��ID, Not Me.mnuViewFileSelf.Checked, Me.mnuViewHistory.Checked, _
                    Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1"
                Me.MousePointer = vbDefault
            Case "ҽ��"
                Me.MousePointer = vbHourglass
                If int������Դ = 2 Then
                    mfrmActive.zlRefresh lng����id, lng��ҳID, 0, 0, False, lngҽ��ID, Not Me.mnuViewAdviceSelf.Checked
                Else
                    mfrmActive.zlRefresh lng����id, str�Һŵ�, 1, 0, lngҽ��ID, Not Me.mnuViewAdviceSelf.Checked
                End If
                Me.MousePointer = vbDefault
                Me.stbThis.Panels(2).Text = ""
            Case "����"
                strMsg = Me.stbThis.Panels(2).Text
                BeginShowProgress "���ڶ�ȡ��"
                Me.MousePointer = vbHourglass
                mfrmActive.zlRefresh Me, lngҽ��ID, lng���ͺ�, mstrPrivs, pgbLoad, Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1"
                Me.MousePointer = vbDefault
                Me.pgbLoad.Visible = False
                Me.stbThis.Panels(2).Text = strMsg
            Case "Ӱ��"
                strMsg = Me.stbThis.Panels(2).Text
                BeginShowProgress "���ڶ�ȡ��"
                Me.MousePointer = vbHourglass
                If mfrmActive.zlRefresh(Me, lngҽ��ID, lng���ͺ�, mstrPrivs, pgbLoad, mnuViewPic.Checked, Me.lvwPati.SelectedItem.ListSubItems(11).Tag = "1", mDispImgs) = True Then
                    mnuExecFunc(10).Enabled = True
                    mnuExecFunc(12).Enabled = True
                Else
                    mnuExecFunc(10).Enabled = False
                    mnuExecFunc(12).Enabled = False
                End If
                Me.MousePointer = vbDefault
                Me.pgbLoad.Visible = False
                Me.stbThis.Panels(2).Text = strMsg
        End Select
    End If
    
    lvwPati.SetFocus
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
'        Case "���"
'            mnuExecFinish_Click
        Case "�˳�"
            mnuFileQuit_Click
        Case "�ɼ�"
            mnuExecFunc_Click 5
        Case "��ӡ"
            mnuFilePrint_Click
        Case "Ԥ��"
            mnuFilePreview_Click
        Case "����"
            mnuRepFunc_Click 0
        Case "���"
            mnuRepFunc_Click 3
        Case "����"
            mnuRepFunc_Click 4
        Case "����"
            mnuHelpTitle_Click
        Case "����"
            mnuViewFilter_Click
        Case Else
            mfrmActive.zlButtonClick Button
    End Select
End Sub

Private Sub mnuToolDiagRef_Click()
'���ܣ�������ϲο�
    Call ShowDiagHelp(0, Me)
End Sub

Private Sub mnuToolItemRef_Click()
'����: �������Ʋο�
    Dim lng������ĿID As Long

    If Me.TabFile.SelectedItem.Key = "ҽ��" Then
        mfrmActive.zlItemRef
    Else
        If Not lvwPati.SelectedItem Is Nothing Then lng������ĿID = Val(lvwPati.SelectedItem.ListSubItems(7).Tag)
        Call ShowClinicHelp(0, Me, lng������ĿID)
    End If
End Sub

Private Sub mnuFilePrintSet_Click()
'���ܣ���ӡ����
    Call zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
'���ܣ������Excel
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
'���ܣ���ӡԤ��
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
'���ܣ���ӡ
    Call OutputList(1)
End Sub

Private Sub SetFuncPrivs()
'���ܣ�������Ȩ����������Ȩ��(Visible)
    Dim i As Integer
    On Error Resume Next
    If InStr(mstrPrivs, "ֱ������") = 0 Then
        Me.mnuExecFunc(0).Visible = False
        Me.mnuExecFunc(1).Visible = False
        Me.mnuExecFunc(3).Visible = False
    End If
    If InStr(mstrPrivs, "Ӱ����") = 0 Then
        Me.mnuExec.Visible = False
    End If
    If InStr(mstrPrivs, "��д����") = 0 Then
'        Me.mnuRepFunc(0).Visible = False
        Me.mnuRepFunc(1).Visible = False
        Me.mnuRepFunc(2).Visible = False
        
        Me.mnuExecFunc(16).Visible = False
        Me.mnuExecFunc(17).Visible = False
        
        Me.tbrMain.Buttons("����").Visible = False
        'û����д����Ȩ��Ҳ���ܴ�ӡ����
        Me.mnuRepFunc(5).Visible = False
        Me.mnuRepFunc(6).Visible = (InStr(mstrPrivs, "�������") > 0) '�����Ȩ�޿��Դ�ӡ
        Me.mnuRepFunc(7).Visible = False
    End If
    If InStr(mstrPrivs, "�������") = 0 And InStr(mstrPrivs, "�������") = 0 Then
        Me.mnuRepFunc(3).Visible = False
        Me.tbrMain.Buttons("���").Visible = False
        
        Me.mnuExecFunc(18).Visible = False
    End If
    If InStr(mstrPrivs, "���沵��") = 0 And InStr(mstrPrivs, "�������") = 0 Then
        Me.mnuRepFunc(4).Visible = False
        Me.tbrMain.Buttons("����").Visible = False
    End If
    If Not Me.mnuRepFunc(3).Visible And Not Me.mnuRepFunc(4).Visible Then
        Me.mnuRepFunc(5).Visible = False
    End If
    If InStr(mstrPrivs, "Ӱ����") = 0 Then
        Me.mnuImageView(0).Visible = False
        Me.mnuImageView(1).Visible = False
        Me.mnuImageView(2).Visible = False
        Me.mnuImageView(3).Visible = False
        
        Me.mnuReqFunc(9).Visible = False
        Me.mnuReqFunc(10).Visible = False
        Me.mnuReqFunc(11).Visible = False
        
        
        Me.tbrMain.Buttons("��Ƭ").Visible = False
        Me.tbrMain.Buttons("ȫѡ").Visible = False
        Me.tbrMain.Buttons("ȫ��").Visible = False
    End If
    If InStr(mstrPrivs, "��д����") = 0 And InStr(mstrPrivs, "�������") = 0 And InStr(mstrPrivs, "���沵��") = 0 And InStr(mstrPrivs, "�������") = 0 Then
        Me.mnuImageView(3).Visible = False
        For i = 0 To mnuRepFunc.Count - 1
            mnuRepFunc(i).Visible = False
        Next
    
        Me.tbrMain.Buttons("Split_Rep").Visible = False
    End If
    If InStr(mstrPrivs, "Ӱ����") = 0 And InStr(mstrPrivs, "��д����") = 0 _
        And InStr(mstrPrivs, "�������") = 0 And InStr(mstrPrivs, "���沵��") = 0 Then Me.mnuRep.Visible = False
'    If InStr(mstrPrivs, "Ӱ����") = 0 Or (InStr(mstrPrivs, "��д����") = 0 And InStr(mstrPrivs, "�������") = 0 And InStr(mstrPrivs, "���沵��") = 0) Then
'        If InStr(mstrPrivs, "��д����") = 0 And InStr(mstrPrivs, "�������") = 0 And InStr(mstrPrivs, "���沵��") = 0 Then Me.mnuRep.Visible = False
'        Me.mnuImageView(0).Visible = False
'        Me.mnuImageView(1).Visible = False
'        Me.mnuViewPic.Visible = False
'
'        Me.tbrMain.Buttons("��Ƭ").Visible = False
'        Me.tbrMain.Buttons("��ʾ").Visible = False
'        Me.tbrMain.Buttons("View_").Visible = False
'    End If
    If InStr(mstrPrivs, "�������") = 0 Then
        Me.mnuMoney.Visible = False
    
        Me.tbrMain.Buttons("����").Visible = False
        Me.tbrMain.Buttons("����").Visible = False
        Me.tbrMain.Buttons("�ķ�").Visible = False
        Me.tbrMain.Buttons("ɾ��").Visible = False
        Me.tbrMain.Buttons("Money_").Visible = False
    End If
    If InStr(mstrPrivs, "ҽ���´�") = 0 Then
        Me.mnuAdvice.Visible = False
    
        Me.tbrMain.Buttons("�¿�").Visible = False
        Me.tbrMain.Buttons("�޸�").Visible = False
        Me.tbrMain.Buttons("ɾ��").Visible = False
        Me.tbrMain.Buttons("����").Visible = False
        Me.tbrMain.Buttons("Advice_").Visible = False
    End If
    If InStr(mstrPrivs, "������д") = 0 Then
        Me.mnuPFile.Visible = False
    
        Me.tbrMain.Buttons("����").Visible = False
        Me.tbrMain.Buttons("�����޸�").Visible = False
        Me.tbrMain.Buttons("ɾ����").Visible = False
        Me.tbrMain.Buttons("File_").Visible = False
    End If
    
    If InStr(mstrPrivs, "��д����") = 0 Then
        'Me.mnuReq.Visible = False
        Me.mnuReqFunc(0).Visible = False
        Me.mnuReqFunc(1).Visible = False
        Me.mnuReqFunc(2).Visible = False
        Me.mnuReqFunc(3).Visible = False
    End If
    
    '���ơ����롱�˵�����ġ���ӡԤ�����͡������ӡ���˵���
    If InStr(mstrPrivs, "�������") = 0 Then
        Me.mnuReqFunc(7).Visible = False
        Me.mnuReqFunc(8).Visible = False
    End If
    
    If InStr(mstrPrivs, "��Ƶ�ɼ�") = 0 Then
        Me.mnuExecFunc(5).Visible = False
        Me.tbrMain.Buttons("�ɼ�").Visible = False
    End If
    If InStr(mstrPrivs, "��ʼ���") = 0 Then
        Me.mnuExecFunc(4).Visible = False
    End If
    If InStr(mstrPrivs, "ȡ�����") = 0 Then
        Me.mnuExecFunc(7).Visible = False
    End If
    If InStr(mstrPrivs, "������ͼ��") = 0 Then
        Me.mnuExecFunc(12).Visible = False
    End If
    
    'ȥ���ָ���
    If InStr(mstrPrivs, "��ʼ���") = 0 And InStr(mstrPrivs, "��Ƶ�ɼ�") = 0 And InStr(mstrPrivs, "ȡ�����") = 0 Then
        Me.mnuExecFunc(8).Visible = False
    End If
    '�ļ�����
    If InStr(mstrPrivs, "�ļ�����") = 0 Then
        mnufileSendImage.Visible = False
    End If
    
End Sub

Private Sub mnuHelpTitle_Click()
'���ܣ����ð�������
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim blnTmp As Boolean
    Dim i As Integer
    Dim ret As Long
    
    Call RestoreWinState(Me, App.ProductName)
    Me.lvwPati.ColumnHeaders(16).Position = 1
    
    InitLocalPars
    mBeforeDays = 2
    
    '���ӷ�����ģ��ı���˵�
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '���ҷ�ʽ
    i = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "FindItem", 0))
    mnuViewFindItem(i).Checked = True
    Call mnuViewFindItem_Click(CInt(i))
    
    Me.Tag = "Loading"
    cboState(0).ListIndex = 0
    Me.Tag = "Loading"
    cboState(1).ListIndex = 0
    Me.Tag = "Loading"
    cboState(2).ListIndex = 0
    Me.Tag = "Loading"
    
    '������������
    mdatFBegin = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
    mdatFEnd = CDate(0)
    mstrFNO = ""
    mlngF����ID = 0
    mstrF��Դ = ""
    mdblF��ʶ�� = 0
    mstrF���￨ = ""
    mstrF���� = ""
    mDatType = 1
    
    '��ʼ�������ʽΪ1
    gintReportFormat = 1
    
    'Ȩ�޴���
    mstrPrivs = gstrPrivs
    Call SetFuncPrivs
    
    mlngPreDept = -1
    mstrPrePati = ""
    mstrFilter = ""
        
    Call InitSysPar '��ʼ��ϵͳ����
    
    AddFileList '���첡���˵�
    LoadBillList
    
    lvwPati.ListItems.Add , , "Temp", , 1
    lvwPati.ListItems.Clear
    
    '��ʼ��ҽ������
    If Not InitDepts Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "���п���") > 0 Then
            MsgBox "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Else
            MsgBox "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If

    TabIndex = 1
    TabFile.Tabs(TabIndex).Selected = True

    Set objImgCapture = CreateObject("zl9ImgCapture.clsImgCapture")
    objImgCapture.InitImgCapture gcnOracle

    '�����ȼ�
    '��¼ԭ����window�����ַ
    If App.LogMode <> 0 Then
        preWinProc = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
        '���Զ���������ԭ����window����
        ret = SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf Wndproc)
    End If
    idHotKey = 1
    Modifiers = MOD_CONTROL     'Ctrl ��
    uVirtKey = vbKey1  '1��
    ret = RegisterHotKey(Me.Hwnd, idHotKey, Modifiers, uVirtKey)

End Sub

Private Sub InitLocalPars()
    mnuViewAdviceSelf.Checked = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ��ʾ�����´��ҽ��", 1)) <> 0
    mnuViewFileSelf.Checked = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ��ʾ������д�Ĳ���", 1)) <> 0
    mnuViewCharge.Checked = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ��ʾ���շѵĲ���", 0)) <> 0
'    mnuViewPic.Checked = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ʾ��ǰ����ͼ��", 0)) <> 0
    
    mstrRoom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰִ�м�")
    blnIfOnlyShow = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ����ǰִ�м���Ŀ", False)
    mDispImgs = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ʾͼ����", 20)
    mblnEmergencyPrint = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�������ʱ��ӡ", 0)
    
    mblnViewImage = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ��Ƭ", 0))
    mblnSample = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�Ǽ�ֱ�Ӽ��", 0))
'    mBeforeDays = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "�����ѯ������", 3)
'    If mBeforeDays <= 0 Then mBeforeDays = 3
End Sub

Public Sub mnuViewRefresh_Click()
    Call LoadPatiList
End Sub

Private Sub cboDept_Click()
    If cboDept.ListIndex = mlngPreDept Then Exit Sub
    mlngPreDept = cboDept.ListIndex
    
    Call LoadPatiList
End Sub

Private Sub cbr_Resize()
    Call Form_Resize
End Sub

Private Sub fraLR_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngMinWidth As Long
    If Button <> 1 Then Exit Sub
    
    lngMinWidth = Me.cmdSeek.Left + Me.cmdSeek.Width + Me.fraState.Left + 150
    fraLR_s.BackColor = RGB(0, 0, 0)
    On Error Resume Next
    If fraLR_s.Left + x < lngMinWidth Then
        fraLR_s.Left = lngMinWidth
    ElseIf Me.ScaleWidth - fraLR_s.Left - x < 2000 Then
        fraLR_s.Left = Me.ScaleWidth - 2000
    Else
        fraLR_s.Left = fraLR_s.Left + x
    End If
End Sub

Private Sub fraLR_s_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraLR_s.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub mnuhelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuFileSetup_Click()
    frmTechnicSetup.Show 1, Me
    If frmTechnicSetup.mblnOK Then
'        Call LoadBillDetail(vsMoney.Row)
        InitLocalPars
    
        Call LoadPatiList
        '��λ������ҵĲ��˼��
        If txt��ʶ��.Text <> "" Then Call SeekNextPati(True)
        If Me.lvwPati.Visible Then
            Me.lvwPati.SetFocus
        End If
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolItem_Click(Index As Integer)
    Dim blnEnabled As Boolean, blnVisible As Boolean, i As Integer
    
    mnuViewToolItem(Index).Checked = Not mnuViewToolItem(Index).Checked
    cbr.Bands(Index + 1).Visible = Not cbr.Bands(Index + 1).Visible

    blnEnabled = False: blnVisible = False
    For i = 1 To cbr.Bands.Count
        'ֻ����һ��ToolBar�ɼ�,��"��ʾ�ı�"�˵��ɼ�
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            If cbr.Bands(i).Visible Then
                blnEnabled = True
            End If
        End If
        'ֻҪ��һ��Band�ɼ�,��CoolBar�ɼ�
        If cbr.Bands(i).Visible Then
            blnVisible = True
        End If
    Next
    mnuViewToolText.Enabled = blnEnabled
    cbr.Visible = blnVisible
    
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer, j As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To cbr.Bands.Count
        If TypeName(cbr.Bands(i).Child) = "Toolbar" Then
            For j = 1 To cbr.Bands(i).Child.Buttons.Count
                cbr.Bands(i).Child.Buttons(j).Caption = IIf(mnuViewToolText.Checked, cbr.Bands(i).Child.Buttons(j).Tag, "")
            Next
            If Not mnuViewToolText.Checked Then
                cbr.Bands(i).Child.TextAlignment = tbrTextAlignBottom
            End If
            cbr.Bands(i).MinHeight = cbr.Bands(i).Child.ButtonHeight
            cbr.Bands(i).Child.Refresh
        End If
    Next
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Hwnd
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, i As Long
    Dim lngMinWidth As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    lngMinWidth = Me.cmdSeek.Left + Me.cmdSeek.Width + Me.fraState.Left + 150
    If Me.fraLR_s.Left > Me.ScaleWidth Then Me.fraLR_s.Left = Me.ScaleWidth - 2000
    If Me.fraLR_s.Left < lngMinWidth Then Me.fraLR_s.Left = lngMinWidth
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    picKind.Left = 0
    picKind.Top = cbrH
    picKind.Height = Me.ScaleHeight - cbrH - staH - Me.fraState.Height + 45
    picKind.Width = fraLR_s.Left
    
    Me.fraState.Left = 0
    Me.fraState.Top = Me.ScaleHeight - staH - fraState.Height
    Me.fraState.Width = Me.picKind.Width
    
    fraLR_s.Top = picKind.Top
    fraLR_s.Height = Me.ScaleHeight - staH - cbrH
    
    With TabFile
        .Left = fraLR_s.Left + fraLR_s.Width: .Top = cbrH
        .Width = Me.ScaleWidth - .Left
    End With
    With picFile
        .Left = fraLR_s.Left + fraLR_s.Width: .Top = TabFile.Top + TabFile.Height '- 140
        .Width = Me.ScaleWidth - .Left: .Height = Me.ScaleHeight - staH - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim objPacsCore As Object
    
    On Error Resume Next
    For i = 0 To 4
        If Not aForms(i) Is Nothing Then
            Unload aForms(i)
            Set aForms(i) = Nothing
        End If
    Next
    If Not mfrmActive Is Nothing Then
        Unload mfrmActive
        Set mfrmActive = Nothing
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ��ʾ���շѵĲ���", IIf(mnuViewCharge.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ��ʾ�����´��ҽ��", IIf(mnuViewAdviceSelf.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "ֻ��ʾ������д�Ĳ���", IIf(mnuViewFileSelf.Checked, 1, 0)
'    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ʾ��ǰ����ͼ��", IIf(mnuViewPic.Checked, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "����δ�������", Me.chk״̬(0).Value = 1
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������ִ�м��", Me.chk״̬(2).Value = 1
'    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "�����ѯ������", mBeforeDays
    '���ҷ�ʽ
    For i = 0 To mnuViewFindItem.UBound
        If mnuViewFindItem(i).Checked Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "FindItem", i
        End If
    Next
    
    Call SaveWinState(Me, App.ProductName)
    
    '�رչ�Ƭվ����
    Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
    objPacsCore.Closefrom
    
    '�رղɼ�վ
    objImgCapture.UnladImgCapture
    Set objImgCapture = Nothing
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    On Error GoTo errH
    
    '��������/סԺҽ������
    str��Դ = "3"
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "1,2,3"
    ElseIf InStr(mstrPrivs, "���ﲡ��") > 0 Then
        str��Դ = "1,3"
    ElseIf InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "2,3"
    End If
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   
    cboDept.Clear
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "," & str��Դ & ",")
    
    str����IDs = GetUser����IDs
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.����ID Then cboDept.ListIndex = cboDept.NewIndex 'ֱ����������
        If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = cboDept.NewIndex
        
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub tbrMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If Button.Key = "����" Then
'        PopupButtonMenu tbrMain, Button, mnuMoneyNew
    End If
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub txt��ʶ��_Change()
    If txt��ʶ��.Text = "" Then txt��ʶ��.Tag = ""
End Sub

Private Sub txt��ʶ��_GotFocus()
    Call zlControl.TxtSelAll(txt��ʶ��)
End Sub

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call Form_KeyDown(vbKeyF3, 0)
    Else
        Select Case Split(Label1.Caption, "(")(0)
            Case "��ʶ��"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "���￨"
                Dim blnCard As Boolean
    
                'ȥ���ſ��������������ַ�
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = InputIsCard(Me.txt��ʶ��, KeyAscii)
                
                'ˢ����ɻ�ȷ������
                If blnCard And Len(Me.txt��ʶ��.Text) = gbytCardLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txt��ʶ��.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txt��ʶ��.Text = Me.txt��ʶ��.Text & Chr(KeyAscii)
                        Me.txt��ʶ��.SelStart = Len(Me.txt��ʶ��.Text)
                    End If
                    KeyAscii = 0
                    Me.txt��ʶ��.Text = UCase(Me.txt��ʶ��)
                    Me.txt��ʶ��.SetFocus
                End If
            Case "���ݺ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txt��ʶ��.Text = "" Or txt��ʶ��.SelLength = Len(txt��ʶ��.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "����"
            
        End Select
    End If
End Sub

Private Sub txt��ʶ��_Validate(Cancel As Boolean)
    If Split(Label1.Caption, "(")(0) = "���ݺ�" Then
        If IsNumeric(txt��ʶ��.Text) Then
            txt��ʶ��.Text = GetFullNO(txt��ʶ��.Text, 0)
        End If
    End If
End Sub

Private Function LoadPatiList(Optional ByVal strKey As String = "") As Boolean
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSQLBak As String, i As Long, j As Long
    Dim objItem As ListItem, strPre As String
    Dim blnDo As Boolean, lngCount As Long
    Dim str��Դ As String
    Dim strFilter As String
    Dim blnMoved As Boolean
    
    If Not lvwPati.SelectedItem Is Nothing And Len(strKey) = 0 Then
        strPre = lvwPati.SelectedItem.Key
    Else
        strPre = strKey
    End If
    blnMoved = MovedByDate(IIf(mdatFBegin = CDate(0), CDate(zlDatabase.Currentdate) - mBeforeDays, mdatFBegin))
    
    '�����������
    mstrPrePati = ""
    lvwPati.ListItems.Clear
'    lblAdvice.Caption = ""
    
    On Error GoTo errH
        
    '������ԴȨ��:(1-����,2-סԺ,3-����,4-���)
    If InStr(mstrPrivs, "���ﲡ��") > 0 And InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "1,2,3,4"
    ElseIf InStr(mstrPrivs, "���ﲡ��") > 0 Then
        str��Դ = "1,4"
    ElseIf InStr(mstrPrivs, "סԺ����") > 0 Then
        str��Դ = "2"
    Else
        str��Դ = "3"
    End If
        
    If Me.chkFilter.Value = 1 Then
        strFilter = " And D.����=[17] "
    Else
        '����ʱ��
        If mdatFEnd <> CDate(0) Then
            strFilter = " And " & IIf(Val(lvwPati.Tag) = 0, "A.����ʱ��", IIf(mDatType = 2, "A.����ʱ��", "A.�״�ʱ��")) & " Between [1] and [2] "
        Else 'ȱʡ��ѯ����
            strFilter = " And " & IIf(Val(lvwPati.Tag) = 0, "A.����ʱ��", IIf(mDatType = 2, "A.����ʱ��", "A.�״�ʱ��")) & " Between [1] and Sysdate "
        End If
        '���ݺ�
        If mstrFNO <> "" Then
            strFilter = strFilter & " And A.NO= [3] "
        End If
        
        '���˿���
        If mlngF����ID <> 0 Then
            strFilter = strFilter & " And B.���˿���ID+0= [4] "
        End If
        
        '������Դ
        If mstrF��Դ <> "" Then
            strFilter = strFilter & " And instr([5],','||B.������Դ||',') > 0 "
        End If
        
        '���˱�ʶ
        
        If mdblF��ʶ�� <> 0 Then
            strFilter = strFilter & " And Decode(B.������Դ,1,D.�����,2,D.סԺ��,NULL)= [6] "
        End If
        
        If mstrF���￨ <> "" Then
            strFilter = strFilter & " And D.���￨�� = [7] "
        End If
        
        If mstrF���� <> "" Then
    '        strFilter = strFilter & " And D.���� = [8] "
            strFilter = strFilter & " And Instr(D.���� , [8])>0 "
        End If
        
        If mstr�걾��λ <> "" Then
            strFilter = strFilter & " And b.�걾��λ = [16]"
        End If
        
        If mdblFChkNO <> 0 Then
            strFilter = strFilter & " And H.����=[13] "
        End If
    End If
    
    '��ҩ�巨,�÷����Զ�����ʾһ��
    '��������,��鲿λִ�п��Ҽ�ʱ��������Ŀ��ͬ,����ʾ
    '��������ִ�п���Ϊ��������Ҫ��ʾ
    '����ҽ������ʾ(��Ȼִ�п���һ�㲻��Ϊҽ������)
'        " And Not (B.������� IN('F','D') And B.���ID is Not NULL)" & _
'        " And Not(B.�������='Z' And Nvl(C.��������,'0')<>'0')" & strWhere &
'        " And X.��¼״̬(+)<>2 And X.ҽ�����(+)=A.ҽ��ID And X.���(+)=1 And C.���='D'" &
    If Len(Trim(frmPACSFilter.cboContent.Text)) = 0 Or Me.chkFilter.Value = 1 Then
        strSQL = _
            " Select Distinct /*����շ���Ŀ*/ X.��¼���� as ��������,X.��¼״̬ as ����״̬," & _
            " A.ҽ��ID,A.���ͺ�,B.���ID,B.���,B.�������,B.������ĿID," & _
            " A.�״�ʱ�� As ���ʱ��,A.����ʱ�� As ����ʱ��,A.NO," & _
            " A.��¼����,A.ִ��״̬,A.�Ʒ�״̬,B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,E.���� as ����,D.����," & _
            " Decode(B.������Դ,1,D.�����,2,D.סԺ��,4,D.�����,NULL) as ��ʶ��,Nvl(D.�ѱ�,'��ͨ') As �ѱ�," & _
            " Decode(B.������Դ,1,'����',2,'סԺ',3,'����',4,'���') as ��Դ,C.���� as ����,A.ִ�м�," & _
            " Nvl(Z.�����ļ�ID,0) As ����ID,Nvl(A.����ID,0) As ����ID,Nvl(A.ִ�й���,0) As ִ�й���," & _
            " B.ҽ������,G.��д��,Decode(A.ִ��״̬,1,Nvl(G.�����,G.������),NULL) As ������,D.���￨��,0 As ת��,Nvl(H.����,'') As ����,Nvl(H.���UID,'') As ���UID,Nvl(A.�������,0) As ����,B.������־,Nvl(H.�Ƿ��ӡ,0) As �Ƿ��ӡ" & _
            " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C,������Ϣ D,���ű� E,���˲�����¼ G,Ӱ�����¼ H,���˷��ü�¼ X,Ӱ������Ŀ Y,���Ƶ���Ӧ�� Z" & _
            " Where A.ҽ��ID=B.ID And A.����ID=G.ID(+) And B.������ĿID=C.ID And B.����ID=D.����ID" & _
            " And B.���˿���ID=E.ID And C.ID=Y.������ĿID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+)" & _
            " And C.ID=Z.������ĿID(+) And (Z.Ӧ�ó���=B.������Դ Or B.������Դ=3 Or Z.������ĿID Is Null)" & _
            " And instr([10],','||B.������Դ||',')> 0 And A.ִ�в���ID+0= [11] " & _
            " And B.���ID is NULL " & _
            strFilter
    Else
        strSQL = _
            " Select Distinct /*����շ���Ŀ*/ X.��¼���� as ��������,X.��¼״̬ as ����״̬," & _
            " A.ҽ��ID,A.���ͺ�,B.���ID,B.���,B.�������,B.������ĿID," & _
            " A.�״�ʱ�� As ���ʱ��,A.����ʱ�� As ����ʱ��,A.NO," & _
            " A.��¼����,A.ִ��״̬,A.�Ʒ�״̬,B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,E.���� as ����,D.����," & _
            " Decode(B.������Դ,1,D.�����,2,D.סԺ��,4,D.�����,NULL) as ��ʶ��,Nvl(D.�ѱ�,'��ͨ') As �ѱ�," & _
            " Decode(B.������Դ,1,'����',2,'סԺ',3,'����',4,'���') as ��Դ,C.���� as ����,A.ִ�м�," & _
            " Nvl(Z.�����ļ�ID,0) As ����ID,Nvl(A.����ID,0) As ����ID,Nvl(A.ִ�й���,0) As ִ�й���," & _
            " B.ҽ������,G.��д��,Decode(A.ִ��״̬,1,Nvl(G.�����,G.������),NULL) As ������,D.���￨��,0 As ת��,Nvl(H.����,'') As ����,Nvl(H.���UID,'') As ���UID,Nvl(A.�������,0) As ����,B.������־,Nvl(H.�Ƿ��ӡ,0) As �Ƿ��ӡ" & _
            " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C,������Ϣ D,���ű� E,���˲�����¼ G,Ӱ�����¼ H,���˷��ü�¼ X,Ӱ������Ŀ Y,���Ƶ���Ӧ�� Z," & _
            " ���˲������� I, ���˲����ı��� J" & _
            " Where A.ҽ��ID=B.ID And A.����ID=G.ID(+) And B.������ĿID=C.ID And B.����ID=D.����ID" & _
            " And B.���˿���ID=E.ID And C.ID=Y.������ĿID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+)" & _
            " And C.ID=Z.������ĿID(+)" & _
            " And G.ID = I.������¼id And I.ID = J.����id And I.�����ı� = [14] AND Instr(J.����,[15])>0 And (Z.Ӧ�ó���=B.������Դ Or B.������Դ=3 Or Z.������ĿID Is Null)" & _
            " And instr([10],','||B.������Դ||',')> 0 And A.ִ�в���ID+0= [11] " & _
            " And B.���ID is NULL " & _
            strFilter
    End If
    Select Case Val(lvwPati.Tag)
        Case 0
'            strSQL = strSQL & " And (A.ִ��״̬=0 Or (A.ִ��״̬=3 And " & IIf(chk״̬(0).Value, _
'                "Nvl(A.ִ�й���,0)<2", "A.ִ�й���=1") & "))"
            strSQL = strSQL & " And ((A.ִ��״̬=3 Or A.ִ��״̬=0) And " & Decode(cboState(0).ListIndex, _
                 1, "Nvl(A.ִ�й���,0)=0)", 2, "Nvl(A.ִ�й���,0)=1)", "Nvl(A.ִ�й���,0)<2)")
        Case 1
'            strSQL = strSQL & " And A.ִ��״̬ =3 And A.ִ�й���=2"
            strSQL = strSQL & " And A.ִ��״̬ =3 And A.ִ�й���=2" & Decode(cboState(1).ListIndex, _
                 1, " And Nvl(A.����ID,0)=0", 2, " And Nvl(A.����ID,0)>0", "")
        Case 2
'            strSQL = strSQL & " And ((A.ִ��״̬ =3 And A.ִ�й���>2) Or " & IIf(chk״̬(2).Value, _
'                "A.ִ��״̬=1", "1=2") & ")"
            strSQL = strSQL & " And " & Decode(cboState(2).ListIndex, _
                1, "A.ִ��״̬ =3 And A.ִ�й��� =3", _
                2, "A.ִ��״̬ =3 And A.ִ�й��� =4", _
                3, "A.ִ��״̬ =3 And A.ִ�й��� =5", _
                4, "A.ִ��״̬ =1", _
                "((A.ִ��״̬ =3 And A.ִ�й���>2) Or A.ִ��״̬=1)")
    End Select
'    strSQL = strSQL & " And A.NO=X.NO(+) And A.��¼����=Decode(X.��¼����(+),0,1,X.��¼����(+))" & _
'        " And X.��¼״̬(+)<>2 And X.ҽ�����(+)=A.ҽ��ID And X.���(+)=1" & _
'        IIf(blnIfOnlyShow, " And A.ִ�м�= [12] ", "")
    strSQL = strSQL & " And A.NO=X.NO(+) And A.��¼����=Decode(X.��¼����(+),0,1,X.��¼����(+))" & _
        " And X.��¼״̬(+)<>2 And X.���(+)=1" & _
        IIf(blnIfOnlyShow, " And A.ִ�м�= [12] ", "")
    '���������ת����Ҫ�����󱸱�
    If blnMoved Then
        strSQLBak = strSQL
        strSQLBak = Replace(strSQLBak, "0 As ת��", "1 As ת��")
        strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
        strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
        strSQLBak = Replace(strSQLBak, "���˲�����¼", "H���˲�����¼")
        strSQLBak = Replace(strSQLBak, "���˷��ü�¼", "H���˷��ü�¼")
        strSQLBak = Replace(strSQLBak, "���˲�������", "H���˲�������")
        strSQLBak = Replace(strSQLBak, "���˲����ı���", "H���˲����ı���")
        strSQL = strSQL & " Union ALL " & strSQLBak
    End If
    strSQL = strSQL & " Order by ���ʱ�� Desc,����ʱ��,����ID,���"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, CDate(Format(mdatFBegin, "yyyy-MM-dd HH:mm:00")), CDate(Format(mdatFEnd, "yyyy-MM-dd HH:mm:59")), _
    mstrFNO, mlngF����ID, "," & mstrF��Դ & ",", mdblF��ʶ��, mstrF���￨, mstrF����, mBeforeDays, "," & str��Դ & ",", _
    cboDept.ItemData(cboDept.ListIndex), mstrRoom, mdblFChkNO, frmPACSFilter.cboItem.Text, frmPACSFilter.cboContent.Text, mstr�걾��λ, mstrPatiName)
    
    lngCount = 0
    For i = 1 To rsTmp.RecordCount
        '�Ƿ�ֻ��ʾ���շѵĲ���
        '1.ֻ��������,���жϸ��ӷ���.Ҳ�����������˷�
        '2.���ʵ��ݵ������շ���ʾ
        '3.����Ʒѻ���δ���������õ�Ҳ��ʾ
        blnDo = True
        If mnuViewCharge.Checked Then
            If Nvl(rsTmp!��������, 0) = 1 And Nvl(rsTmp!����״̬, 0) <> 1 Then
                blnDo = False '0-δ�շѻ�δ���,3-���˷ѻ�������;
            End If
        End If
        
        If blnDo Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!ҽ��ID & "_" & rsTmp!���ͺ�, Nvl(rsTmp!��Դ), , IIf(Len(rsTmp!���UID) > 0, 5, 0))
            objItem.SubItems(1) = Nvl(rsTmp!NO)
            objItem.SubItems(2) = LTrim(RTrim(Nvl(rsTmp!����)))
            objItem.SubItems(3) = Nvl(rsTmp!����)
            objItem.SubItems(4) = IIf(rsTmp!ִ��״̬ = 1, "ִ����", _
                IIf(Nvl(rsTmp!ִ�й���, 0) = 2 And Nvl(rsTmp!����ID, 0) > 0, "����", _
                Decode(Nvl(rsTmp!ִ�й���, 0), 0, "δ����", 1, "�ѱ���", 2, "�����", 3, "�����", 4, "����", 5, "����", 6, "������")))
            objItem.SubItems(5) = Nvl(rsTmp!����)
            objItem.SubItems(6) = Nvl(rsTmp!��ʶ��)
            objItem.SubItems(7) = Nvl(rsTmp!�ѱ�)
            objItem.SubItems(8) = Format(rsTmp!���ʱ��, "yy-MM-dd HH:mm")
            objItem.SubItems(9) = Nvl(rsTmp!ִ�м�)
            objItem.SubItems(10) = rsTmp!ҽ��ID '& "_" & rsTmp!���ͺ�
'            objItem.SmallIcon = IIf(objItem.SubItems(4) = "��ִ��", "δִ��", objItem.SubItems(4))
            objItem.SubItems(11) = GetPart(Nvl(rsTmp!ҽ������))
            objItem.SubItems(12) = Nvl(rsTmp!��д��)
            objItem.SubItems(13) = Nvl(rsTmp!������)
            objItem.SubItems(14) = Nvl(rsTmp!����)
            objItem.SubItems(15) = IIf(Nvl(rsTmp!������־, 0) = 1, "��", "")
            objItem.SubItems(16) = IIf(Nvl(rsTmp!�Ƿ��ӡ, 0) = 0, "", "��")
            objItem.SubItems(17) = Format(rsTmp!����ʱ��, "yy-MM-dd HH:mm")
            
            Select Case objItem.SubItems(4)
            Case "�ѱ���", "�����", "�����"
                objItem.ForeColor = 0 '��ɫ
            Case "δ����", "ִ����" '��ɫ
                objItem.ForeColor = &H808080
            Case "����" '��ɫ
                objItem.ForeColor = &H40C0&
            Case "����", "������" '��ɫ
                objItem.ForeColor = &HC00000
            End Select
            For j = 1 To Me.lvwPati.ColumnHeaders.Count - 1
                objItem.ListSubItems(j).ForeColor = objItem.ForeColor
            Next
            
            '��Ÿ�������
            objItem.Tag = rsTmp!����ID & "_" & Nvl(rsTmp!��ҳID, 0) & "_" & Nvl(rsTmp!�Һŵ�)
            
            objItem.ListSubItems(1).Tag = Nvl(rsTmp!���ID, 0)
            objItem.ListSubItems(2).Tag = rsTmp!�������
            objItem.ListSubItems(3).Tag = Nvl(rsTmp!ִ��״̬, 0)
            objItem.ListSubItems(4).Tag = Nvl(rsTmp!�Ʒ�״̬, 0)
            objItem.ListSubItems(5).Tag = Nvl(rsTmp!��¼����, 1)
            objItem.ListSubItems(6).Tag = Nvl(rsTmp!���˿���ID, 0)
            objItem.ListSubItems(7).Tag = Nvl(rsTmp!������ĿID, 0)
            objItem.ListSubItems(8).Tag = Nvl(rsTmp!����ID, 0) & "|" & Nvl(rsTmp!����ID, 0)
            objItem.ListSubItems(9).Tag = Nvl(rsTmp!ִ�й���, 0)
            objItem.ListSubItems(10).Tag = Nvl(rsTmp!���￨��)
            objItem.ListSubItems(11).Tag = Nvl(rsTmp!ת��, 0)
            
            objItem.ListSubItems(2).ReportIcon = IIf(rsTmp!���� = 1, 6, 7)
            
            If objItem.Key = strPre Then objItem.Selected = True
            
            lngCount = lngCount + 1
        End If
        rsTmp.MoveNext
    Next
    
    If Not lvwPati.SelectedItem Is Nothing Then
        Call lvwPati_ItemClick(lvwPati.SelectedItem)
        lvwPati.SelectedItem.EnsureVisible
        If (lvwPati.SelectedItem.Index <> lvwPati.ListItems.Count) Then
            lvwPati.ListItems(lvwPati.SelectedItem.Index + 1).EnsureVisible
        End If
    Else
        Call tabFile_Click
    End If
    
    Me.stbThis.Panels(2).Text = IIf(lngCount = 0, "û��", "���У�" & lngCount & " ��") & Decode(lvwPati.Tag, _
        "0", "��ִ�еļ��", _
        "1", "����ִ�еļ��", _
        "2", "����ɵļ��")
        
    mstr�걾��λ = ""
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowMenu()
    Dim blnEnable As Boolean, i As Integer
    Dim int������Դ As Integer
    
    On Error Resume Next
    blnEnable = False
    
    mnuExecFunc(0).Enabled = True
    If Not lvwPati.SelectedItem Is Nothing Then
        Me.mnuExec.Enabled = True
        Me.mnuRep.Enabled = True
        Me.mnuPFile.Enabled = True
        Me.mnuReq.Enabled = True
        Me.mnuAdvice.Enabled = True
        Me.mnuMoney.Enabled = True
        With lvwPati.SelectedItem
'            mnuExecFunc(4).Enabled = Not (InStr("1", .ListSubItems(3).Tag) > 0 Or _
'                 (.ListSubItems(3).Tag = 3 And Val(.ListSubItems(9).Tag) > 1))
            mnuExecFunc(1).Enabled = InStr("0", .ListSubItems(3).Tag) > 0
            
            mnuExecFunc(4).Enabled = Not (InStr("1", .ListSubItems(3).Tag) > 0)
            
            mnuExecFunc(4).Caption = IIf(.ListSubItems(3).Tag = 3 And Val(.ListSubItems(9).Tag) > 1, "�޸���Ϣ(&A)", "��ʼ���(&A)")
            
            mnuExecFunc(5).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            tbrMain.Buttons("�ɼ�").Enabled = mnuExecFunc(5).Enabled
            
            mnuExecFunc(6).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            
            mnuExecFunc(7).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            
            mnuExecFunc(9).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(10).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(11).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(12).Enabled = Not (InStr("0,3", .ListSubItems(3).Tag) = 0 Or InStr("1,2", .ListSubItems(9).Tag) = 0)
            
            mnuExecFunc(14).Enabled = Not (.ListSubItems(3).Tag <> "3" Or .ListSubItems(9).Tag <> "2")
            
            mnuExecFunc(15).Enabled = Not (.ListSubItems(3).Tag <> "3" Or Val(.ListSubItems(9).Tag) < 3)
            
            mnuExecFunc(17).Enabled = mnuRepFunc(0).Enabled
                
'            mnuRepFunc(0).Enabled = Not (InStr("1,3", .ListSubItems(3).Tag) = 0 Or _
'                InStr("3,4,5,6", .ListSubItems(9).Tag) = 0)
'            mnuRepFunc(3).Enabled = Not (.ListSubItems(3).Tag <> "3" Or _
'                InStr("4", .ListSubItems(9).Tag) = 0)
            mnuRepFunc(3).Enabled = Not (.ListSubItems(3).Tag <> "3")
            mnuExecFunc(18).Enabled = mnuRepFunc(3).Enabled
'            mnuRepFunc(4).Enabled = Not (InStr("1,3", .ListSubItems(3).Tag) = 0 Or _
'                InStr("4,6", .ListSubItems(9).Tag) = 0)
            mnuRepFunc(4).Enabled = Not InStr("1,3", .ListSubItems(3).Tag) = 0
        
            If Val(Split(.Tag, "_")(1)) = 0 And Len(Split(.Tag, "_")(2)) = 0 Then Me.mnuAdvice.Visible = False
            int������Դ = IIf(.Text = "סԺ", 2, 1)
            If TabFile.SelectedItem.Key = "ҽ��" Then
                mnuAdviceFunc(4).Visible = int������Դ = 2
                mnuAdviceFunc(5).Visible = int������Դ = 2
                mnuAdviceFunc(6).Visible = Not (int������Դ = 2)
                mnuAdviceFunc(7).Visible = Not (int������Դ = 2)
            Else
                mnuAdvice.Visible = False
            End If
        End With
        Me.mnuViewInfo.Enabled = True
    Else
        Me.mnuExec.Enabled = True
        For i = 1 To mnuExecFunc.Count - 1
            mnuExecFunc(i).Enabled = False
        Next
        Me.mnuRep.Enabled = False
        Me.mnuPFile.Enabled = False
        Me.mnuReq.Enabled = False
        Me.mnuAdvice.Enabled = False
        Me.mnuMoney.Enabled = False
        Me.mnuViewInfo.Enabled = False
    
        tbrMain.Buttons("�ɼ�").Enabled = False
    End If
    
    With Me.tbrMain.Buttons
        For i = 1 To .Count
            Select Case .Item(i).Description
                Case "����"
                    .Item(i).Enabled = mnuMoney.Enabled
                Case "ҽ��"
                    .Item(i).Enabled = mnuAdvice.Enabled
                    .Item(i).Visible = mnuAdvice.Visible
                Case "����"
                    .Item(i).Enabled = mnuPFile.Enabled
                Case "Ӱ��"
                    .Item(i).Enabled = mnuRep.Enabled
            End Select
        Next
    End With

    Me.tbrMain.Buttons("����").Enabled = mnuRepFunc(0).Enabled
    Me.tbrMain.Buttons("���").Enabled = mnuRepFunc(3).Enabled
    Me.tbrMain.Buttons("����").Enabled = mnuRepFunc(4).Enabled
End Sub

Private Sub OutputList(bytStyle As Byte)
'����: ������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrintLvw

    On Error Resume Next
    If lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Me.TabFile.SelectedItem.Key
        Case "����"
            mfrmActive.zlPrint bytStyle
            Exit Sub
        Case "ҽ��"
            Select Case bytStyle
                Case 1
                    mfrmActive.zlPrint
                Case 2
                    mfrmActive.zlPreview
                Case 3
                    mfrmActive.zlExcel
            End Select
            Exit Sub
    End Select

    Set objOut.Body.objData = Me.lvwPati
    objOut.Title.Text = Decode(Val(lvwPati.Tag), 0, "��ִ��", 1, "������", 2, "�����") & _
        "����嵥"
    If bytStyle = 1 Then
        bytStyle = zlPrintAsk(objOut)
        If bytStyle <> 0 Then zlPrintOrViewLvw objOut, bytStyle
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
End Sub

Private Sub BeginShowProgress(ByVal strCaption As String)
    With pgbLoad
        .Left = stbThis.Panels(2).Left + Me.TextWidth(strCaption) + 200
        .Top = stbThis.Top + (stbThis.Height - .Height) / 2
        .Width = stbThis.Panels(2).Width + stbThis.Panels(2).Left - .Left
        .Value = 0
        
        stbThis.Panels(2).Text = strCaption
        .Visible = Me.stbThis.Visible: Me.Refresh
    End With
End Sub

'��������ӵĲ����ļ��˵�
Private Sub AddFileList()
    Dim rsFileList As ADODB.Recordset
    Dim i As Integer, iNum As Integer
    
    '����ļ��嵥
    iNum = FileList.Count
    FileList(0).Visible = True
    For i = 1 To iNum - 1
        Unload FileList(i)
    Next
    
    Set rsFileList = GetPatientFileList(UserInfo.����ID, 0)
    If Not rsFileList Is Nothing Then
        i = 1
        Do While Not rsFileList.EOF
            Load FileList(FileList.Count)
            With FileList(FileList.Count - 1)
                .Caption = "&" & i & " " & rsFileList("����")
                .Tag = "O" & rsFileList("ID")
                .Enabled = True
                .Visible = True
            End With
            
            i = i + 1
            rsFileList.MoveNext
        Loop
        
        On Error Resume Next
        FileList(0).Visible = False
    End If
    Set rsFileList = GetPatientFileList(UserInfo.����ID, 1)
    If Not rsFileList Is Nothing Then
        i = 1
        Do While Not rsFileList.EOF
            Load FileList(FileList.Count)
            With FileList(FileList.Count - 1)
                .Caption = "&" & i & " " & rsFileList("����")
                .Tag = "I" & rsFileList("ID")
                .Enabled = True
                .Visible = True
            End With
            
            i = i + 1
            rsFileList.MoveNext
        Loop
        
        On Error Resume Next
        FileList(0).Visible = False
    End If
End Sub

Private Sub ShowAddFileMenu(ByVal SrcType As Integer)
    Dim i As Integer, iNum As Integer
    Dim blnOutVisible As Boolean
    
    If SrcType = 1 Then blnOutVisible = True
    
    '����ļ��嵥
    iNum = FileList.Count
    For i = 1 To iNum - 1
        FileList(i).Visible = True
    Next
    For i = 1 To iNum - 1
        If FileList(i).Tag Like "O*" Then
            FileList(i).Visible = blnOutVisible
        Else
            FileList(i).Visible = Not blnOutVisible
        End If
    Next
End Sub

Private Function LoadBillList() As Boolean
'���ܣ���ȡ��ǰ���õĸ��ﵥ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objMenu As Menu
    
    On Error GoTo errH
    
    '������е����嵥
    For i = Me.ReqList.UBound To 0 Step -1
        Me.ReqList(i).Tag = ""
        If i = 0 Then
            ReqList(i).Caption = "<�޿��õ���>"
        Else
            Unload ReqList(i)
        End If
    Next
    
    '���ؿ��õ���
    strSQL = "Select Distinct A.ID,A.���,A.����,A.˵��" & _
        " From �����ļ�Ŀ¼ A,�����ļ���� B" & _
        " Where A.����=5 And A.ǰ�� IN(2,3)" & _
        " And A.ID=B.�����ļ�ID And B.��дʱ�� IN(1,2)" & _
        " Order by A.���"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If i <> 1 Then Load ReqList(ReqList.UBound + 1)
        Set objMenu = ReqList(ReqList.UBound)
        objMenu.Caption = rsTmp!����
        If i <= 10 Then
            objMenu.Caption = objMenu.Caption & "(&" & i - 1 & ")"
        ElseIf i <= 36 Then
            objMenu.Caption = objMenu.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
        End If
        objMenu.Tag = rsTmp!ID: objMenu.Enabled = True
        rsTmp.MoveNext
    Next
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPart(ByVal strAdvice As String) As String
'���ܣ�����ҽ�����ݻ�ȡ�걾��λ
    Dim iPos As Integer, iPos1 As Integer
    iPos = 0
    Do While True
        iPos1 = InStr(iPos + 1, strAdvice, "(")
        If iPos1 = 0 Then Exit Do
        iPos = iPos1
    Loop
    If iPos > 0 Then
        GetPart = Mid(strAdvice, iPos + 1, Len(strAdvice) - iPos - 1)
    Else
        GetPart = ""
    End If
End Function
Public Function GetReprotFrm() As Form
    Set GetReprotFrm = mfrmRepEdit
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.Hwnd)
End Sub


Private Sub subCancelSeriesRelate(lngAdviceNo As Long, lngSendNO As Long, strSeriesNo As String)
'-----------------------------------------------------------------------------
'����:ȡ������ͼ��Ĺ���
'�޸���:�ƽ�
'�޸�����:2007-1-30
'-----------------------------------------------------------------------------
    
    Dim mcnFTP As New clsFtp
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim strCacheFileName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    Dim strNewStudyUID As String    '�����ɵļ��UID
    Dim strOldStudyUID As String    'ͼ������ԭ���ļ��UID
    Dim strDBStudyUID As String     '���ݿ��б���ļ��UID����ͼ��洢·�����
    Dim strMoveFiles As String  '�洢��Ҫ�ƶ���ͼ���ļ�����ʹ�á�|���ָ�
    Dim blnNoImage As Boolean   '1û��ͼ��ֱ�Ӷ�ȡ���ݿ���Ϣ��0��ͼ��ʹ��ͼ����Ϣ
    
    'ͼ���еĲ��˻�����Ϣ
    Dim strModality As String
    Dim strPatientID As String
    Dim strPatientName As String
    Dim strSex As String
    Dim strAge As String
    Dim strDateOfBirth As String
    Dim strManufacturer As String
    Dim strReceiveDateTime As String
    
    
    
    '���������е�һ��ͼ��� ����ID��Ӣ�������Ա����䣬�������ڣ����UID������豸������ʱ��
    strCachePath = App.Path & "\TmpImage\"
    strSQL = "Select A.ͼ���,D.�û��� As User1,D.���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1,c.���uid," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2 , e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.����UID= [1] Order By A.ͼ���"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strSeriesNo)
    
    If Not rsTmp.EOF Then   '�����д���ͼ��
        strDBStudyUID = Nvl(rsTmp("���uid"))
        '����ͼ��
        If rsTmp("�豸��1") <> "" Then
            mcnFTP.FuncFtpConnect rsTmp("Host1"), rsTmp("User1"), rsTmp("Pwd1")
            strCacheFileName = strCachePath & objFile.GetFileName(rsTmp("URL1"))
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strCacheFileName, objFile.GetFileName(rsTmp("URL1"))
            mcnFTP.FuncFtpDisConnect
        ElseIf rsTmp("�豸��2") <> "" Then
            mcnFTP.FuncFtpConnect rsTmp("Host2"), rsTmp("User2"), rsTmp("Pwd2")
            strCacheFileName = strCachePath & objFile.GetFileName(rsTmp("URL2"))
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strCacheFileName, objFile.GetFileName(rsTmp("URL2"))
            mcnFTP.FuncFtpDisConnect
        End If
        '��ȡͼ��
        If Dir(strCacheFileName) <> vbNullString Then
            Set img = imgs.ReadFile(strCacheFileName)
            '-----------�Ƿ�ʹ�ñ�����ͼ�������Ϣ��ȡ������
            strOldStudyUID = img.StudyUID
            strModality = GetImageAttribute(img.Attributes, ATTR_Ӱ�����)
            strPatientID = img.PatientID
            strPatientName = img.Name
            strSex = img.Sex
            If IsDate(img.DateOfBirthAsDate) Then
                strAge = CStr(Year(Date) - Year(img.DateOfBirthAsDate))
                strDateOfBirth = Format(img.DateOfBirthAsDate, "YYYY-MM-DD")
            Else
                strAge = "": strDateOfBirth = ""
            End If
            strManufacturer = GetImageAttribute(img.Attributes, ATTR_����豸)
            strReceiveDateTime = GetImageAttribute(img.Attributes, ATTR_�������) & " " & _
                        Format(GetImageAttribute(img.Attributes, ATTR_���ʱ��), "HH:MM")
            'ɾ����ʱͼ��
            Set img = Nothing
            imgs.Remove (1)
            objFile.DeleteFile strCacheFileName
        Else
            '�����һ��ͼ�����ز���ȷ����ȡ���ݿ���Ϣ
            blnNoImage = True
        End If
    Else
        '������û��ͼ��ֻ��ʹ�ñ����������ݿ��е�ֵ
        blnNoImage = True
    End If
    
    '����û��ͼ����Ϣ�ɶ�ȡ�������ֱ�Ӷ�ȡ���ݿ��е���Ϣ
    If blnNoImage = True Then
        strSQL = "select a.Ӱ�����,a.����,a.����,a.Ӣ����,a.�Ա�,a.����,a.��������,a.���uid," & _
                " a.����豸,a.�������� from Ӱ�����¼ a where a.ҽ��id =[1] and a.���ͺ� =[2]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceNo, lngSendNO)
        If Not rsTmp.EOF Then
            strOldStudyUID = Nvl(rsTmp("���uid"))
            strDBStudyUID = Nvl(rsTmp("���uid"))
            strModality = Nvl(rsTmp("Ӱ�����"))
            strPatientID = Nvl(rsTmp("����"))
            strPatientName = Nvl(rsTmp("Ӣ����"))
            strSex = Nvl(rsTmp("�Ա�"))
            strAge = Nvl(rsTmp("����"))
            strDateOfBirth = Nvl(rsTmp("��������"), "1899-12-30")
            strManufacturer = Nvl(rsTmp("����豸"))
            strReceiveDateTime = Nvl(rsTmp("��������"))
        End If
    End If
    '��֯ͼ���ļ����ƴ�
    strSQL = "select ͼ��UID from Ӱ�������� a,Ӱ����ͼ�� b where a.����UID =[1] and a.����UID = b.����UID"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strSeriesNo)
    If Not rsTmp.EOF Then
        strMoveFiles = rsTmp(0)
        rsTmp.MoveNext
        While Not rsTmp.EOF
            strMoveFiles = strMoveFiles & "|" & rsTmp(0)
            rsTmp.MoveNext
        Wend
    End If
    
    '������UID�����ݿ����ִ�ļ��UID��ͬ���򴴽��µļ��UID�����޸�ͼ��FTP·��
    strNewStudyUID = funGetStudyUID(strOldStudyUID)
    If strNewStudyUID <> strDBStudyUID Then
        Call MergeImageFiles(strDBStudyUID, strNewStudyUID, Format(strReceiveDateTime, "YYYY-MM-DD"), strMoveFiles)
    End If
    
    '�޸����ݿ⣬������¼ת����ʱ��¼
        strSQL = "ZL_Ӱ����_PhotoCancel(" & lngAdviceNo & "," & lngSendNO & ",'" & strNewStudyUID & "','" & _
                  strSeriesNo & "','" & strModality & "'," & Val(strPatientID) & ",'" & _
                  strPatientName & "','" & strSex & "','" & strAge & "'," & _
                  IIf(Len(strDateOfBirth) = 0, "null", "to_date('" & strDateOfBirth & "','YYYY-MM-DD')") & _
                  ",'" & strManufacturer & "',to_date('" & strReceiveDateTime & "','YYYY-MM-DD HH24:MI:SS'))"
                  
        ExecuteProc strSQL, Me.Caption
End Sub

Private Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'����:��ѯ���ݿ⣬�жϵ�ǰͼ��ļ��UID�Ƿ��Ѿ����������������ʱ���У�
'     ������ڣ����ڼ��UID�������Ӻ�׺����������ֱ�ӷ�������ļ��UID
'�޸���:�ƽ�
'�޸�����:2007-1-27
'-----------------------------------------------------------------------------
    '
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select ���UID from Ӱ�����¼ where ���UID = [1]" & _
              " Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = [1]"
    Set rsMatch = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strOldStudyUID)
    If Not rsMatch.EOF Then
        '����һ���µļ��UID
        gstrSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsMatch = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function


Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
'-----------------------------------------------------------------------------
'����:��ȡDICOM���Լ��е�ָ������ֵ
'�޸���:�ƽ�
'�޸�����:2007-2-6
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        GetImageAttribute = Nvl(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Value)
    End If
End Function
