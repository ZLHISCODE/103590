VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeptBilling 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "���ҷ�ɢ����"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9870
   Icon            =   "frmDeptBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmDeptBilling.frx":08CA
   ScaleHeight     =   6195
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2730
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   7140
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3930
      Width           =   7140
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   5835
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDeptBilling.frx":0A58
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmDeptBilling.frx":0DCC
            Text            =   "״̬˵��"
            TextSave        =   "״̬˵��"
            Key             =   "state"
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9870
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6195
      MinHeight1      =   720
      Width1          =   4500
      NewRow1         =   0   'False
      Caption2        =   "���˿���"
      Child2          =   "cboDept"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1800
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   7785
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   6615
         _ExtentX        =   11668
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
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Billing"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Billing"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BillingBilling"
                     Object.Tag             =   "���ʵ�"
                     Text            =   "���ʵ�"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BillingTable"
                     Object.Tag             =   "���ʱ�"
                     Text            =   "���ʱ�"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "BillingSimple"
                     Object.Tag             =   "�򵥼���"
                     Text            =   "�򵥼���"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "�Ե�ǰѡ�е�������"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "Style"
               Description     =   "�鿴"
               Object.ToolTipText     =   "�鿴"
               Object.Tag             =   "�鿴"
               ImageKey        =   "Style"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "��ͼ��"
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "Сͼ��"
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "�б�"
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "��ϸ����"
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   2670
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5025
      ScaleWidth      =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   810
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4575
      Left            =   -15
      TabIndex        =   0
      ToolTipText     =   "����ʱ,Ĭ����ʾ7�����ڵĲ���"
      Top             =   1260
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   8070
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "��Ա��"
         Text            =   "סԺ��"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "�Ա�"
         Text            =   "�Ա�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Key             =   "��Ժ����"
         Text            =   "��Ժ����"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Key             =   "��Ժ����"
         Text            =   "��Ժ����"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "��ǰ����"
         Text            =   "��ǰ����"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Key             =   "סԺ"
         Text            =   "סԺ"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "ҽ�Ƹ��ʽ"
         Text            =   "ҽ�Ƹ��ʽ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "��ǰ����ID"
         Text            =   "��ǰ����ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "��������"
         Text            =   "��������"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":0F6A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":1184
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":139E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":15B8
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":1D32
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":1F4C
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":2166
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":2380
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":259A
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":27B4
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":2EAE
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":35A8
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":3CA2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":3EBC
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":40D6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":42F0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":450A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":4724
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":4E9E
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":50B8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":52D2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":54EC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":5706
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":5920
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":601A
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":6714
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":6E0E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":7028
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   975
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   635
      TabFixedWidth   =   2290
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      TabMinWidth     =   882
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ǰ��Ժ(&1)"
            Key             =   "InHos"
            Object.ToolTipText     =   "��ǰ��Ժ�Ĳ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ժ����(&2)"
            Key             =   "OutHos"
            Object.ToolTipText     =   "�ڼ��ڳ�Ժ�Ĳ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ת������(&3)"
            Key             =   "ת��"
            Object.Tag             =   "ת��"
            Object.ToolTipText     =   "ת������"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3105
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":7242
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":7B1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3690
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":83F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptBilling.frx":8CD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2955
      Left            =   2700
      TabIndex        =   1
      Top             =   960
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   5212
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmDeptBilling.frx":95AA
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   2715
      TabIndex        =   2
      Top             =   3960
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3307
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmDeptBilling.frx":98C4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00808080&
      Caption         =   " ���˷��øſ�"
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   2775
      TabIndex        =   10
      Top             =   765
      Width           =   6990
   End
   Begin VB.Label lbl_s 
      BackColor       =   &H00808080&
      Caption         =   " ʱ��:2001-01-01��2001-01-01"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   30
      TabIndex        =   9
      ToolTipText     =   "�ڸ�ʱ�䷶Χ�ڵ�סԺ����"
      Top             =   750
      Width           =   2580
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditBilling 
         Caption         =   "���ʵ�(&B)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditTable 
         Caption         =   "���ʱ�(&T)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSimple 
         Caption         =   "�򵥼���(&S)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditCust 
         Caption         =   "�Զ�����ʵ�(&U)"
         Begin VB.Menu mnuEditCustBill 
            Caption         =   "(��)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "�޸ĵ���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "���Ƶ���(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "����ʱ��(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditAdjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "��������(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDelApply 
         Caption         =   "��������(&Q)"
      End
      Begin VB.Menu mnuEditDelAudit 
         Caption         =   "�������(&H)"
      End
      Begin VB.Menu mnuEditDel_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "���ĵ���(&V)"
      End
      Begin VB.Menu mnuEditPrint 
         Caption         =   "��ӡ����(&P)"
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
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "סԺ����(&U)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "��ʾ���˷�ʽ(&M)"
         Begin VB.Menu mnuViewByDept 
            Caption         =   "��������ʾ(&U)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewByDept 
            Caption         =   "��������ʾ(&U)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "��ͼ��(&I)"
         Index           =   0
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView_6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "ˢ�·�ʽ(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "������Ҫˢ������(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "��������ʾ�Ƿ�ˢ��(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
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
Attribute VB_Name = "frmDeptBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset
Private mrsPati As ADODB.Recordset

Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Operator As String
    FeeItems As String
    IncomeItems As String
    lngHospNo As Long 'סԺ��
    strPatiName As String
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String
Private mintBedLen As Integer
Private mdtBegin As Date, mdtEnd As Date
Private mbln���� As Boolean, mbln���� As Boolean
Private mstrҽ����Ч As String

Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long

Private mblnFirst As Boolean, mblnMax As Boolean
Private mlngDeptID As Long, mlngUnitID As Long
Private mstrPage As String

Private mstrPrivs As String     '���浱ǰģ�����Ȩ����
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mlngModul As Long
Private mblnNOMoved As Boolean '��¼��ǰѡ��ĵ����Ƿ����ں����ݱ���
Private mblnNotClick As Boolean
Private mrsDept As ADODB.Recordset
'���˺� ����:27380 ����:2010-01-22 15:11:16
Private Type Ty_Para
    blnת������ As Boolean
    intת������ As Integer
End Type
Private mTy_Modul_Para As Ty_Para

Private Sub zlSetPatiPages()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת����ҳ�����ʾ
    '����:���˺�
    '����:2010-01-27 09:48:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnת�� As Boolean, blnHaveData As Boolean, i As Long
    
    blnת�� = mTy_Modul_Para.blnת������
    blnת�� = blnת�� And IIf(mnuViewByDept(0).Checked, mlngUnitID > 0, mlngDeptID > 0)
    
    blnHaveData = False
    For i = 1 To tbs.Tabs.Count
        If tbs.Tabs(i).Key = "ת��" Then
            blnHaveData = True: Exit For
        End If
    Next
    If blnת�� Then
        If blnHaveData = False Then
            tbs.Tabs.Add , "ת��", "ת������(&3)"
        End If
    Else
        If blnHaveData Then
            '�Ƴ�
            If tbs.SelectedItem.Index = i Then
               tbs.Tabs(1).Selected = True
            End If
            tbs.Tabs.Remove i
        End If
    End If
End Sub
Private Sub cboDept_Click()
    Dim strTmp As String
    If mnuViewByDept(0).Checked Then
        If cboDept.ItemData(cboDept.ListIndex) = mlngUnitID Then Exit Sub
        mlngUnitID = cboDept.ItemData(cboDept.ListIndex)
        '��ǰ������ѡ���Ĳ��˵Ŀ�����ȷ��
    Else
        If cboDept.ItemData(cboDept.ListIndex) = mlngDeptID Then Exit Sub
        mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
        If mlngDeptID = 0 Then
            mlngUnitID = 0
        Else
            mlngUnitID = Get����ID(mlngDeptID)
        End If
    End If
    mstrPage = ""
    Call zlSetPatiPages
        
    If Visible Then Call tbs_Click
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    If KeyAscii <> 13 Then Exit Sub
    
    If cboDept.ListIndex <> -1 Then
        ZLCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Call InitUnits
    
    
    Dim strRootCaption As String
    strRootCaption = ""
    If InStr(mstrPrivs, ";���в���;") > 0 Then strRootCaption = IIf(mnuViewByDept(1).Checked, "���п���", "���в���")
    
    
    If zlSelectDept(Me, mlngModul, cboDept, mrsDept, cboDept.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    Dim lngID As Long
    
    If cboDept.ListIndex >= 0 Then Exit Sub
    
   If mnuViewByDept(0).Checked Then
        lngID = mlngUnitID
        '��ǰ������ѡ���Ĳ��˵Ŀ�����ȷ��
   Else
       lngID = mlngDeptID
   End If
   zlControl.CboLocate cboDept, lngID, True
   If cboDept.ListIndex < 0 And cboDept.ListCount <> 0 Then cboDept.ListIndex = 0
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    Call InitLocPar(mlngModul)
    If mblnFirst Then
        '����:29435:��Ҫ��Ҫ����ת��ҳ��
        Call zlSetPatiPages
        If lvw.Visible And lvw.Enabled Then lvw.SetFocus
        mshList_GotFocus
        
        mblnFirst = False
    End If
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2010-01-22 15:12:34
    '����:27380
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim i As Long, blnHaveData As Boolean
    
    strTemp = zlDatabase.GetPara("���ת������", glngSys, mlngModul, "0|3")
    mTy_Modul_Para.intת������ = Val(Split(strTemp & "|", "|")(1))
    mTy_Modul_Para.blnת������ = IIf(Val(Split(strTemp & "|", "|")(0)) = 1, True, False)
    blnHaveData = False
    For i = 1 To tbs.Tabs.Count
        If tbs.Tabs(i).Key = "ת��" Then
            blnHaveData = True: Exit For
        End If
    Next
    If blnHaveData Then
        tbs.Tabs("ת��").ToolTipText = "��ʾ" & mTy_Modul_Para.intת������ & "��ת���Ĳ���"
    End If
End Sub


Private Sub mnuEditAdjust_Click()
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Ե�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Զ����ʵ���ֹ����
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) = 3 Then
        MsgBox "�õ���Ϊ�Զ����ʵ�,���ܵ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    '�Ѿ�������(����)�ĵ��ݲ��������
    If BillExistDelete(strNO, 2) Then
        MsgBox "�õ��ݰ�������������,�����������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ��Ѿ�����
    If HaveBilling(2, strNO) <> 0 Then
        Select Case gbytBillOpt
            Case 0
            Case 1
                If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Case 2
                MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,���ܵ�����", vbExclamation, gstrSysName: Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Err.Clear
        
    If BillisBatch(strNO) Then '��������
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 2
        frmBillings.mstrInNO = strNO
        frmBillings.mbytUseType = 1
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmBillings.mlng����ID = CLng(lvw.SelectedItem.Tag)
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 2
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmSimpleBilling.mlng����ID = CLng(lvw.SelectedItem.Tag)
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        Dim lng����ID  As Long
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 2
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 1
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng����ID = CLng(lvw.SelectedItem.Tag)
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            If Not lvw.SelectedItem Is Nothing Then lng����ID = CLng(lvw.SelectedItem.Tag)
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 1, 2, strNO, mlngUnitID, mlngDeptID, lng����ID, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuEditBilling_Click()
    Dim cur��� As Currency, blnOut As Boolean
        
    '��Ժ���˼���Ȩ��
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    
    If blnOut Then
        cur��� = Get�������(CLng(lvw.SelectedItem.Tag), 0)
        If cur��� = 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur��� <> 0 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mbytUseType = 1
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    
    If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng����ID = CLng(lvw.SelectedItem.Tag)
    
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditCopy_Click()
    Dim strNO As String, cur��� As Currency, blnOut As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Ը��ƣ�", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '��Ժ���˼���Ȩ��
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    
    If blnOut Then
        cur��� = Get�������(CLng(lvw.SelectedItem.Tag), 0)
        If cur��� = 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur��� <> 0 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytUseType = 1
    frmCharge.mbytInState = 0
    frmCharge.mblnCopyBill = True
    frmCharge.mstrInNO = strNO
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    
    If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng����ID = CLng(lvw.SelectedItem.Tag)
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditCustBill_Click(Index As Integer)
    '�Զ������
    Dim lng����ID As Long, varTemp As Variant
    Dim cur��� As Currency, blnOut As Boolean
    
    '��Ժ���˼���Ȩ��
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    If blnOut Then
        cur��� = Get�������(CLng(lvw.SelectedItem.Tag), 0)
        If cur��� = 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur��� <> 0 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�������������ǣ�
    '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs��blnViewCancel
    
    If Not lvw.SelectedItem Is Nothing Then lng����ID = CLng(lvw.SelectedItem.Tag)
    
    varTemp = Array(mnuEditCustBill(Index).Tag, 1, 0, "", mlngUnitID, mlngDeptID, lng����ID, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '����ֵ
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditDelApply_Click()
    Dim lngPatientID As Long
    
    If mlngUnitID = 0 Then
        MsgBox "����ѡ���˲���!", vbInformation, gstrSysName
        cboDept.SetFocus
        Exit Sub
    End If
    If Not lvw.SelectedItem Is Nothing Then lngPatientID = Val(lvw.SelectedItem.Tag)
    
    With frmReCharge
        .mlngDeptID = mlngUnitID
        .mbytUseType = 0
        .mbytFun = 0
        .mstrPrivs = mstrPrivs
        .mlngPatientID = lngPatientID
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditDelAudit_Click()
    If mlngUnitID = 0 Then
        MsgBox "����ѡ���˲���!", vbInformation, gstrSysName
        cboDept.SetFocus
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngUnitID
        .mbytUseType = 0
        .mbytFun = 1
        .mstrPrivs = mstrPrivs
        .Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End With
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditPrint_Click()
    Dim strNO As String, strTime As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If

    If Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) = 3 Then
        MsgBox "�õ���Ϊ�Զ����ʵ�,�������ܼ�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    If mshList.TextMatrix(mshList.Row, GetColNum("����")) <> 1 Then
        MsgBox "�õ���Ϊ���ʵ��ݻ��ѱ����ʣ������ٴ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1134", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1134", Me, "NO=" & strNO, "�Ǽ�ʱ��=" & strTime, "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=1", 2)
    End If
End Sub

Private Sub mnuEditSimple_Click()
    Dim cur��� As Currency, blnOut As Boolean
    
    '��Ժ���˼���Ȩ��
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    If blnOut Then
        cur��� = Get�������(CLng(lvw.SelectedItem.Tag), 0)
        If cur��� = 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur��� <> 0 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mbytUseType = 1
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    If Not lvw.SelectedItem Is Nothing Then frmSimpleBilling.mlng����ID = CLng(lvw.SelectedItem.Tag)
    
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditTable_Click()
    Dim cur��� As Currency, blnOut As Boolean
    
    '��Ժ���˼���Ȩ��
    If tbs.SelectedItem.Index = 2 And Not lvw.SelectedItem Is Nothing Then
        blnOut = True
    ElseIf tbs.SelectedItem.Index = 1 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    ElseIf tbs.SelectedItem.Index = 3 And Not lvw.SelectedItem Is Nothing Then
        If Val(lvw.SelectedItem.ListSubItems(1).Tag) = 3 Then blnOut = True
    End If
    If blnOut Then
        cur��� = Get�������(CLng(lvw.SelectedItem.Tag), 0)
        If cur��� = 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷����Ѿ�����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf cur��� <> 0 And InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") = 0 Then
            MsgBox "�ó�Ժ(��Ԥ��Ժ)���˷�����δ����,��û��Ȩ�޶Ըò��˼��ʣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mbytUseType = 1
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    
    If Not lvw.SelectedItem Is Nothing Then frmBillings.mlng����ID = CLng(lvw.SelectedItem.Tag)
    
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuEditModi_Click()
    Dim strNO As String, strInfo As String, strUnitIDs As String
    Dim strInsure As String, arrInsure As Variant
    Dim i As Long
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ����޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
                
    'δȫ����˻�����˵Ĳ������޸�
    If Not BillIdentical(strNO) Then
        MsgBox "�����а�������δ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
        Exit Sub
    End If
                
    '�����޸�Ȩ��
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))), "�޸�", strNO) Then Exit Sub
        
    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, mstrPrivsOpt)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��Ժ���˲���Ȩ���ж�
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "�޸�") Then Exit Sub
    
    
    'ȫԺ����
    If InStr(mstrPrivsOpt, ";ȫԺ����;") = 0 Then
        If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
        
        If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(mshList.Row, GetColNum("��������ID"))) & ",") = 0 Then
            MsgBox "��û��Ȩ�޶��������ҵĵ�������,�������޸ĸõ��ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
                    
    'ȥ����ҽ������ƥ����
                
    '����������ʱ��ҩƷ�ĵ��ݽ�ֹ�޸�
    If Not BillCanModi(strNO, 2) Then
        MsgBox "���ŵ����а���������ʱ��ҩƷ,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
                
    '�Ѿ�������(����)�ĵ��ݲ������޸�
    If BillExistDelete(strNO, 2) Then
        MsgBox "�õ��ݰ��������ʷ���,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�����������ִ�л�ȫ��ִ�е���Ŀ,��һ������ȫ������,�������޸�
    If HaveExecute(2, strNO, 2) Then
        MsgBox "�õ����а�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '�Ƿ��Ѿ����ʵ�
    If HaveBilling(2, strNO) <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , arrInsure(i)) Then
                        'ҽ�����˵ĵ��ݹ̶�Ϊ�ѽ��ʾͽ�ֹ�޸�
                        MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ�޸���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        Case 2
                            MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
                    End Select
                End If
            Next
        End If
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
        
    gbytBilling = 0 '�����޸�
    If BillisBatch(strNO) Then '��������
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 0
        frmBillings.mstrInNO = strNO
        frmBillings.mbytUseType = 1
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmBillings.mlng����ID = CLng(lvw.SelectedItem.Tag)
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 0
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        If Not lvw.SelectedItem Is Nothing Then frmSimpleBilling.mlng����ID = CLng(lvw.SelectedItem.Tag)
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        Dim lng����ID  As Long
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 0
            frmCharge.mstrInNO = strNO
            frmCharge.mbytUseType = 1
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            
            If Not lvw.SelectedItem Is Nothing Then frmCharge.mlng����ID = CLng(lvw.SelectedItem.Tag)
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            If Not lvw.SelectedItem Is Nothing Then lng����ID = CLng(lvw.SelectedItem.Tag)
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 1, 0, strNO, mlngUnitID, mlngDeptID, lng����ID, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,�޸ĺ�ĵ��ݺ�Ϊ:[" & gstrModiNO & "],Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ShowBills(mstrFilter)
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                Call ShowBills(mstrFilter)
            End If
        End If
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim bln�������� As Boolean
    Dim blnסԺ���� As Boolean
    Dim blnסԺ��λ As Boolean
    Dim blnת�� As Boolean
    
    bln�������� = gbln��������
    blnסԺ���� = gblnסԺ����
    blnסԺ��λ = gblnסԺ��λ
    
    blnת�� = mTy_Modul_Para.blnת������
    
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytUseType = 1
    frmSetExpence.mbytInFun = 0
    frmSetExpence.Show 1, Me
    If gblnOK Then
        '����:27380
        Call InitPara
        Call zlSetPatiPages
        
        If bln�������� <> gbln�������� Or blnסԺ���� <> gblnסԺ���� Then
            mlngDeptID = -1: mstrPage = ""
            Call InitUnits
        ElseIf blnסԺ��λ <> gblnסԺ��λ Then
            If Not (mshList.Rows = 2 And mshList.TextMatrix(1, GetColNum("���ݺ�")) = "") Then
                Call mnuViewReFlash_Click
                blnת�� = mTy_Modul_Para.blnת������
            End If
        End If
        If blnת�� <> mTy_Modul_Para.blnת������ And tbs.SelectedItem.Index = 3 Then
             Call mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����ID As Long, lng��ҳID As Long, strסԺ�� As String
    Dim strNO As String
    
    If Not lvw.SelectedItem Is Nothing Then
        '����:29444
        lng����ID = Val(Split(lvw.SelectedItem.Key, "_")(1))
        lng��ҳID = Val(Split(lvw.SelectedItem.Key, "_")(2))
        
        strסԺ�� = lvw.SelectedItem.SubItems(1)
        
        strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
        If strNO = "" Then
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, "����=" & mlngUnitID, "���˿���=" & mlngDeptID, "סԺ��=" & strסԺ��)
        Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, "����=" & mlngUnitID, "���˿���=" & mlngDeptID, _
                "NO=" & strNO, "������=" & mshList.TextMatrix(mshList.Row, GetColNum("������")), "סԺ��=" & strסԺ��)
        End If
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "����=" & mlngUnitID, "���˿���=" & mlngDeptID)
    End If
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    mlngDeptID = 0: mlngUnitID = 0
    Call InitUnits
End Sub

Private Sub mnuViewFilter_Click()
    With frmDeptFilter
        .mstrPrivs = mstrPrivs
        If .mlngDeptID <> mlngDeptID Then
            .mlngDeptID = mlngDeptID
            .mlngUnitID = mlngUnitID
            .GetOperator    '����ʽ����form_load�¼�
        End If
        .Show 1, Me
        If gblnOK Then
            mdtBegin = Format(.dtpB.Value, "yyyy-MM-dd 00:00:00")
            mdtEnd = Format(.dtpE.Value, "yyyy-MM-dd 23:59:59")
            
            mstrFilter = .mstrFilter
            mbln���� = .chkType(0).Value = 1
            mbln���� = .chkType(1).Value = 1
            
             If .chkBill(chkBills.��������).Value = 1 And .chkBill(chkBills.��������).Value = 1 Then
                If .chkBill(chkBills.��ͨ����).Value = 0 And .chkBill(chkBills.�Զ�����).Value = 0 Then
                    mstrҽ����Ч = " And D.ҽ����Ч In(0,1)"
                ElseIf .chkBill(chkBills.��ͨ����).Value = 0 Then
                    mstrҽ����Ч = " And (A.��¼����=2 And D.ҽ����Ч In(0,1) Or A.��¼����=3)"
                Else
                    mstrҽ����Ч = ""
                End If
            ElseIf .chkBill(chkBills.��������).Value = 1 Then
                If .chkBill(chkBills.��ͨ����).Value = 0 And .chkBill(chkBills.�Զ�����).Value = 0 Then
                    mstrҽ����Ч = " And D.ҽ����Ч=1"
                ElseIf .chkBill(chkBills.��ͨ����).Value = 0 Then
                    mstrҽ����Ч = " And (A.��¼����=2 And D.ҽ����Ч=1 Or A.��¼����=3)"
                Else
                    mstrҽ����Ч = " And (D.ҽ����Ч=1 Or D.ҽ����Ч is Null)"
                End If
            ElseIf .chkBill(chkBills.��������).Value = 1 Then
                If .chkBill(chkBills.��ͨ����).Value = 0 And .chkBill(chkBills.�Զ�����).Value = 0 Then
                    mstrҽ����Ч = " And D.ҽ����Ч=0"
                ElseIf .chkBill(chkBills.��ͨ����).Value = 0 Then
                    mstrҽ����Ч = " And (A.��¼����=2 And D.ҽ����Ч=0 Or A.��¼����=3)"
                Else
                    mstrҽ����Ч = " And (D.ҽ����Ч=0 Or D.ҽ����Ч is Null)"
                End If
            Else
                mstrҽ����Ч = " And D.ҽ����Ч is Null"
            End If
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.Operator = zlStr.NeedName(.cbo����Ա.Text)
            SQLCondition.FeeItems = .mstrFeeItems
            SQLCondition.IncomeItems = .mstrIncomeItems
            
            '�����: 51625�޸���:���˺�,�޸�ʱ��:2012-12-10 18:21:16
            SQLCondition.lngHospNo = Val(.txtHospitalNO.Text)
            SQLCondition.strPatiName = Trim(.txtName.Text)
            mnuViewReFlash_Click
        End If
    End With
End Sub


Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mshDetail_EnterCell()
    mshDetail.ForeColorSel = mshDetail.CellForeColor
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long, strTime As String, blnDel As Boolean
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, strTime, blnDel, True)
    End If
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEditView.Enabled Then mnuEditView_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, strTime As String
    Dim lng���ʵ�ID As Long, blnDo As Boolean, blnDel As Boolean
        
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "��" & Val(lbl_s.Tag) & "������,��ǰ:" & lvw.SelectedItem.Text & ",סԺ��:" & _
                lvw.SelectedItem.SubItems(1) & ", �� " & Nvl(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(Nvl(mrsTotal!���, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2

    mnuEditAdjust.Enabled = Not blnDel
    '�Զ����ʵ���ҽ�����ɵļ��ʵ��������޸�
    mnuEditModi.Enabled = Not blnDel And Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) <> 3 _
                            And mshList.TextMatrix(mshList.Row, GetColNum("��������")) = "��ͨ����"
    mnuEditDel.Enabled = Not blnDel
    tbr.Buttons("Modi").Enabled = mnuEditModi.Enabled
    tbr.Buttons("Del").Enabled = mnuEditDel.Enabled
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNO, strTime, blnDel)
    
    '���ÿɷ��Ƶ���
    blnDo = True
    If blnDel Then blnDo = False '���ʵ���
    If blnDo Then '�Զ����ʵ�
        If Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) = 3 Then blnDo = False
    End If
    If blnDo Then '�Զ�����ʵ�
        If Val(mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))) <> 0 Then blnDo = False
    End If
    If blnDo Then If BillisBatch(strNO) Then blnDo = False '���ʱ�
    If blnDo Then If BillisSimple(strNO) Then blnDo = False '�򵥼���
    
    mnuEditCopy.Enabled = blnDo
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEditDel.Enabled And mnuEditDel.Visible Then Call mnuEditDel_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then
                If Me.ActiveControl Is lvw Then
                    Call FindNextPati
                Else
                    Call SeekBill(False)
                End If
            End If
        Case vbKeyReturn
            If Not Me.ActiveControl Is cboDept Then
                If mnuEditView.Enabled Then mnuEditView_Click
            End If
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEditDel_Click()
    Dim strNO As String, strTime As String
    Dim blnBat As Boolean, intTmp As Integer
    Dim str����IDs As String, strInfo As String, i As Long, intInsure As Integer
    Dim strInsure As String, arrInsure As Variant, bytType As Byte, blnFlagPrint As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ������ʣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    bytType = Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����")))
    
    'Ȩ���ж�
    If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), CDate(strTime), "����", strNO, , bytType) Then Exit Sub
        
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, bytType, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
        
    '��Ŀ����Ȩ��
    If Not CheckDelPriv(strNO, mstrPrivsOpt, strTime, bytType) Then Exit Sub
        
    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, mstrPrivsOpt, strTime, bytType)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '�Ƿ���ִ��
    blnBat = Val(mshList.TextMatrix(mshList.Row, GetColNum("�ಡ�˵�"))) <> 0
    i = BillCanDelete(strNO, bytType, blnBat, strTime, mstrPrivsOpt, blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '�õ��ݲ�����
                MsgBox "ָ�������е����ݲ�����,������û������շ���Ŀ������Ȩ�ޣ�", vbInformation, gstrSysName
            Case 2 '�Ѿ�ȫ����ȫִ��
                MsgBox "ָ�������е������Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
            Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                MsgBox "ָ�������е�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�", vbInformation, gstrSysName
        End Select
        Exit Sub
    End If
    If blnFlagPrint Then
        If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '��Ժ���˲���Ȩ���ж�
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "����", strTime, str����IDs, bytType) Then Exit Sub
    
    '�Ƿ��Ѿ�����
    intTmp = HaveBilling(2, strNO, False, strTime, bytType)
    If intTmp <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True, bytType)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , arrInsure(i)) Then
                        'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ����
                        If intTmp = 1 Then
                            MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            MsgBox "��ҽ�����ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbExclamation, gstrSysName
                        End If
                    End If
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�ü��ʵ��ݰ����Ѿ����ʵ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        Case 2
                            If intTmp = 1 Then
                                MsgBox "�ü��ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                                Exit Sub
                            Else
                                MsgBox "�ü��ʵ��ݰ����Ѿ����ʵ�����,ֻ�ܶ�δ���ʲ��ֽ������ʣ�", vbExclamation, gstrSysName
                            End If
                    End Select
                End If
            Next
        End If
    End If
    
    intInsure = BillExistInsure(strNO, , , bytType) '�ж��Ƿ���ҽ�����˼ǵ���,���ʱ�������ֻҪ��ҽ������
    'ҽ�����ʲ�����Ը�����¼��������
    If intInsure <> 0 Then
        If CheckNONegative(strNO, bytType) Then
            MsgBox "�õ��ݴ��ڸ������ʼ�¼,���������ҽ�����ʲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
        
    '�Ƿ������������¼
    If CheckRecalcRecord(strNO) Then
        MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
            "����ǰ�밴�ѱ�������ã������˽����������ʵ��ݵĴ����Żݽ�", vbInformation, Me.Caption
    End If
    
    On Error Resume Next
    Err.Clear
        
    If blnBat Then '��������
        frmBillings.mbytUseType = 1
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 3
        frmBillings.mstrInNO = strNO
        frmBillings.mstrTime = strTime
        frmBillings.mstr����IDs = str����IDs
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO, bytType) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 3
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long, varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 1
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 3
            frmCharge.mstrInNO = strNO
            frmCharge.mbytNOType = bytType
            frmCharge.mstrTime = strTime
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 1, 3, strNO, 0, 0, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ShowBills(mstrFilter)
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            Call ShowBills(mstrFilter)
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEditView_Click()
    Dim strNO As String, strTime As String, blnDel As Boolean
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Բ��ģ�", vbInformation, gstrSysName
        Exit Sub
    End If

    If Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) = 3 Then
        MsgBox "�õ���Ϊ�Զ����ʵ�,�������ܼ�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
        
    On Error Resume Next
    Err.Clear
    
    If BillisBatch(strNO) Then '��������
        frmBillings.mbytUseType = 1
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 1
        frmBillings.mstrInNO = strNO
        frmBillings.mblnNOMoved = mblnNOMoved   '�Ƿ�Ӻ󱸱���ȡ��
        frmBillings.mstrTime = strTime
        frmBillings.mblnDelete = blnDel
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 1
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnNOMoved = mblnNOMoved   '�Ƿ�Ӻ󱸱���ȡ��
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mblnDelete = blnDel
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 1
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.mblnNOMoved = mblnNOMoved   '�Ƿ�Ӻ󱸱���ȡ��
            frmCharge.mstrTime = strTime
            frmCharge.mblnDelete = blnDel
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 1, 1, strNO, 0, 0, 0, mstrPrivs, blnDel)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    mstrPage = ""
    Call tbs_Click
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).minHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If lbl_s.Width + X < 2580 Or mshList.Width - X < 3500 Then Exit Sub
        pic.Left = pic.Left + X
        lbl_s.Width = lbl_s.Width + X
        tbs.Width = tbs.Width + X
        lvw.Width = lvw.Width + X
        
        lblMoney.Left = lblMoney.Left + X
        lblMoney.Width = lblMoney.Width + X
        
        mshList.Left = mshList.Left + X
        mshList.Width = mshList.Width - X
        
        mshDetail.Left = mshDetail.Left + X
        mshDetail.Width = mshDetail.Width - X
        
        picHsc.Left = picHsc.Left + X
        picHsc.Width = picHsc.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lvw.SetFocus
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Text = "������ɫ" Then Call zlDatabase.ShowPatiColorTip(Me)
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "View"
            mnuEditView_Click
        Case "Billing"
            mnuEditBilling_Click
        Case "Modi"
            mnuEditModi_Click
        Case "Del"
            mnuEditDel_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "BillingBilling"
            mnuEditBilling_Click
        Case "BillingTable"
            mnuEditTable_Click
        Case "BillingSimple"
            mnuEditSimple_Click
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '��ͷ
    objOut.Title.Text = "סԺ���ʵ����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmDeptFilter
        objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " �� " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEditModi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    mnuEditCopy.Enabled = blnUsed
    mnuEditAdjust.Enabled = blnUsed
    
    mnuEditDel.Enabled = blnUsed
    mnuEditView.Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
End Sub

Private Sub SetCustBill()
'�������Զ�����ʵ���ص�����
    Dim rsTmp As New ADODB.Recordset
    Dim lngCount As Long, lngSum As Long
    On Error Resume Next
    
    If gobjCustBill Is Nothing Then
        Set gobjCustBill = CreateObject("zl9CustAcc.clsCustAcc")
    End If
    If InStr(mstrPrivsOpt, ";ר�����;") = 0 Then
        mnuEditCust.Visible = False
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    '��������ɹ����ٶ�����Ӧ�Ĳ˵�
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,���� From �շѼ��ʵ� Where substr(���÷�Χ,3,1)='1' Order by ���"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        lngSum = rsTmp.RecordCount
    End If
    
    If lngSum > 0 Then
        For lngCount = 1 To lngSum
            '���ӵ����˵���
            If lngCount > 1 Then
                Load mnuEditCustBill(lngCount)
            End If
            mnuEditCustBill(lngCount).Caption = rsTmp("����") & "(&" & lngCount & ")"
            mnuEditCustBill(lngCount).Tag = rsTmp("ID")
            
            rsTmp.MoveNext
        Next
    Else
        mnuEditCustBill(1).Enabled = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long, lngTmp As Long
    
    mstrPrivs = gstrPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call SetCustBill
    Call RestoreWinState(Me, App.ProductName)
    Set stbThis.Panels("state").Picture = Me.Picture
        
    lngTmp = Val(zlDatabase.GetPara("��ʾ���˷�ʽ", glngSys, mlngModul, 0))
    For i = 0 To mnuViewByDept.UBound
        mnuViewByDept(i).Checked = (i = lngTmp)
    Next
    
    i = IIf(zlDatabase.GetPara("ҳ��", glngSys, mlngModul, "1") = "1", 1, 2)
    tbs.Tabs(i).Selected = True
        
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '���ݱ����б�ʽ���ò˵�
    Call SetView(lvw.View)

    mlngCurRow = 1: mlngTopRow = 1
    mblnFirst = True
    
    'Ȩ������
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 Then
        mnuEditBilling.Visible = False
        mnuEditTable.Visible = False
        mnuEditSimple.Visible = False
        mnuEditCust.Visible = False
        mnuEditCopy.Visible = False
        mnuEditBilling_.Visible = False
        
        tbr.Buttons("Billing").Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 Then
        mnuEditDel.Visible = False
        If InStr(mstrPrivsOpt, ";ҩƷ��������;") = 0 _
            And InStr(mstrPrivsOpt, ";������������;") = 0 _
            And InStr(mstrPrivsOpt, ";������������;") = 0 _
            And InStr(mstrPrivsOpt, ";�������;") = 0 Then
            mnuEditDel_.Visible = False
        End If
        tbr.Buttons("Del").Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";ҩƷ��������;") = 0 _
        Or InStr(mstrPrivsOpt, ";������������;") = 0 _
        Or InStr(mstrPrivsOpt, ";������������;") = 0 _
        Or InStr(1, mstrPrivsOpt, ";��������;") = 0 Then
        mnuEditDelApply.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";�������;") = 0 Then
        mnuEditDelAudit.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 Then
        mnuEditModi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";��¼����;") = 0 Then
        mnuEditAdjust.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 _
        And InStr(mstrPrivsOpt, ";��¼����;") = 0 _
        And InStr(mstrPrivsOpt, ";סԺ����;") = 0 Then
        mnuEditAdjust_.Visible = False
    End If
    
    '55380
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 And InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 _
        And (InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0) Then
        tbr.Buttons("Del_").Visible = False
    End If
    
    Call InitPara
        
    If Not InitUnits Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        MsgBox "û�з�������������,���㲻�������в���Ȩ��,����ʹ�ÿ��ҷ�ɢ���ʣ�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
        
    mbln���� = True
    mbln���� = False
    
    mdtEnd = zlDatabase.Currentdate + 7
    mdtBegin = DateAdd("m", -1, mdtEnd)
    mstrPage = tbs.SelectedItem.Key
    
    
    Call LoadPatients '�����Ѱ���Call SetDetail Call SetHeader  Call SetMenu
    
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, sngVsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    
    If mblnMax Then
        sngVsc = 0.3: mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    lbl_s.Left = Me.ScaleLeft
    lbl_s.Top = Me.ScaleTop + cbrH + 45
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = lbl_s.Top + lbl_s.Height + 45
    tbs.Width = lbl_s.Width
    
    lvw.Left = Me.ScaleLeft
    lvw.Top = tbs.Top + tbs.Height - 75
    lvw.Width = lbl_s.Width
    lvw.Height = Me.ScaleHeight - staH - cbrH - lbl_s.Height - tbs.Height - 15
    
    pic.Left = lvw.Left + lvw.Width
    pic.Top = Me.ScaleTop + cbrH
    pic.Height = Me.ScaleHeight - cbrH - staH
    
    lblMoney.Left = pic.Left + pic.Width
    lblMoney.Top = Me.ScaleTop + cbrH + 45
    lblMoney.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
    
    mshList.Left = pic.Left + pic.Width
    mshList.Top = lblMoney.Top + lblMoney.Height + 15
    mshList.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
    mshList.Height = (Me.ScaleHeight - cbrH - staH - lblMoney.Height - picHsc.Height - 60) * (1 - sngVsc)
    
    picHsc.Left = mshList.Left
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Width = mshList.Width
    
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Left = mshList.Left
    mshDetail.Width = mshList.Width
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - lblMoney.Height - picHsc.Height - mshList.Height - 60
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, lngTmp As Long
    Dim blnHavePara As Boolean
    blnHavePara = InStr(1, mstrPrivs, ";��������;") > 0
    mstrFilter = ""
    mlngDeptID = 0
    mlngUnitID = 0
    
    Set mrsPati = Nothing
    Unload frmDeptFilter
    Unload frmDeptGo
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "ҳ��", tbs.SelectedItem.Index, glngSys, mlngModul, blnHavePara
    
        
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, blnHavePara
            Exit For
        End If
    Next
    
    '��ʾ���˷�ʽ
    lngTmp = 0
    For i = 0 To mnuViewByDept.UBound
        If mnuViewByDept(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "��ʾ���˷�ʽ", lngTmp, glngSys, mlngModul, blnHavePara
    
End Sub

Private Sub mnuViewGo_Click()
    Dim blnPati As Boolean
    blnPati = Me.ActiveControl Is lvw
    
    If blnPati Then
        '��λ����
        With frmDeptGo
            .fraBill.Visible = False
            .fraPati.Visible = True
            .Height = 2490
            .fraPati.Width = 3100
            .Width = .fraPati.Width + 600
            .cmdCancel.Left = .fraPati.Left + .fraPati.Width - .cmdCancel.Width - 100
            .cmdOk.Left = .cmdCancel.Left - .cmdOk.Width - 100
        End With
    Else
        '��λ����
        With frmDeptGo
            .fraBill.Visible = True
            .fraPati.Visible = False
            .Height = 1770
            .Width = .fraBill.Width + 600
            .cmdCancel.Left = .fraBill.Left + .fraBill.Width - .cmdCancel.Width - 100
            .cmdOk.Left = .cmdCancel.Left - .cmdOk.Width - 100
        End With
    End If
    frmDeptGo.Show 1, Me
    If gblnOK Then
        If blnPati Then
            Call FindPati
        Else
            Call SeekBill(frmDeptGo.optHead)
        End If
    End If
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, bln As Boolean, intRows As Integer
    Dim blnFill As Boolean, j As Long
    Dim strCurNO As String
    
    If frmDeptGo.txtNO.Text = "" Then Exit Sub
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmDeptGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ݺ�")) = .txtNO.Text
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            
            Call mshList_EnterCell
            mlngGo = i + 1
            
            stbThis.Panels(2).Text = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, GetColNum("���ݺ�")) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "��������,1,900|���ݺ�,1,850|��������,1,850|������,1,800|�ѱ�,1,900|Ӧ�ս��,7,850|ʵ�ս��,7,850|����Ա,1,800|�Ǽ�ʱ��,1,1850|˵��,1,850|����,1,0|��¼����,1,0|�ಡ�˵�,1,0|���ʵ�ID,1,0|��������ID,1,0"
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
        
        i = GetColNum("����"): mshList.ColWidth(i) = 0
        i = GetColNum("��¼����"): mshList.ColWidth(i) = 0
        i = GetColNum("�ಡ�˵�"): mshList.ColWidth(i) = 0
        i = GetColNum("���ʵ�ID"): mshList.ColWidth(i) = 0
        
        '�鿴ҽ����Ȩ��
        i = GetColNum("������")
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then
            mshList.ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            mshList.ColWidth(i) = 800
        End If
        
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
                
        Call mshList_EnterCell
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, j As Long, k As Long
    Dim strSql As String, lng����ID As Long, lng��ҳID As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
        
    If Not blnSort Then
        Call ZLCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        
        'ȱʡ��������(һ����)
        SQLCondition.Default = (strIF = "")
        If strIF = "" Then
            strIF = " And �Ǽ�ʱ��>Sysdate-1 And ��¼����=2 And ��¼״̬ IN(1,3) And ����Ա����||''=[5]"
            mstrҽ����Ч = ""   'ȱʡΪ��ͨ����+����+����
        End If
        
        If lvw.SelectedItem Is Nothing Then
            strIF = strIF & " And Rownum<1"
        Else
            lng����ID = Val(Split(lvw.SelectedItem.Key, "_")(1))
            lng��ҳID = Val(Split(lvw.SelectedItem.Key, "_")(2))
            strIF = strIF & " And ����ID=[6] And Nvl(��ҳID,0)=[7]"
        End If
        
               
        strIF = " Where �����־=2 And ����Ա���� is Not NULL " & strIF
        
        'ɸѡʱ��ʱ�������һ��ת��֮ǰ,��Ժ�ͳ�Ժ���˶����ܴ��ڵ��ݱ�ת��
        If frmDeptFilter.mblnDateMoved Then
            strIF = zlGetFullFieldsTable("סԺ���ü�¼", 2, strIF, False)
        Else
            strIF = zlGetFullFieldsTable("סԺ���ü�¼", 0, strIF, False)
        End If
        
        '���ݺ�,��������,������,�ѱ�,Ӧ�ս��,ʵ�ս��,����Ա,�Ǽ�ʱ��,˵��,����,��¼����,�ಡ�˵�,���ʵ�ID
        'Sign(ִ��״̬):������������ʱ����ͬʱ�б�Ҫ,���Զ�����
        strSql = _
            "Select Decode(A.��¼����,3,'�Զ�����',Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����')) as ��������,A.NO as ���ݺ�," & _
            " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,B.����) as ��������," & _
            " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.������) as ������," & _
            " A.�ѱ�,To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
            " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
            " A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
            " Decode(A.��¼����,3,Decode(Max(A.��¼״̬),2,'�Զ�����','�Զ�����'),Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼')) as ˵��," & _
            " Max(A.��¼״̬) as ����,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,A.��������ID" & _
            " From (" & strIF & ") A,���ű� B,����ҽ����¼ D" & _
            " Where A.��������ID=B.ID And A.ҽ�����=D.id(+) " & mstrҽ����Ч & _
            " Group by Sign(Decode(Nvl(A.ִ��״̬,0),0,1,Nvl(A.ִ��״̬,0))),A.NO," & _
            " Decode(A.��¼����,3,'�Զ�����',Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����'))," & _
            " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,B.����),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.������)," & _
            " A.�ѱ�,A.����Ա����,A.�Ǽ�ʱ��,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,A.��������ID" & _
            " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        With SQLCondition
            If .Default Then .Operator = UserInfo.����
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, lng����ID, lng��ҳID, .FeeItems, .IncomeItems)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = stbThis.Panels(2).Text & ",��ǰ����û�й��˳��ò�����ص��κε���"
        Call SetMenu(False)
    Else
        '��ʵ�պϼƽ��
        If Not blnSort Then
            strSql = "Select Sum(ʵ�ս��) as ���,Count(Distinct NO) as ���� From (" & Replace(strIF, "��¼״̬ IN(1,3)", "��¼״̬ IN(1,2,3)") & ")"
            With SQLCondition
                Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, lng����ID, lng��ҳID, .FeeItems, .IncomeItems)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    '������ɫ
    If mbln���� And Not mbln���� Then
        mshList.ForeColor = &HC0
    Else
        mshList.ForeColor = ForeColor
        k = GetColNum("����")
        For i = 1 To mshList.Rows - 1
            If Val(mshList.TextMatrix(i, k)) = 2 Then
                '���ʼ�¼�ú�ɫ
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '�������ʵ�����ɫ
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
        Next
    End If
    
    Call SetHeader
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�")) = "" Then Call SetDetail
        
    mshList.Redraw = True
    
    If Not lvw.SelectedItem Is Nothing And Not blnSort Then
        lng����ID = Val(lvw.SelectedItem.Tag)
        mrsPati.Filter = "����id=" & lng����ID
        Set rsTmp = GetMoneyInfo(CLng(lvw.SelectedItem.Tag), , Not IsNull(mrsPati!����), 2)
        mrsPati.Filter = ""
        
        If Not rsTmp Is Nothing Then
            lblMoney.Caption = " " & lvw.SelectedItem.Text & "  Ԥ���" & Format(rsTmp!Ԥ�����, "0.00") & _
                ",δ����ã�" & Format(rsTmp!�������, gstrDec) & ",ʣ��" & Format(rsTmp!Ԥ����� - rsTmp!�������, "0.00")
        Else
            lblMoney.Caption = " " & lvw.SelectedItem.Text & "  Ԥ���0.00,δ����ã�" & gstrDec & ",ʣ��0.00"
        End If
    End If
    
    If Not blnSort Then Call ZLCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbs_Click()
    If Not Visible Then Exit Sub
    If tbs.SelectedItem.Key = mstrPage Then Exit Sub
    
    '��ȡ����
    mstrPage = tbs.SelectedItem.Key
    Call LoadPatients
    'lvw.SetFocus
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSql As String, strServiceRange As String
    Dim blnByDept As Boolean
    
    On Error GoTo errH
    blnByDept = mnuViewByDept(1).Checked
    cbr.Bands(2).Caption = IIf(blnByDept, "���˿���", "���˲���")
    cboDept.Clear
    If InStr(mstrPrivs, ";���в���;") > 0 Then cboDept.AddItem IIf(blnByDept, "���п���", "���в���")
    
    '��Ȩ����ʾ����۲��Ҷ�Ӧ���ٴ�����,סԺ������סԺ��ͬ
    If InStr(mstrPrivsOpt, ";�������ۼ���;") And gbln�������� Then
        strServiceRange = "1,2,3"
    Else
        strServiceRange = "2,3"
    End If
    Set mrsDept = GetUnit(InStr(mstrPrivs, ";���в���;") = 0, strServiceRange, IIf(blnByDept, "�ٴ�", "����"), True)
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cboDept.AddItem mrsDept!���� & "-" & mrsDept!����
            cboDept.ItemData(cboDept.NewIndex) = mrsDept!ID
            If UserInfo.����ID = mrsDept!ID Then cboDept.ListIndex = cboDept.NewIndex
                
            mrsDept.MoveNext
        Next
        If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    ElseIf InStr(mstrPrivs, ";���в���;") > 0 Then
        MsgBox "û�з���" & IIf(blnByDept, "�ٴ�", "����") & "������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetView(bytStyle As Byte)
'���ܣ�������λ�б���ʾ��ʽ
'������bytstyle=0-��ͼ��,1-Сͼ��,2-�б�,3-��ϸ����
    mnuViewStyle(0).Checked = False
    mnuViewStyle(1).Checked = False
    mnuViewStyle(2).Checked = False
    mnuViewStyle(3).Checked = False
    mnuViewStyle(bytStyle).Checked = True
    lvw.View = bytStyle
End Sub

Private Function LoadPatients() As Boolean
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim objItem As ListItem, strSql As String
    Dim i As Long, j As Long, strCount As String
    Dim blnByDept As Boolean, strWhere As String
    Dim strWhere�䶯 As String
    Dim blnFind As Boolean
    
    On Error GoTo errH
    
    Call ZLCommFun.ShowFlash("���ڶ�ȡסԺ�����嵥,���Ժ� ...", Me)
    DoEvents
    
    Me.Refresh
    blnByDept = mnuViewByDept(1).Checked
    If blnByDept Then
        mintBedLen = GetMaxBedLen(mlngDeptID, True)
    Else
        mintBedLen = GetMaxBedLen(mlngUnitID, False)
    End If
    
    '���۲�������
    If InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� _
        And InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strWhere = " And Nvl(B.��������,0) IN (0,1,2)"
    ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
        strWhere = " And Nvl(B.��������,0) IN (0,1)"
    ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
        strWhere = " And Nvl(B.��������,0) IN (0,2)"
    Else
        strWhere = " And Nvl(B.��������,0)=0"
    End If
    
    '�����: 51625�޸���:���˺�,�޸�ʱ��:2012-12-10 18:21:16
    If SQLCondition.lngHospNo <> 0 Then
         strWhere = " And  B.סԺ�� =[6]"
    End If
    If SQLCondition.strPatiName <> "" Then
         strWhere = " And A.���� Like [7]"
    End If
    
    strWhere�䶯 = ""
    If blnByDept Then
        Select Case tbs.SelectedItem.Index
        Case 1, 2
            strWhere = strWhere & IIf(mlngDeptID > 0, " And B.��Ժ����ID" & IIf(tbs.SelectedItem.Index = 2, "+0", "") & "=[2]", "")
        Case Else
            strWhere = strWhere & IIf(mlngDeptID > 0, " And C.����ID =[2]", "")
            strWhere�䶯 = IIf(mlngDeptID > 0, " And ����ID =[2]", "")
        End Select
    Else
        Select Case tbs.SelectedItem.Index
        Case 1, 2
            strWhere = strWhere & IIf(mlngUnitID > 0, " And B.��ǰ����ID" & IIf(tbs.SelectedItem.Index = 2, "+0", "") & "=[1]", "")
        Case Else
            strWhere = strWhere & IIf(mlngDeptID > 0, " And C.����ID =[1]", "")
            strWhere�䶯 = IIf(mlngDeptID > 0, " And ����ID =[1]", "")
        End Select
    End If
    
    Select Case tbs.SelectedItem.Index
    Case 1
        '�ò�����ҳ�ĳ�Ժ����ID��������,����Ϊ�������������۲���(���ۿ��һ���û�д�λ),���Բ��ܴӴ�λ״����¼��ȥ�Ҳ�
        '��ǰ��Ժ�Ĳ���
        strSql = _
            "Select   A.����ID,B.��ҳID,A.סԺ��, " & _
            "   Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(b.����, a.����) as ����,B.ҽ�Ƹ��ʽ," & _
            "   B.��Ժ����,B.��Ժ����,LPAD(B.��Ժ����," & mintBedLen & ",' ') as ����," & _
            "   C.���� as ��ǰ����,B.����,B.��������,B.״̬,B.��Ժ����ID ��ǰ����ID,B.��������" & _
            " From ������Ϣ A,������ҳ B,���ű� C,��Ժ���� ZY " & _
            " Where A.����ID=B.����ID And B.����ID=ZY.����ID And B.��Ժ����ID=C.ID" & strWhere & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
            " And B.��Ժ���� is NULL And Nvl(B.��ҳID,0)<>0  " & _
             IIf(mlngDeptID = 0, " Order by  A.סԺ�� Desc", " Order by   LPAD(����,10,' ')")
    Case 2
        '���ڼ��Ժ�Ĳ���
        strSql = _
            "Select A.����ID,B.��ҳID,A.סԺ��," & _
            "   Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(b.����, a.����) as ����,B.ҽ�Ƹ��ʽ," & _
            "   B.��Ժ����,B.��Ժ����,LPAD(B.��Ժ����," & mintBedLen & ",' ') as ����," & _
            "   C.���� as ��ǰ����,B.����,B.��������,B.״̬,B.��Ժ����ID ��ǰ����ID,B.��������" & _
            " From ������Ϣ A,������ҳ B,���ű� C" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=C.ID" & strWhere & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
            " And B.��Ժ����<=[4] And B.��Ժ���� Between [3] And [4]" & _
            IIf(mlngDeptID = 0, " Order by A.סԺ�� Desc", " Order by LPAD(����,10,' ')")
    Case 3
        '���ܴ���ͬһ����һ���Χ�ڵ����������ϵ�ת��,�������һ��Ϊ׼.
        'And C.��ֹʱ�� =(Select Max(��ֹʱ��)  From ���˱䶯��¼  Where ����ID=C.����ID And ��ҳID=C.��ҳID And  ��ֹԭ��=3  And Nvl(���Ӵ�λ,0)=0 " & strWhere�䶯 & ")"
        '����:29435
 
        strSql = "" & _
        " Select /*+ RULE */ Distinct A.����ID,B.��ҳID,B.סԺ��, " & _
        "       Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(b.����, a.����) as ����,B.ҽ�Ƹ��ʽ," & _
        "           B.��Ժ����,B.��Ժ����,LPAD(C.����," & mintBedLen & ",' ') as ����," & _
        "           D.���� as ��ǰ����,B.����,B.��������,B.״̬,C.����ID as ��ǰ����ID,B.��������" & _
        " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D " & _
        " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And C.����ID=D.ID " & _
        "       And Nvl(B.״̬,0)<>2 " & IIf(blnByDept, " And B.��Ժ����ID<>[1] ", " And B.��ǰ����ID<>[1] ") & _
        "       And B.����ID=C.����ID And B.��ҳID=C.��ҳID  " & _
        "       And C.��ֹԭ��=3 And (C.��ֹʱ��   Between Sysdate-[5] And Sysdate )  And Nvl(C.���Ӵ�λ,0)=0" & _
        "       And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL  " & strWhere & _
        "       And C.��ֹʱ�� =(Select Max(��ֹʱ��)  From ���˱䶯��¼  Where ����ID=C.����ID And ��ҳID=C.��ҳID And  ��ֹԭ��=3  And Nvl(���Ӵ�λ,0)=0 " & strWhere�䶯 & ")"
            
            '��������˹鵵��
            ',�շ���ĿĿ¼ E
            '    "       Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
            '    ",C.����ҽʦ as סԺҽʦ,B.����״̬," & _
            '    " E.���� as ����ȼ�,B.�ѱ�,B.��������,B.״̬,B.����,A.���￨��"
        strSql = "Select * FROM ( " & strSql & ") " & vbCrLf & IIf(mlngDeptID = 0, " Order by   סԺ�� Desc", " Order by   LPAD(����,10,' ')")
    Case Else
        Exit Function
    End Select
    
    mdtBegin = CDate(Format(mdtBegin, "yyyy-MM-dd 00:00:00"))
    mdtEnd = CDate(Format(mdtEnd, "yyyy-MM-dd 23:59:59"))
    Set mrsPati = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnitID, mlngDeptID, mdtBegin, _
        mdtEnd, mTy_Modul_Para.intת������, SQLCondition.lngHospNo, SQLCondition.strPatiName & "%")
  
    lvw.ListItems.Clear
    
    If Not mrsPati.EOF Then
        For i = 1 To mrsPati.RecordCount
            If IIf(IsNull(mrsPati!��������), 0, mrsPati!��������) = 0 Then
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!����ID & "_" & mrsPati!��ҳID, mrsPati!����, 1, 1)
            Else
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!����ID & "_" & mrsPati!��ҳID, mrsPati!����, 2, 2)
            End If
            objItem.SubItems(1) = IIf(IsNull(mrsPati!סԺ��), "", mrsPati!סԺ��)
            objItem.SubItems(2) = IIf(IsNull(mrsPati!����) And mrsPati!״̬ = 0, "��ͥ", Nvl(mrsPati!����, " "))
            objItem.SubItems(3) = IIf(IsNull(mrsPati!�Ա�), "", mrsPati!�Ա�)
            objItem.SubItems(4) = IIf(IsNull(mrsPati!����), "", mrsPati!����)
            objItem.SubItems(5) = Format(mrsPati!��Ժ����, "yyyy-MM-dd")
            objItem.SubItems(6) = Format(IIf(IsNull(mrsPati!��Ժ����), "", mrsPati!��Ժ����), "yyyy-MM-dd")
            objItem.SubItems(7) = IIf(IsNull(mrsPati!��ǰ����), "", mrsPati!��ǰ����)
            objItem.SubItems(8) = mrsPati!��ҳID
            objItem.SubItems(9) = Nvl(mrsPati!ҽ�Ƹ��ʽ)
            objItem.SubItems(10) = Val("" & mrsPati!��ǰ����id)
            objItem.SubItems(11) = "" & mrsPati!��������
            objItem.Tag = mrsPati!����ID
            objItem.ListSubItems(1).Tag = Nvl(mrsPati!״̬)
            
            objItem.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsPati!��������))
            For j = 1 To objItem.ListSubItems.Count
                objItem.ListSubItems(j).ForeColor = zlDatabase.GetPatiColor(Nvl(mrsPati!��������))
            Next
            
            If InStr(strCount & ",", "," & mrsPati!����ID & ",") = 0 Then strCount = strCount & "," & mrsPati!����ID
            mrsPati.MoveNext
        Next
        lbl_s.Tag = UBound(Split(Mid(strCount, 2), ",")) + 1
        If tbs.SelectedItem.Index = 1 Then
            lbl_s.Caption = " ��ǰ��Ժ�Ĳ���,����:" & Val(lbl_s.Tag)
        ElseIf tbs.SelectedItem.Index = 2 Then
            lbl_s.Caption = " ʱ��:" & Format(mdtBegin, "yyyy-MM-dd") & "��" & Format(mdtEnd, "yyyy-MM-dd") & ",����:" & Val(lbl_s.Tag)
        ElseIf tbs.SelectedItem.Index = 3 Then
            lbl_s.Caption = "��ʾ" & mTy_Modul_Para.intת������ & "����ת���Ĳ���"
        End If
        Me.Refresh
        Call ClearFeeList
    Else
        lbl_s.Tag = ""
        stbThis.Panels(2).Text = ""
        Call ShowBills 'û�в��˾�û�е���
    End If
    Call ZLCommFun.StopFlash
    Exit Function
errH:
    Call ZLCommFun.StopFlash
    If ErrCenter() = 1 Then
        Call ZLCommFun.ShowFlash("���ڶ�ȡסԺ�����嵥,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub ClearFeeList()
    '�����������Ϣ
    With mshDetail
            .Clear
            .Rows = 2
            .Cols = 2
    End With
    With mshList
        .Clear
        .Rows = 2
        .Cols = 2
    End With
    Call SetHeader
End Sub
Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If mblnNotClick Then Exit Sub
    stbThis.Panels(2).Text = "��" & Val(lbl_s.Tag) & "������,��ǰ:" & Item.Text & ",סԺ��:" & Item.SubItems(1)
    If mnuViewByDept(0).Checked Then mlngDeptID = Val(Item.SubItems(10))
    
    '��ȡ����
    Call ShowBills(mstrFilter)
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw.Sorted = True
    With lvw
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
    lvw.SortKey = ColumnHeader.Index - 1
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub FindPati()
    Dim strFilter As String
    Dim strBed As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    With frmDeptGo
        If .txtסԺ��.Text <> "" Then strFilter = strFilter & " Or סԺ��=" & .txtסԺ��.Text
        If .txt����.Text <> "" Then strFilter = strFilter & " Or ���� Like '%" & .txt����.Text & "%'"
        If .txt����.Text <> "" Then
            strBed = .txt����.Text
            If mintBedLen - ZLCommFun.ActualLen(strBed) > 0 Then
                strBed = String(mintBedLen - ZLCommFun.ActualLen(strBed), " ") & strBed
            End If
            strFilter = strFilter & " Or ����='" & strBed & "'"
        End If
    End With
    If strFilter = "" Then Exit Sub
    mrsPati.Filter = 0
    mrsPati.Filter = Mid(strFilter, 5)
    
    If mrsPati.EOF Then
        stbThis.Panels(2).Text = "û�з��ָò��ˣ�"
    Else
        lvw.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub FindNextPati()
    On Error Resume Next
    If mrsPati Is Nothing Then Exit Sub
    If mrsPati.RecordCount = 0 Then Exit Sub
    If mrsPati.Filter = 0 Then Exit Sub
    If mrsPati.EOF Then
        mrsPati.MoveFirst
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then mrsPati.MoveFirst
    End If
    lvw.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID).Selected = True
    lvw.SelectedItem.EnsureVisible
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "��������,1,850|������,1,800|���,1,650|����,1,1600" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|���,1,1000|��λ,4,500|����,7,850|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ͳ����,7,850|ִ�п���,1,850|����,1,850|˵��,1,1000|��¼״̬,1,0"
    
    With mshDetail
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        '���˺�:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(.Cols - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowDetail(Optional ByVal strNO As String, Optional ByVal strTime As String, _
    Optional ByVal blnDel As Boolean, Optional ByVal blnSort As Boolean)
    Dim strSql As String, i As Long, j As Long
    Dim blnBat As Boolean, bytFlag As Byte
    
    On Error GoTo errH
        
    blnBat = Val(mshList.TextMatrix(mshList.Row, GetColNum("�ಡ�˵�"))) <> 0
    bytFlag = mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))
    
    If Not blnSort Then
         If frmDeptFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved("סԺ���ü�¼", strNO, , 2, Me.Caption)
        Else
            mblnNOMoved = False   '����Ҫ����һ��
        End If
        
        strSql = _
        " Select E.���� as ��������,A.������,C.���� as ���,Nvl(F.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
                IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ��λ," & _
        "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ",'9999990.00000') as ����, " & _
        "       To_Char(Sum(A.��׼����)" & _
                IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ͳ����),'9999999" & gstrDec & "') as ͳ����, " & _
        "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
        "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��, A.��¼״̬" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼  A") & "," & _
        "       �շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,�շ���Ŀ���� F,ҩƷ��� X" & _
                  IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
        " Where A.�շ�ϸĿID=B.ID And A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID" & _
        "       And A.NO=[1] And A.��¼����=[2] And A.�����־=2" & _
        "       And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & _
        "       And A.�շ�ϸĿID=X.ҩƷID(+) And A.����ID+0=[3]" & _
        "       And A.�շ�ϸĿID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & IIf(strTime <> "", " And A.�Ǽ�ʱ��=[4]", "") & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),E.����,A.������,C.����,Nvl(F.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & "B.���,A.���㵥λ," & _
        "       D.����,Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID,X.סԺ��λ,Nvl(X.סԺ��װ,1)" & _
        " Order by Nvl(A.�۸񸸺�,A.���)"
        
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag, Val(lvw.SelectedItem.Tag), CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag, Val(lvw.SelectedItem.Tag))
        End If
    End If
        
    mshDetail.Redraw = False
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    mshDetail.ForeColor = IIf(blnDel, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    
    '������ɫ
    If blnDel Then
        '�˷�ֱ��Ϊ��ɫ
        mshDetail.ForeColor = &HC0
    Else
        'ԭʼ�����˹���Ϊ��ɫ
        mshDetail.ForeColor = ForeColor
        For i = 1 To mshDetail.Rows - 1
            If Val(mshDetail.TextMatrix(i, mshDetail.Cols - 1)) = 3 Then
                mshDetail.Row = i
                For j = 0 To mshDetail.Cols - 1
                    mshDetail.Col = j
                    mshDetail.CellForeColor = &HC00000
                Next
            End If
        Next
    End If

    Call SetDetail
        
    '���ʱ�Ҫ��ʾ������Ϣ
    If blnBat Then
        '��������,������
        mshDetail.ColWidth(0) = 850
        mshDetail.ColWidth(1) = 800
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then mshDetail.ColWidth(1) = 0
    End If
        
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

