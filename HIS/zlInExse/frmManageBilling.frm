VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBilling 
   AutoRedraw      =   -1  'True
   Caption         =   "סԺ���ʹ���"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9960
   Icon            =   "frmManageBilling.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   Picture         =   "frmManageBilling.frx":08CA
   ScaleHeight     =   6225
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9960
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinWidth1       =   6600
      MinHeight1      =   720
      Width1          =   4995
      NewRow1         =   0   'False
      Caption2        =   "���˲���"
      Child2          =   "cboDept"
      MinWidth2       =   1800
      MinHeight2      =   300
      Width2          =   1800
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8070
         TabIndex        =   2
         Text            =   "cboDept"
         Top             =   240
         Width           =   1800
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   6900
         _ExtentX        =   12171
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
            NumButtons      =   19
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
               Caption         =   "����"
               Key             =   "Price"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Price"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PriceBilling"
                     Object.Tag             =   "���ʵ�"
                     Text            =   "���ʵ�"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PriceTable"
                     Object.Tag             =   "���ʱ�"
                     Text            =   "���ʱ�"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PriceSimple"
                     Object.Tag             =   "�򵥼���"
                     Text            =   "�򵥼���"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Auditing"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageKey        =   "Auditing"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingBilling"
                     Object.Tag             =   "���ʵ�"
                     Text            =   "���ʵ�"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingTable"
                     Object.Tag             =   "���ʱ�"
                     Text            =   "���ʱ�"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingSimple"
                     Object.Tag             =   "�򵥼���"
                     Text            =   "�򵥼���"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Tag             =   "-"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingPati"
                     Object.Tag             =   "���������"
                     Text            =   "���������"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AuditingBatch"
                     Object.Tag             =   "�������"
                     Text            =   "�������"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Billing_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ҩ"
               Key             =   "Give"
               Description     =   "��ҩ"
               Object.ToolTipText     =   "��ҩ"
               Object.Tag             =   "��ҩ"
               ImageKey        =   "Give"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Give_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "�Ե�ǰѡ�е�������"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   5865
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBilling.frx":0A58
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8731
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
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmManageBilling.frx":0DCC
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
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9855
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3915
      Width           =   9855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   2805
      Left            =   15
      TabIndex        =   0
      Top             =   1080
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   4948
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
      MouseIcon       =   "frmManageBilling.frx":0F6A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1875
      Left            =   0
      TabIndex        =   1
      Top             =   3990
      Width           =   9945
      _ExtentX        =   17542
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
      MouseIcon       =   "frmManageBilling.frx":1284
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5205
      Top             =   270
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
            Picture         =   "frmManageBilling.frx":159E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":17B8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":19D2
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":1BEC
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2366
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2580
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":279A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":29B4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2BCE
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":2DE8
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":34E2
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":3BDC
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":42D6
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":44F0
            Key             =   "Give"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   4620
      Top             =   270
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
            Picture         =   "frmManageBilling.frx":470A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4924
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4B3E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":4D58
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":54D2
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":56EC
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5906
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5B20
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5D3A
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":5F54
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":664E
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":6D48
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":7442
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBilling.frx":765C
            Key             =   "Give"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   750
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   688
      TabWidthStyle   =   2
      TabFixedWidth   =   2293
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ʵ���(&1)"
            Key             =   "Auditing"
            Object.ToolTipText     =   "��ʾֱ�Ӽ��ʻ򻮼ۺ�����˵ļ��ʵ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���۵���(&2)"
            Key             =   "Price"
            Object.ToolTipText     =   "��ʾ���ۺ�δ��˵ļ��ʵ���"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
         Caption         =   "סԺ����(&B)"
         Begin VB.Menu mnuEditBillingBilling 
            Caption         =   "���ʵ�(&B)"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuEditBillingTable 
            Caption         =   "���ʱ�(&T)"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuEditBillingSimple 
            Caption         =   "�򵥼���(&S)"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuEditBillingCust 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEditPrice 
         Caption         =   "סԺ����(&R)"
         Begin VB.Menu mnuEditPriceBilling 
            Caption         =   "���ʵ�(&B)"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuEditPriceTable 
            Caption         =   "���ʱ�(&T)"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuEditPriceSimple 
            Caption         =   "�򵥼���(&S)"
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu mnuEditAuditing 
         Caption         =   "�������(&A)"
         Begin VB.Menu mnuEditAuditingBilling 
            Caption         =   "���ʵ�(&B)"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnuEditAuditingTable 
            Caption         =   "���ʱ�(&T)"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuEditAuditingSimple 
            Caption         =   "�򵥼���(&S)"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu mnuEditAuditing_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditAuditingPati 
            Caption         =   "���������(&P)"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuEditAuditingBatch 
            Caption         =   "�������(&A)"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuEditBilling_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGive 
         Caption         =   "���ݷ�ҩ(&G)"
      End
      Begin VB.Menu mnuEditGive_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditModi 
         Caption         =   "�޸ĵ���(&M)"
         Shortcut        =   ^M
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
      Begin VB.Menu mnuEditDelBat 
         Caption         =   "��������(&B)"
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
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
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
Attribute VB_Name = "frmManageBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mrsTotal As ADODB.Recordset
Private mrsDetail As ADODB.Recordset
Private mstrFilter As String

'��ʿ����վ����ʱ�Ĺ�������
'��3�������
'1.ֻ������ʾָ�����ݺŵĵ��ݣ�.���ݺ�,.ҽ��ID
'2.������ʾһ��ҽ����¼�ĵ��ݣ�.ҽ��ID,.���ͺ�
'3.������ʾĳ�η��͵����е��ݣ�.���ͺ�
Private Type TYPE_NurseStation
    Nurse As Boolean '������ǰ������ʾ�Ƿ�ǿ��ʹ�û�ʿ����վ���õ�����
    ����ID As Long '��ʿ����վ��ǰ���˵Ĳ���
    ����ID As Long '��ʿ����վ��ǰ���˵Ŀ���
    ���ͺ� As Long
    ҽ��ID As Long 'һ��ҽ����¼��ID��Nvl(���ID,ID)
    ���ݺ� As String
    ���� As Boolean 'ȱʡ��λ��ҳ���Ƿ񻮼�(���������������,�Ի�ʿ����վ����ʱ�ĵ�ǰ���ݵ����Ϊȱʡ)
    ReLoad As Boolean '�Ƿ���������(������ʾ�������)
    Mode As Boolean '�Ƿ��ģ̬���ڷ���
End Type
Private mvNurseFilter As TYPE_NurseStation

'���ʹ�����Ĺ�������
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Operator As String
    InPatientID As Double   '34512
    Patient As String
    FeeItems As String
    IncomeItems As String
End Type
Private SQLCondition As Type_SQLCondition

Private mbln���� As Boolean, mbln���� As Boolean
Private mstr����Ա As String, mstrҽ����Ч As String

Private mstrPage As String
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngDeptID As Long, mlngUnitID As Long
Private mblnMax As Boolean
Private mrsDept As ADODB.Recordset

Private mstrPrivs As String
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mlngModul As Long
Private mblnNOMoved As Boolean '��¼��ǰѡ��ĵ����Ƿ����ں����ݱ���

Public Function ShowMeByNurse(frmMain As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, ByVal strNO As String, ByVal bln���� As Boolean) As Object
'���ܣ��ɻ�ʿ����վ���ò��Զ�������ʾ��ָ��ҽ�������ķ��õ��ݣ�Ŀ���ǳ�����Щ����
'��������Ӧ������TYPE_NurseStation�е��ֶζ���
'���أ����ӷ�ģ̬���ڷ���ʱ,���ؼ��ʹ�����,���ڸ��ٹر��¼�(��ģ̬��ʾʱ��ˢ������)
    With mvNurseFilter
        .Nurse = True
        .����ID = lng����ID
        .����ID = lng����ID
        .���ͺ� = lng���ͺ�
        .ҽ��ID = lngҽ��ID
        .���ݺ� = strNO
        .���� = bln����
        .ReLoad = False
        .Mode = False
    End With
    
    On Error Resume Next
    If mstrPrivs <> "" Then '�Ѵ�
        mvNurseFilter.ReLoad = True
        Call Form_Load
    End If
    Me.Show , frmMain '�Է�ģ̬��ʾ
    Err.Clear
    
    If Not mvNurseFilter.Mode Then
        Set ShowMeByNurse = Me
    End If
End Function

Private Sub cboDept_Click()
    Dim strTmp As String
    
    If Not mvNurseFilter.Nurse Then
        If tbs.SelectedItem.Key = "Auditing" Then
            If InStr(mstrPrivs, ";�鿴���ʵ�;") <= 0 Then Exit Sub
        Else
            If InStr(mstrPrivs, ";�鿴���۵�;") <= 0 Then Exit Sub
        End If
    End If
    
    If cboDept.ItemData(cboDept.ListIndex) = mlngUnitID Then Exit Sub
    mlngUnitID = cboDept.ItemData(cboDept.ListIndex)
    
    If mlngUnitID = 0 Then
        mlngDeptID = 0
    Else
        strTmp = Get����IDs(mlngUnitID)
        If InStr(1, strTmp, ",") > 0 Then
            mlngDeptID = Split(strTmp, ",")(0)
        Else
            mlngDeptID = Val(strTmp)
        End If
    End If
        
    If Visible Then
        If mvNurseFilter.Nurse Then
            If Not mvNurseFilter.ReLoad Then Call ShowBillsByNurse
        Else
            Call ShowBills(mstrFilter)
        End If
    End If
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
    If InStr(";" & mstrPrivs, ";���в���") > 0 Then strRootCaption = "���в���"
    
    If zlSelectDept(Me, mlngModul, cboDept, mrsDept, cboDept.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub

End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
        
    If cboDept.ListIndex >= 0 Then Exit Sub
    zlControl.CboLocate cboDept, mlngUnitID, True
    If cboDept.ListIndex < 0 And cboDept.ListCount <> 0 Then cboDept.ListIndex = 0

End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
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
    
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("�ಡ�˵�"))) = 1 Then '��������
        frmBillings.mbytUseType = 0
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 2
        frmBillings.mlngModule = mlngModul
        frmBillings.mstrInNO = strNO
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 0
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 2
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 0
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 2
            frmCharge.mstrInNO = strNO
            frmCharge.mlngModule = mlngModul
            
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 0, 2, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
        End If
    End If
End Sub

Private Sub mnuEditAuditingBatch_Click()
    Dim rsWarn As ADODB.Recordset
    Dim blnTrans As Boolean, Curdate As Date
    Dim strSql As String, str���ʱ�� As String, strNO As String, strInfo As String
    Dim lngCOL��� As Long, lngCOLNO As Long, lngCOLInsure As Long
    Dim i As Long, j As Long, intInsure As Integer
    
    lngCOL��� = GetColNum("���")
    lngCOLNO = GetColNum("���ݺ�")
    lngCOLInsure = GetColNum("����")
        
    For i = 1 To mshList.Rows - 1
        If mshList.TextMatrix(i, lngCOL���) = "��" And mshList.TextMatrix(i, lngCOLNO) <> "" Then
            j = j + 1
        End If
    Next
    If j = 0 Then
        MsgBox "û��ѡ��Ҫ��˵Ļ��۵��ݣ�", vbExclamation, gstrSysName
        Exit Sub
    Else
        If MsgBox("ȷʵҪ��ѡ���" & j & "�Ż��۵��ݽ��������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    Set rsWarn = GetUnitWarn

    'ÿ�ŵ���һ������������,���������ʱ��������
    For i = 1 To mshList.Rows - 1
        strNO = mshList.TextMatrix(i, lngCOLNO)
        If mshList.TextMatrix(i, lngCOL���) = "��" And strNO <> "" Then
            '���ñ���
            If AuditingWarn(mstrPrivsOpt, rsWarn, strNO, "") Then
                If str���ʱ�� = "" Then
                    Curdate = zlDatabase.Currentdate
                    str���ʱ�� = "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                intInsure = Val(mshList.TextMatrix(i, lngCOLInsure))
                strSql = "zl_סԺ���ʼ�¼_Verify('" & strNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "',NULL,NULL," & str���ʱ�� & ")"
                
                gcnOracle.BeginTrans
                    blnTrans = True
                    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                    
                    If intInsure <> 0 Then 'ҽ��ʵʱ�ϴ�,���������ϸ
                        If gclsInsure.GetCapability(support�����ϴ�, , intInsure) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                            strInfo = ""
                            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strInfo, , intInsure) Then
                                gcnOracle.RollbackTrans
                                If strInfo <> "" Then MsgBox strInfo, vbInformation, gstrSysName
                                Call mnuViewReFlash_Click
                                Exit Sub        'ֻҪ��һ��ʧ�ܾ��˳�
                            End If
                        End If
                    End If
                gcnOracle.CommitTrans
                blnTrans = False
                
                If intInsure <> 0 Then 'ҽ���Ӻ��ϴ�,���������ϸ
                    If gclsInsure.GetCapability(support�����ϴ�, , intInsure) And gclsInsure.GetCapability(support������ɺ��ϴ�, , intInsure) Then
                        strInfo = ""
                        If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strInfo, , intInsure) Then
                            If strInfo <> "" Then
                                MsgBox strInfo, vbInformation, gstrSysName
                            Else
                                MsgBox "����""" & strNO & """��������ҽ������ʧ��,�õ�������ˣ�", vbInformation, gstrSysName
                            End If
                            Call mnuViewReFlash_Click
                            Exit Sub 'ֻҪ��һ��ʧ�ܾ��˳�
                        End If
                    End If
                End If
                                
                If gbln��˴�ӡ Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & strNO, "�Ǽ�ʱ��=" & Format(Curdate, "yyyy-MM-dd HH:mm:ss"), "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "�ش�=0", 2)
                End If
            End If
        End If
    Next
    
    Call mnuViewReFlash_Click
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call mnuViewReFlash_Click       'ִ�й����г�����Ҫˢ��
End Sub

Private Sub mnuEditBillingCust_Click(Index As Integer)
    '�Զ������
    Dim varTemp As Variant
            
    '�������������ǣ�
    '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs��blnViewCancel
    varTemp = Array(mnuEditBillingCust(Index).Tag, 0, 0, "", mlngUnitID, mlngDeptID, 0, mstrPrivs)
    gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
    
    gblnOK = varTemp '����ֵ
    
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

Private Function GetDelSerial(ByVal strNO As String, strTime As String) As String
'���ܣ���ָ�����ʵ���δ��ȫִ�м���ʣ���������к�,������������
'������strTime=�Ǽ�ʱ��,���ڲ�����˵ļ��ʵ�
'���أ���=��ʾû�п������ʵ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    
    strSql = _
        " Select ���,Sum(Nvl(����,1)*����) as ����" & _
        " From סԺ���ü�¼" & _
        " Where ��¼����=2 And NO=[1] And �Ǽ�ʱ��=[2] And Nvl(ִ��״̬,0)<>1 And �۸񸸺� is NULL" & _
        " Group by ��� Having Nvl(Sum(Nvl(����,1)*����),0)<>0"
    On Error GoTo errH
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)
    End If
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!���
        rsTmp.MoveNext
    Loop
    GetDelSerial = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditDelApply_Click()
    Dim strMsg As String
    If mlngUnitID = 0 Then
        If cboDept.Visible Then
            strMsg = "����ѡ���˲���!"
            cboDept.SetFocus
        Else
            strMsg = "����ѡ���˲���!" & vbCrLf & "(��ʾ����ѡ���б�:�鿴-������-סԺ����)"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    With frmReCharge
        .mlngDeptID = mlngUnitID
        .mbytUseType = 0
        .mbytFun = 0
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

Private Sub mnuEditDelBat_Click()
    Dim arrSQL() As Variant, blnTrans As Boolean
    Dim i As Long, j As Long, intInsure As Integer
    Dim blnBat As Boolean, blnBilling As Boolean, blnFlagPrint As Boolean
    Dim strNO As String, blnDo As Boolean, bytType As Byte
    Dim strInfo As String, strTime As String, str��� As String, strRebateNOS As String, strUnitIDs As String, strUnDelNOs As String
    Dim lngCol���ݺ� As Long, lngCol�Ǽ�ʱ�� As Long, lngCol��¼���� As Long, lngCol��������ID As Long
    
    If MsgBox("�������ʲ������ɻָ���ȷʵҪ����ǰ�б��еĵ���ȫ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If MsgBox("ȷʵҪ����ǰ�б��еĵ���ȫ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    lngCol���ݺ� = GetColNum("���ݺ�")
    lngCol�Ǽ�ʱ�� = GetColNum("�Ǽ�ʱ��")
    lngCol��¼���� = GetColNum("��¼����")
    lngCol��������ID = GetColNum("��������ID")
    
    arrSQL = Array()
    j = 0: blnBilling = True
    For i = 1 To mshList.Rows - 1
        blnDo = True
        strNO = mshList.TextMatrix(i, lngCol���ݺ�)
        strTime = mshList.TextMatrix(i, lngCol�Ǽ�ʱ��)
        bytType = Val(mshList.TextMatrix(i, lngCol��¼����))
        
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
            '�Ƿ���ת������ݱ���
            '���ʻ��۵������ں󱸱���,��¼����ֻȡ2
            '��ʱ���ܸ���showdetailʱ��mblnNOMoved���ж�,��Ϊû�е��bill��,mblnNOMoved���ǵ�ǰ��ѡ���ݵ�����
            '���Ա���Ҫ���ж�
            If zlDatabase.NOMoved("סԺ���ü�¼", strNO, , bytType, Me.Caption) Then
                If Not ReturnMovedExes(strNO, bytType, Me.Caption) Then blnDo = False
            End If
        End If
        
        '�Ƿ����ʼ�¼
        blnDo = mshList.TextMatrix(i, GetColNum("����")) <> 2
            
        'Ȩ���ж�
        If blnDo Then
            If tbs.SelectedItem.Key = "Price" Then
                If Not BillOperCheck(5, mshList.TextMatrix(i, GetColNum("������")), CDate(strTime), "����", strNO) Then blnDo = False
            Else
                If Not BillOperCheck(5, mshList.TextMatrix(i, GetColNum("������")), CDate(strTime), "����", strNO, , bytType) Then blnDo = False
            End If
        End If
        
        '��Ŀ����Ȩ��
        If blnDo Then
            If Not CheckDelPriv(strNO, mstrPrivsOpt, strTime, bytType, 0) Then Screen.MousePointer = 0: Exit Sub  '���ټ�����ȡ������,������ܲ��ϵ�����ʾ
        End If
        
        'ȫԺ����
        If blnDo And InStr(mstrPrivsOpt, ";ȫԺ����;") = 0 Then
            If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
            
            If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(i, lngCol��������ID)) & ",") = 0 Then
                strUnDelNOs = strUnDelNOs & "," & strNO
                blnDo = False
            End If
        End If
            
        '���۲���Ȩ��
        If blnDo Then
            strInfo = Check���۲���(strNO, mstrPrivsOpt, strTime, bytType)
            If strInfo <> "" Then
                Screen.MousePointer = 0
                MsgBox "����""" & strNO & """�а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
                Exit Sub '���ټ�����ȡ������
            End If
        End If
        
        '�Ƿ���ִ��
        If blnDo Then
            blnBat = Val(mshList.TextMatrix(i, GetColNum("�ಡ�˵�"))) = 1
            If BillCanDelete(strNO, bytType, blnBat, strTime, , blnFlagPrint) <> 0 Then blnDo = False
            If blnFlagPrint Then
                If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        
        '��Ժ���˲���Ȩ���ж�
        If blnDo Then
            If Not BillCanBeOperate(strNO, mstrPrivsOpt, "��������", strTime, , bytType) Then Screen.MousePointer = 0: Exit Sub
        End If
        
        '�Ƿ��Ѿ�����(�еĻ�ֻ��һ��)
        If blnDo Then
            If gbytBillOpt <> 0 Then
                If HaveBilling(2, strNO, True, strTime, bytType) <> 0 Then
                    If gbytBillOpt = 2 Then
                        blnDo = False
                    ElseIf gbytBillOpt = 1 Then
                        If j = 0 Then
                            j = j + 1
                            If MsgBox("�����Ѿ����ʵĵ���Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                blnDo = False: blnBilling = False
                            End If
                        Else
                            blnDo = blnBilling
                        End If
                    End If
                End If
            End If
        End If
        
        '�Ƿ���ڴ��۳����¼
        If blnDo Then
            If CheckRecalcRecord(strNO) Then
                strRebateNOS = strRebateNOS & strNO & ","
                If (UBound(Split(strRebateNOS, ",")) Mod 8) = 0 Then strRebateNOS = strRebateNOS & vbCrLf
            End If
        End If
                
        'ȡ�����ʵ��к�(������˵ĵ���)
        str��� = ""
        If blnDo Then
            If Not BillIdentical(strNO) Then            '�����ж��Զ����ʵ�
                str��� = GetDelSerial(strNO, strTime)
                If str��� = "" Then blnDo = False
            End If
        End If
        
        'ҽ�����˵ķ��ò�������������
        If blnDo And tbs.SelectedItem.Key = "Auditing" Then '��������ʱ����
            intInsure = BillExistInsure(strNO, , , bytType) '�ж��Ƿ�ҽ�����˼ǵ���
            If intInsure > 0 Then
                Screen.MousePointer = 0
                MsgBox "ҽ�����˵ļ��ʷ��ò������������ʣ�", vbInformation, gstrSysName
                mshList.Row = i: mshList.TopRow = i
                Call mshList_EnterCell: Exit Sub
            End If
        End If

        '����SQL
        If blnDo Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_סԺ���ʼ�¼_Delete('" & strNO & "','" & str��� & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & bytType & ")"
        End If
    Next
    Screen.MousePointer = 0
    
    If UBound(arrSQL) = -1 Then
        MsgBox "û�п������ʵļ�¼���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If strRebateNOS <> "" Then
       MsgBox "�������µ��ݴ��ڰ��ѱ�����Ĵ��۳����¼:" & vbCrLf & Mid(strRebateNOS, 1, InStrRev(strRebateNOS, ",") - 1) & vbCrLf & _
                "����ǰ�����Щ���ݵĲ���������ã������˽����ܵ�������ǰ�Ĵ����Żݽ�", vbInformation, Me.Caption
    End If
    
    'ִ�й���
    Call ZLCommFun.ShowFlash("����ִ����������,���Ժ� ...", Me)
    DoEvents
    Me.Refresh
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    
    Call ZLCommFun.StopFlash
    Me.Refresh
    
    If strUnDelNOs <> "" Then
        MsgBox "��û��[ȫԺ����]��Ȩ��,�����������ҵĵ���δ����." & vbCrLf & Mid(strUnDelNOs, 2), vbInformation, gstrSysName
    End If
    
    Call mnuViewReFlash_Click
    Exit Sub
errH:
    Call ZLCommFun.StopFlash
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditGive_Click()
    Dim rsTmp As ADODB.Recordset
    Dim arrSQL() As String, i As Long
    Dim strSql  As String, blnTran As Boolean, bln���ʱ� As Boolean
    Dim strNO As String, strTime As String
    Dim str�������� As String, strDate As String, str���ܺ� As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ��Է�ҩ��", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    str�������� = mshList.TextMatrix(mshList.Row, GetColNum("��������"))
    bln���ʱ� = Len(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 0
    On Error GoTo errH
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    'ֻ����ָ��ʱ����˲��ݵ�����
    Set rsTmp = Get����ҩ�嵥(strNO, strTime, bln���ʱ�)
    
    If rsTmp.EOF Then
        MsgBox "����""" & strNO & """��ǰ������û�п��Է��ŵ�ҩƷ��", vbInformation, gstrSysName
        Exit Sub
    Else
        If IsNull(rsTmp!�ⷿID) Then
            MsgBox "���ŵ��ݵ�ǰ����δȷ��ִ��ҩ�������������﷢ҩ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪ�Ե���""" & strNO & """��ǰ���ݷ�ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        ReDim arrSQL(rsTmp.RecordCount - 1)
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        str���ܺ� = zlDatabase.GetNextNo(20)
        
        For i = 0 To rsTmp.RecordCount - 1
            arrSQL(i) = "ZL_ҩƷ�շ���¼_���ŷ�ҩ(" & rsTmp!�ⷿID & "," & rsTmp!ID & ",'" & UserInfo.���� & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),Null,Null,Null," & str���ܺ� & ")"
            rsTmp.MoveNext
        Next
    End If
    
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
        
    Call mshList_EnterCell
    
    '��ӡ��ҩ�嵥
    '�����,ĿǰZL1_BILL_1133_2�в�û���õ�����:��������,������Ϊ��ĳ���û��ķ�ҩ�����������
    If MsgBox("����""" & strNO & """��ҩ��ɣ�Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133_2", Me, "���ݺ�=" & strNO, "�Ǽ�ʱ��=" & strTime, str��������, 1)
    End If
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPriceBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 1
    frmCharge.mbytUseType = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditPriceSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 1
    frmSimpleBilling.mbytUseType = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditPriceTable_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 1
    frmBillings.mbytUseType = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditAuditingBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 2
    frmCharge.mbytUseType = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditAuditingPati_Click()
    Err.Clear: On Error Resume Next
    If Not frmBillingAuditing.zlCardShow(Me, mlngModul, mstrPrivs, mlngUnitID) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditAuditingSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 2
    frmSimpleBilling.mbytUseType = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditAuditingTable_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 2
    frmBillings.mbytUseType = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditBillingBilling_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mbytUseType = 0
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInState = 0
    frmCharge.mlngDeptID = mlngDeptID
    frmCharge.mlngUnitID = mlngUnitID
    frmCharge.mlngModule = mlngModul
    frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditBillingSimple_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmSimpleBilling.mbytUseType = 0
    frmSimpleBilling.mstrPrivs = mstrPrivs
    frmSimpleBilling.mbytInState = 0
    frmSimpleBilling.mlngDeptID = mlngDeptID
    frmSimpleBilling.mlngUnitID = mlngUnitID
    frmSimpleBilling.mlngModule = mlngModul
    frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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

Private Sub mnuEditBillingTable_Click()
    Err.Clear
    On Error Resume Next
    
    gbytBilling = 0
    frmBillings.mbytUseType = 0
    frmBillings.mstrPrivs = mstrPrivs
    frmBillings.mbytInState = 0
    frmBillings.mlngDeptID = mlngDeptID
    frmBillings.mlngUnitID = mlngUnitID
    frmBillings.mlngModule = mlngModul
    frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
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
        MsgBox "�����а�������δ��ȫ��˻�ֶ����˵����ݣ��������޸ġ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'Ȩ���ж�
    If tbs.SelectedItem.Key = "Price" Then
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("������")), _
            CDate(mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))), "�޸�", strNO) Then Exit Sub
    Else
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("������")), _
            CDate(mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))), "�޸�", strNO) Then Exit Sub
    End If
    
    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, mstrPrivsOpt)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '��Ժ���˲���Ȩ���ж�
    If Not BillCanBeOperate(strNO, mstrPrivsOpt, "�޸�") Then Exit Sub
    
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
    
    'ȫԺ����
    If InStr(mstrPrivsOpt, ";ȫԺ����;") = 0 Then
        If strUnitIDs = "" Then strUnitIDs = GetUserUnits(True)
        
        If InStr("," & strUnitIDs & ",", "," & Val(mshList.TextMatrix(mshList.Row, GetColNum("��������ID"))) & ",") = 0 Then
            MsgBox "��û��Ȩ�޶��������ҵĵ�������,�������޸ĸõ��ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�����������ִ�л�ȫ��ִ�е���Ŀ,��һ������ȫ������,�������޸�
    If HaveExecute(2, strNO, 2) Then
        MsgBox "�õ����а�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '�ѽ��ʵ����ж�
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
    
    '�Ƿ������������¼
    If CheckRecalcRecord(strNO) Then
        MsgBox "���ָü��ʵ��ݴ��ڰ��ѱ�����Ĵ��۳����¼!" & vbCrLf & _
            "����ǰ�밴�ѱ�������ã������˽����ܵ����޸�ǰ�Ĵ����Żݽ�", vbInformation, Me.Caption
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
        
    If tbs.SelectedItem.Key = "Auditing" Then
        gbytBilling = 0 '�����޸�
    Else
        gbytBilling = 1 '�����޸�
    End If
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("�ಡ�˵�"))) = 1 Then '��������
        frmBillings.mbytUseType = 0
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 0
        frmBillings.mstrInNO = strNO
        frmBillings.mlngDeptID = mlngDeptID
        frmBillings.mlngUnitID = mlngUnitID
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mbytUseType = 0
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 0
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mlngDeptID = mlngDeptID
        frmSimpleBilling.mlngUnitID = mlngUnitID
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mbytUseType = 0
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 0
            frmCharge.mstrInNO = strNO
            frmCharge.mlngDeptID = mlngDeptID
            frmCharge.mlngUnitID = mlngUnitID
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 0, 0, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK Then
        If gstrModiNO <> "" Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,�޸ĺ�ĵ��ݺ�Ϊ:[" & gstrModiNO & "],Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        Else
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
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
        
    If InStr(",0,1,", Val(mshList.TextMatrix(mshList.Row, GetColNum("����")))) = 0 Then
        MsgBox "�õ���Ϊ���ʵ��ݻ��ѱ����ʣ������ٴ�ӡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1133", Me, "NO=" & strNO, "�Ǽ�ʱ��=" & strTime, "ҩƷ��λ=" & IIf(gblnסԺ��λ, 1, 0), "PrintEmpty=0", "�ش�=1", 2)
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim bln�������� As Boolean
    Dim blnסԺ��λ As Boolean
    
    bln�������� = gbln��������
    blnסԺ��λ = gblnסԺ��λ
    
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytInFun = 0
    frmSetExpence.mbytUseType = 0
    frmSetExpence.Show 1, Me
    If gblnOK Then
        If bln�������� <> gbln�������� Then
            '���۲���
            mlngDeptID = -1: mlngUnitID = 0: mstrPage = ""
            Call InitUnits
        ElseIf blnסԺ��λ <> gblnסԺ��λ Then
            If Not (mshList.Rows = 2 And mshList.TextMatrix(1, GetColNum("���ݺ�")) = "") Then
                Call mnuViewReFlash_Click
            End If
        End If
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "����=" & mlngUnitID, "���˿���=" & mlngDeptID)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����=" & mlngUnitID, "���˿���=" & mlngDeptID, "NO=" & strNO, _
                "סԺ��=" & .TextMatrix(.Row, GetColNum("סԺ��")), _
                "����ID=" & .TextMatrix(.Row, GetColNum("����ID")), _
                "��ҳID=" & .TextMatrix(.Row, GetColNum("��ҳID")), _
                "������=" & .TextMatrix(.Row, GetColNum("������")))
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    With frmBillingFilter
        .mstrPrivs = mstrPrivs
        If .mlngDeptID <> mlngDeptID Then
            .mlngDeptID = mlngDeptID
            .mlngUnitID = mlngUnitID
            .LoadOper
        End If
        
        If tbs.SelectedItem.Key = "Auditing" Then
            .lbl����Ա.Caption = "������"
        Else
            .lbl����Ա.Caption = "������"
        End If
        
        .Show 1, Me
        If gblnOK Then
            mvNurseFilter.Nurse = False '�ֹ����˺���ʹ�û�ʿ����վ���õ�����
            
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
            
            
            mstr����Ա = ""
            If .cbo����Ա.ListIndex <> -1 Then
                If .cbo����Ա.ItemData(.cbo����Ա.ListIndex) <> 0 Then
                    mstr����Ա = zlStr.NeedName(.cbo����Ա.Text)
                End If
            End If
        
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.Operator = mstr����Ա
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.InPatientID = Val(.txtסԺ��.Text)
            SQLCondition.Patient = gstrLike & UCase(.txt����.Text) & "%"
            SQLCondition.FeeItems = .mstrFeeItems
            SQLCondition.IncomeItems = .mstrIncomeItems
            
            mnuViewReFlash_Click
        End If
    End With
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
                
        '����Ҫ�Ǽ�ʱ��(�˷ѻ򲿷����)
        strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, strTime, blnDel, True)
    End If
End Sub

Private Sub mshList_DblClick()
    Dim lngCOL��� As Long, i As Long
    Dim lngCOLNO As Long
        
    If tbs.SelectedItem.Key = "Price" Then
        With mshList
            If .MouseRow > 0 Then
                If .TextMatrix(.Row, GetColNum("���ݺ�")) <> "" Then
                    lngCOL��� = GetColNum("���")
                    If .MouseCol = lngCOL��� Then
                        If .TextMatrix(.Row, lngCOL���) = "��" Then
                            .TextMatrix(.Row, lngCOL���) = ""
                        Else
                            .TextMatrix(.Row, lngCOL���) = "��"
                        End If
                    Else
                        If mnuEditView.Enabled Then mnuEditView_Click
                    End If
                End If
            ElseIf .MouseRow = 0 And .Rows > 1 Then
                lngCOL��� = GetColNum("���")
                If .MouseCol = lngCOL��� Then
                    lngCOLNO = GetColNum("���ݺ�")
                    For i = 1 To mshList.Rows - 1
                        If .TextMatrix(i, lngCOLNO) <> "" Then
                            If .MouseCol = lngCOL��� Then
                                If .TextMatrix(i, lngCOL���) = "��" Then
                                    .TextMatrix(i, lngCOL���) = ""
                                Else
                                    .TextMatrix(i, lngCOL���) = "��"
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End With
    Else
        If mnuEditView.Enabled Then mnuEditView_Click
    End If
End Sub

Private Sub mshList_EnterCell()
    Dim strNO As String, strTime As String, blnDel As Boolean
        
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If mshList.Row = 0 Or strNO = "" Then Exit Sub
    
    stbThis.Panels(2).Text = "�� " & Nvl(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(Nvl(mrsTotal!���, 0), gstrDec)
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    '����Ҫ�Ǽ�ʱ��(�˷ѻ򲿷����)
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2
    
    mnuEditAdjust.Enabled = Not blnDel
    '�Զ����ʵ���ҽ�����ɵļ��ʵ��������޸�
    mnuEditModi.Enabled = Not blnDel And Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����"))) <> 3 _
                            And mshList.TextMatrix(mshList.Row, GetColNum("��������")) = "��ͨ����"
    mnuEditGive.Enabled = Not blnDel And tbs.SelectedItem.Key = "Auditing"
    mnuEditDel.Enabled = Not blnDel
    mnuEditDelBat.Enabled = Not blnDel
    
    tbr.Buttons("Modi").Enabled = mnuEditModi.Enabled
    tbr.Buttons("Give").Enabled = mnuEditGive.Enabled
    tbr.Buttons("Del").Enabled = mnuEditDel.Enabled
        
        
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 And tbs.SelectedItem.Key <> "Auditing" Then
        mnuEditDel.Enabled = False
        mnuEditDelBat.Enabled = False
        tbr.Buttons("Del").Enabled = False
    End If
        
    mshList.ForeColorSel = mshList.CellForeColor
    
    Call ShowDetail(strNO, strTime, blnDel)
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
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If Me.ActiveControl Is cboDept Then
            Else
                If mnuEditView.Enabled Then mnuEditView_Click
            End If
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEditDel_Click()
    Dim strNO As String, strTime As String, blnBat As Boolean, blnFlagPrint As Boolean
    Dim strInfo As String, str����IDs As String
    Dim strInsure As String, arrInsure As Variant, intInsure As Integer
    Dim intTmp As Integer, i As Long, bytType As Byte   '��¼����
    
    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�е��ݿ������ʣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    strTime = mshList.TextMatrix(mshList.Row, GetColNum("�Ǽ�ʱ��"))
    bytType = Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼����")))
        
    'Ȩ���ж�
    If tbs.SelectedItem.Key = "Price" Then
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("������")), CDate(strTime), "����", strNO) Then Exit Sub
    Else
        If Not BillOperCheck(5, mshList.TextMatrix(mshList.Row, GetColNum("������")), CDate(strTime), "����", strNO, , bytType) Then Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 2, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    '��Ŀ����Ȩ��
    If Not CheckDelPriv(strNO, mstrPrivsOpt, strTime) Then Exit Sub
    
    '���۲���Ȩ��
    strInfo = Check���۲���(strNO, mstrPrivsOpt, strTime)
    If strInfo <> "" Then
        MsgBox "�����а���" & strInfo & ",��û��Ȩ�޶Ըõ��ݽ��в�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ִ��
    blnBat = Val(mshList.TextMatrix(mshList.Row, GetColNum("�ಡ�˵�"))) = 1
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
    
    '�Ƿ��Ѿ�����:0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    intTmp = HaveBilling(2, strNO, False, strTime)
    If intTmp <> 0 Then
        Call GetBillInsures(strInsure, strNO, , , True, bytType)
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If arrInsure(i) <> 0 Then
                    If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, , arrInsure(i)) Then
                        'ҽ�����˵ĵ���,�̶�Ϊ�ѽ��ʵĽ�ֹ���ʡ�
                        If intTmp = 1 Then
                            MsgBox "��ҽ�����ʵ���δ���ʲ����Ѿ�����,�������ʣ�", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            '����˵��Ϊҽ�����˱���ȫ������,����Ӧ�ò�������������
                            '���ܳ�����ҽ������ͨ���˻�ϵļ��ʱ�,δ��ȷ����
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
    
    '����ģʽ(�ڼ��ʻ򻮼�ʱֱ�ӵ��뵥�����ʵ�ģʽ�����ڱ����ģʽ)
    If tbs.SelectedItem.Key = "Auditing" Then
        gbytBilling = 0 '��������
    Else
        gbytBilling = 1 '��������
    End If
    If blnBat Then '��������
        frmBillings.mbytUseType = 0
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
        frmSimpleBilling.mbytUseType = 0
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
            If mvNurseFilter.Nurse And mvNurseFilter.ҽ��ID <> 0 Then
                frmCharge.mlngҽ��ID = mvNurseFilter.ҽ��ID
            End If
            
            frmCharge.mbytUseType = 0
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
            varTemp = Array(lng����ID, 0, 3, strNO, mlngUnitID, mlngDeptID, 0, mstrPrivs)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
            
            gblnOK = varTemp
        End If
    End If

    If gblnOK And Visible Then '��ʿվ����
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
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
    
    If tbs.SelectedItem.Key = "Auditing" Then
        gbytBilling = 0 '���ʲ���
    Else
        gbytBilling = 1 '���۲���
    End If
    
    If Val(mshList.TextMatrix(mshList.Row, GetColNum("�ಡ�˵�"))) = 1 Then '��������
        frmBillings.mstrPrivs = mstrPrivs
        frmBillings.mbytInState = 1
        frmBillings.mstrInNO = strNO
        frmBillings.mblnNOMoved = mblnNOMoved
        frmBillings.mstrTime = strTime
        frmBillings.mblnDelete = blnDel
        frmBillings.mlngModule = mlngModul
        frmBillings.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    ElseIf BillisSimple(strNO) Then '�򵥼���
        frmSimpleBilling.mstrPrivs = mstrPrivs
        frmSimpleBilling.mbytInState = 1
        frmSimpleBilling.mstrInNO = strNO
        frmSimpleBilling.mblnNOMoved = mblnNOMoved
        frmSimpleBilling.mstrTime = strTime
        frmSimpleBilling.mblnDelete = blnDel
        frmSimpleBilling.mlngModule = mlngModul
        frmSimpleBilling.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else '���ʵ�
        Dim lng����ID As Long
        Dim varTemp As Variant
        
        lng����ID = mshList.TextMatrix(mshList.Row, GetColNum("���ʵ�ID"))
        
        If lng����ID = 0 Or gobjCustBill Is Nothing Then
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mstrInNO = strNO
            frmCharge.mblnNOMoved = mblnNOMoved
            frmCharge.mstrTime = strTime
            frmCharge.mblnDelete = blnDel
            frmCharge.mlngModule = mlngModul
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            '����ID��bytUseType��bytInState��strInNO��lngUnitID��lngDeptID��lng����ID��mstrPrivs
            varTemp = Array(lng����ID, 0, 1, strNO, 0, 0, 0, mstrPrivs, blnDel)
            gobjCustBill.CodeMan glngSys, -1, gcnOracle, Me, gstrDBUser, varTemp
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
    If mvNurseFilter.Nurse Then
        Call ShowBillsByNurse
    Else
        Call ShowBills(mstrFilter)
    End If
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
            mnuEditBillingBilling_Click
        Case "Price"
            mnuEditPriceBilling_Click
        Case "Auditing"
            mnuEditAuditingBilling_Click
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
        Case "Give"
            mnuEditGive_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim lngCount As Integer
    Dim str���ʵ�ID As String
    
    Select Case ButtonMenu.Key
        Case "BillingBilling"
            mnuEditBillingBilling_Click
        Case "BillingTable"
            mnuEditBillingTable_Click
        Case "BillingSimple"
            mnuEditBillingSimple_Click
        Case "PriceBilling"
            mnuEditPriceBilling_Click
        Case "PriceTable"
            mnuEditPriceTable_Click
        Case "PriceSimple"
            mnuEditPriceSimple_Click
        Case "AuditingBilling"
            mnuEditAuditingBilling_Click
        Case "AuditingTable"
            mnuEditAuditingTable_Click
        Case "AuditingSimple"
            mnuEditAuditingSimple_Click
        Case "AuditingPati"
            mnuEditAuditingPati_Click
        Case "AuditingBatch"
            mnuEditAuditingBatch_Click
        Case Else
            '�Զ������
            str���ʵ�ID = Mid(ButtonMenu.Key, 2)
            For lngCount = mnuEditBillingCust.LBound To mnuEditBillingCust.UBound
                If str���ʵ�ID = mnuEditBillingCust(lngCount).Tag Then
                    Call mnuEditBillingCust_Click(lngCount)
                    Exit Sub
                End If
            Next
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
    With frmBillingFilter
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
    
    mnuEditAdjust.Enabled = blnUsed
    mnuEditModi.Enabled = blnUsed
    tbr.Buttons("Modi").Enabled = blnUsed
    
    mnuEditGive.Enabled = blnUsed And tbs.SelectedItem.Key = "Auditing"
    tbr.Buttons("Give").Enabled = mnuEditGive.Enabled
    
    mnuEditDel.Enabled = blnUsed
    mnuEditDelBat.Enabled = blnUsed
    mnuEditView.Enabled = blnUsed
    mnuEditPrint.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
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
        mnuEditBillingCust(0).Visible = False
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    '��������ɹ����ٶ�����Ӧ�Ĳ˵�
    If Not gobjCustBill Is Nothing Then
        gstrSQL = "Select ID,���� From �շѼ��ʵ� Where substr(���÷�Χ,2,1)='1' Order by ���"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        lngSum = rsTmp.RecordCount
    End If
    
    If lngSum > 0 Then
        For lngCount = 1 To lngSum
            '���ӵ����˵���
            Load mnuEditBillingCust(lngCount)
            mnuEditBillingCust(lngCount).Caption = rsTmp("����") & "(&" & lngCount & ")"
            mnuEditBillingCust(lngCount).Tag = rsTmp("ID")
            '�����������˵���
            If lngCount = 1 Then
                tbr.Buttons("Billing").ButtonMenus.Add , , "-"
            End If
            tbr.Buttons("Billing").ButtonMenus.Add , "C" & rsTmp("ID"), rsTmp("����")
            
            rsTmp.MoveNext
        Next
    Else
        mnuEditBillingCust(0).Visible = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
'˵������Ϊ�����屻��ʿվ��ģ̬���ã�����ǿ���ظ�ִ��Form_Load���г�ʼ��,�����Щ���ǰ��Visible�����ж�
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    mlngModul = glngModul
    
    If Not Visible Then
        Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
        Call SetCustBill  '�����Զ�����ʵ�
        Call RestoreWinState(Me, App.ProductName)
        Set stbThis.Panels(5).Picture = Me.Picture
    
        'ˢ�·�ʽ
        For i = 0 To mnuViewRefeshOptionItem.UBound
            If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
                mnuViewRefeshOptionItem(i).Checked = True
            Else
                mnuViewRefeshOptionItem(i).Checked = False
            End If
        Next
    End If
    
    If mvNurseFilter.Nurse Then
        tbs.Tabs(IIf(mvNurseFilter.����, "Price", "Auditing")).Selected = True
    ElseIf Not Visible Then
        i = IIf(zlDatabase.GetPara("ҳ��", glngSys, mlngModul, "1") = "1", 1, 2)
        tbs.Tabs(i).Selected = True
    End If
            
    mlngCurRow = 1: mlngTopRow = 1
    
    'Ȩ������
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 Then
        mnuEditBilling.Visible = False
        tbr.Buttons("Billing").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";סԺ����;") = 0 Then
        mnuEditPrice.Visible = False
        tbr.Buttons("Price").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";�������;") = 0 Then
        mnuEditAuditing.Visible = False
        tbr.Buttons("Auditing").Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";ҩƷ��ҩ;") = 0 Then
        mnuEditGive.Visible = False
        mnuEditGive_.Visible = False
        tbr.Buttons("Give").Visible = False
        tbr.Buttons("Give_").Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 Then
        mnuEditModi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivsOpt, ";��¼����;") = 0 Then
        mnuEditAdjust.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";��¼�޸�;") = 0 _
        And InStr(mstrPrivsOpt, ";��¼����;") = 0 Then
        mnuEditAdjust_.Visible = False
    End If
    '55380
    If InStr(mstrPrivsOpt, ";ҩƷ����;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 _
        And InStr(mstrPrivsOpt, ";��������;") = 0 Then
        mnuEditDel.Visible = False
        mnuEditDelBat.Visible = False
        '55380
        If InStr(mstrPrivsOpt, ";ҩƷ��������;") = 0 _
            And InStr(mstrPrivsOpt, ";������������;") = 0 _
            And InStr(mstrPrivsOpt, ";������������;") = 0 _
            And InStr(mstrPrivsOpt, ";�������;") = 0 Then
            mnuEditDel_.Visible = False
        End If
        
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Del_").Visible = False
        
        '��ʿ����վ����ʱ,Ȩ�޲�����ʾ
        If mvNurseFilter.Nurse Then
            MsgBox "�㲻����סԺ���ʹ���ģ���Ӧ������Ȩ�ޡ�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    If InStr(mstrPrivsOpt, ";ҩƷ��������;") = 0 _
        Or InStr(mstrPrivsOpt, ";������������;") = 0 _
        Or InStr(mstrPrivsOpt, ";������������;") = 0 _
        Or mvNurseFilter.Nurse _
        Or InStr(1, mstrPrivsOpt, "��������") = 0 Then
        mnuEditDelApply.Visible = False
    End If
    If InStr(mstrPrivsOpt, ";�������;") = 0 Or mvNurseFilter.Nurse Then
        mnuEditDelAudit.Visible = False
    End If
    
    If InStr(mstrPrivsOpt, ";�ش򵥾�;") = 0 Then
        mnuEditPrint.Visible = False
    End If
    
    
    '��������ҳ�ʼ
    If Not InitUnits Then Unload Me: Exit Sub
    If cboDept.ListIndex = -1 Then
        MsgBox "û�з�������������,���㲻�������в���Ȩ��,����ʹ��סԺ���ʹ���", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If

    mstrPage = tbs.SelectedItem.Key
    
    mbln���� = True
    mbln���� = False
    mstr����Ա = UserInfo.����
    
    Call SetHeader
    Call SetDetail
    Call SetMenu(False)
    
    If mvNurseFilter.Nurse Then
        Call mnuViewReFlash_Click
        If mvNurseFilter.ReLoad Then
            If Me.WindowState = 1 Then Me.WindowState = 0
            mvNurseFilter.ReLoad = False
        End If
        
        '����ָ������ʱ���Զ��������г���
        If mvNurseFilter.���ݺ� <> "" Then
            Call mnuEditDel_Click
            If Not Visible Then
                mvNurseFilter.Mode = True
                Unload Me: Exit Sub
            End If
        End If
    Else
        stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
    End If
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
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = Me.ScaleTop + cbrH + 15
    
    mshList.Left = 0
    mshList.Top = tbs.Top + tbs.TabFixedHeight + 30
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = Me.ScaleLeft
    picHsc.Width = Me.ScaleWidth
    
    mshDetail.Left = Me.ScaleLeft
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Width = Me.ScaleWidth
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - (tbs.TabFixedHeight + 45) - picHsc.Height - mshList.Height
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    mvNurseFilter.Nurse = False
    mstrPrivs = ""
    mstrPrivsOpt = ""
    mstrFilter = ""
    mlngUnitID = 0
    mlngDeptID = 0
    
    mstrҽ����Ч = ""
    Unload frmBillingFilter
    Unload frmBillingGo
    Call SaveWinState(Me, App.ProductName)
    zlDatabase.SetPara "ҳ��", tbs.SelectedItem.Index, glngSys, mlngModul, blnHavePrivs
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
End Sub

Private Sub mnuViewGo_Click()
    frmBillingGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmBillingGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, j As Long, blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents

        '�Ƚ�����
        blnFill = True
        With frmBillingGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ݺ�")) = .txtNO.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("����")) = .txt����.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
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
        If mshList.MouseCol = GetColNum("���") Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        If mvNurseFilter.Nurse Then
            Call ShowBillsByNurse(True)
        Else
            Call ShowBills(, True)
        End If
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    If tbs.SelectedItem.Key = "Auditing" Then
        strHead = "����,1,0|��������,1,900|���ݺ�,1,850|סԺ��,1,750|����,1,500|����,1,700|�ѱ�,1,900|ҽ�Ƹ��ʽ,1,1400|Ӧ�ս��,7,850|ʵ�ս��,7,850" & _
                "|��������,1,1000|������,1,800|������,1,800|������,1,800|�Ǽ�ʱ��,1,1850|˵��,1,850|����,1,0|��¼����,1,0|�ಡ�˵�,1,0|���ʵ�ID,1,0|����ID,1,0|��ҳID,1,0|��������ID,1,0"
    Else
        strHead = "���,1,450|����,1,0|��������,1,900|���ݺ�,1,850|סԺ��,1,750|����,1,500|����,1,700|�ѱ�,1,900|ҽ�Ƹ��ʽ,1,1400|Ӧ�ս��,7,850|ʵ�ս��,7,850" & _
                "|��������,1,1000|������,1,800|������,1,800|������,1,800|�Ǽ�ʱ��,1,1850|˵��,1,850|����,1,0|��¼����,1,0|�ಡ�˵�,1,0|���ʵ�ID,1,0|����ID,1,0|��ҳID,1,0|��������ID,1,0"
    End If
    
    With mshList
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or (tbs.SelectedItem.Key <> mstrPage) Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or (tbs.SelectedItem.Key <> mstrPage) Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 320
        
        i = GetColNum("����"): mshList.ColWidth(i) = 0
        i = GetColNum("��¼����"): mshList.ColWidth(i) = 0
        i = GetColNum("�ಡ�˵�"): mshList.ColWidth(i) = 0
        i = GetColNum("���ʵ�ID"): mshList.ColWidth(i) = 0
        
        If tbs.SelectedItem.Key = "Auditing" Then
            mshList.ColWidth(GetColNum("������")) = 800
        Else
            mshList.ColWidth(GetColNum("������")) = 0
        End If
        
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
    Dim strSql As String, strҽ����Ч As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        Call ZLCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        
        'ȱʡ��������(һ����)
        If strIF = "" Then
            strIF = " And �Ǽ�ʱ�� Between trunc(sysdate) And trunc(sysdate+1)-1/24/60/60"
            If tbs.SelectedItem.Key = "Auditing" Then
                strIF = strIF & "  And ��¼����=2 And ��¼״̬ IN(1,3)"
            Else
                strIF = strIF & " And ��¼����=2 And ��¼״̬=0"
            End If
            mstrҽ����Ч = ""   'ȱʡΪ��ͨ����+����+����
        End If
        
        '����Ա��������
        If mstr����Ա <> "" Then
            If tbs.SelectedItem.Key = "Auditing" Then
                strIF = strIF & " And ����Ա����||''=[7]"
            Else
                strIF = strIF & " And ������||''=[7]"
            End If
        End If
        
        '����������,���в���ʱ���˲���ID=0
        If mlngUnitID > 0 Then strIF = strIF & " And ���˲���ID+0=[8]"
        
        '��¼����(�Զ����ʵ�)
        strIF = "Where �����־=2 " & strIF

        
        '���ݺ�,סԺ��,����,����,�ѱ�,Ӧ�ս��,ʵ�ս��,��������,������,������,������,�Ǽ�ʱ��,˵��,����,��¼����,�ಡ�˵�,���ʵ�ID
        'Sign(ִ��״̬):������������ʱ����ͬʱ�б�Ҫ,���Զ�����
        If tbs.SelectedItem.Key = "Auditing" Then
            strIF = strIF & " And ����Ա���� IS NOT NULL"
            
            'ɸѡʱ��ʱ�������һ��ת��֮ǰ,�ҵ�ǰ�б��ǻ��۵�
            If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
                strIF = zlGetFullFieldsTable("סԺ���ü�¼", 2, strIF, False)
            Else
                strIF = zlGetFullFieldsTable("סԺ���ü�¼", 0, strIF, False)
            End If
            
            strSql = _
                "Select Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.����) ����,Decode(A.��¼����,3,'�Զ�����',Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����')) as ��������, A.NO as ���ݺ�," & _
                " To_Number(Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.��ʶ��)) as סԺ��," & _
                " To_Char(Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.��Ժ����)) as ����," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.����) as ����," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.�ѱ�) as �ѱ�,Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,B.����) as ��������," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.������) as ������," & _
                " A.������,A.����Ա���� as ������,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
                " Decode(A.��¼����,3,Decode(Max(A.��¼״̬),2,'�Զ�����','�Զ�����'),Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼')) as ˵��," & _
                " Max(A.��¼״̬) as ����,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,Decode(Nvl(A.�ಡ�˵�,0),1,0,A.����ID) ����ID,Decode(Nvl(A.�ಡ�˵�,0),1,0,A.��ҳID) ��ҳID,A.��������ID" & _
                " From (" & strIF & ") A,���ű� B,������ҳ C,����ҽ����¼ D" & _
                " Where A.��������ID=B.ID And A.����ID=C.����ID And A.��ҳID=C.��ҳID " & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                " And A.ҽ�����=D.id(+) " & mstrҽ����Ч & _
                " Group by Sign(Decode(Nvl(A.ִ��״̬,0),0,1,Nvl(A.ִ��״̬,0))),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.����),Decode(A.��¼����,3,'�Զ�����',Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����')),A.NO," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,B.����),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.������)," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.��ʶ��),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.��Ժ����)," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.����),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.�ѱ�)," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.ҽ�Ƹ��ʽ)," & _
                " A.������,A.����Ա����,A.�Ǽ�ʱ��,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,0,A.����ID),Decode(Nvl(A.�ಡ�˵�,0),1,0,A.��ҳID),A.��������ID" & _
                " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        Else
            strIF = strIF & " And ����Ա���� IS NULL And ������ is Not NULL"
            
                    'ɸѡʱ��ʱ�������һ��ת��֮ǰ,�ҵ�ǰ�б��ǻ��۵�
            If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
                strIF = zlGetFullFieldsTable("סԺ���ü�¼", 2, strIF, False)
            Else
                strIF = zlGetFullFieldsTable("סԺ���ü�¼", 0, strIF, False)
            End If
        
            strSql = _
                "Select '��' ���,Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.����) ����,Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����') as ��������,A.NO as ���ݺ�," & _
                " To_Number(Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.��ʶ��)) as סԺ��," & _
                " To_Char(Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.��Ժ����)) as ����," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.����) as ����," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.�ѱ�) as �ѱ�,Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
                " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,B.����) as ��������," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.������) as ������," & _
                " A.������,A.����Ա���� as ������,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
                " Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼') as ˵��,Max(A.��¼״̬) as ����,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,Decode(Nvl(A.�ಡ�˵�,0),1,0,A.����ID) ����ID,Decode(Nvl(A.�ಡ�˵�,0),1,0,A.��ҳID) ��ҳID,A.��������ID" & _
                " From (" & strIF & ") A,���ű� B,������ҳ C,����ҽ����¼ D" & _
                " Where A.��������ID=B.ID And A.����ID=C.����ID And A.��ҳID=C.��ҳID" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                " And A.ҽ�����=D.id(+) " & mstrҽ����Ч & _
                " Group by Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.����),Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����'),A.NO," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,B.����),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.������)," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.��ʶ��),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.��Ժ����)," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.����),Decode(Nvl(A.�ಡ�˵�,0),1,NULL,A.�ѱ�)," & _
                " Decode(Nvl(A.�ಡ�˵�,0),1,NULL,C.ҽ�Ƹ��ʽ)," & _
                " A.������,A.����Ա����,A.�Ǽ�ʱ��,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,Decode(Nvl(A.�ಡ�˵�,0),1,0,A.����ID),Decode(Nvl(A.�ಡ�˵�,0),1,0,A.��ҳID),A.��������ID" & _
                " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        End If
        With SQLCondition
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, mstr����Ա, mlngUnitID, .FeeItems, .IncomeItems)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        '��ʵ�պϼƽ��
        If Not blnSort Then
            strSql = "Select Sum(ʵ�ս��) as ���,Count(Distinct NO) as ���� From (" & _
                Replace(strIF, "��¼״̬ IN(1,3)", "��¼״̬ IN(1,2,3)") & ") A,���ű� B Where A.��������ID = B.ID" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)"
            With SQLCondition
            Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .InPatientID, .Patient, mstr����Ա, mlngUnitID, .FeeItems, .IncomeItems)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "�� " & Nvl(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(Nvl(mrsTotal!���, 0), gstrDec)
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    '������ɫ
    If mbln���� And Not mbln���� And tbs.SelectedItem.Key = "Auditing" Then
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
    
    If Not blnSort Then Call ZLCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowBillsByNurse(Optional blnSort As Boolean)
'����:����ʿ����վ����������ȡҪ�����ĵ����б�
'����:blnSort=�����¶�ȡ����,��������ʾ�����������
'˵��:��ΪҪ�������ʲ���,ֻ��ȡ��������;��ʿ����վ����֮ǰӦ��������ж�
    Dim strSql As String, strIF As String
    Dim i As Long, j As Long, k As Long
    
    On Error GoTo errH
    
    If Not blnSort Then
        '��ʿ����վ������
        If mvNurseFilter.���ݺ� <> "" Then
            strIF = " Where �����־=2 And ��¼����=2 And NO=[1]"
        ElseIf mvNurseFilter.ҽ��ID <> 0 Then
            strIF = "Select ID From ����ҽ����¼ Where ID=[2] Or ���ID=[2]"
            strIF = "Select NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[3] And ҽ��ID IN(" & strIF & ")"
            strIF = "Where �����־=2 And ��¼����=2 And NO IN(" & strIF & ")"
        ElseIf mvNurseFilter.���ͺ� <> 0 Then
            strIF = "Select Distinct NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[3]"
            strIF = "Where �����־=2 And ��¼����=2 And NO IN(" & strIF & ")"
        End If
                
        '���ʻ򻮼�
        If tbs.SelectedItem.Key = "Auditing" Then
            strIF = strIF & " And ��¼״̬ IN(1,3)"
        Else
            strIF = strIF & " And ��¼״̬=0"
        End If
        
        '���������
        If mlngDeptID > 0 Then strIF = strIF & " And ���˿���ID+0=[4]"
        strIF = zlGetFullFieldsTable("סԺ���ü�¼", 0, strIF, False)
        
        strSql = _
            "Select " & IIf(tbs.SelectedItem.Key = "Price", " NULL as ���,", "") & _
            " C.����,Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����') as ��������,A.NO as ���ݺ�," & _
            " A.��ʶ�� as סԺ��,A.����,A.����,A.�ѱ�,C.ҽ�Ƹ��ʽ," & _
            " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
            " To_Char(Sum(Decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��," & _
            " B.���� as ��������,A.������,A.������,A.����Ա���� as ������,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��," & _
            " Decode(Max(A.��¼״̬),2,'���ʼ�¼','���ʼ�¼') as ˵��," & _
            " Max(A.��¼״̬) as ����,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,A.����ID,A.��ҳID" & _
            " From (" & strIF & ") A,���ű� B,������ҳ C,����ҽ����¼ D" & _
            " Where A.��������ID=B.ID(+) And A.����ID=C.����ID And A.��ҳID=C.��ҳID And A.ҽ�����=D.ID" & _
            " Group by C.����,Decode(D.ҽ����Ч,1,'��������',0,'��������','��ͨ����')," & _
            " A.NO,B.����,A.������,A.��ʶ��,A.����,A.����,A.�ѱ�,C.ҽ�Ƹ��ʽ," & _
            " A.������,A.����Ա����,A.�Ǽ�ʱ��,A.��¼����,A.�ಡ�˵�,A.���ʵ�ID,A.����ID,A.��ҳID" & _
            " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
        With mvNurseFilter
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .���ݺ�, .ҽ��ID, .���ͺ�, mlngDeptID)
        End With
    End If
    
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        '��ʵ�պϼƽ��
        If Not blnSort Then
            strSql = "Select Sum(ʵ�ս��) as ���,Count(Distinct NO) as ���� From (" & Replace(strIF, "��¼״̬ IN(1,3)", "��¼״̬ IN(1,2,3)") & ")"
            With mvNurseFilter
                Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .���ݺ�, .ҽ��ID, .���ͺ�)
            End With
        End If
    
        Set mshList.DataSource = mrsList
        stbThis.Panels(2).Text = "�� " & Nvl(mrsTotal!����, 0) & " �ŵ���,�ϼ�:" & Format(Nvl(mrsTotal!���, 0), gstrDec)
        Call SetMenu(True)
    End If

    mshList.Redraw = False
    
    '������ɫ
    mshList.ForeColor = ForeColor
    k = GetColNum("����")
    For i = 1 To mshList.Rows - 1
        If Val(mshList.TextMatrix(i, k)) = 3 Then
            '�������ʵ�����ɫ
            mshList.Row = i
            For j = 0 To mshList.Cols - 1
                mshList.Col = j
                mshList.CellForeColor = &HC00000
            Next
        End If
    Next
        
    Call SetHeader
    If mrsList.EOF Then Call SetDetail
    
    mshList.Redraw = True

    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbs_Click()
    Dim blnVisible As Boolean

    If (tbs.SelectedItem.Key = mstrPage) And Visible Then Exit Sub      '����ʱҪ��������ж�Ȩ��
    
    
    If mvNurseFilter.Nurse Then
        blnVisible = True
    Else
        If tbs.SelectedItem.Key = "Auditing" Then
            blnVisible = InStr(mstrPrivs, ";�鿴���ʵ�;") > 0
        Else
            blnVisible = InStr(mstrPrivs, ";�鿴���۵�;") > 0
        End If
    End If
    
    mnuEditView.Visible = blnVisible
    tbr.Buttons("View").Visible = blnVisible
    mnuViewFilter.Visible = blnVisible
    tbr.Buttons("Filter").Visible = blnVisible
    mnuViewreFlash.Visible = blnVisible
    
    If Not blnVisible Then
        mnuViewRefeshOptionItem(0).Checked = True '��ˢ��
        mnuViewRefeshOptionItem(1).Checked = False
        mnuViewRefeshOptionItem(2).Checked = False
        mnuViewRefeshOptionItem(1).Enabled = False
        mnuViewRefeshOptionItem(2).Enabled = False
        
        mshList.Clear
        mshList.Rows = 2
        mshDetail.Clear
        mshDetail.Rows = 2
        
        Call SetHeader
        Call SetDetail
        Call SetMenu(False)
        
        mstrPage = tbs.SelectedItem.Key
        Exit Sub
    End If
    
    mstrFilter = ""   ' ��¼���ʱ��ˣ�Ҫ�������
    
    If Visible Then
        If mvNurseFilter.Nurse Then
            If Not mvNurseFilter.ReLoad Then Call ShowBillsByNurse
        Else
            Call ShowBills(mstrFilter) '���봰��ʱȱʡ����ʾ�κε���
        End If
    End If
    
    If mshList.Visible And mshList.Enabled Then mshList.SetFocus
    
    mstrPage = tbs.SelectedItem.Key
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ����
    Dim i As Long
    Dim strServiceRange As String
    
    On Error GoTo errH
    cboDept.Clear
    If InStr(";" & mstrPrivs, ";���в���") > 0 Then cboDept.AddItem "���в���"
        
    '��Ȩ����ʾ����۲��Ҷ�Ӧ���ٴ�����,סԺ������סԺ��ͬ
    If InStr(mstrPrivsOpt, ";�������ۼ���;") And gbln�������� Then
        strServiceRange = "1,2,3"
    Else
        strServiceRange = "2,3"
    End If
    Set mrsDept = GetUnit(InStr(mstrPrivs, ";���в���;") = 0, strServiceRange, "����", True)
    If Not mrsDept.EOF Then
        For i = 1 To mrsDept.RecordCount
            cboDept.AddItem mrsDept!���� & "-" & mrsDept!����
            cboDept.ItemData(cboDept.NewIndex) = mrsDept!ID
            
            'ȷ��ȱʡ�Ĳ���
            If mvNurseFilter.Nurse Then
                If mrsDept!ID = mvNurseFilter.����ID Then cboDept.ListIndex = cboDept.NewIndex
            Else
                If UserInfo.����ID = mrsDept!ID Then cboDept.ListIndex = cboDept.NewIndex
            End If
            
            mrsDept.MoveNext
        Next
        If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    ElseIf InStr(";" & mstrPrivs, ";���в���;") > 0 Then
        MsgBox "û�з���סԺ������Ϣ,���ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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
    
    strHead = "סԺ��,1,750|����,1,500|����,1,700|�ѱ�,1,750|��������,1,1000|������,1,700|���,1,650|����,1,1600" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|���,1,1000|��λ,4,500|����,7,850|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ͳ����,7,850|ִ�п���,1,850|����,1,850|˵��,1,1000|��¼״̬,1,0"
    
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
        '���˺�:27990 2010-02-22 17:29:47
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 320
        .ColWidth(.Cols - 1) = 0
        
        'סԺ��,����,����,�ѱ�,��������,������
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        
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
        If frmBillingFilter.mblnDateMoved And tbs.SelectedItem.Key = "Auditing" Then
            '���ʻ��۵�������Ƿ��ں󱸱���,��Ϊ����ת�����󱸱�
            mblnNOMoved = zlDatabase.NOMoved("סԺ���ü�¼", strNO, , CStr(bytFlag), Me.Caption)
        Else
            mblnNOMoved = False   '����Ҫ����һ��
        End If
        
        strSql = _
        " Select A.��ʶ�� as סԺ��,A.����,A.����,A.�ѱ�,F.���� as ��������,A.������," & _
        "       C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
                IIf(gblnסԺ��λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X.סԺ��λ)", "A.���㵥λ") & " as ��λ," & _
        "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                IIf(gblnסԺ��λ, "/Nvl(X.סԺ��װ,1)", "") & ",'9999990.00000') as ����, " & _
        "       To_Char(Sum(A.��׼����)" & IIf(gblnסԺ��λ, "*Nvl(X.סԺ��װ,1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
        "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ͳ����),'9999999" & gstrDec & "') as ͳ����, " & _
        "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
        "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��, A.��¼״̬" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("סԺ���ü�¼"), "סԺ���ü�¼ A") & " ," & _
        "       �շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,���ű� F,ҩƷ��� X" & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
        " Where A.�շ�ϸĿID=B.ID And A.�շ����=C.����" & _
        "       And A.��������ID=F.ID(+) And A.ִ�в���ID=D.ID(+)" & _
        "       And A.NO=[1] And A.��¼����=[2] And A.�����־=2" & _
        "       And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼״̬" & IIf(blnDel, "=2", " IN(0,1,3)") & IIf(strTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),A.��ʶ��,A.����,A.����,A.�ѱ�,F.����,A.������," & _
        "       C.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & " B.���,A.���㵥λ,D.����,Nvl(A.��������,B.��������)," & _
        "       A.ִ��״̬,A.��¼״̬,X.ҩƷID,X.סԺ��λ,Nvl(X.סԺ��װ,1)" & _
        " Order by " & IIf(blnBat, "LPAD(A.����,10,' '),A.��ʶ��,Nvl(A.�۸񸸺�,A.���)", "Nvl(A.�۸񸸺�,A.���),A.��ʶ��,LPAD(A.����,10,' ')")
        
        If strTime <> "" Then
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag, CDate(strTime))
        Else
            Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO, bytFlag)
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
        'סԺ��,����,����,�ѱ�,��������,������
        mshDetail.ColWidth(0) = 850
        mshDetail.ColWidth(1) = 800
        mshDetail.ColWidth(2) = 700
        mshDetail.ColWidth(3) = 500
        mshDetail.ColWidth(4) = 1000
        mshDetail.ColWidth(5) = 700
        If InStr(mstrPrivsOpt, ";ҽ����ѯ;") = 0 Then mshDetail.ColWidth(4) = 0
    End If
    
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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

