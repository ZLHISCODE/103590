VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManagePatient 
   Caption         =   "������Ϣ����"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmManagePatient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboNodeList 
      Height          =   300
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   870
      Width           =   2100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   5325
      Left            =   2880
      TabIndex        =   6
      Top             =   1065
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   9393
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorSel    =   12632256
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePatient.frx":06EA
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip TabPatiState 
      Height          =   5655
      Left            =   2865
      TabIndex        =   5
      Top             =   750
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���в���"
            Key             =   "T_���в���"
            Object.Tag             =   "���в���"
            Object.ToolTipText     =   "���в���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ժ����"
            Key             =   "T_��Ժ����"
            Object.Tag             =   "��Ժ����"
            Object.ToolTipText     =   "��Ժ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ժ����"
            Key             =   "T_��Ժ����"
            Object.Tag             =   "��Ժ����"
            Object.ToolTipText     =   "��Ժ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ﲡ��"
            Key             =   "T_���ﲡ��"
            Object.Tag             =   "���ﲡ��"
            Object.ToolTipText     =   "���ﲡ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���۲���"
            Key             =   "T_���۲���"
            Object.Tag             =   "���۲���"
            Object.ToolTipText     =   "���۲���"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6390
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManagePatient.frx":0A04
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "PatiColor"
            Object.Tag             =   "PatiColor"
            Object.ToolTipText     =   "��������˵��"
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
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   2730
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5595
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   720
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvwDist_s 
      Height          =   5175
      Left            =   -15
      TabIndex        =   3
      Top             =   1230
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   9128
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9975
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9855
         _ExtentX        =   17383
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
            NumButtons      =   21
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
               Caption         =   "�Ǽ�"
               Key             =   "Add"
               Description     =   "�Ǽ�"
               Object.ToolTipText     =   "�Ǽ��²�����Ϣ"
               Object.Tag             =   "�Ǽ�"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ĵ�ǰѡ�в�����Ϣ"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰѡ�в�����Ϣ"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ϲ�"
               Key             =   "Merge"
               Description     =   "�ϲ�"
               Object.ToolTipText     =   "����ǰѡ���˵���Ϣ�ϲ�������һ��������"
               Object.Tag             =   "�ϲ�"
               ImageKey        =   "Merge"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Merge_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ƭ"
               Key             =   "View"
               Description     =   "��Ƭ"
               Object.ToolTipText     =   "�Կ�Ƭ��ʽ���ĵ�ǰ������Ϣ"
               Object.Tag             =   "��Ƭ"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "�ڵ�ǰ�����嵥�й������������Ĳ���"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������Ĳ�����"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�շ�����"
               Object.Tag             =   "����"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Family"
               Description     =   "����"
               Object.ToolTipText     =   "�����Ǽ�"
               Object.Tag             =   "����"
               ImageKey        =   "Family"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FamilyAdd"
                     Text            =   "�����Ǽ�"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "FamilyView"
                     Text            =   "������Ϣ"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��չ"
               Key             =   "PlugIn"
               Object.ToolTipText     =   "��չ����"
               Object.Tag             =   "��չ"
               ImageKey        =   "PlugIn"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "-"
               Key             =   "FamilySplit"
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   240
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
            Picture         =   "frmManagePatient.frx":1296
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":14B0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":16CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":18E4
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AFE
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1D18
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":2412
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":2B0C
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3206
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3420
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":363A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3854
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":3A6E
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":D405
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":13C67
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1260
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1A4C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   120
      Top             =   240
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
            Picture         =   "frmManagePatient.frx":1A623
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1A83D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AA57
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AC71
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1AE8B
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1B0A5
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1B79F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1BE99
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1C593
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1C7AD
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1C9C7
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1CBE1
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1CDFB
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":1D4F5
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePatient.frx":23D57
            Key             =   "PlugIn"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNode 
      AutoSize        =   -1  'True
      Caption         =   "վ��"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_Preview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintMed 
         Caption         =   "��ӡ����(&M)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "�շ�����(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInsure 
         Caption         =   "�������(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "�Ǽ�(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPatiInfo 
         Caption         =   "������Ϣ����(&J)"
      End
      Begin VB.Menu mnuEdit_Split1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelCard 
         Caption         =   "ȡ�����Ű�(&C)"
      End
      Begin VB.Menu mnuEditBlackList 
         Caption         =   "���ⲡ��(&T)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit_ToInPati 
         Caption         =   "תΪסԺ����(&I)"
      End
      Begin VB.Menu mnuEdit_Merge 
         Caption         =   "���˺ϲ�(&G)"
      End
      Begin VB.Menu mnuEdit_Surety 
         Caption         =   "������Ϣ(&B)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuEdit_Merge_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Stop 
         Caption         =   "ͣ�ò���(&S)"
      End
      Begin VB.Menu mnuEdit_Restore 
         Caption         =   "ȡ��ͣ��(&R)"
      End
      Begin VB.Menu mnuEdit_Restore_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_QueryPass 
         Caption         =   "���ò�ѯ����(&P)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "��ݿ�Ƭ(&V)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzReCalc 
         Caption         =   "���ѱ������������(&F)"
      End
      Begin VB.Menu mnuEdit_Family 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_FamilyAdd 
         Caption         =   "�����Ǽ�"
      End
      Begin VB.Menu mnuEdit_FamilyView 
         Caption         =   "������Ϣ"
      End
      Begin VB.Menu mnuEdit_PlugIn 
         Caption         =   "��չ(&E)"
         Begin VB.Menu mnuEdit_PlugItem 
            Caption         =   "����"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "��ѯ(&Q)"
      Begin VB.Menu mnuQuery_ChangeLog 
         Caption         =   "������Ϣ�䶯��־(&C)"
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
         Begin VB.Menu mnuViewToolDist 
            Caption         =   "���˷ֲ�(&D)"
            Checked         =   -1  'True
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStop 
         Caption         =   "��ʾͣ�ò���(&P)"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "��ʾ���˷�ʽ(&M)"
         Begin VB.Menu mnuViewByDept 
            Caption         =   "��������ʾ(&U)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewByDept 
            Caption         =   "��������ʾ(&D)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuView_3 
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
Attribute VB_Name = "frmManagePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsPati As ADODB.Recordset
Private mblnMax As Boolean, mblnUnLoad As Boolean
Private mblnDown As Boolean, mblnGo As Boolean
Private mstrFilter As String, mstrFilterInfo As String, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mlngCardType As Long  'ȱʡҽ�ƿ����
Private mbln�Ƿ�ȡ���� As Boolean '�����Ƿ����ִ��ȡ�����Ű󶨲�����ֻ�е��������ſ���ȡ���󶨿�������
Private mblnInitGrid As Boolean '�Ƿ���ɱ���ʼ��

Private mstrUserUnitIDs As String

Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    �Ǽ�ʱ��B As Date
    �Ǽ�ʱ��E As Date
    ����ʱ��B As Date
    ����ʱ��E As Date
    ��Ժʱ��B As Date
    ��Ժʱ��E As Date
    ��Ժʱ��B As Date
    ��Ժʱ��E As Date
    סԺ�� As String
    �Ա� As String
    �ѱ� As String
    ���� As String
    ҽ�Ƹ��ʽ As String
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��

Private Sub cboNodeList_Click()
    Call InitUnits
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub mnuEdit_FamilyAdd_Click()
'����:���˼�������
    If Not CreatePublicPatient Then Exit Sub
    Call gobjPublicPatient.MakePatiFamily(Me, 0, 2, mlngModul) '�༭
End Sub

Private Sub mnuEdit_FamilyView_Click()
    Dim lng����ID As Long
    
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID"))) Then
            MsgBox "û�пͻ���Ϣ���Բ鿴������Ϣ��", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
            MsgBox "û�в�����Ϣ���Բ鿴������Ϣ��", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    
    If Not CreatePublicPatient Then Exit Sub
    Call gobjPublicPatient.MakePatiFamily(Me, lng����ID, 1, mlngModul) '�鿴
End Sub

Private Sub mnuEdit_PlugItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuEdit_PlugItem(Index).Tag)
End Sub

Private Sub mnuEdit_QueryPass_Click()
    Dim strFirstPassWord As String, strSecPassWord As String, rsTemp As ADODB.Recordset, lng����ID As Long
    Dim strPassWord As String
    Dim strPassInput As String
    
    On Error GoTo errH
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If

    If InStr(mstrPrivs, "ǿ�Ƹ��Ĳ�ѯ����") <= 0 Then '��"ǿ�Ƹ��Ĳ�ѯ����"Ȩ������У�������
        If frmInput.InputVal(Me, "ԭ��ѯ����", "������ԭ��ѯ����,�����ԭ����ֱ��ȷ����", strFirstPassWord, 3, 10, True, False, False, "*") Then
                strPassInput = zlCommFun.zlStringEncode(strFirstPassWord)
                
                'У��ԭ����
                gstrSQL = "select ��ѯ���� from ������Ϣ where ����ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", lng����ID)
                If Not rsTemp.EOF Then
                    If strPassInput <> Nvl(rsTemp!��ѯ����) Then
                        MsgBox "ԭ��ѯ����������󣬽�ֹ�޸ģ����飡", vbExclamation, gstrSysName: Exit Sub
                    End If
                End If
        Else
            Exit Sub
        End If
    End If
    
    '����������
    strFirstPassWord = "": strSecPassWord = ""
    If frmInput.InputVal(Me, "�²�ѯ����", "�������ѯ�����룬���볤��0��10λ��" & vbCrLf & "����������ͨ��������������ѯ�����з��ò�ѯ��", strFirstPassWord, 3, 10, True, False, False, "*") Then
        strPassWord = zlCommFun.zlStringEncode(strFirstPassWord)
        '�ٴ�ȷ��������
        If frmInput.InputVal(Me, "ȷ��������", "���ٴ�����������,��ȷ��������" & vbCrLf & "����������ͨ��������������ѯ�����з��ò�ѯ��", strSecPassWord, 3, 10, True, False, False, "*") Then
            strPassInput = zlCommFun.zlStringEncode(strSecPassWord)
            If strPassWord <> strPassInput Then
                '�������벻һ��
                MsgBox "ǰ����������������벻һ�����˴���������δ��Ч�����飡", vbExclamation, gstrSysName: Exit Sub
            Else
                gstrSQL = "ZL_������Ϣ_UpdatePass(" & lng����ID & ",'" & strPassWord & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "")
                MsgBox "�����޸ĳɹ���", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Restore_Click()
    Dim intRow As Long, lng����ID As Long
    Dim strSQL As String, i As Long
    Dim blnTrans As Boolean
    
    intRow = mshPati.Row
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(intRow, GetColNum("�ͻ�ID")))
        If lng����ID = 0 Then
            MsgBox "û�пͻ���Ϣ����ȡ��ͣ�ã�", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        lng����ID = Val(mshPati.TextMatrix(intRow, GetColNum("����ID")))
        If lng����ID = 0 Then
            MsgBox "û�в�����Ϣ����ȡ��ͣ�ã�", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    If MsgBox("ȷʵҪȡ��ͣ��""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    strSQL = "zl_������Ϣ_Restore(" & lng����ID & ")"
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '��ֱ�Ӵ���
    mshPati.TextMatrix(intRow, GetColNum("ͣ��ʱ��")) = ""
    mshPati.Redraw = False
    For i = 0 To mshPati.Cols - 1
        mshPati.Col = i
        mshPati.CellForeColor = Me.ForeColor
    Next
    mshPati.Redraw = True
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    Call mshPati_EnterCell
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Stop_Click()
    Dim intRow As Long, lng����ID As Long, int������� As Integer
    Dim strSQL As String, i As Long
    Dim blnTrans As Boolean
    
    intRow = mshPati.Row
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(intRow, GetColNum("�ͻ�ID")))
        If lng����ID = 0 Then
            MsgBox "û�пͻ���Ϣ����ͣ�ã�", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        lng����ID = Val(mshPati.TextMatrix(intRow, GetColNum("����ID")))
        If lng����ID = 0 Then
            MsgBox "û�в�����Ϣ����ͣ�ã�", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    int������� = GetColNum("�������")
    If int������� <> -1 Then
        int������� = Val(mshPati.TextMatrix(intRow, int�������))
        If int������� > 0 Then
            MsgBox """" & mshPati.TextMatrix(intRow, GetColNum("����")) & """�Ѿ���Ժ����� " & int������� & " �Σ�������ͣ�á�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("ȷʵҪͣ��""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    strSQL = "zl_������Ϣ_Stop(" & lng����ID & ")"
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '��ֱ�Ӵ���
    mshPati.TextMatrix(intRow, GetColNum("ͣ��ʱ��")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If Not mnuViewStop.Checked Then
        If mshPati.Rows > 2 Then
            mshPati.RemoveItem intRow
        Else
            With mshPati
                For i = 0 To .Cols - 1
                    .TextMatrix(intRow, i) = ""
                Next
            End With
        End If
        
        If intRow <= mshPati.Rows - 1 Then
            mshPati.Row = intRow
        Else
            mshPati.Row = mshPati.Rows - 1
        End If
    Else
        mshPati.Redraw = False
        For i = 0 To mshPati.Cols - 1
            mshPati.Col = i
            mshPati.CellForeColor = &HC0&
        Next
        mshPati.Redraw = True
    End If
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    Call mshPati_EnterCell
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_ToInPati_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strסԺ�� As String, str���� As String
    Dim strSQL As String, strNote As String
    Dim rsTemp As New ADODB.Recordset
        
    lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) 'ҩ��ϵͳ������Ժ����
    lng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�������")))
    strסԺ�� = mshPati.TextMatrix(mshPati.Row, GetColNum("סԺ��"))
    
    If lng����ID = 0 Then
        MsgBox "û�в��˿���תΪסԺ���ˡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    strSQL = "Select Nvl(״̬,0) ״̬ From ������ҳ Where ����ID=[1] And ��ҳID=[2] And ��������=2"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    
    If rsTemp!״̬ = 1 Then
        MsgBox "���˵�ǰ��δ���,����תΪסԺ���ˡ����Ƚ�������ƺ����ԡ�", vbInformation, gstrSysName
        Exit Sub
    ElseIf rsTemp!״̬ = 2 Then
        MsgBox "���˵�ǰ����ת��,����תΪסԺ���ˡ����Ƚ�����ת�ƻ�ȡ��ת�ƺ����ԡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪ����סԺ���۲���תΪסԺ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    'û��סԺ�������һ��
    If strסԺ�� = "" Then
        strסԺ�� = zlDatabase.GetNextNo(2)
        strNote = "�����۲��� " & str���� & " תΪסԺ����֮ǰ������Ϊ�ò���ȷ��һ��סԺ�š�"
        If Not frmInput.InputVal(Me, "סԺ��", strNote, strסԺ��, 1, 10, False) Then Exit Sub
    End If
    
    
    strSQL = "ZL_���˱䶯��¼_תסԺ(" & lng����ID & "," & lng��ҳID & "," & strסԺ�� & ")"
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Call mnuViewReFlash_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditBlackList_Click()
    frmBlackList.mstrPrivs = mstrPrivs
    frmBlackList.Show 1, Me
End Sub

Private Sub mnuEditDelCard_Click()
    Dim strSQL As String, lng����ID As Long
    
    lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    
    If CheckBindCard(lng����ID) = False Then
        '���˺�:24537
        MsgBox "�ò��˵Ŀ��Ų��ǰ󶨿�,�뵽ҽ�ƿ����Ź����������˿�����!", vbInformation, gstrSysName
        Exit Sub
    Else
        If MsgBox("��ȷ��Ҫȡ����ǰ���˵Ŀ��Ű���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
     'Zl_ҽ�ƿ��䶯_Insert
       strSQL = "Zl_ҽ�ƿ��䶯_Insert("
      '      �䶯����_In   Number,
      '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
      strSQL = strSQL & "" & 14 & ","
      '      ����id_In     סԺ���ü�¼.����id%Type,
      strSQL = strSQL & "" & lng����ID & ","
      '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
      strSQL = strSQL & "" & mlngCardType & ","
      '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "NULL,"
      '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "'" & mshPati.TextMatrix(mshPati.Row, GetColNum("���￨��")) & "',"
      '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
      strSQL = strSQL & "'ȡ�����Ű�',"
      '      ����_In       ������Ϣ.����֤��%Type,
      strSQL = strSQL & "NULL,"
      '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
      strSQL = strSQL & "NULL,"
      '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
      'strSQL = strSQL & "to_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
      '      Ic����_In     ������Ϣ.Ic����%Type := Null,
      strSQL = strSQL & "NULL,"
      '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
      strSQL = strSQL & "NULL)"
    
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    mshPati.TextMatrix(mshPati.Row, GetColNum("���￨")) = ""
    mshPati.TextMatrix(mshPati.Row, GetColNum("���￨��")) = ""
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckBindCard(ByVal lng����ID As Long) As Boolean
'���ܣ���鲡���Ƿ��о��￨��¼
'�����:52133
    Dim rsTmp As ADODB.Recordset, strSQL As String
    'by lesfeng 2009-12-30 �����  ���˷��ü�¼ --��סԺ���ü�¼ ����ֻ��סԺ ��¼���� = 5���ھ��￨����,�����￨������סԺ���ü�¼��
    strSQL = "Select Count(*) As �Ƿ���� From ����ҽ�ƿ��䶯 Where ����ID=[1] And �����ID=[2] And ����=[3] And �䶯���=11"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, mlngCardType, Trim(mshPati.TextMatrix(mshPati.Row, GetColNum("���￨��"))))
    If rsTmp Is Nothing Then CheckBindCard = False: Exit Function
    If rsTmp.RecordCount = 0 Then CheckBindCard = False: Exit Function
    CheckBindCard = rsTmp!�Ƿ���� > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditMzReCalc_Click()
    Dim lng����ID As Long
    Dim str���� As String, strSQL As String
    
    '���ѱ�����������ʷ���
    '����:41034
    If mshPati.Row <= 0 Then Exit Sub
    lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    If lng����ID = 0 Then
        MsgBox "��ѡ����Ҫ��������Ĳ��ˣ�", vbExclamation, gstrSysName: Exit Sub
    End If
    str���� = mshPati.TextMatrix(mshPati.Row, GetColNum("����"))
    If MsgBox("��ȷ��Ҫ��[" & str���� & "]��δ���������ʷ��ð���ǰ�ѱ�������?" & vbCrLf & vbCrLf & _
        "�������������˵�ǰ�ѱ��Ӧ���Żݱ��ʶ�δ��������½��д��ۼ���!", vbInformation + vbYesNo + vbDefaultButton1, App.ProductName) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo errH
    strSQL = "Zl_����δ���������_Recalc(" & lng����ID & ")"
    zlDatabase.ExecuteProcedure strSQL, App.ProductName
    MsgBox "��������ɹ�!", vbOKOnly + vbInformation, gstrSysName
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditPatiInfo_Click()
    Dim lng����ID As Long, lng����ID As Long
    Dim strInfo As String
    Dim blnOK As Boolean
    
    If CreatePublicPatient = False Then Exit Sub
    '65802:������
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    If lng����ID <> 0 Then
        Select Case TabPatiState.SelectedItem.Key
            Case "T_��Ժ����", "T_��Ժ����", "T_���۲���"
                lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
            Case Else
                lng����ID = 0
        End Select
    Else
        lng����ID = 0
    End If
    '������Ϣ����
    blnOK = gobjPublicPatient.ModiPatiBaseInfo(Me, "������Ϣ����", lng����ID, lng����ID, 2)
    If blnOK = True And lng����ID <> 0 Then Call mnuViewReFlash_Click
End Sub

Private Sub mnuFileLocalSet_Click()
    Call frmLocalSet.zlSetPara(Me, mstrPrivs, mlngModul)
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFilePrintMed_Click()
    Dim lng����ID As Long
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    If lng����ID = 0 Then Exit Sub
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1101", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1101", Me, "����ID=" & lng����ID, 2)
    End If
End Sub

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuQuery_ChangeLog_Click()
    Dim lng����ID As Long
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    
    Call frmPatiInfoChangeLog.ShowMe(Me, mstrPrivs, lng����ID)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim str����ID As String
    
    str����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    If str����ID <> "" Then
        With mshPati
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����ID=" & str����ID)
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    Call InitUnits
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub mnuViewFilter_Click()
    frmPatiFilter.mbytType = Val(mshPati.Tag)
    frmPatiFilter.Show 1, Me
    If gblnOK Then
        With frmPatiFilter
            mstrFilter = .mstrFilter
            mstrFilterInfo = .mstrFilterInfo
            SQLCondition.�Ǽ�ʱ��B = .dtp�Ǽ�B
            SQLCondition.�Ǽ�ʱ��E = .dtp�Ǽ�E
            SQLCondition.����ʱ��B = .dtp����B
            SQLCondition.����ʱ��E = .dtp����E
            
            SQLCondition.��Ժʱ��B = .dtp��ԺB
            SQLCondition.��Ժʱ��E = .dtp��ԺE
            SQLCondition.��Ժʱ��B = .dtp��ԺB
            SQLCondition.��Ժʱ��E = .dtp��ԺE
            
            SQLCondition.סԺ�� = Trim(.txtסԺ��.Text)
            SQLCondition.�Ա� = zlCommFun.GetNeedName(.cbo�Ա�.Text)
            SQLCondition.�ѱ� = zlCommFun.GetNeedName(.cbo�ѱ�.Text)
            SQLCondition.���� = zlCommFun.GetNeedName(.txt����.Text)
            SQLCondition.ҽ�Ƹ��ʽ = zlCommFun.GetNeedName(.cboPayPlan.Text)
            
            '59340:������,2013-04-23,����ƥ�����gstrLike
            If .PatiIdentify.GetCurCard.���� = "����" And .mlngPatiId = 0 And (.chk�Ǽ�.Value = 1 Or .chk��Ժ.Value = 1 Or .chk��Ժ.Value = 1) Then     '����
                SQLCondition.Patient = gstrLike & Trim(.PatiIdentify.Text) & "%"
            Else
                SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
            End If
        End With
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmPatiFind.mbytType = Val(mshPati.Tag)
    frmPatiFind.Show 1, Me
    If gblnOK Then Call SeekPati(frmPatiFind.optHead)
End Sub

Private Sub mnuViewStop_Click()
    mnuViewStop.Checked = Not mnuViewStop.Checked
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub mnuEdit_Surety_Click()
    Dim lng����ID As Long, lngRow As Long
    Dim bln��Ժ���� As Boolean
    
    lngRow = mshPati.Row
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(lngRow, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(lngRow, GetColNum("����ID")))
    End If
    
    If GetColNum("����") <> -1 Then
        bln��Ժ���� = Trim(mshPati.TextMatrix(lngRow, GetColNum("����"))) <> ""
    End If
    
    If lng����ID <> 0 Then
        frmSurety.mlng����ID = lng����ID
        frmSurety.mbln��Ժ���� = bln��Ժ����
        frmSurety.mstrPrivs = mstrPrivs
        frmSurety.Show 1, Me
    End If
End Sub

Private Sub mnuViewToolDist_Click()
    mnuViewToolDist.Checked = Not mnuViewToolDist.Checked
    tvwDist_s.Visible = mnuViewToolDist.Checked
    pic.Visible = tvwDist_s.Visible
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub mshPati_DblClick()
    If glngSys Like "8??" Then
        If mshPati.MouseRow = 0 Or mshPati.TextMatrix(mshPati.MouseRow, GetColNum("�ͻ�ID")) = "" Then Exit Sub
    Else
        If mshPati.MouseRow = 0 Or mshPati.TextMatrix(mshPati.MouseRow, GetColNum("����ID")) = "" Then Exit Sub
    End If
    mnuEdit_View_Click
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    ElseIf Button = 1 Then
        mblnDown = True
    End If
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then
        Unload Me
    Else
        Call InitLocPar(mlngModul)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekPati(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, Curdate As Date, lngTmp As Long
    Dim blnHavePrivs As Boolean
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mblnInitGrid = False
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    lngTmp = Val(zlDatabase.GetPara("��ʾ���˷�ʽ", glngSys, mlngModul, 0))
    For i = 0 To mnuViewByDept.UBound
        mnuViewByDept(i).Checked = (i = lngTmp)
    Next
   
    '��ʼ��վ���б�
    
    
    '�ָ����Բ����嵥����
    mshPati.Tag = zlDatabase.GetPara("��������", glngSys, mlngModul, 1) 'mshPati.Tag�б�������ʵ�������0-����,1-��Ժ,2-��Ժ,3-����,4-����
    
    TabPatiState.Tabs(Val(mshPati.Tag) + 1).Selected = True
    TabPatiState.Tag = TabPatiState.SelectedItem.Key
    mnuViewToolDist.Enabled = TabPatiState.SelectedItem.Key = "T_��Ժ����" Or TabPatiState.SelectedItem.Key = "T_��Ժ����"
    If mnuViewToolDist.Enabled Then InitUnits
    Call InitFace
    
    If glngSys Like "8??" Then
        Me.Caption = "�ͻ���Ϣ����"
        mnuEditBlackList.Visible = False
        mnuEdit_Merge.Caption = "�ͻ��ϲ�(&G)"
        mnuEdit_Stop.Caption = "ͣ�ÿͻ�(&S)"
        For i = 1 To tbr.Buttons.Count
            tbr.Buttons(i).ToolTipText = Replace(tbr.Buttons(i).ToolTipText, "����", "�ͻ�")
        Next
        
        mshPati.Tag = 3
        TabPatiState.Tabs.Remove 5
        TabPatiState.Tabs.Remove 3
        TabPatiState.Tabs.Remove 2
        TabPatiState.Tabs.Remove 1
    End If
    
    RestoreWinState Me, App.ProductName
    
    mblnUnLoad = False
    
    'Ȩ������
    If InStr(mstrPrivs, ";�޸�;") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    
    If InStr(mstrPrivs, ";����;") = 0 Then
         mnuEdit_Add.Visible = False
         tbr.Buttons("Add").Visible = False
    End If
    
    If InStr(mstrPrivs, ";ɾ��;") = 0 Then
        mnuEdit_Del.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    
    If InStr(mstrPrivs, ";��ͣ;") = 0 Then
        mnuEdit_Stop.Visible = False
        mnuEdit_Restore.Visible = False
        mnuEdit_Restore_.Visible = False
    End If
    
    If InStr(mstrPrivs, ";����;") = 0 And InStr(mstrPrivs, ";�޸�;") = 0 And InStr(mstrPrivs, ";ɾ��;") = 0 Then
        mnuEdit_.Visible = False
        tbr.Buttons("Edit_").Visible = False
    End If
    
    If Not (InStr(mstrPrivs, ";����;") = 0 And InStr(mstrPrivs, ";�޸�;") = 0 And InStr(mstrPrivs, ";ɾ��;") = 0 And InStr(mstrPrivs, ";��ͣ;") = 0) Then
        If gstr�ſ�ID <> "" Then
            Call UpdateShareID(mlngModul, gstr�ſ�ID, 5)
        End If
    End If
    
    If InStr(mstrPrivs, "��ݺϲ�") = 0 Then
        mnuEdit_Merge.Visible = False
        mnuEdit_Merge_.Visible = False
        tbr.Buttons("Merge").Visible = False
        tbr.Buttons("Merge_").Visible = False
    End If
    
    If InStr(mstrPrivs, "סԺ����תסԺ") = 0 Then
        mnuEdit_ToInPati.Visible = False
    End If
    If InStr(mstrPrivs, "ȡ�����Ű�") = 0 Then
        mnuEditDelCard.Visible = False
    End If
    '�շ����ʹ���
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";����;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("����").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    
    '����:41034
    mnuEditMzReCalc.Visible = InStr(1, mstrPrivs, ";�����������;") > 0
    mnuEditSplit.Visible = InStr(1, mstrPrivs, ";�����������;") > 0
    
    mstrUserUnitIDs = GetUserUnits '���ڿ�����������+�������ڲ���
    '����ʱȱʡ����ʾ����
    Call SetHeader(gblnMyStyle)
    Call mshPati_EnterCell
    
    '��Ժ���Ժ���˲���ʾ���ҷֲ�
    If TabPatiState.SelectedItem.Key = "T_��Ժ����" Or TabPatiState.SelectedItem.Key = "T_��Ժ����" Then
        tvwDist_s.Visible = mnuViewToolDist.Enabled
    Else
        tvwDist_s.Visible = False
    End If
    
    With frmPatiFilter
        .mbytType = Val(mshPati.Tag)
        Call .MakeFilter
        
        mstrFilter = .mstrFilter
        mstrFilterInfo = .mstrFilterInfo
        SQLCondition.�Ǽ�ʱ��B = .dtp�Ǽ�B
        SQLCondition.�Ǽ�ʱ��E = .dtp�Ǽ�E
        SQLCondition.����ʱ��B = .dtp����B
        SQLCondition.����ʱ��E = .dtp����E
        
        SQLCondition.��Ժʱ��B = .dtp��ԺB
        SQLCondition.��Ժʱ��E = .dtp��ԺE
        SQLCondition.��Ժʱ��B = .dtp��ԺB
        SQLCondition.��Ժʱ��E = .dtp��ԺE
        
        SQLCondition.סԺ�� = Trim(.txtסԺ��.Text)
        SQLCondition.�Ա� = zlCommFun.GetNeedName(.cbo�Ա�.Text)
        SQLCondition.�ѱ� = zlCommFun.GetNeedName(.cbo�ѱ�.Text)
        SQLCondition.���� = zlCommFun.GetNeedName(.txt����.Text)
        '59340:������,2013-04-23,����ƥ�����gstrLike
        If .PatiIdentify.GetCurCard.���� = "����" And .mlngPatiId = 0 And (.chk�Ǽ�.Value = 1 Or .chk��Ժ.Value = 1 Or .chk��Ժ.Value = 1) Then     '����
            SQLCondition.Patient = gstrLike & Trim(.PatiIdentify.Text) & "%"
        Else
            SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
        End If
    End With
    
    '��ʼ��������Ϣ��������
    Call CreatePublicPatient
    '��չ����
    Call LoadPlugInMnu

End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    Dim DisW As Long '���˷ֲ�����
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshPati.MousePointer = 0
    
    mshPati.Redraw = False
    
    If mblnMax Then
        tvwDist_s.Width = 2500
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    DisW = IIf(tvwDist_s.Visible, tvwDist_s.Width + pic.Width, 0)
    
    pic.Visible = tvwDist_s.Visible
    
    With tvwDist_s
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop + cbrH + IIf(cboNodeList.Visible, cboNodeList.Height + 100, 0)
        .Height = Me.ScaleHeight - staH - cbrH - IIf(cboNodeList.Visible, cboNodeList.Height + 100, 0)
    End With
    With pic
        .Left = tvwDist_s.Left + tvwDist_s.Width
        '.Top = tvwDist_s.Top
        .Top = IIf(cboNodeList.Visible, Me.ScaleTop + cbrH, tvwDist_s.Top)
        '.Height = tvwDist_s.Height
        .Height = IIf(cboNodeList.Visible, Me.ScaleHeight - staH - cbrH, tvwDist_s.Height)
    End With
    
    With TabPatiState
        .Left = DisW
        '.Top = tvwDist_s.Top
        .Top = IIf(cboNodeList.Visible, Me.ScaleTop + cbrH, tvwDist_s.Top)
        .Width = Me.ScaleWidth - DisW
        '.Height = tvwDist_s.Height
        .Height = IIf(cboNodeList.Visible, Me.ScaleHeight - staH - cbrH, tvwDist_s.Height)
    End With
    With mshPati
        .Left = TabPatiState.ClientLeft
        .Top = TabPatiState.ClientTop
        .Height = TabPatiState.ClientHeight
        .Width = TabPatiState.ClientWidth
    End With
    cboNodeList.Width = tvwDist_s.Width - 600
    mshPati.Redraw = True
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTmp As Long, i As Long
    
    If Not gobjPublicPatient Is Nothing Then
        Set gobjPublicPatient = Nothing
    End If
    
    mstrFilter = ""
    mstrFilterInfo = ""
    zlDatabase.SetPara "��������", Val(mshPati.Tag), glngSys, mlngModul
    SaveWinState Me, App.ProductName
    
    '��ʾ���˷�ʽ
    lngTmp = 0
    For i = 0 To mnuViewByDept.UBound
        If mnuViewByDept(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "��ʾ���˷�ʽ", lngTmp, glngSys, mlngModul
    
    Unload frmPatiFind
    Unload frmPatiFilter
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strSQL As String, intRow As Long, i As Long
    Dim strSQL1 As String
    Dim blnTrans As Boolean
    
    intRow = mshPati.Row
    
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("�ͻ�ID"))) Then
            MsgBox "û�пͻ���Ϣ����ɾ����", vbExclamation, gstrSysName: Exit Sub
        End If
        If MsgBox("�ò�����ɾ���Ϳͻ�""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """��ص�������Ϣ�����Ҳ��ɻָ���Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        strSQL1 = "Zl_������Ƭ_Delete(" & mshPati.TextMatrix(intRow, GetColNum("�ͻ�ID")) & ")"
        strSQL = "zl_������Ϣ_DELETE(" & mshPati.TextMatrix(intRow, GetColNum("�ͻ�ID")) & ")"
    Else
        If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("����ID"))) Then
            MsgBox "û�в�����Ϣ����ɾ����", vbExclamation, gstrSysName: Exit Sub
        End If
        If MsgBox("�ò�����ɾ���Ͳ���""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """��ص�������Ϣ�����Ҳ��ɻָ���Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        strSQL1 = "Zl_������Ƭ_Delete(" & mshPati.TextMatrix(intRow, GetColNum("����ID")) & ")"
        strSQL = "zl_������Ϣ_DELETE(" & mshPati.TextMatrix(intRow, GetColNum("����ID")) & ")"
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
'    Call SQLTest(App.ProductName, Me.Caption, strSQL)
'    gcnOracle.Execute strSQL, , adCmdStoredProc
'    Call SQLTest
    zlDatabase.ExecuteProcedure strSQL1, Me.Caption
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '��ֱ�Ӵ���
    If mshPati.Rows > 2 Then
        mshPati.RemoveItem intRow
    Else
        With mshPati
            For i = 0 To .Cols - 1
                .TextMatrix(intRow, i) = ""
            Next
        End With
    End If
    
    If intRow <= mshPati.Rows - 1 Then
        mshPati.Row = intRow
    Else
        mshPati.Row = mshPati.Rows - 1
    End If
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    Call mshPati_EnterCell
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Merge_Click()
    Dim lng����ID As Long
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
        If lng����ID = 0 Then
            MsgBox "û�пͻ���Ϣ�ɹ��ϲ���", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
        If lng����ID = 0 Then
            MsgBox "û�в�����Ϣ�ɹ��ϲ���", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    If ExistFeeInsurePatient(lng����ID) Then
        MsgBox "��ҽ�����˴���δ�����,���Ƚ�����ٺϲ���", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    
    frmMergePatient.mstrPrivs = mstrPrivs
    frmMergePatient.mlng����ID = lng����ID
    frmMergePatient.Show 1, Me
    
    If gblnOK Then
        If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID"))) Then
            MsgBox "û�пͻ���Ϣ�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
        End If
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
            MsgBox "û�в�����Ϣ�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
        End If
    End If
    
    On Error Resume Next
    Err.Clear
    
    If glngSys Like "8??" Then
        frmPatient.mlng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        frmPatient.mlng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    frmPatient.mlngModul = mlngModul
    frmPatient.mstrPrivs = mstrPrivs
    frmPatient.mbytInState = 1
    frmPatient.mbytView = Val(mshPati.Tag)
    frmPatient.Show 1, Me
    If gblnOK Then
        If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Err.Clear
    
    frmPatient.mlngModul = mlngModul
    frmPatient.mstrPrivs = mstrPrivs
    frmPatient.mbytInState = 0
    frmPatient.mbytView = Val(mshPati.Tag)
    frmPatient.Show 1, Me
    If gblnOK Then
        If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuEdit_View_Click()
    Dim lng����ID  As Long
    Dim lng��ҳID  As Long
    
    On Error Resume Next
    If glngSys Like "8??" Then
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID"))) Then
            MsgBox "û�пͻ���Ϣ���Բ鿴��", vbExclamation, gstrSysName: Exit Sub
        End If
        frmPatient.mlngModul = mlngModul
        frmPatient.mstrPrivs = mstrPrivs
        frmPatient.mbytInState = 2
        frmPatient.mbytView = Val(mshPati.Tag)
        frmPatient.mlng����ID = CLng(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
        frmPatient.Show 1, Me
        mshPati.Refresh
    Else
        If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
            MsgBox "û�в�����Ϣ���Բ鿴��", vbExclamation, gstrSysName: Exit Sub
        End If
        lng����ID = CLng(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
        lng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID"))) 'CLNG���մ� ���� 13-���Ͳ�ƥ��
        If CreatePublicPatient Then
            Call gobjPublicPatient.ReadPatiDegreeCard(Me, lng����ID, lng��ҳID)
        End If
        
        mshPati.Refresh
    End If
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pic.Left + X < 1000 Or TabPatiState.Width - X < 2000 Then Exit Sub
        pic.Left = pic.Left + X
        tvwDist_s.Width = tvwDist_s.Width + X
        TabPatiState.Left = TabPatiState.Left + X
        TabPatiState.Width = TabPatiState.Width - X
        mshPati.Left = TabPatiState.ClientLeft
        mshPati.Width = TabPatiState.ClientWidth
        cboNodeList.Width = tvwDist_s.Width - 600
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PatiColor" Then
        zlDatabase.ShowPatiColorTip Me
    End If
End Sub

Private Sub TabPatiState_Click()
    If TabPatiState.Tag = TabPatiState.SelectedItem.Key Then Exit Sub
    '35632:������,2013-07-29
    If mblnInitGrid = True Then SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    mshPati.Tag = TabPatiState.SelectedItem.Index - 1 '�洢�������0-����,1-��Ժ,2-��Ժ,3-����,4-����
    cboNodeList.Visible = TabPatiState.SelectedItem.Index <> 1 And TabPatiState.SelectedItem.Index <> 4 And TabPatiState.SelectedItem.Index <> 5 And cboNodeList.ListCount > 0
    mnuViewToolDist.Enabled = TabPatiState.SelectedItem.Key = "T_��Ժ����" Or TabPatiState.SelectedItem.Key = "T_��Ժ����"
    mnuViewPatiMode.Enabled = mnuViewToolDist.Enabled
    If mnuViewToolDist.Enabled Then InitUnits
    
    Unload frmPatiFilter
    Unload frmPatiFind
    If TabPatiState.Tag <> "" Then '�������ʱ��������
        With frmPatiFilter
            .mbytType = Val(mshPati.Tag)
            Call .MakeFilter
        
        '�л���������ʱ�����ָ�Ϊ��(ʹ��ȱʡ����)
            mstrFilter = .mstrFilter
            mstrFilterInfo = .mstrFilterInfo
            SQLCondition.�Ǽ�ʱ��B = .dtp�Ǽ�B
            SQLCondition.�Ǽ�ʱ��E = .dtp�Ǽ�E
            SQLCondition.����ʱ��B = .dtp����B
            SQLCondition.����ʱ��E = .dtp����E
            
            SQLCondition.��Ժʱ��B = .dtp��ԺB
            SQLCondition.��Ժʱ��E = .dtp��ԺE
            SQLCondition.��Ժʱ��B = .dtp��ԺB
            SQLCondition.��Ժʱ��E = .dtp��ԺE
            
            SQLCondition.סԺ�� = Trim(.txtסԺ��.Text)
            SQLCondition.�Ա� = zlCommFun.GetNeedName(.cbo�Ա�.Text)
            SQLCondition.�ѱ� = zlCommFun.GetNeedName(.cbo�ѱ�.Text)
            SQLCondition.���� = zlCommFun.GetNeedName(.txt����.Text)
            '59340:������,2013-04-23,����ƥ�����gstrLike
            If .PatiIdentify.GetCurCard.���� = "����" And .mlngPatiId = 0 And (.chk�Ǽ�.Value = 1 Or .chk��Ժ.Value = 1 Or .chk��Ժ.Value = 1) Then      '����
                SQLCondition.Patient = gstrLike & Trim(.PatiIdentify.Text) & "%"
            Else
                SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
            End If
        End With
        mshPati.Clear: mshPati.Rows = 2
        Call SetHeader(gblnMyStyle)
        Select Case TabPatiState.SelectedItem.Key
            Case "T_���в���"  '���в���
                tvwDist_s.Visible = False
            Case "T_��Ժ����" '��Ժ����
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_��Ժ����"  '��Ժ����
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_���ﲡ��" '���ﲡ��
                tvwDist_s.Visible = False
            Case "T_���۲���"    '���۲���
                tvwDist_s.Visible = False
        End Select
        Call Form_Resize
        
        Dim blnAutoRefresh As Boolean
        '54701:������,2012-10-19
        blnAutoRefresh = (Val(zlDatabase.GetPara("�Զ�ˢ������", glngSys, mlngModul, 0)) = 1)
        If blnAutoRefresh = True Then
            Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo) 'ǿ�����ò��ָ��п�
        Else
            tvwDist_s.Tag = ""
        End If
    End If
    TabPatiState.Tag = TabPatiState.SelectedItem.Key
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Go"
            mnuViewGo_Click
        Case "Filter"
            mnuViewFilter_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Add"
            mnuEdit_Add_Click
        Case "Merge"
            mnuEdit_Merge_Click
        Case "View"
            mnuEdit_View_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "����"
            mnuFileRollingCurtain_Click
        Case "Family"
            mnuEdit_FamilyAdd_Click
        Case "PlugIn"
            PopupMenu mnuEdit_PlugIn, vbPopupMenuRightButton
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "FamilyAdd"
       mnuEdit_FamilyAdd_Click
    Case "FamilyView"
       mnuEdit_FamilyView_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub tvwDist_s_NodeClick(ByVal Node As MSComctlLib.Node)
    '��ͬ������ٴ���
    If tvwDist_s.Tag = Node.Key Then Exit Sub
    tvwDist_s.Tag = Node.Key
    
    Call ShowPatis(mstrFilter, , gblnMyStyle, mstrFilterInfo)
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
    
    intRow = mshPati.Row
    
    '��ͷ
    If glngSys Like "8??" Then
        objOut.Title.Text = "�ͻ��嵥"
    Else
        objOut.Title.Text = "�����嵥"
    End If
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    If Not glngSys Like "8??" Then
        Select Case TabPatiState.SelectedItem.Key
            Case "T_���в���"
                objRow.Add "���ࣺ���в���"
            Case "T_��Ժ����"
                objRow.Add "���ࣺ��Ժ����"
                objRow.Add "���ţ�" & tvwDist_s.SelectedItem.Text
            Case "T_��Ժ����"
                objRow.Add "���ࣺ��Ժ����"
            Case "T_���ﲡ��"
                objRow.Add "���ࣺ���ﲡ��"
            Case "T_���۲���"
                objRow.Add "���ࣺ���۲���"
        End Select
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshPati.Redraw = False
    Set objOut.Body = mshPati
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshPati.Row = intRow
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = True
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled Then mnuEdit_Del_Click
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub InitFace()
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node, i As Integer
    Dim strPreKey  As String, strSQL As String, strUnitIDs As String
    Dim blnLimitIn As Boolean, blnByDept As Boolean, blnLimitUnit As Boolean
    
    On Error GoTo errHandle
    
    blnLimitIn = TabPatiState.SelectedItem.Key = "T_��Ժ����"
    blnByDept = mnuViewByDept(1).Checked
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    
    '����վ��ѡ��
    strSQL = "SELECT DISTINCT a.վ��, c.����" & vbNewLine & _
            " FROM ���ű� a, ��������˵�� b, Zlnodelist c" & vbNewLine & _
            " WHERE a.Id = b.����id AND a.վ�� = c.��� AND (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') OR a.����ʱ�� IS NULL) AND" & vbNewLine & _
            "      b.�������� = [1] " & vbNewLine & _
            IIf(blnLimitIn, " And ID In (Select Distinct " & IIf(blnByDept, "����id", "����id") & " From ��λ״����¼ Where ����id Is Not Null)", "") & vbNewLine & _
            IIf(blnLimitUnit, " And A.ID In (" & strUnitIDs & ")", "") & _
            " ORDER BY a.վ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnByDept, "�ٴ�", "����"))
    cboNodeList.Clear
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            cboNodeList.AddItem rsTmp!վ�� & "-" & rsTmp!����
            cboNodeList.ItemData(rsTmp.AbsolutePosition - 1) = rsTmp!վ��
            rsTmp.MoveNext
        Wend
        Call cbo.Locate(cboNodeList, gstrNodeNo, True)
    Else
        lblNode.Visible = False
        cboNodeList.Visible = False
        Form_Resize
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ�����˲������ҷֲ��б�
'˵�����Բ���-���ҷֲ�,���в����������ڵ�ǰ��Ժ����֮�л��
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node, i As Integer
    Dim strPreKey  As String, strSQL As String, strUnitIDs As String
    Dim blnLimitIn As Boolean, blnByDept As Boolean, blnLimitUnit As Boolean
    Dim strNodeNo
    
    On Error GoTo errH
    
    blnLimitIn = TabPatiState.SelectedItem.Key = "T_��Ժ����"
    blnByDept = mnuViewByDept(1).Checked
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
             
    strPreKey = ""
    If Not tvwDist_s.SelectedItem Is Nothing Then strPreKey = tvwDist_s.SelectedItem.Key
    
    tvwDist_s.Nodes.Clear
    Set objNode = tvwDist_s.Nodes.Add(, , "Root", IIf(blnByDept, "���п���", "���в���"), 1)
    objNode.Expanded = True
    If objNode.Key = strPreKey Then objNode.Selected = True
    'by lesfeng 2010-03-08 �����Ż�
    strSQL = "Select A.ID, A.����, A.����" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where A.ID = B.����id And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
            "      B.�������� =[1] " & _
            IIf(blnLimitIn, " And ID In (Select Distinct " & IIf(blnByDept, "����id", "����id") & " From ��λ״����¼ Where ����id Is Not Null)", "") & vbNewLine & _
            IIf(blnLimitUnit, " And A.ID In (" & strUnitIDs & ")", "") & _
            " And (A.վ��=[2] Or A.վ�� is Null)" & _
            "Order By A.����"
'    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnByDept, "�ٴ�", "����"), strNodeNo)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvwDist_s.Nodes.Add("Root", 4, "D" & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, 1)
            
            If rsTmp!ID = UserInfo.����ID Then objNode.Selected = True
            If objNode.Key = strPreKey Then objNode.Selected = True
            
            objNode.Expanded = True
            rsTmp.MoveNext
        Next
    End If
    If tvwDist_s.SelectedItem Is Nothing Then
        tvwDist_s.Nodes(IIf(tvwDist_s.Nodes.Count > 1, 2, 1)).Selected = True
    End If
    
    InitUnits = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowPatis(Optional ByVal strIF As String, Optional blnSort As Boolean, Optional blnSet As Boolean, Optional ByVal strIFInfo As String)
'���ܣ����ݵ�ǰ�˵����Ҫ��(�Զ���������),��ȡ������Ϣ
'������strIF=" And ...."��ʽ�Ĺ�������
    Dim strSQL As String, strInfo As String, strCard As String
    Dim i As Double, j As Double, lngFamily As Long, lngDeptID As Long
    Dim lngCol���� As Long, lngColͣ�� As Long, lngPreRow As Long
    Dim blnByDept As Boolean
    Dim str����SQL As String, strFileds As String
    Dim rsTemp As Recordset
    
    On Error GoTo errH
    
    If Not blnSort Then
        blnByDept = mnuViewByDept(1).Checked
                
        '��һ��(ÿ���л�)����ȱʡ����
        If strIF = "" Then
            Select Case TabPatiState.SelectedItem.Key
                Case "T_���в���", "T_���ﲡ��", "T_���۲���" '���в���,���ﲡ�˻����۲���(����)
                    strIF = " And A.�Ǽ�ʱ�� Between trunc(Sysdate) And Sysdate"
                Case "T_��Ժ����" '��Ժ����
                    'strIF = " And P.��Ժ���� Between trunc(Sysdate) And Sysdate"
                Case "T_��Ժ����" '��Ժ����(����)
                    strIF = " And P.��Ժ���� Between trunc(Sysdate) And Sysdate"
            End Select
        End If
        If strIFInfo = "" Then
            Select Case TabPatiState.SelectedItem.Key
                Case "T_���в���", "T_���ﲡ��", "T_���۲���" '���в���,���ﲡ�˻����۲���(����)
                    strIFInfo = " And A.�Ǽ�ʱ�� Between trunc(Sysdate) And Sysdate"
            End Select
        End If
        
        If Not mnuViewStop.Checked Then strIF = strIF & " And A.ͣ��ʱ�� is NULL"
        
        '���￨����ʾ
        '55849:������,2012-11-21,��ԭ��Decode�жϵķ�ʽ��Ϊ�̶���ȡ�ֶ�,
        '��ΪDecode��һ������ʹ�ó�����ָ������ȡ�ֶ����ݣ����ܵ��µ��²鲻����������߷��صļ�¼�����ʳ���E-FAIL���󣬹�����ADO��Oracle�����Ե�Bug�����ض���Decode���ӱ��ѯͬʱʹ��ʱ����֣���û����ȷ�Ĺ��ɡ�
        'strCard = "Decode(" & IIf(gblnShowCard, 1, 0) & ",1,H.���￨��,LPAD('*',Length(H.���￨��),'*')) as ���￨,H.���￨�� as ���￨��,"
        If gblnShowCard = True Then
            strCard = "A.���￨�� as ���￨,A.���￨�� as ���￨��,"
        Else
            strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨,A.���￨�� as ���￨��,"
        End If
        
        Select Case TabPatiState.SelectedItem.Key
            Case "T_��Ժ����", "T_��Ժ����" '��Ժ���˻��Ժ����
                lngDeptID = Val(Mid(tvwDist_s.SelectedItem.Key, 2)) '���в�������ʱ,Ϊroot��0
                If lngDeptID <> 0 Then
                    If TabPatiState.SelectedItem.Key = "T_��Ժ����" Then
                        If blnByDept Then
                            strIF = strIF & " And R.����ID=[1]"
                        Else
                            strIF = strIF & " And R.����ID=[1]"
                        End If
                    Else
                        If blnByDept Then
                            strIF = strIF & " And P.��Ժ����ID=[1]"
                        Else
                            strIF = strIF & " And P.��ǰ����ID=[1]"
                        End If
                    End If
                ElseIf InStr(mstrPrivs, "���в���") = 0 Then
                    If blnByDept Then
                        strIF = strIF & " And (A.��ǰ����id Is NULL Or A.��ǰ����id In(Select ����ID From �������Ҷ�Ӧ Where Instr(','||[2]||',',','||����ID||',')>0))"
                    Else
                        strIF = strIF & " And (A.��ǰ����ID Is NULL Or Instr(','||[2]||',',','||A.��ǰ����ID||',')>0)"
                    End If
                End If
            Case "T_���в���"
                If InStr(mstrPrivs, "���в���") = 0 Then
                    If blnByDept Then
                        strIF = strIF & " And (A.��ǰ����id Is NULL Or A.��ǰ����id In(Select ����ID From �������Ҷ�Ӧ Where Instr(','||[2]||',',','||����ID||',')>0))"
                    Else
                        strIF = strIF & " And (A.��ǰ����ID Is NULL Or Instr(','||[2]||',',','||A.��ǰ����ID||',')>0)"
                    End If
                End If
        End Select
        
        '�����:51223
        
        '�����:52133
        '��ȡȱʡҽ�ƿ����
        mlngCardType = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , , True))
        '�����:53807
        If mlngCardType = 0 Then '��û������ȱʡ��������ʱ,ȱʡȱ���￨
            strSQL = "Select ID From ҽ�ƿ���� A Where A.����='���￨'"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ�ƿ����ID")
            If rsTemp.EOF = False Then mlngCardType = rsTemp!ID
            Set rsTemp = Nothing
        End If
        strSQL = "" & _
        "   Select �Ƿ����� From ҽ�ƿ���� Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ�ƿ����", mlngCardType)
        If rsTemp Is Nothing Then mbln�Ƿ�ȡ���� = False
        If rsTemp.RecordCount = 0 Then
            mbln�Ƿ�ȡ���� = False
        Else
            mbln�Ƿ�ȡ���� = rsTemp!�Ƿ����� = 0
        End If
        '���˿���
        str����SQL = "(Select f_List2str(Cast(COLLECT(G.����) as t_Strlist))" & _
            " From ����ҽ�ƿ���Ϣ G, ҽ�ƿ���� H" & _
            " Where G.����ID = A.����ID And G.�����ID = H.ID and G.״̬ = 0 And H.ID=[16]) ���￨��,"
         
        'סԺ����תסԺ
        If TabPatiState.SelectedItem.Key = "T_��Ժ����" Or TabPatiState.SelectedItem.Key = "T_���۲���" Then '��Ժ����
            mnuEdit_ToInPati.Visible = InStr(mstrPrivs, "סԺ����תסԺ") > 0
        Else
            mnuEdit_ToInPati.Visible = False
        End If
        mnuViewStop.Enabled = TabPatiState.SelectedItem.Key = "T_���в���" Or TabPatiState.SelectedItem.Key = "T_���ﲡ��"
        '�����:521333
        Select Case TabPatiState.SelectedItem.Key
            Case "T_���в���"  '���в���
                strFileds = "A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.����ѱ�,A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,A.����,A.����,A.����,A.����,A.��Ժʱ��,A.סԺĿ��,A.��Ժʱ��,A.סԺ����,A.��������," & _
                    " A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,A.�Ǽ�ʱ��, A.�Ǽ���," & _
                    " A.ͣ��ʱ��,A.��������,A.��������,A.��ҳID,A.�������"
                    
                strSQL = "Select A.����ID,A.�����,A.סԺ��," & str����SQL & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�,Nvl(P.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ," & _
                    " Nvl(A.ҽ����,E.��Ϣֵ) as ҽ����,X.���� as ����," & _
                    " B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,P.סԺĿ��," & _
                    " To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������," & _
                    " A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��, p.�Ǽ���," & _
                    " To_Char(A.ͣ��ʱ��,'YYYY-MM-DD') as ͣ��ʱ��,0 as ��������,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������,P.��ҳID,A.��ҳID �������" & _
                    " From ������ҳ P,������Ϣ A,������ҳ�ӱ� E,���ű� B,���ű� C,������� X" & _
                    " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.����=X.���(+)" & _
                    " And A.����ID=P.����ID(+) And A.��ҳID=P.��ҳID(+) " & strIF & _
                    " And A.����ID=E.����ID(+) And Nvl(A.��ҳID,0)=E.��ҳID(+) And E.��Ϣ��(+)='ҽ����'" & _
                    " Order by A.�Ǽ�ʱ�� Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "���ڶ�ȡ���в����嵥,���Ժ� ..."
                tvwDist_s.Visible = False
            Case "T_��Ժ����" '��Ժ����
                strFileds = "A.����ID,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.סԺ�ѱ�,A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,A.����,A.����,A.����,A.����,A.��Ժʱ��,A.סԺĿ��," & _
                    " A.סԺ����,A.��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,A.�Ǽ�ʱ��, A.�Ǽ���, " & _
                    " A.ͣ��ʱ��,A.��������,A.��������,A.��ҳID,A.�������"
                '58842,������,2013-02-25,��Ժ���˶�ȡ(����Ժ�����ж�ȡ)
                strSQL = "Select A.����ID,A.סԺ��," & str����SQL & "NVL(P.����,A.����) ����,NVL(P.�Ա�,A.�Ա�) �Ա�,NVL(P.����,A.����) ����,P.�ѱ� as סԺ�ѱ�,P.ҽ�Ƹ��ʽ," & _
                    " E.��Ϣֵ as ҽ����,X.���� as ����," & _
                    " B.���� as ����,C.���� as ����,Decode(P.״̬,1,P.��Ժ����,Nvl(P.��Ժ����,'��ͥ')) as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,P.סԺĿ��," & _
                    " A.סԺ����,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��, p.�Ǽ���, " & _
                    " To_Char(A.ͣ��ʱ��,'YYYY-MM-DD') as ͣ��ʱ��,Nvl(P.��������,0) as ��������,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������,P.��ҳID,A.��ҳID �������" & _
                    " From ������ҳ P,������Ϣ A,������ҳ�ӱ� E,���ű� B,���ű� C,������� X,��Ժ���� R" & _
                    " Where P.��ǰ����ID=B.ID(+) And P.��Ժ����ID=C.ID And A.����=X.���(+)" & _
                    " And A.����ID=P.����ID And A.��ҳID=P.��ҳID And Nvl(P.��ҳID,0)<>0 " & strIF & _
                    " And E.��Ϣ��(+)='ҽ����' And P.����ID=E.����ID(+) And P.��ҳID=E.��ҳID(+) And R.����ID=A.����ID" & _
                    " Order by A.��Ժʱ�� Desc,A.סԺ�� Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "���ڶ�ȡ��Ժ�����嵥,���Ժ� ..."
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_��Ժ����"  '��Ժ����
                '����28813 by lesfeng 2010-04-07 A.סԺ�� A.��Ժʱ�� A.��Ժʱ�� A.סԺ����
                strFileds = "A.����ID,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.סԺ�ѱ�,A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,A.����,A.��Ժʱ��,A.סԺĿ��,A.��Ժʱ��," & _
                    " A.סԺ����,A.��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,A.�Ǽ�ʱ��, A.�Ǽ���, " & _
                    " A.ͣ��ʱ��,A.��������,A.��������,A.��ҳID,A.�������"
                    
                strSQL = "Select A.����ID,P.סԺ��," & str����SQL & "NVL(P.����,A.����) ����,NVL(P.�Ա�,A.�Ա�) �Ա�,NVL(P.����,A.����) ����,P.�ѱ� as סԺ�ѱ�,P.ҽ�Ƹ��ʽ," & _
                    " E.��Ϣֵ as ҽ����,X.���� as ����," & _
                    " To_Char(P.��Ժ����,'YYYY-MM-DD') as ��Ժʱ��,P.סԺĿ��,To_Char(P.��Ժ����,'YYYY-MM-DD') as ��Ժʱ��," & _
                    " P.��ҳID as סԺ����,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��, p.�Ǽ���, " & _
                    " To_Char(A.ͣ��ʱ��,'YYYY-MM-DD') as ͣ��ʱ��,Nvl(P.��������,0) as ��������,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������,P.��ҳID,A.��ҳID �������" & _
                    " From ������ҳ P,������Ϣ A,������ҳ�ӱ� E,������� X" & _
                    " Where A.����ID=P.����ID And Nvl(P.��ҳID,0)<>0" & _
                    " And P.��Ժ���� is Not NULL And A.����=X.���(+)" & strIF & _
                    " And P.����ID=E.����ID(+) And NVL(P.��ҳID,0)=E.��ҳID(+) And E.��Ϣ��(+)='ҽ����'" & _
                    " Order by A.��Ժʱ�� Desc,A.סԺ��"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "���ڶ�ȡ��Ժ�����嵥,���Ժ� ..."
                tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
                tvwDist_s.Visible = mnuViewToolDist.Enabled
            Case "T_���ﲡ��" '���ﲡ��
                strFileds = "A.����ID,A.�����," & strCard & "A.����,A.�Ա�,A.����," & _
                     IIf(glngSys Like "8??", "A.��Ա�ȼ�", "A.����ѱ�") & ",A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,A.����,A.��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,A.�Ǽ�ʱ��,A.ͣ��ʱ��,A.��������,A.��������," & _
                    " A.סԺ����,A.��ҳID,A.�������"
                    
                strSQL = "Select A.����ID,A.�����," & str����SQL & "A.����,A.�Ա�,A.����," & _
                    " A.�ѱ� as " & IIf(glngSys Like "8??", "��Ա�ȼ�", "����ѱ�") & ",A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,X.���� as ����," & _
                    " To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
                    " To_Char(A.ͣ��ʱ��,'YYYY-MM-DD') as ͣ��ʱ��,0 as ��������,Decode(A.����,Null,'��ͨ����','ҽ������') ��������," & _
                    " NULL סԺ����,NULL ��ҳID,NULL �������" & _
                    " From ������Ϣ A,������� X" & _
                    " Where A.��ǰ����ID is NULL And A.��ǰ����ID Is NULL" & _
                    " And A.��ҳID IS NULL And A.����=X.���(+)" & strIF & _
                    " Order by A.�Ǽ�ʱ�� ,A.����� Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                If glngSys Like "8??" Then
                    strInfo = "���ڶ�ȡ�ͻ��嵥,���Ժ� ..."
                Else
                    strInfo = "���ڶ�ȡ���ﲡ���嵥,���Ժ� ..."
                End If
                tvwDist_s.Visible = False
            Case "T_���۲���"    '���۲���
                strFileds = "A.����ID,A.����,A.�����,A.סԺ��,A.סԺ����," & strCard & "A.����,A.�Ա�,A.����," & _
                     IIf(glngSys Like "8??", "A.��Ա�ȼ�", "A.����ѱ�") & ",A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,A.����,A.��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,A.�Ǽ�ʱ��, A.�Ǽ���, " & _
                    " A.ͣ��ʱ��,A.��������,A.��������,A.��ҳID,A.�������"
                    
                strSQL = "Select Distinct A.����ID,Decode(P.��������,1,'��������','סԺ����') as ����,A.�����,A.סԺ��,NULL as סԺ����," & str����SQL & "NVL(P.����,A.����) ����,NVL(P.�Ա�,A.�Ա�) �Ա�,NVL(P.����,A.����) ����," & _
                    " A.�ѱ� as " & IIf(glngSys Like "8??", "��Ա�ȼ�", "����ѱ�") & ",A.ҽ�Ƹ��ʽ," & _
                    " A.ҽ����,X.���� as ����," & _
                    " To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                    " A.���֤��,A.�ֻ���,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��, p.�Ǽ���, " & _
                    " To_Char(A.ͣ��ʱ��,'YYYY-MM-DD') as ͣ��ʱ��,Nvl(P.��������,0) as ��������,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������,P.��ҳID,Decode(P.��������,2,A.��ҳID,NULL) �������" & _
                    " From ������ҳ P,������Ϣ A,������� X" & _
                    " Where A.����ID=P.����ID And A.��ҳID=P.��ҳID And P.��������<>0 And P.סԺ�� Is Null" & _
                    " And A.����=X.���(+)" & strIF & _
                    " Order by ����,�Ǽ�ʱ�� Desc"
                strSQL = "Select " & strFileds & " From (" & strSQL & ") A"
                strInfo = "���ڶ�ȡ���۲����嵥,���Ժ� ..."
                tvwDist_s.Visible = False
        End Select
        
        Call Form_Resize
        
        Call zlCommFun.ShowFlash(strInfo, Me)
        DoEvents
        Me.Refresh
        
        With SQLCondition
            Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID, mstrUserUnitIDs, .�Ǽ�ʱ��B, .�Ǽ�ʱ��E, .����ʱ��B, .����ʱ��E, _
                .��Ժʱ��B, .��Ժʱ��E, .��Ժʱ��B, .��Ժʱ��E, .סԺ��, .�Ա�, .����, .�ѱ�, .Patient, mlngCardType, .ҽ�Ƹ��ʽ)
        End With
    End If
    
    '35632:������,2013-07-29
    If mblnInitGrid = True Then SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    
    mshPati.Clear
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call SetHeader(blnSet)
        If glngSys Like "8??" Then
            stbThis.Panels(2).Text = "��ǰ����û�й��˳��κοͻ�"
        Else
            stbThis.Panels(2).Text = "��ǰ����û�й��˳��κβ���"
        End If
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(blnSet)
        
        lngFamily = 0
        lngCol���� = GetColNum("����")
        lngColͣ�� = GetColNum("ͣ��ʱ��")
        lngPreRow = mshPati.Row
        
        mshPati.Redraw = False
        For i = 1 To mshPati.Rows - 1
            If TabPatiState.SelectedItem.Key = "T_��Ժ����" Then '��Ժ����ͳ�Ƽ�ͥ��������
                If mshPati.TextMatrix(i, lngCol����) = "��ͥ" Then
                    lngFamily = lngFamily + 1
                End If
            End If
            If mshPati.TextMatrix(i, lngColͣ��) <> "" Then 'ͣ�ò��˺�ɫ��ʾ
                mshPati.Row = i
                For j = 0 To mshPati.Cols - 1
                    mshPati.Col = j
                    mshPati.CellForeColor = &HC0&
                Next
            End If
        Next
        mshPati.Row = lngPreRow: mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
        mshPati.Redraw = True
        
        If glngSys Like "8??" Then
            stbThis.Panels(2) = "�� " & mrsPati.RecordCount & " ���ͻ�"
        Else
            If TabPatiState.SelectedItem.Key = "T_��Ժ����" Then
                stbThis.Panels(2) = "�� " & mrsPati.RecordCount & " ������,���м�ͥ���� " & lngFamily & " ��"
            Else
                stbThis.Panels(2) = "�� " & mrsPati.RecordCount & " ������"
            End If
        End If
    End If
    Call mshPati_EnterCell
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetHeader(Optional blnSet As Boolean)
    Dim strHead As String
    Dim i As Integer
    
    mblnInitGrid = False
    
    Select Case TabPatiState.SelectedItem.Key
        Case "T_���в���" '���в���
            strHead = "����ID,1,750|�����,1,750|סԺ��,1,750|���￨,1,850|���￨��,1,0|����,1,800|�Ա�,1,500|����,1,800|����ѱ�,1,850|ҽ�Ƹ��ʽ,1,1400|" & _
                "ҽ����,1,1200|����,1,1500|����,1,850|����,1,850|����,1,500|��Ժʱ��,1,1000|סԺĿ��,1,800|��Ժʱ��,1,1000|סԺ����,4,850|��������,1,1000|" & _
                "����,1,500|����,1,800|����,1,600|ѧ��,1,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|�ֻ���,1,1100|��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,1,1000|�Ǽ���,1,800|" & _
                "ͣ��ʱ��,1,0|��������,1,0|��������,1,1000|��ҳID,1,0|�������,1,0"
        Case "T_��Ժ����" '��Ժ����
            strHead = "����ID,1,750|סԺ��,1,750|���￨,1,850|���￨��,1,0|����,1,800|�Ա�,1,500|����,1,800|סԺ�ѱ�,1,850|ҽ�Ƹ��ʽ,1,1400|" & _
                "ҽ����,1,1200|����,1,1500|����,1,850|����,1,850|����,1,500|��Ժʱ��,1,1000|סԺĿ��,1,800|סԺ����,4,850|��������,1,1000|" & _
                "����,1,500|����,1,800|����,1,600|ѧ��,1,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|�ֻ���,1,1100|��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,1,1000|�Ǽ���,1,800|" & _
                "ͣ��ʱ��,1,0|��������,1,0|��������,1,1000|��ҳID,1,0|�������,1,0"
        Case "T_��Ժ����" '��Ժ����
            strHead = "����ID,1,750|סԺ��,1,750|���￨,1,850|���￨��,1,0|����,1,800|�Ա�,1,500|����,1,800|סԺ�ѱ�,1,850|ҽ�Ƹ��ʽ,1,1400|" & _
                "ҽ����,1,1200|����,1,1500|��Ժʱ��,1,1000|סԺĿ��,1,800|��Ժʱ��,1,1000|סԺ����,4,850|��������,1,1000|����,1,500|����,1,800|����,1,600|" & _
                "ѧ��,1,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|�ֻ���,1,1100|��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,1,1000|�Ǽ���,1,800|ͣ��ʱ��,1,0|��������,1,0|��������,1,1000|��ҳID,1,0|�������,1,0"
        Case "T_���ﲡ��" '���ﲡ��
            If glngSys Like "8??" Then
                strHead = "�ͻ�ID,1,750|�ͻ���,1,0|��Ա��,1,850|����,1,800|�Ա�,1,500|����,1,800|��Ա�ȼ�,1,850|ҽ�Ƹ��ʽ,1,1400|" & _
                    "ҽ����,1,1200|����,1,1500|��������,1,1000|����,1,500|����,1,800|����,1,600|ѧ��,1,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|�ֻ���,1,1100|" & _
                    "��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,1,1000|ͣ��ʱ��,1,0|��������,1,0|��������,1,1000|סԺ����,1,0|��ҳID,1,0|�������,1,0"
            Else
                strHead = "����ID,1,750|�����,1,750|���￨,1,850|���￨��,1,0|����,1,800|�Ա�,1,500|����,1,800|����ѱ�,1,850|ҽ�Ƹ��ʽ,1,1400|" & _
                    "ҽ����,1,1200|����,1,1500|��������,1,1000|����,1,500|����,1,800|����,1,600|ѧ��,1,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|�ֻ���,1,1100|" & _
                    "��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,1,1000|ͣ��ʱ��,1,0|��������,1,0|��������,1,1000|סԺ����,1,0|��ҳID,1,0|�������,1,0"
            End If
        Case "T_���۲���" '���۲���
            strHead = "����ID,1,750|����,1,1000|�����,1,750|סԺ��,1,750|סԺ����,1,750|���￨,1,850|���￨��,1,0|����,1,800|�Ա�,1,500|����,1,800|����ѱ�,1,850|ҽ�Ƹ��ʽ,1,1400|" & _
                    "ҽ����,1,1200|����,1,1500|��������,1,1000|����,1,500|����,1,800|����,1,600|ѧ��,1,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|�ֻ���,1,1100|" & _
                    "��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,1,1000|�Ǽ���,1,800|ͣ��ʱ��,1,0|��������,1,0|��������,1,1000|��ҳID,1,0|�������,1,0"
    End Select
    
    With mshPati
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or blnSet Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or blnSet Then Call RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        
        If glngSys Like "8??" Then .ColWidth(1) = 0
        
        .RowHeight(0) = 320
        
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
        Call mshPati_EnterCell

        .Redraw = True
    End With
    mblnInitGrid = True
End Sub

Private Sub mshPati_EnterCell()
    mshPati.ForeColorSel = mshPati.CellForeColor
    Call SetMenuEnabled
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = 0
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '˫�����ʱ��ִ��
        mblnDown = False
        
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If glngSys Like "8??" Then
            If mshPati.TextMatrix(1, GetColNum("�ͻ�ID")) = "" Then Exit Sub
        Else
            If mshPati.TextMatrix(1, GetColNum("����ID")) = "" Then Exit Sub
        End If
        
        Set mshPati.DataSource = Nothing
        
        If glngSys Like "8??" Then
            Select Case mshPati.TextMatrix(0, lngCol)
                Case "�ͻ�ID"
                    mrsPati.Sort = "����ID" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
                Case "��Ա��"
                    mrsPati.Sort = "���￨" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
                Case Else
                    mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
            End Select
        Else
            mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        End If
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        
        Call ShowPatis(, True, gblnMyStyle)
    End If
End Sub

Private Sub SetMenuEnabled()
'���ܣ����ݵ�ǰ��¼������ò˵�����״̬
    Dim lng����ID As Long, byt�������� As Byte, strͣ��ʱ�� As String, lng������� As Long, lng��ҳID As Long
    Dim strCard As String
    Dim blnPrivs As Boolean
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    lng������� = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�������")))
    lng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
    strCard = Trim(mshPati.TextMatrix(mshPati.Row, GetColNum("���￨")))
    
    byt�������� = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��������")))
    strͣ��ʱ�� = mshPati.TextMatrix(mshPati.Row, GetColNum("ͣ��ʱ��"))
        
    mnuEdit_Stop.Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And mnuViewStop.Enabled And lng������� = lng��ҳID 'ͣ��
    mnuEdit_Restore.Enabled = lng����ID <> 0 And strͣ��ʱ�� <> "" And mnuViewStop.Enabled And lng������� = lng��ҳID 'ȡ��ͣ��
    mnuEdit_ToInPati.Enabled = lng����ID <> 0 And byt�������� = 2 And lng������� = lng��ҳID                        'תΪסԺ����
    '----
    mnuFile_Print.Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                               '��ӡ
    mnuFile_Preview.Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                             'Ԥ��
    mnuFile_Excel.Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                               'excel����
    tbr.Buttons("Print").Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                        '��ӡ
    tbr.Buttons("Preview").Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                      'Ԥ��
    
    mnuEdit_Modi.Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And lng������� = lng��ҳID                           '�޸�
    If mnuEdit_Modi.Enabled And TabPatiState.SelectedItem.Key = "T_��Ժ����" Then                                    '��Ժ
        mnuEdit_Modi.Enabled = InStr(mstrPrivs, "�޸ĳ�Ժ����") > 0
    End If
    mnuEdit_Del.Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And lng������� = lng��ҳID                            'ɾ��
    mnuEdit_Merge.Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And lng������� = lng��ҳID                          '�ϲ�
    mnuEdit_View.Enabled = lng����ID <> 0                                                                            '��ݿ�Ƭ
    mnuEditDelCard.Enabled = lng����ID <> 0 And strCard <> "" And mbln�Ƿ�ȡ����  '�����:52133                    'ȡ�����Ű�
    
    tbr.Buttons("Modi").Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And lng������� = lng��ҳID                    '�޸�
    If tbr.Buttons("Modi").Enabled And TabPatiState.SelectedItem.Key = "T_��Ժ����" Then                                    '��Ժ
        tbr.Buttons("Modi").Enabled = InStr(mstrPrivs, "�޸ĳ�Ժ����") > 0
    End If
    tbr.Buttons("Del").Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And lng������� = lng��ҳID                     'ɾ��
    tbr.Buttons("Merge").Enabled = lng����ID <> 0 And strͣ��ʱ�� = "" And lng������� = lng��ҳID                   '�ϲ�
    tbr.Buttons("View").Enabled = lng����ID <> 0                                           '��ݿ�Ƭ
    
    mnuViewGo.Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                                   '��λ
    tbr.Buttons("Go").Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                           '��λ
    mnuEditMzReCalc.Enabled = lng����ID <> 0
    mnuEdit_Surety.Enabled = lng����ID <> 0 And lng������� = lng��ҳID                                              '����
    mnuEdit_QueryPass.Enabled = lng����ID <> 0 And lng������� = lng��ҳID
    
    blnPrivs = InStr(";" & GetPrivFunc(glngSys, 9003) & ";", ";���˼���;") > 0
    mnuEdit_Family.Visible = blnPrivs
    mnuEdit_FamilyView.Visible = blnPrivs
    mnuEdit_FamilyAdd.Visible = blnPrivs
    mnuEdit_FamilyView.Enabled = lng����ID <> 0
    tbr.Buttons("FamilySplit").Visible = blnPrivs
    tbr.Buttons("Family").Visible = blnPrivs
    tbr.Buttons("Family").ButtonMenus.Item("FamilyView").Enabled = lng����ID <> 0
    '������Ϣ����
    mnuEditPatiInfo.Visible = InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;")
    If lng����ID <> 0 Then
        mnuEditPatiInfo.Enabled = strͣ��ʱ�� = "" And mnuEditPatiInfo.Visible
    Else
        mnuEditPatiInfo.Enabled = mnuEditPatiInfo.Visible
    End If
    
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    If glngSys Like "8??" Then
        stbThis.Panels(2).Text = "���ڶ�λ���������Ŀͻ�,��ESC��ֹ ..."
    Else
        stbThis.Panels(2).Text = "���ڶ�λ���������Ĳ���,��ESC��ֹ ..."
    End If
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmPatiFind
            If .txt����ID.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("�ͻ�ID")) = .txt����ID.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����ID")) = .txt����ID.Text
                End If
            End If
            If .txt���￨.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("��Ա��")) = .txt���￨.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���￨")) = .txt���￨.Text
                End If
            End If
            If .txt�����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("�����")) = .txt�����.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����")) = .txt����.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshPati.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If .txt���֤.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���֤��")) = .txt���֤.Text
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
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

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub LoadPlugInMnu()
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim blnHave As Boolean
    
    If CreatePlugInOK(glngModul) Then
        blnHave = True
    End If
    
    If glngSys Like "8??" Then blnHave = False
    
    mnuEdit_PlugIn.Visible = blnHave
    tbr.Buttons("PlugIn").Visible = blnHave
    
    If blnHave Then
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul)
        Call zlPlugInErrH(Err, "GetFuncNames")
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then Exit Sub
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuEdit_PlugItem(i)
            End If
            mnuEdit_PlugItem(i).Caption = CStr(arrTmp(i))
            mnuEdit_PlugItem(i).Tag = CStr(arrTmp(i))
            
            If i <= 9 Then
                mnuEdit_PlugItem(i).Caption = CStr(arrTmp(i)) & "(&" & IIf(i = 9, 0, i + 1) & ")"
            End If
        Next
    End If
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lngPatiId As Long
    Dim lngPageID As Long
    
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
        MsgBox "δѡ���κβ��ˣ�����ִ�д˲�����", vbExclamation, gstrSysName: Exit Sub
    End If
        
    If CreatePlugInOK(glngModul) Then
        lngPatiId = CLng(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
        lngPageID = CLng(Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID"))))
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, glngModul, strFunName, lngPatiId, lngPageID, 0)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub
