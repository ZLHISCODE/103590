VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageExamine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "���˷�������"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   540
   ClientWidth     =   11040
   Icon            =   "frmManageExamine.frx":0000
   KeyPreview      =   -1  'True
   Picture         =   "frmManageExamine.frx":1601A
   ScaleHeight     =   6480
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TabStrip tbsClass 
      Height          =   340
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   609
      TabFixedWidth   =   2290
      TabFixedHeight  =   529
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫ��(&0)"
            Key             =   "ȫ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�г�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11040
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
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   7785
         _ExtentX        =   13732
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
            NumButtons      =   12
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
               Key             =   "Line_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�༭"
               Key             =   "Edit"
               Description     =   "������Ŀ�༭"
               Object.ToolTipText     =   "�༭"
               Object.Tag             =   "�༭"
               ImageKey        =   "Billing"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Edit"
                     Object.Tag             =   "�༭"
                     Text            =   "�༭"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line_2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "���˳�Ժ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Object.ToolTipText     =   "��λ����������Ŀ"
               Object.Tag             =   "��λ"
               ImageKey        =   "Go"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8955
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   2670
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5145
      ScaleWidth      =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   690
      Width           =   45
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
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
      NumItems        =   10
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
         Text            =   "����"
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
            Picture         =   "frmManageExamine.frx":161A8
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":163C2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":165DC
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":167F6
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":16F70
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1718A
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":173A4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":175BE
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":177D8
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":179F2
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":180EC
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":187E6
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":18EE0
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":190FA
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
            Picture         =   "frmManageExamine.frx":19314
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1952E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":19748
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":19962
            Key             =   "View"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A0DC
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A2F6
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A510
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A72A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1A944
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1AB5E
            Key             =   "Billing"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1B258
            Key             =   "Price"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1B952
            Key             =   "Auditing"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1C04C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1C266
            Key             =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   0
      TabIndex        =   2
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
         NumTabs         =   2
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
            Picture         =   "frmManageExamine.frx":1C480
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1CD5A
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
            Picture         =   "frmManageExamine.frx":1D634
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageExamine.frx":1DF0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   6120
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageExamine.frx":1E7E8
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
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
   Begin VSFlex8Ctl.VSFlexGrid vsExist 
      Height          =   4515
      Left            =   2760
      TabIndex        =   10
      Top             =   1300
      Width           =   8265
      _cx             =   14579
      _cy             =   7964
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16574424
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmManageExamine.frx":1F07A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00808080&
      Caption         =   " ��ǰ�������������շ���Ŀ"
      ForeColor       =   &H00C0FFFF&
      Height          =   180
      Left            =   2775
      TabIndex        =   7
      Top             =   765
      Width           =   6990
   End
   Begin VB.Label lbl_s 
      BackColor       =   &H00808080&
      Caption         =   " ʱ��:2001-01-01��2001-01-01"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   30
      TabIndex        =   6
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_EditItem 
         Caption         =   "������Ŀ����(&A)"
      End
      Begin VB.Menu mnuEdit_split 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_EditTemplet 
         Caption         =   "��Ŀģ�����(&B)"
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
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "���˳�Ժ����(&T)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "��λ������Ŀ(&C)"
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ����(&G)"
      End
      Begin VB.Menu mnuView_7 
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
      Begin VB.Menu mnuViewFindPati 
         Caption         =   "δ����������Ŀ����(&W)"
      End
      Begin VB.Menu mnuViewPatiMode 
         Caption         =   "��ʾ���˷�ʽ(&K)"
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
      Begin VB.Menu mnuView_8 
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
Attribute VB_Name = "frmManageExamine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mlngCurRow As Long, mlngTopRow As Long
Public mdtBegin As Date, mdtEnd As Date

Private mblnFirst As Boolean

Private mlngDeptID As Long

Private mstrPrePati As String

Private mstrPage As String

Public mstrPrivs  As String

Private mintBedLen  As Integer
Private mblnUnLoad As Boolean
Private mrsPati As New ADODB.Recordset
Public mrsExistItem As New ADODB.Recordset

Private Enum ColNum
    ��� = 0: ����: ����: ���: ����: ��λ: ˵��: ������: ����ʱ��
End Enum
Private Sub cboDept_Click()
    If cboDept.ItemData(cboDept.ListIndex) = mlngDeptID Then Exit Sub
    mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
    
    mstrPage = ""
    vsExist.Rows = 1
    If Visible Then Call tbs_Click
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    If mblnUnLoad = True Then Unload Me: Exit Sub

End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    lbl_s.Left = Me.ScaleLeft
    lbl_s.Top = Me.ScaleTop + cbrH + 45
    lbl_s.Width = pic.Left
    
    tbs.Left = Me.ScaleLeft
    tbs.Top = lbl_s.Top + lbl_s.Height + 45
    tbs.Width = lbl_s.Width
    
    lvw.Left = Me.ScaleLeft
    lvw.Top = tbs.Top + tbs.Height - 75
    lvw.Width = lbl_s.Width
    lvw.Height = Me.ScaleHeight - staH - cbrH - lbl_s.Height - tbs.Height - 15
    
    pic.Top = Me.ScaleTop + cbrH
    pic.Height = Me.ScaleHeight - cbrH - staH
    
    lblMoney.Left = pic.Left + pic.Width
    lblMoney.Top = Me.ScaleTop + cbrH + 45
    lblMoney.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
    
    If tbsClass.Visible = True Then
        tbsClass.Left = pic.Left + pic.Width
        tbsClass.Top = lblMoney.Top + lblMoney.Height
        tbsClass.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
        
        vsExist.Left = tbsClass.Left
        vsExist.Top = tbsClass.Top + tbsClass.Height - 50
        vsExist.Width = tbsClass.Width
        vsExist.Height = pic.Height - tbsClass.Height - lblMoney.Height + 50
        vsExist.ZOrder
    Else
         vsExist.Left = pic.Left + pic.Width
         vsExist.Top = lblMoney.Top + lblMoney.Height
         vsExist.Width = Me.ScaleWidth - lbl_s.Width - pic.Width
          vsExist.Height = pic.Height - lblMoney.Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_DblClick()
    If mnuEdit_EditItem.Visible = True And tbs.SelectedItem.Index = 1 Then
        Call mnuEdit_EditItem_Click
    End If
End Sub

Private Sub mnuEdit_EditItem_Click()
    Dim lng����ID As Long, lng��ҳID As Long, lng���� As Long
    Dim arrTmp As Variant
    
    If tbs.SelectedItem.Index = 2 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    arrTmp = Split(Mid(lvw.SelectedItem.Key, 2), "_")
    lng����ID = Val(arrTmp(0))
    lng��ҳID = Val(arrTmp(1))
    lng���� = Val(arrTmp(2))

    If InStr(mstrPrivs, "����������Ŀ") > 0 Then
        Call frmExamineEdit.ExamineEdit(lng����ID, lng��ҳID, lng����, False, False)
    ElseIf InStr(mstrPrivs, "ɾ��������Ŀ") > 0 Then
        Call frmExamineEdit.ExamineEdit(lng����ID, lng��ҳID, lng����, True, False)
    End If
    If lvw.SelectedItem Is Nothing Then Exit Sub
    mstrPrePati = ""
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub mnuEdit_EditTemplet_Click()
    frmExamineEdit.ExamineEdit 0, 0, 0, False, True
End Sub

Private Sub mnuFile_Excel_Click()
    zlRptPrint 3
End Sub

Private Sub mnuFile_PreView_Click()
    zlRptPrint 2
End Sub

Private Sub mnuFile_Print_Click()
    zlRptPrint 1
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    mlngDeptID = 0
    Call InitUnits
End Sub

Private Sub mnuViewFilter_Click()
    frmSetExamine.EditWhere Me
    If tbs.SelectedItem Is Nothing Then Exit Sub
    If tbs.SelectedItem.Index = 2 Then
        Call LoadPatients
    End If
End Sub


Private Sub mnuViewFind_Click()
    Dim lngRow As Long
    Dim strOld As String
    
    If vsExist.Rows = 1 Then Exit Sub
    If tbsClass.Visible = True Then
        For lngRow = 1 To tbsClass.Tabs.Count
            frmExamineFind.cbo���.AddItem tbsClass.Tabs.Item(lngRow).Key, lngRow - 1
        Next lngRow
        strOld = tbsClass.SelectedItem.Key
    Else
        frmExamineFind.cbo���.AddItem vsExist.TextMatrix(1, ColNum.���), 0
    End If
    If frmExamineFind.cbo���.ListCount > 0 Then frmExamineFind.cbo���.ListIndex = 0
    Set frmExamineFind.mrsfind = mrsExistItem
    
    frmExamineFind.Show 1, Me
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
            If mintBedLen - zlCommFun.ActualLen(strBed) > 0 Then
                strBed = String(mintBedLen - zlCommFun.ActualLen(strBed), " ") & strBed
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
        lvw.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID & "_" & mrsPati!����).Selected = True
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub mnuViewFindPati_Click()
    mnuViewFindPati.Checked = Not mnuViewFindPati.Checked
    Call LoadPatients
End Sub

Private Sub mnuViewGo_Click()
    Dim blnPati As Boolean

    frmDeptGo.Show 1, Me
    If gblnOK = True Then Call FindPati
End Sub

Private Sub mnuViewStyle_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuFile_Quit_Click()
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
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
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
        If lbl_s.Width + X < 2580 Or vsExist.Width - X < 3500 Then Exit Sub
        pic.Left = pic.Left + X
        Call Form_Resize
        Me.Refresh
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lvw.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_Quit_Click
        Case "Edit"
            mnuEdit_EditItem_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Filter"
             mnuViewFilter_Click
        Case "Go"
            mnuViewFind_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Style"
            Call SetView((lvw.View + 1) Mod 4)
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
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

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub Form_Load()
    Dim i As Long, datTmp As Date
    mblnUnLoad = False
    mstrPrivs = gstrPrivs
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    Call RestoreWinState(Me, App.ProductName)
                
    
    '���ݱ����б�ʽ���ò˵�
    Call SetView(lvw.View)

    mlngCurRow = 1: mlngTopRow = 1
    mblnFirst = True
    
    'Ȩ������

    
    If InStr(mstrPrivs, "����������Ŀ") > 0 Or InStr(mstrPrivs, "ɾ��������Ŀ") > 0 Then
        mnuEdit_EditItem.Visible = True
        tbr.Buttons.Item("Edit").Visible = True
    Else
        mnuEdit_EditItem.Visible = False
        tbr.Buttons.Item("Edit").Visible = False
        mnuEdit_split.Visible = False
    End If
    
    If InStr(mstrPrivs, "ģ�����") > 0 Then
        mnuEdit_EditTemplet.Visible = True
        mnuEdit_split.Visible = mnuEdit_EditItem.Visible
        tbr.Buttons.Item("Line_2").Visible = mnuEdit_EditItem.Visible
    Else
        
        mnuEdit_EditTemplet.Visible = False
        mnuEdit_split.Visible = False
    End If
    
    '����(����Ա������������)
    If Not InitUnits Then mblnUnLoad = True: Exit Sub
    If cboDept.ListIndex = -1 Then
        MsgBox "û�з�������������,���㲻�������в���Ȩ��,����ʹ�ò��˷���������", vbInformation, gstrSysName
       mblnUnLoad = True: Exit Sub
    End If
        
    datTmp = zlDatabase.Currentdate
    mdtBegin = Format(DateAdd("m", -1, datTmp), "YYYY-MM-DD")
    mdtEnd = Format(datTmp, "YYYY-MM-DD")
    
    Call LoadPatients '�����Ѱ���Call SetDetail Call SetHeader  Call SetMenu
    
End Sub

Private Sub tbs_Click()
    If Not Visible Then Exit Sub
    If tbs.SelectedItem.Key = mstrPage Then Exit Sub
    If tbs.SelectedItem.Index = 2 Then
         mnuEdit_EditItem.Enabled = False
         tbr.Buttons.Item("Edit").Enabled = False
         mnuViewFilter.Enabled = True
         tbr.Buttons.Item("Filter").Enabled = True
    Else
         mnuEdit_EditItem.Enabled = True
         tbr.Buttons.Item("Edit").Enabled = True
         mnuViewFilter.Enabled = False
         tbr.Buttons.Item("Filter").Enabled = False
    End If
    '��ȡ����
    mstrPage = tbs.SelectedItem.Key
    Call LoadPatients
    lvw.SetFocus
End Sub

Private Sub ReadExistsItem(lng����ID As Long, lng��ҳID As Long, lng���� As Long)
    Dim strSQL As String
    Dim lngRow As Long
    Dim strClass As String, strOld As String
    Dim arrClass As Variant
    Dim blnClass As Boolean
    Dim objTab As MSComctlLib.Tab
    Dim i As Integer
    
    strSQL = " Select C.���� ���, A.����, A.����, B.ʹ������, B.��������, A.���, A.����, A.���㵥λ, A.˵��,B.������,B.����ʱ��" & _
             " From �շ���ĿĿ¼ A,����������Ŀ B, �շ���Ŀ��� C" & _
             " Where A.��� = C.���� And A.ID=B.��ĿID And B.����ID=[1] AND B.��ҳID=[2]" & _
             " Order by ���,����"
    
    On Error GoTo errHandle
    Set mrsExistItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    
    Set vsExist.DataSource = mrsExistItem
    
    If mrsExistItem.RecordCount = 0 Then
        With vsExist
            .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = 4
        End With
    End If
    
    While Not mrsExistItem.EOF
        If mrsExistItem!��� <> strOld Then
            strClass = strClass & "," & mrsExistItem!���
            strOld = mrsExistItem!���
        End If
        mrsExistItem.MoveNext
    Wend
    
    For i = tbsClass.Tabs.Count To 2 Step -1
        tbsClass.Tabs.Remove i
    Next
    
    arrClass = Split(Mid(strClass, 2), ",")
    
    If UBound(arrClass) > 0 Then
        tbsClass.Visible = True
        tbsClass.ZOrder
        Call Form_Resize
        For i = 0 To UBound(arrClass)
            If i < 9 Then
                '��Alt��ݼ������޷�����
                Set objTab = tbsClass.Tabs.Add(, arrClass(i), arrClass(i) & "(&" & i + 1 & ")")
            Else
                Set objTab = tbsClass.Tabs.Add(, arrClass(i), arrClass(i), 2)
            End If
            objTab.Tag = arrClass(i)
        Next
    Else
        tbsClass.Visible = False
    End If
'    If vsExist.Tag <> "" Then
'        Set tbsClass.SelectedItem = tbsClass.Tabs.Item(Int(vsExist.Tag))
'        Call tbsClass_Click
'    End If
'    '�ָ���˳��:Ӧ����������֮ǰ
'    Call RestoreColPosition
'    '������:������,�Ա���洦���к�
'    Call RestoreColSort
    Call Form_Resize
    If vsExist.Rows > 1 Then
        mnuViewFind.Enabled = True
        tbr.Buttons.Item("Go").Enabled = True
        
        If tbsClass.SelectedItem.Index = 1 Then
            mnuEdit_EditItem.Enabled = True
            tbr.Buttons.Item("Edit").Enabled = True
        End If
    Else
        mnuViewFind.Enabled = False
        tbr.Buttons.Item("Go").Enabled = False
        
        If InStr(mstrPrivs, "����������Ŀ") = 0 And InStr(mstrPrivs, "ɾ��������Ŀ") > 0 Then
            mnuEdit_EditItem.Enabled = False
            tbr.Buttons.Item("Edit").Enabled = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String, blnByDept As Boolean, blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    On Error GoTo errH
    blnByDept = mnuViewByDept(1).Checked
    cbr.Bands(2).Caption = IIf(blnByDept, "���˿���", "���˲���")
    
    '��Ȩ����ʾ��������۲��Ҷ�Ӧ���ٴ�����,סԺ������סԺ��ͬ
    cboDept.Clear
    If InStr(mstrPrivs, "���в���") > 0 Then cboDept.AddItem IIf(blnByDept, "���п���", "���в���")
    
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    'by lesfeng 2010-03-08 �����Ż�
    strSQL = _
         " Select A.ID,A.����,A.����" & _
         " From ���ű� A,��������˵�� B" & _
         " Where B.����ID = A.ID And B.������� IN(1,2,3) And B.�������� IN([1])" & _
         " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
         IIf(blnLimitUnit, " And A.ID In (" & strUnitIDs & ")", "") & _
         " And (A.վ��=[2] Or A.վ�� is Null)" & _
         " Order by A.����"
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(blnByDept, "�ٴ�", "����"), gstrNodeNo)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
            If UserInfo.����ID = rsTmp!ID Then cboDept.ListIndex = cboDept.NewIndex
            
            rsTmp.MoveNext
        Next
        If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    ElseIf InStr(mstrPrivs, "���в���") > 0 Then
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
    Dim objItem As ListItem, strSQL As String
    Dim i As Long, j As Long, strCount As String
    Dim blnByDept As Boolean
    
    On Error GoTo errH
    
    Call zlCommFun.ShowFlash("���ڶ�ȡסԺ�����嵥,���Ժ� ...", Me)
    DoEvents
    blnByDept = mnuViewByDept(1).Checked
    Me.Refresh
    
    mintBedLen = GetMaxBedLen(mlngDeptID, blnByDept)
    
    If tbs.SelectedItem.Index = 1 Then
        If blnByDept Then
            strSQL = strSQL & IIf(mlngDeptID > 0, " And E.����ID=[1]", "")
        Else
            strSQL = strSQL & IIf(mlngDeptID > 0, " And E.����ID=[1]", "")
        End If
    Else
        If blnByDept Then
            strSQL = strSQL & IIf(mlngDeptID > 0, " And B.��Ժ����ID=[1]", "")
        Else
            strSQL = strSQL & IIf(mlngDeptID > 0, " And B.��ǰ����ID=[1]", "")
        End If
    End If
    
    If tbs.SelectedItem.Index = 1 Then
        '��ǰ��Ժ�Ĳ���
        '58842,������,2013-02-25,��Ժ���˶�ȡ(����Ժ�����ж�ȡ)
        If mnuViewFindPati.Checked = False Then
            strSQL = _
                "Select A.����ID,B.��ҳID,A.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.ҽ�Ƹ��ʽ," & _
                " B.��Ժ����,B.��Ժ����,LPAD(B.��Ժ����," & mintBedLen & ",' ') as ����," & _
                " C.���� as ��ǰ����,B.����,D.���� ҽ������,B.��������,B.״̬" & _
                " From ������Ϣ A,������ҳ B,���ű� C,������� D,��Ժ���� E" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��Ժ����ID=C.ID And A.����ID=E.����ID " & strSQL & _
                " And Nvl(B.��ҳID,0)<>0 AND B.���� Is Not Null And B.����=D.���" & _
                IIf(mlngDeptID = 0, " Order by A.סԺ�� Desc", " Order by ����")
        Else
             strSQL = _
                "Select A.����ID,B.��ҳID,A.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.ҽ�Ƹ��ʽ," & _
                " B.��Ժ����,B.��Ժ����,LPAD(B.��Ժ����," & mintBedLen & ",' ') as ����," & _
                " C.���� as ��ǰ����,B.����,D.���� ҽ������,B.��������,B.״̬" & _
                " From ������Ϣ A,������ҳ B,���ű� C,������� D,��Ժ���� E" & _
                " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��Ժ����ID=C.ID And A.����ID=E.����ID " & strSQL & _
                " And Nvl(B.��ҳID,0)<>0 AND B.���� Is Not Null And B.����=D.���" & _
                " And NOT Exists(Select D.����ID,D.��ҳID from ����������Ŀ D WHERE B.����ID=D.����Id and B.��ҳid=D.��ҳid)" & _
                IIf(mlngDeptID = 0, " Order by A.סԺ�� Desc", " Order by ����")
           
        End If
    ElseIf tbs.SelectedItem.Index = 2 Then
        '���ڼ��Ժ�Ĳ���
        If mnuViewFindPati.Checked = False Then
            strSQL = _
                "Select A.����ID,B.��ҳID,A.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.ҽ�Ƹ��ʽ," & _
                " B.��Ժ����,B.��Ժ����,LPAD(B.��Ժ����," & mintBedLen & ",' ') as ����," & _
                " C.���� as ��ǰ����,B.����,D.���� ҽ������,B.��������,B.״̬" & _
                " From ������Ϣ A,������ҳ B,���ű� C,������� D" & _
                " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=C.ID" & strSQL & _
                " And B.���� Is Not Null And B.����=D.��� AND B.��Ժ����<=[3]" & _
                " And B.��Ժ���� Between [2] And [3]" & _
                IIf(mlngDeptID = 0, " Order by A.סԺ�� Desc", " Order by ����")
        Else
            strSQL = _
                "Select A.����ID,B.��ҳID,A.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.ҽ�Ƹ��ʽ," & _
                " B.��Ժ����,B.��Ժ����,LPAD(B.��Ժ����," & mintBedLen & ",' ') as ����," & _
                " C.���� as ��ǰ����,B.����,D.���� ҽ������,B.��������,B.״̬" & _
                " From ������Ϣ A,������ҳ B,���ű� C,������� D" & _
                " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.��Ժ����ID=C.ID" & strSQL & _
                " And NOT Exists(Select D.����ID,D.��ҳID from ����������Ŀ D WHERE B.����ID=D.����Id and B.��ҳid=D.��ҳid)" & _
                " And B.���� Is Not Null And B.����=D.��� AND B.��Ժ����<=[3]" & _
                " And B.��Ժ���� Between [2] And [3]" & _
                IIf(mlngDeptID = 0, " Order by A.סԺ�� Desc", " Order by ����")
        End If
    End If
    
    mdtBegin = CDate(Format(mdtBegin, "yyyy-MM-dd 00:00:00"))
    mdtEnd = CDate(Format(mdtEnd, "yyyy-MM-dd 23:59:59"))
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, mdtBegin, mdtEnd)
  
    lvw.ListItems.Clear
    
    If Not mrsPati.EOF Then
        For i = 1 To mrsPati.RecordCount
            If IIf(IsNull(mrsPati!��������), 0, mrsPati!��������) = 0 Then
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!����ID & "_" & mrsPati!��ҳID & "_" & mrsPati!����, mrsPati!����, 1, 1)
            Else
                Set objItem = lvw.ListItems.Add(, "_" & mrsPati!����ID & "_" & mrsPati!��ҳID & "_" & mrsPati!����, mrsPati!����, 2, 2)
            End If
            objItem.SubItems(1) = IIf(IsNull(mrsPati!סԺ��), "", mrsPati!סԺ��)
            objItem.SubItems(2) = IIf(IsNull(mrsPati!����) And mrsPati!״̬ = 0, "��ͥ", Nvl(mrsPati!����, " "))
            objItem.SubItems(3) = IIf(IsNull(mrsPati!�Ա�), "", mrsPati!�Ա�)
            objItem.SubItems(4) = IIf(IsNull(mrsPati!����), "", mrsPati!����)
            objItem.SubItems(5) = Format(mrsPati!��Ժ����, "yyyy-MM-dd")
            objItem.SubItems(6) = Format(IIf(IsNull(mrsPati!��Ժ����), "", mrsPati!��Ժ����), "yyyy-MM-dd")
            objItem.SubItems(7) = IIf(IsNull(mrsPati!��ǰ����), "", mrsPati!��ǰ����)
            objItem.SubItems(8) = mrsPati!��ҳID
            objItem.SubItems(9) = Nvl(mrsPati!ҽ������)
            objItem.Tag = mrsPati!����ID
            objItem.ListSubItems(1).Tag = Nvl(mrsPati!״̬)
            If objItem.Tag = mstrPrePati Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
            
            If InStr(strCount & ",", "," & mrsPati!����ID & ",") = 0 Then strCount = strCount & "," & mrsPati!����ID
            mrsPati.MoveNext
        Next
        
        lbl_s.Tag = UBound(Split(Mid(strCount, 2), ",")) + 1
        If tbs.SelectedItem.Index = 1 Then
            lbl_s.Caption = " ��ǰ��Ժ�Ĳ���,����:" & Val(lbl_s.Tag)
        ElseIf tbs.SelectedItem.Index = 2 Then
            lbl_s.Caption = " ʱ��:" & Format(mdtBegin, "yyyy-MM-dd") & "��" & Format(mdtEnd, "yyyy-MM-dd") & ",����:" & Val(lbl_s.Tag)
        End If
        
        Me.Refresh
        mstrPrePati = ""
    Else
        lbl_s.Tag = ""
        stbThis.Panels(2).Text = ""
        mstrPrePati = ""
        If tbs.SelectedItem.Index = 1 Then
            lbl_s.Caption = " ��ǰ��Ժ�Ĳ���,����:0"
        ElseIf tbs.SelectedItem.Index = 2 Then
            lbl_s.Caption = " ʱ��:" & Format(mdtBegin, "yyyy-MM-dd") & "��" & Format(mdtEnd, "yyyy-MM-dd") & ",����:0"
        End If
    End If
    Call zlCommFun.StopFlash
    
    If lvw.ListItems.Count > 0 Then
        Set lvw.SelectedItem = lvw.ListItems.Item(1)
        If tbs.SelectedItem.Index = 1 Then
            mnuEdit_EditItem.Enabled = True
            tbr.Buttons.Item("Edit").Enabled = True
        End If
        mnuViewGo.Enabled = True
    Else
        mnuEdit_EditItem.Enabled = False
        tbr.Buttons.Item("Edit").Enabled = False
        mnuViewFind.Enabled = False
        tbr.Buttons.Item("Go").Enabled = False
        mnuViewGo.Enabled = False
        vsExist.Rows = 1
        Set mrsExistItem = Nothing
        For i = tbsClass.Tabs.Count To 2 Step -1
            tbsClass.Tabs.Remove i
        Next
    End If
    If Not lvw.SelectedItem Is Nothing Then Call lvw_ItemClick(lvw.SelectedItem)
    Exit Function
errH:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Call zlCommFun.ShowFlash("���ڶ�ȡסԺ�����嵥,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMaxBedLen(Optional lng����ID As Long, Optional bln���� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
'      blnռ��=�Ƿ�ֻ�ܱ�ռ�õĴ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln���� Or lng����ID = 0 Then
        strSQL = "Select Nvl(Max(Lengthb(����)),0) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=[1]")
    Else
        strSQL = "Select Nvl(Max(Lengthb(����)),0) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=[1]")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lng����ID As Long, lng��ҳID As Long, lng���� As Long
    Dim arrTmp As Variant
    
    If Item.Key = mstrPrePati Then Exit Sub
    
    stbThis.Panels(2).Text = "��" & Val(lbl_s.Tag) & "������,��ǰ:" & Item.Text & ",סԺ��:" & Item.SubItems(1)
        
    arrTmp = Split(Mid(Item.Key, 2), "_")
    lng����ID = Val(arrTmp(0))
    lng��ҳID = Val(arrTmp(1))
    lng���� = Val(arrTmp(2))
    
    Call ReadExistsItem(lng����ID, lng��ҳID, lng����)

    mstrPrePati = Item.Key
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
    lvw.ListItems("_" & mrsPati!����ID & "_" & mrsPati!��ҳID & "_" & mrsPati!����).Selected = True
    lvw.SelectedItem.EnsureVisible
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub tbsClass_Click()
    If tbsClass.SelectedItem.Index <> 1 Then
        mrsExistItem.Filter = "���='" & tbsClass.SelectedItem.Tag & "'"
    Else
        mrsExistItem.Filter = 0
    End If
    Set vsExist.DataSource = mrsExistItem
    If tbsClass.SelectedItem.Index <> 1 Then
        vsExist.ColHidden(ColNum.���) = True
    Else
        vsExist.ColHidden(ColNum.���) = False
    End If
    
    If InStr("�в�ҩ,�г�ҩ,����ҩ,����", tbsClass.SelectedItem.Tag) = 0 Then
        vsExist.ColHidden(ColNum.����) = True
    Else
        vsExist.ColHidden(ColNum.����) = False
    End If
    
    vsExist.Tag = tbsClass.SelectedItem.Index
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.vsExist.Rows = 1 Then Exit Sub
    
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsExist
    
    objPrint.Title.Text = lvw.SelectedItem.SubItems(1) & "-" & lvw.SelectedItem.Text & "���������Ŀ�嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim str����ID As String, str��ҳID As String, strסԺ�� As String
    Dim blnByDept As Boolean
    
    blnByDept = mnuViewByDept(1).Checked
    If Not lvw.SelectedItem Is Nothing Then
        str����ID = Val(lvw.SelectedItem.Tag)
        strסԺ�� = Val(lvw.SelectedItem.SubItems(1))
        str��ҳID = Val(lvw.SelectedItem.SubItems(8))
        
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����ID=" & str����ID, "��ҳID=" & str��ҳID, "����=" & IIf(blnByDept, 0, mlngDeptID), "���˿���=" & IIf(blnByDept, mlngDeptID, 0), "סԺ��=" & strסԺ��)
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            "����=" & IIf(blnByDept, 0, mlngDeptID), "���˿���=" & IIf(blnByDept, mlngDeptID, 0))
    End If
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

