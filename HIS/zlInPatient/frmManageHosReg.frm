VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageHosReg 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "������Ժ����"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10980
   Icon            =   "frmManageHosReg.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picFind 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3870
      ScaleHeight     =   420
      ScaleWidth      =   3705
      TabIndex        =   10
      Top             =   810
      Width           =   3705
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   360
         Left            =   570
         TabIndex        =   11
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmManageHosReg.frx":058A
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   2
         ShowSortName    =   -1  'True
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   480
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   3795
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4320
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   840
      Width           =   45
   End
   Begin VB.ComboBox cboNodeList 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   645
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   2430
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6765
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageHosReg.frx":066D
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "PatiColor"
            Object.Tag             =   "PatiColor"
            Object.ToolTipText     =   "������ɫ˵��"
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
      TabIndex        =   4
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   10980
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   8355
      NewRow1         =   0   'False
      Child2          =   "chkOnly"
      MinWidth2       =   1200
      MinHeight2      =   300
      Width2          =   1065
      NewRow2         =   0   'False
      Begin VB.CheckBox chkOnly 
         Caption         =   "ֻ��ʾ��������ԤԼ"
         Height          =   300
         Left            =   8550
         TabIndex        =   7
         Top             =   240
         Width           =   2340
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
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
            NumButtons      =   20
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
               Caption         =   "��Ժ"
               Key             =   "Add"
               Description     =   "��Ժ"
               Object.ToolTipText     =   "��סԺ���˽�����Ժ�Ǽ�"
               Object.Tag             =   "��Ժ"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Keep"
               Description     =   "����"
               Object.ToolTipText     =   "�����۲��˽��еǼ�"
               Object.Tag             =   "����"
               ImageKey        =   "Keep"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "OutKeep"
                     Object.Tag             =   "�������۵Ǽ�"
                     Text            =   "�������۵Ǽ�"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "InKeep"
                     Object.Tag             =   "סԺ���۵Ǽ�"
                     Text            =   "סԺ���۵Ǽ�"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Keep_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ԤԼ"
               Key             =   "PreAdd"
               Description     =   "ԤԼ"
               Object.ToolTipText     =   "ԤԼ��Ժ�Ǽ�"
               Object.Tag             =   "ԤԼ"
               ImageKey        =   "PreAdd"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Confirm"
               Description     =   "����"
               Object.ToolTipText     =   "ԤԼ��Ժ����"
               Object.Tag             =   "����"
               ImageKey        =   "Confirm"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Confirm0"
                     Text            =   "����ΪסԺ����"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Confirm1"
                     Text            =   "����Ϊ��������"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Confirm2"
                     Text            =   "����ΪסԺ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Confirm_"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ĵ�ǰ��Ժ�ǼǼ�¼"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "Del"
               Description     =   "ȡ��"
               Object.ToolTipText     =   "ȡ����ǰ��Ժ�ǼǼ�¼"
               Object.Tag             =   "ȡ��"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ��Ժ�ǼǼ�¼"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "���������������������Ĳ���"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������Ĳ�����"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
               Key             =   "FamilySplit"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.TreeView tvwDist_s 
      Height          =   4290
      Left            =   0
      TabIndex        =   1
      Top             =   1575
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   7567
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   420
      Top             =   2565
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
            Picture         =   "frmManageHosReg.frx":0F01
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":111B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":1335
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":154F
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":1769
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":1983
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":20FD
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2317
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2531
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":274B
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2965
            Key             =   "Keep"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":2B7F
            Key             =   "PreAdd"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":3279
            Key             =   "Confirm"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":3973
            Key             =   "Family"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1005
      Top             =   2565
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
            Picture         =   "frmManageHosReg.frx":A1D5
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":A3EF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":A609
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":A823
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":AA3D
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":AC57
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":B3D1
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":B5EB
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":B805
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":BA1F
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":BC39
            Key             =   "Keep"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":BE53
            Key             =   "PreAdd"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":C54D
            Key             =   "Confirm"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageHosReg.frx":CC47
            Key             =   "Family"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1650
      Top             =   2565
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
            Picture         =   "frmManageHosReg.frx":134A9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsType 
      Height          =   1065
      Left            =   30
      TabIndex        =   0
      Top             =   1200
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1879
      TabWidthStyle   =   2
      TabFixedWidth   =   2646
      TabFixedHeight  =   563
      HotTracking     =   -1  'True
      TabMinWidth     =   1235
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȷ�ϵǼ�(&1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ԤԼ�Ǽ�(&2)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   4530
      Left            =   3870
      TabIndex        =   2
      Top             =   1290
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7990
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmManageHosReg.frx":13603
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblNode 
      AutoSize        =   -1  'True
      Caption         =   "վ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   9
      Top             =   900
      Width           =   480
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
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_PrintMed 
         Caption         =   "��ӡ����(&M)"
      End
      Begin VB.Menu mnuFile_PrintWristlet 
         Caption         =   "��ӡ���(&W)"
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
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Add 
         Caption         =   "������Ժ�Ǽ�(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditOutKeep 
         Caption         =   "�������۵Ǽ�(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEditInKeep 
         Caption         =   "סԺ���۵Ǽ�(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPreAdd 
         Caption         =   "ԤԼ��Ժ�Ǽ�(&P)"
      End
      Begin VB.Menu mnuEditConfirm 
         Caption         =   "ԤԼ��Ժ����(&C)"
         Begin VB.Menu mnuEditConfirmType 
            Caption         =   "����ΪסԺ����(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuEditConfirmType 
            Caption         =   "����Ϊ��������(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuEditConfirmType 
            Caption         =   "����ΪסԺ����(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "�޸ĵǼ�(&M)"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "ȡ���Ǽ�(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditToKeep 
         Caption         =   "��Ϊ����(&K)"
      End
      Begin VB.Menu mnuEditToIn 
         Caption         =   "סԺ����תΪסԺ(&P)"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "���ĵǼ�(&V)"
      End
      Begin VB.Menu mnuEdit_Surety 
         Caption         =   "������Ϣ(&B)"
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
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInBed 
         Caption         =   "��ʾ��ס����(&I)"
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
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColmunSet 
         Caption         =   "�Զ�����ʾ��(&C)"
      End
      Begin VB.Menu mnuView_6 
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
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "ˢ�·�ʽ(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "������Ҫˢ������(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "��������ʾ�Ƿ�ˢ��(&2)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Index           =   2
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
Attribute VB_Name = "frmManageHosReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsPati As ADODB.Recordset
Private mblnMax As Boolean, mblnUnload As Boolean
Private mblnDown As Boolean, mblnGo As Boolean
Private mstrFilter As String, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mintBedLen As Integer '������󳤶�
Private mcllFilterA As Collection
Private mblnPassShowCard As Boolean '�����Ƿ�������ʾ
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��
Private mstrHead As String
Private Enum PATIVSF_COLUMN
    COL_�������� = 0
    COL_�Ǽ����� = 1
    COL_����ID = 2
    COL_����� = 3
    COL_סԺ�� = 4
    COL_���ۺ� = 5
    COL_���￨ = 6
    COL_���� = 7
    COL_���� = 8
    COL_�Ա� = 9
    COL_���� = 10
    COL_�ѱ� = 11
    COL_ҽ�Ƹ��ʽ = 12
    COL_ҽ���� = 13
    COL_���� = 14
    COL_��Ժʱ�� = 15
    COL_��Ժ���� = 16
    
    COL_��Ժ���� = 17
    COL_����ȼ� = 18
    COL_���� = 19
    Col_��Ժ���� = 20
    COL_��Ժ��ʽ = 21
    COL_סԺĿ�� = 22
    COL_�������� = 23
    COL_���� = 24
    COL_���� = 25
    COL_ѧ�� = 26
    COL_ְҵ = 27
    COL_��� = 28
    COL_���֤�� = 29
    COL_�ֻ��� = 30
    COL_���� = 31
    COL_������λ = 32
    COL_��ͥ��ַ = 33
    COL_��ͥ�绰 = 34
    COL_������� = 35
    COL_��ע = 36
    COL_�Ǽ�Ա = 37
    COL_״̬ = 38
    COL_��ҳID = 39
    COL_�������� = 40
End Enum

'by lesfeng 2010-1-11 �����Ż�
Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:
    '����:
    '����:
    '����:lesfeng
    '����:2010-01-11 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilterA = New Collection
    mcllFilterA.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "��Ժ����"
    mcllFilterA.Add Array("", ""), "סԺ��"
    '����17122 by lesfeng 2010-02-02
    mcllFilterA.Add "", "��������"
    mcllFilterA.Add "", "�Ǽ���"
    mcllFilterA.Add "", "�����"
    mstrFilter = ""
End Sub

Private Sub cboNodeList_Click()
    Call InitUnits
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub chkOnly_Click()
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub mnuEdit_FamilyAdd_Click()
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

Private Sub mnuEdit_Surety_Click()
    '56964:������,2013-04-23
    Dim lng����ID As Long, lngRow As Long
    Dim bln��Ժ���� As Boolean
    
    lngRow = mshPati.Row
    
    If lngRow >= mshPati.FixedRows And lngRow < mshPati.Rows Then
        lng����ID = Val(mshPati.TextMatrix(lngRow, GetColNum("����ID")))
    Else
        lng����ID = 0
    End If

    frmSurety.mlng����ID = lng����ID
    frmSurety.mbln��Ժ���� = True
    frmSurety.mstrPrivs = mstrPrivs
    frmSurety.Show 1, Me
End Sub

Private Sub mnuEditConfirmType_Click(Index As Integer)
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
        MsgBox "û��ԤԼ�Ǽǿ��Խ��ա�", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 2 '����ԤԼ
    frmHosReg.mbytKind = Index '0-����ԤԼ,1-��������,2-סԺ����
    frmHosReg.mbytInState = 0 '��Ϊ����
    frmHosReg.mlng����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    frmHosReg.mlng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditInKeep_Click()
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 0
    frmHosReg.mbytKind = 2
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 1 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditOutKeep_Click()
    On Error Resume Next
    Err.Clear
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 0
    frmHosReg.mbytKind = 1
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 1 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditPreAdd_Click()
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 1 'ԤԼ�Ǽ�
    frmHosReg.mbytKind = 0 '���ṩ���۵�ԤԼ
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 2 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEditToIn_Click()
'��סԺ���۲���תΪסԺ����
    Dim lng����ID As Long, lng��ҳID As Long, intRow As Long
    Dim strסԺ�� As String, str���� As String
    Dim strSQL As String, strNote As String
    Dim lng���� As Long
    Dim rsTemp As New ADODB.Recordset
    
    intRow = mshPati.Row
    lng���� = GetColNum("��������")
    If Val(mshPati.TextMatrix(intRow, lng����)) <> 2 Then Exit Sub
        
        
    lng����ID = Val(mshPati.TextMatrix(intRow, GetColNum("����ID")))
    lng��ҳID = Val(mshPati.TextMatrix(intRow, GetColNum("��ҳID")))
    strסԺ�� = mshPati.TextMatrix(intRow, GetColNum("סԺ��"))
            
    strSQL = "Select Nvl(״̬,0) ״̬ From ������ҳ Where ����ID=[1] And ��ҳID=[2] And ��������=2"
    On Error GoTo errH
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        If rsTemp!״̬ = 1 Then
            MsgBox "���˵�ǰ��δ���,����תΪסԺ���ˡ����Ƚ�������ƺ����ԡ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf rsTemp!״̬ = 2 Then
            MsgBox "���˵�ǰ����ת��,����תΪסԺ���ˡ����Ƚ�����ת�ƻ�ȡ��ת�ƺ����ԡ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("ȷʵҪ����סԺ���۲���תΪסԺ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '60500:������,2013-05-09,���۵Ǽ�û��ȷ��סԺ�ţ�תΪסԺ���������ʹ��ͳһסԺ�ţ�Ӧ�ñ��ֺ�֮ǰסԺһ��
    If strסԺ�� = "" And gblnÿ��סԺ��סԺ�� = False Then
        strSQL = " SELECT Nvl(a.סԺ��," & vbNewLine & _
            "            (SELECT סԺ��" & vbNewLine & _
            "             FROM ������ҳ" & vbNewLine & _
            "             WHERE ����id = a.����id AND" & vbNewLine & _
            "                   ��ҳid = (SELECT MAX(��ҳid) FROM ������ҳ WHERE ����id = a.����id AND סԺ�� IS NOT NULL))) סԺ��" & vbNewLine & _
            " FROM ������Ϣ a" & vbNewLine & _
            " WHERE ����id = [1]"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If Not rsTemp.EOF Then
            strסԺ�� = NVL(rsTemp!סԺ��)
        End If
    End If
    'û��סԺ�������һ��
    If strסԺ�� = "" Or gblnÿ��סԺ��סԺ�� Then
        strסԺ�� = zlDatabase.GetNextNo(2)
        str���� = mshPati.TextMatrix(intRow, GetColNum("����"))
        strNote = "�����۲��� " & str���� & " תΪסԺ����֮ǰ������Ϊ�ò���ȷ��һ��סԺ�š�"
        If Not frmInput.InputVal(Me, "סԺ��", strNote, strסԺ��, 1, 10, False, InStr(mstrPrivs, ";�޸�סԺ��;") <> 0) Then Exit Sub
    End If
        
    strSQL = "ZL_���˱䶯��¼_תסԺ(" & lng����ID & "," & lng��ҳID & "," & strסԺ�� & ")"
    On Error GoTo errH
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    '��ֱ�Ӵ���
    mshPati.TextMatrix(intRow, lng����) = "0"
    mshPati.TextMatrix(intRow, GetColNum("�Ǽ�����")) = "סԺ����"
    mshPati.TextMatrix(intRow, GetColNum("סԺ��")) = strסԺ��
    
    Call mshPati_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mnuEditToKeep_Click()
'��סԺ���˳���ΪסԺ���۲���
    Dim intRow As Integer, i As Integer
    Dim lng����ID As Long, lng��ҳID As Long, int���סԺ�� As Integer
    Dim strSQL As String
    
    intRow = mshPati.Row
    
    If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("����ID"))) Then
        MsgBox "û�в��˿��Գ�Ϊ���۲��ˣ�", vbExclamation, gstrSysName: Exit Sub
    End If
    
    lng����ID = Val(mshPati.TextMatrix(intRow, GetColNum("����ID")))
    lng��ҳID = Val(mshPati.TextMatrix(intRow, GetColNum("��ҳID")))
            
    'ȥ����ҽ������ƥ����
    
    If MsgBox("ȷʵҪ������""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """����ΪסԺ���۲�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If lng��ҳID = 1 Then
        If MsgBox("ͬʱ����ò��˵�סԺ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then int���סԺ�� = 1
    End If
    strSQL = "zl_��Ժ������ҳ_DELETE(" & lng����ID & "," & lng��ҳID & ",1," & int���סԺ�� & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    
    '��ֱ�Ӵ���
    mshPati.TextMatrix(mshPati.Row, GetColNum("��������")) = "2"
    mshPati.TextMatrix(mshPati.Row, GetColNum("�Ǽ�����")) = "סԺ����"
    If int���סԺ�� = 1 Then mshPati.TextMatrix(mshPati.Row, GetColNum("סԺ��")) = ""
    
    Call mshPati_EnterCell
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFile_PrintMed_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    
    lng����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    lng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, 2)
    End If
End Sub

Private Sub mnuFile_PrintWristlet_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    
    lng����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    lng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, 2)
    End If
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    frmSetPar.mlngModul = mlngModul
    frmSetPar.mstrPrivs = mstrPrivs
    frmSetPar.Show 1, Me
End Sub

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strTmp As String, str����ID As String
    
    str����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    If Left(tvwDist_s.Tag, 1) = "U" Then
        strTmp = "����="
    Else    'δѡ��ʱ,���ɿ���
        strTmp = "���˿���="
    End If
    
    If str����ID <> "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            strTmp & Mid(tvwDist_s.Tag, 2), _
            "����ID=" & str����ID, _
            "סԺ��=" & mshPati.TextMatrix(mshPati.Row, GetColNum("סԺ��")))
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
            strTmp & Mid(tvwDist_s.Tag, 2))
    End If
End Sub

Private Sub mnuViewByDept_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewByDept.Count - 1
        mnuViewByDept(i).Checked = (i = Index)
    Next
    Call tbsType_Click
End Sub

Private Sub mnuViewColmunSet_Click()
    Call frmColumnSet.ShowMe(Me, mshPati, mstrHead)

End Sub

Private Sub mnuViewFilter_Click()
    frmHosRegFilter.Show 1, Me
    If gblnOK Then
        mstrFilter = frmHosRegFilter.mstrFilter
        'by lesfeng 2010-1-11 �����Ż�
        Set mcllFilterA = frmHosRegFilter.mcllFilter
        If mcllFilterA("�����") <> "" Then tvwDist_s.Nodes(1).root.Selected = True
        InitNode
        mnuViewreFlash_Click
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmHosRegFind.Show 1, Me
    If gblnOK Then Call SeekPati(frmHosRegFind.optHead)
End Sub

Private Sub mnuViewInBed_Click()
    mnuViewInBed.Checked = Not mnuViewInBed.Checked
    Call ShowPatis(mstrFilter)
End Sub

Private Sub mnuViewToolDist_Click()
    mnuViewToolDist.Checked = Not mnuViewToolDist.Checked
    tbsType.Visible = mnuViewToolDist.Checked
    tvwDist_s.Visible = mnuViewToolDist.Checked
    pic.Visible = tvwDist_s.Visible
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub mshPati_Click()
    If tbsType.SelectedItem.Index = 1 Then Exit Sub
    If mshPati.RowSel = 0 Then Exit Sub
    If (mshPati.TextMatrix(mshPati.RowSel, GetColNum("�Ǽ�����")) = "סԺ����" And InStr(mstrPrivs, "����סԺԤԼ") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("�Ǽ�����")) = "��������" And InStr(mstrPrivs, "������������ԤԼ") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("�Ǽ�����")) = "סԺ����" And InStr(mstrPrivs, "����סԺ����ԤԼ") = 0) Then
        
        mnuEditConfirm.Enabled = False
        tbr.Buttons("Confirm").Enabled = False
    End If
End Sub

Private Sub mshPati_DblClick()
    If mshPati.MouseRow = 0 Or mshPati.TextMatrix(mshPati.MouseRow, GetColNum("����ID")) = "" Then Exit Sub
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
    If mblnUnload Then
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
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, Curdate As Date
    Dim lngTmp As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    '�շ�����ģ��Ȩ��
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    On Error GoTo errHandle
    mstrHead = "��������,1,0|�Ǽ�����,1,1150|����ID,1,1050|�����,1,1050|סԺ��,1,1050|���ۺ�,1,1050|���￨,1,1150|����,1,800|����,1,1100|�Ա�,1,800|����,1,800|�ѱ�,1,800|ҽ�Ƹ��ʽ,1,1500|" & _
            "ҽ����,1,1300|����,1,1800|��Ժʱ��,1,1300|��Ժ����,1,1850|��Ժ����,1,1850|����ȼ�,1,1150|����,4,800|" & _
            "��Ժ����,1,1150|��Ժ��ʽ,1,1150|סԺĿ��,1,1150|��������,1,1300|" & _
            "����,1,800|����,1,1300|ѧ��,1,800|ְҵ,1,1300|���,1,1050|���֤��,1,2300|�ֻ���,1,1500|����,1,800|" & _
            "������λ,1,2300|��ͥ��ַ,1,2300|��ͥ�绰,1,1500|�������,1,4300|��ע,1,2300|�Ǽ�Ա,1,1050|״̬,1,0|��ҳID,1,0|��������,1,1300|�Һ�ID,1,0|����,1,0"
    
    '80509:������,2014-12-09,��Ӳ��˲��ҡ�����
    If Not gobjSquare Is Nothing Then Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
    
    strSQL = "Select �������� From ҽ�ƿ���� where ����='���￨' and �Ƿ�̶�=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mblnPassShowCard = NVL(rsTemp!��������) <> ""
    End If
    'by lesfeng 2010-1-11 �����Ż�
    Call InitFilter
    
    mstrPrivs = ";" & gstrPrivs & ";"
    mlngModul = glngModul
    
    '�ָ����Բ����嵥����
    RestoreWinState Me, App.ProductName
    mnuViewInBed.Checked = zlDatabase.GetPara("��ʾ��ס����", glngSys, mlngModul, "0")
    'ˢ�·�ʽ
    lngTmp = zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, "1")
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = (i = lngTmp)
    Next
    lngTmp = zlDatabase.GetPara("��ʾ���˷�ʽ", glngSys, mlngModul, "0")
    For i = 0 To mnuViewByDept.UBound
        mnuViewByDept(i).Checked = (i = lngTmp)
    Next
    
    mblnUnload = False
    
    'Ȩ������
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    '��ʼ��վ���б�
    Call InitNode
    
    '�����Ǽ�
    If InStr(mstrPrivs, ";����Ǽ�;") = 0 Then '�������������ۣ�סԺ����
        mnuEdit_Add.Visible = False
        mnuEditInKeep.Visible = False
        mnuEditOutKeep.Visible = False
        mnuEdit_1.Visible = False
        
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Keep").Visible = False
        tbr.Buttons("Keep_").Visible = False
    End If
            
    'ԤԼ�ͽ���
    If InStr(mstrPrivs, ";ԤԼ�Ǽ�;") = 0 Then '���ṩ���۲���ԤԼ�Ǽ�
        mnuEditPreAdd.Visible = False
        tbr.Buttons("PreAdd").Visible = False
    End If
    If InStr(mstrPrivs, ";����ԤԼ;") = 0 Then '�������������ۣ�סԺ����
        mnuEditConfirm.Visible = False
        tbr.Buttons("Confirm").Visible = False
    Else
        If InStr(mstrPrivs, ";����סԺԤԼ;") = 0 And InStr(mstrPrivs, ";������������ԤԼ;") = 0 And InStr(mstrPrivs, ";����סԺ����ԤԼ;") = 0 Then
            mnuEditConfirm.Enabled = False
            tbr.Buttons("Confirm").Enabled = False
        Else
            If InStr(mstrPrivs, ";����סԺԤԼ;") = 0 Then
                mnuEditConfirmType(0).Visible = False
                tbr.Buttons("Confirm").ButtonMenus.Item(1).Visible = False
            End If
            
            If InStr(mstrPrivs, ";������������ԤԼ;") = 0 Then
                mnuEditConfirmType(1).Visible = False
                tbr.Buttons("Confirm").ButtonMenus.Item(2).Visible = False
            End If
            
            If InStr(mstrPrivs, ";����סԺ����ԤԼ;") = 0 Then
                mnuEditConfirmType(2).Visible = False
                tbr.Buttons("Confirm").ButtonMenus.Item(3).Visible = False
            End If
        End If
    End If
    
    If InStr(mstrPrivs, ";ԤԼ�Ǽ�;") = 0 _
        And InStr(mstrPrivs, ";����ԤԼ;") = 0 Then
        mnuEditPreAdd.Visible = False
        mnuEditConfirm.Visible = False
        mnuEdit_2.Visible = False
        tbr.Buttons("PreAdd").Visible = False
        tbr.Buttons("Confirm").Visible = False
        tbr.Buttons("Confirm_").Visible = False
    End If
                            
    '���۲�����Ȩ��:�����ǼǺ�ԤԼ�Ǽǵ�
    If InStr(mstrPrivs, ";סԺ���˵Ǽ�;") = 0 And InStr(mstrPrivs, ";�������۵Ǽ�;") = 0 And InStr(mstrPrivs, ";סԺ���۵Ǽ�;") = 0 Then
        mnuEdit_Add.Visible = False
        tbr.Buttons("Add").Visible = False
        mnuEditOutKeep.Visible = False
        tbr.Buttons("Keep").ButtonMenus("OutKeep").Visible = False
        mnuEditInKeep.Visible = False
        tbr.Buttons("Keep").ButtonMenus("InKeep").Visible = False
        mnuEditConfirm.Visible = False
        tbr.Buttons("Confirm").Visible = False
    Else
        If InStr(mstrPrivs, ";סԺ���˵Ǽ�;") = 0 Then
            mnuEdit_Add.Visible = False
            tbr.Buttons("Add").Visible = False
            mnuEditConfirmType(0).Visible = False
            tbr.Buttons("Confirm").ButtonMenus("Confirm0").Visible = False
        End If
        If InStr(mstrPrivs, ";�������۵Ǽ�;") = 0 Then
            mnuEditOutKeep.Visible = False
            tbr.Buttons("Keep").ButtonMenus("OutKeep").Visible = False
            mnuEditConfirmType(1).Visible = False
            tbr.Buttons("Confirm").ButtonMenus("Confirm1").Visible = False
        End If
        If InStr(mstrPrivs, ";סԺ���۵Ǽ�;") = 0 Then
            mnuEditInKeep.Visible = False
            tbr.Buttons("Keep").ButtonMenus("InKeep").Visible = False
            mnuEditConfirmType(2).Visible = False
            tbr.Buttons("Confirm").ButtonMenus("Confirm2").Visible = False
        End If
    End If
    If InStr(mstrPrivs, ";�������۵Ǽ�;") = 0 _
        And InStr(mstrPrivs, ";סԺ���۵Ǽ�;") = 0 Then
        mnuEdit_1.Visible = False
        tbr.Buttons("Keep").Visible = False
        tbr.Buttons("Keep_").Visible = False
    End If
                        
    '�޸�Ȩ��
    If InStr(mstrPrivs, ";����Ǽ�;") = 0 _
        And InStr(mstrPrivs, ";ԤԼ�Ǽ�;") = 0 _
        And InStr(mstrPrivs, ";����ԤԼ;") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
                        
    '������ȡ����Ժ,ȡ��ԤԼ�Ĺ���
    If InStr(mstrPrivs, ";ȡ����Ժ;") = 0 Then
        mnuEdit_Del.Visible = False
        mnuEditToKeep.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    
    'סԺ����תסԺ
    If InStr(mstrPrivs, ";סԺ����תסԺ;") = 0 Then
        mnuEditToIn.Visible = False
    End If
    Call tbsType_Click
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        tvwDist_s.width = 3780
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    DisW = IIf(tvwDist_s.Visible, tvwDist_s.width + pic.width, 0)
    
    pic.Visible = tvwDist_s.Visible
    lblNode.Visible = cboNodeList.Visible
    
    cboNodeList.Top = Me.ScaleTop + cbrH + 15
    lblNode.Top = cboNodeList.Top
    If cboNodeList.Height - lblNode.Height > 0 Then
        lblNode.Top = lblNode.Top + (cboNodeList.Height - lblNode.Height) \ 2
    End If
    
    With tbsType
        .Top = Me.ScaleTop + cbrH + 15 + IIf(cboNodeList.Visible, cboNodeList.Height + 100, 0)
        .Left = Me.ScaleLeft + 30
        .width = tvwDist_s.width - 45
    End With
    With tvwDist_s
        .Left = Me.ScaleLeft
        .Top = tbsType.Top + 330
        .Height = Me.ScaleHeight - staH - .Top
    End With
    With pic
        .Left = tvwDist_s.Left + tvwDist_s.width
        .Top = tvwDist_s.Top
        .Height = tvwDist_s.Height
    End With
    With picFind
        .Left = DisW
        .Top = Me.ScaleTop + cbrH
    End With
    With mshPati
        .Left = DisW
        .Top = picFind.Top + picFind.Height ' Me.ScaleTop + cbrH
        .Height = Me.ScaleHeight - cbrH - staH - picFind.Height
        .width = Me.ScaleWidth - DisW
    End With
    cboNodeList.width = tvwDist_s.width - 600
    mshPati.Redraw = True
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, lngTmp As Long
    
    mstrFilter = ""
    
    SaveWinState Me, App.ProductName
    zlDatabase.SetPara "��ʾ��ס����", mnuViewInBed.Checked, glngSys, mlngModul
    
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "ˢ�·�ʽ", lngTmp, glngSys, mlngModul
    
    '��ʾ���˷�ʽ
    lngTmp = 0
    For i = 0 To mnuViewByDept.UBound
        If mnuViewByDept(i).Checked Then
            lngTmp = i
            Exit For
        End If
    Next
    zlDatabase.SetPara "��ʾ���˷�ʽ", lngTmp, glngSys, mlngModul
    
    
    Unload frmHosRegFind
    Unload frmHosRegFilter
End Sub

Private Sub mnuEdit_Del_Click()
    Dim intRow As Integer, i As Integer
    Dim lng����ID As Long, lng��ҳID As Long, lng�Һ�ID As Long
    Dim strSQL As String, int���� As Integer
    Dim rsTmp As ADODB.Recordset
    Dim blnNotCommit As Boolean
    Dim blnTrans As Boolean
    intRow = mshPati.Row
    
    If Not IsNumeric(mshPati.TextMatrix(intRow, GetColNum("����ID"))) Then
        MsgBox "û�в��˿���ȡ���Ǽǡ�", vbExclamation, gstrSysName: Exit Sub
    End If
    
    lng����ID = mshPati.TextMatrix(intRow, GetColNum("����ID"))
    lng��ҳID = Val(mshPati.TextMatrix(intRow, GetColNum("��ҳID")))
    lng�Һ�ID = Val(mshPati.TextMatrix(intRow, GetColNum("�Һ�ID")))
    'ȥ����ҽ������ƥ����
    
    If MsgBox("ȷʵҪȡ������""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """�ĵǼ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '����:31635
    blnNotCommit = False
    int���� = 0
    On Error GoTo errH
    '����22073 by lesfeng 2010-08-02  ��֤�Ƿ���д���Ӳ���
    If GetCaseHistory(lng����ID, lng��ҳID) Then
        MsgBox "�Ѿ��Բ���""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """��д���Ӳ���������ȡ����Ժ��", vbExclamation, gstrSysName: Exit Sub
    End If
    
    Set rsTmp = GetMoneyInfo(lng����ID, , , 2)
    If Not rsTmp Is Nothing Then
        If NVL(rsTmp!Ԥ�����) <> 0 Then '����û��Ԥ�����з������
            If MsgBox("����""" & mshPati.TextMatrix(intRow, GetColNum("����")) & """��Ԥ����δ�ˣ��Ƿ�Ҫ��������ȡ����Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
        
    If tbsType.SelectedItem.Index = 1 Then
        'ҽ��(ȡ����Ժ:ʵ����ĳЩҽ����ִ�г�Ժ����)
        If isYBPati(lng����ID, , int����) Then
            If Not gclsInsure.ComeInDelSwap(lng����ID, lng��ҳID, int����) Then
                gcnOracle.RollbackTrans: blnTrans = False: Exit Sub
            End If
        End If
        '����:31635
        blnNotCommit = True
        
        strSQL = "zl_סԺһ�η���_Delete(" & lng����ID & "," & lng��ҳID & ")"
       
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    strSQL = "zl_��Ժ������ҳ_DELETE(" & lng����ID & "," & lng��ҳID & ",0," & IIf(gblnÿ��סԺ��סԺ��, "1", "0") & ")" '"��ҳID=0"��ʾԤԼ�Ǽ�
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    gcnOracle.CommitTrans: blnTrans = False
    
    If lng��ҳID = 0 Then
        '����ԤԼϵͳ�ӿ�{"�Һ�id_In": "�Һ�ID","״̬_In": "״̬" ---�ѽ��գ�δ���գ����˳�}
        Call Sys.NewSystemSvr("ԤԼ����", "��ס����סȡ��", "{""�Һ�id_In"": """ & lng�Һ�ID & """,""״̬_In"": ""���˳�""}", "")
    End If
     '����:31635
    If int���� > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ComeInDelSwap, True, int����)
    
    On Error GoTo 0
    
    '��ֱ�Ӵ���
    If mshPati.Rows > 2 Then
        mshPati.RemoveItem intRow
        Call SetMenu(True)
    Else
        With mshPati
            For i = 0 To .Cols - 1
                .TextMatrix(intRow, i) = ""
            Next
        End With
        Call SetMenu(False)
    End If
    
    If intRow <= mshPati.Rows - 1 Then
        mshPati.Row = intRow
    Else
        mshPati.Row = mshPati.Rows - 1
    End If
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    
    Call mshPati_EnterCell
    Call mshPati_Click
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
     '����:31635
    If int���� > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ComeInDelSwap, False, int����)
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Modi_Click()
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
        MsgBox "û�в�����Ϣ�����޸ģ�", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = tbsType.SelectedItem.Index - 1 '������ԤԼ
    frmHosReg.mbytKind = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��������")))
    frmHosReg.mbytInState = 1
    frmHosReg.mlng����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    frmHosReg.mlng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = 0
    frmHosReg.mbytKind = 0
    frmHosReg.mbytInState = 0
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOK And tbsType.SelectedItem.Index = 1 Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewreFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewreFlash_Click
        End If
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_View_Click()
    If Not IsNumeric(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))) Then
        MsgBox "û�в�����Ϣ���Բ鿴��", vbExclamation, gstrSysName: Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmHosReg.mstrPrivs = mstrPrivs
    frmHosReg.mlngModul = mlngModul
    frmHosReg.mbytMode = tbsType.SelectedItem.Index - 1 '������ԤԼ
    frmHosReg.mbytKind = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��������")))
    frmHosReg.mbytInState = 2
    frmHosReg.mlng����ID = mshPati.TextMatrix(mshPati.Row, GetColNum("����ID"))
    frmHosReg.mlng��ҳID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID")))
    frmHosReg.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    mshPati.Refresh
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewreFlash_Click()
    Call tbsType_Click
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

Private Sub mshPati_RowColChange()
    If tbsType.SelectedItem.Index = 1 Then Exit Sub
    If mshPati.RowSel = 0 Then Exit Sub
    If (mshPati.TextMatrix(mshPati.RowSel, GetColNum("�Ǽ�����")) = "סԺ����" And InStr(mstrPrivs, "����סԺԤԼ") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("�Ǽ�����")) = "��������" And InStr(mstrPrivs, "������������ԤԼ") = 0) Or _
        (mshPati.TextMatrix(mshPati.RowSel, GetColNum("�Ǽ�����")) = "סԺ����" And InStr(mstrPrivs, "����סԺ����ԤԼ") = 0) Then
        
        mnuEditConfirm.Enabled = False
        tbr.Buttons("Confirm").Enabled = False
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    Dim strTag As String
    
    blnCancel = False
    If objHisPati Is Nothing Then blnCancel = True
    If blnCancel = False Then
        If objHisPati.����ID = 0 Then blnCancel = True
    End If
    
    If tbsType.SelectedItem.Index = 1 Then
        strTag = "��Ժ����"
    Else
        strTag = "ԤԼ����"
    End If
            
    If blnCancel Then
        MsgBox "û���ҵ����������Ĳ��ˣ���ȷ��Ҫ���ҵĲ����Ƿ�����" & strTag & "��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ʼ�������б���Ѱ�Ҳ���
    If FindPatiInfo(objHisPati.����ID, objHisPati) = True Then Exit Sub
    '��ȡ������Ϣ����
    If GetPatiInfo(objHisPati.����ID, objHisPati) = False Then blnCancel = True: Exit Sub
    If mshPati.Enabled And mshPati.Visible Then mshPati.SetFocus
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strPati As String, vRect As RECT, strName As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim strTag As String
    Dim lng����ID As Long
    
    strName = Trim(PatiIdentify.Text)
    
    On Error GoTo ErrHand
    blnCancel = False
    If Not tvwDist_s.SelectedItem Is Nothing And mnuViewByDept(0).Checked = True Then
        PatiIdentify.���˲���ID = Val(Mid(tvwDist_s.SelectedItem.Key, 2))
    Else
        PatiIdentify.���˲���ID = 0
    End If
            
    If objCard.���� Like "*��*��*" And blnCard = False And strName <> "" And InStr("-*+/", Left(Trim(PatiIdentify.Text), 1)) = 0 Then
       
        If gblnSeekName = False Then '��������ģ������
            MsgBox "��ˢ��������[-����ID]��[+סԺ��]��[*�����]�ȷ�ʽ��ȡ���˵���Ϣ��", vbInformation, gstrSysName
            blnCancel = True
            Exit Sub
        End If
        
        If tbsType.SelectedItem.Index = 1 Then
            strIF = " And A.��ҳID=C.��ҳID And Nvl(C.��ҳID,0)<>0"
            strTag = "��Ժ����"
        Else
            strIF = " And Nvl(C.��ҳID,0)=0"
            strTag = "ԤԼ����"
        End If
    
        
        strPati = "Select 1 As ����id, a.����id As Id, a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.סԺ����, Trunc(c.��Ժ����, 'dd') As ��Ժ����, a.��������," & vbNewLine & _
            "       a.���֤��, a.��ͥ��ַ, a.������λ, c.��������" & vbNewLine & _
            " From ������Ϣ a, ������ҳ c" & vbNewLine & _
            " Where a.ͣ��ʱ�� Is Null And a.����id = c.����id " & strIF & " And c.��Ժ���� Is Null  And a.���� Like [1] " & _
            IIf(gintNameDays = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])") & " And Rownum < 101"
        strPati = strPati & " Order by ����ID,����,��Ժ���� Desc"
        
        vRect = zlControl.GetControlRect(PatiIdentify.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, PatiIdentify.Height, blnCancel, False, True, strName & "%", gintNameDays)
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!ID) = 0 Then
                blnCancel = True: Exit Sub
            Else '�Բ���ID��ȡ
                lng����ID = NVL(rsTmp!ID)
            End If
        Else 'ȡ��ѡ��
            If blnCancel = False Then
                MsgBox "û���ҵ����������Ĳ��ˣ���ȷ��Ҫ���ҵĲ����Ƿ�����" & strTag & "��", vbInformation, gstrSysName
            End If
            blnCancel = True: Exit Sub
        End If
        
        '��ʼ�������б���Ѱ�Ҳ���
        If FindPatiInfo(lng����ID, objCardData) = True Then blnFindPatied = True: blnCancel = True: Exit Sub
        '��ȡ������Ϣ����
        If GetPatiInfo(lng����ID, objCardData) = False Then
            blnCancel = True: Exit Sub
        Else
            blnFindPatied = True: blnCancel = True
        End If
        If mshPati.Enabled And mshPati.Visible Then mshPati.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FindPatiInfo(ByVal lngPatiID As Long, objCardData As zlIDKind.PatiInfor) As Boolean
'����:���ݲ���ID,��λ������
    Dim i As Long, blnFind As Boolean
    
    If lngPatiID = 0 Then Exit Function
    
    For i = 1 To mshPati.Rows - 1
        If Val(mshPati.TextMatrix(i, GetColNum("����ID"))) = lngPatiID Then
            blnFind = True
            Exit For
        End If
    Next i
    
    If objCardData Is Nothing Then
        Set objCardData = New zlIDKind.PatiInfor
    End If
    
    If blnFind = True Then
        mshPati.Row = i: mshPati.TopRow = i
        mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
        If mshPati.Enabled And mshPati.Visible Then mshPati.SetFocus
        objCardData.����ID = lngPatiID: objCardData.���� = mshPati.TextMatrix(i, GetColNum("����"))
        FindPatiInfo = True
    End If
End Function

Private Function GetPatiInfo(ByVal lngPatiID As Long, objCardData As zlIDKind.PatiInfor) As Boolean
    Dim i As Long, strSQL As String
    Dim strCard As String, strIF As String
    Dim strNodeNo As String
    Dim rsPati As ADODB.Recordset
    Dim strTag As String
    
    On Error GoTo errH
    '��ȡվ���
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
    
    strIF = " And A.����ID=[1]"
    If tbsType.SelectedItem.Index = 1 Then
        strIF = strIF & " And A.��ҳID=B.��ҳID And Nvl(B.��ҳID,0)<>0"
        strTag = "��Ժ����"
    Else
        strIF = strIF & " And Nvl(B.��ҳID,0)=0"
        strTag = "ԤԼ����"
    End If
    
    If mblnPassShowCard = True Then
        strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨,"
    Else
        strCard = "A.���￨�� as ���￨,"
    End If

    
    mintBedLen = GetMaxBedLen

    strSQL = _
        "Select ��������,Decode(B.��������,1,'��������',2,'סԺ����','סԺ����') as �Ǽ�����," & _
        " A.����ID, A.�����, B.סԺ��,B.���ۺ�," & strCard & "Decode(Nvl(B.״̬,0),1,NULL,LPad(B.��Ժ����," & mintBedLen & ", ' ')) as ����," & _
        " NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.�ѱ�,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,Nvl(A.ҽ����,F.��Ϣֵ) as ҽ����,X.���� as ����," & _
        " To_Char(B.��Ժ����,'YYYY-MM-DD HH24:MI:SS') as ��Ժʱ��," & _
        " C.���� as ��Ժ����,D.���� as ��Ժ����,E.���� as ����ȼ�,A.סԺ���� as ����,B.��Ժ����," & _
        " B.��Ժ��ʽ,B.סԺĿ��,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����," & _
        " A.ѧ��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.����״�� as ����,A.������λ,A.��ͥ��ַ, A.��ͥ�绰, g.�������, B.��ע,B.�Ǽ��� as �Ǽ�Ա,B.״̬,B.��ҳID," & _
        " Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������" & _
        " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,�շ���ĿĿ¼ E,������ҳ�ӱ� F," & _
        "       (Select distinct  ����ID, first_value(��¼) OVER (PARTITION BY ����ID ORDER BY ��¼���� DESC) AS �������" & vbNewLine & _
        "       From (SELECT a.����id,q.������� ��¼,q.��¼����" & vbNewLine & _
        "               FROM ������Ϣ a, ������ҳ b, ���˹Һż�¼ p, ������ϼ�¼ q" & vbNewLine & _
        "               Where a.����id = b.����id AND B.��Ժ���� is NULL AND b.����id=p.����id(+) And p.����id = q.����id(+)" & vbNewLine & _
        "               AND p.Id = q.��ҳid AND p.��¼����=1 and p.��¼״̬=1 and ��¼��Դ(+) = 3 AND �������(+) = 1" & vbNewLine & _
        "               AND ��ϴ���(+) = 1 " & strIF & ")) g,������� X" & _
        " Where A.����ID=B.����ID And B.��Ժ���� is NULL And B.��Ժ����ID=C.ID(+)" & _
        " And B.��Ժ����ID=D.ID " & IIf(cboNodeList.Visible, "And (d.վ��=" & strNodeNo & " Or d.վ�� Is Null)", "") & " And B.����ȼ�ID=E.ID(+)" & _
        " And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) And B.����ID = G.����ID(+)" & _
        " And F.��Ϣ��(+)='ҽ����' And B.����=X.���(+)" & strIF
    
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID)
    
    If rsPati.EOF Then
        MsgBox "û���ҵ����������Ĳ��ˣ���ȷ��Ҫ���ҵĲ����Ƿ�����" & strTag & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not tvwDist_s.SelectedItem Is Nothing Then
            For i = 1 To tvwDist_s.Nodes.Count
                If tvwDist_s.Nodes(i).Selected = True Then
                    tvwDist_s.Nodes(i).Selected = False
                End If
            Next i
            tvwDist_s.Tag = ""
            Set tvwDist_s.SelectedItem = Nothing
    End If
    
    mshPati.Clear
    mshPati.ClearStructure
    mshPati.Rows = 2
    
    Set mshPati.DataSource = rsPati
    Call setHeader(mstrHead)          '�����е�enter_cell���ѵ���SetMenu(false)
    If mnuViewInBed.Checked Then Call SetInBed
    stbThis.Panels(2) = "�� " & rsPati.RecordCount & " ������"
    Call SetMenu(True)
    
    mshPati_Click
    
    If objCardData Is Nothing Then Set objCardData = New zlIDKind.PatiInfor
    objCardData.����ID = lngPatiID: objCardData.���� = mshPati.TextMatrix(mshPati.Row, GetColNum("����"))
    
    Me.Refresh
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pic.Left + X < 1000 Or mshPati.width - X < 2000 Or mshPati.width - X < picFind.width Then Exit Sub
        pic.Left = pic.Left + X
        tbsType.width = tbsType.width + X
        tvwDist_s.width = tvwDist_s.width + X
        picFind.Left = picFind.Left + X
        mshPati.Left = mshPati.Left + X
        mshPati.width = mshPati.width - X
        cboNodeList.width = tvwDist_s.width - 600
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PatiColor" Then
        zlDatabase.ShowPatiColorTip Me
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
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
        Case "Keep"
            If Not Button.ButtonMenus("OutKeep").Visible _
                And Button.ButtonMenus("InKeep").Visible Then
                mnuEditInKeep_Click
            ElseIf Not Button.ButtonMenus("InKeep").Visible _
                And Button.ButtonMenus("OutKeep").Visible Then
                mnuEditOutKeep_Click
            End If
        Case "PreAdd"
            mnuEditPreAdd_Click
        Case "Confirm"
            '���ݲ������ʾ���ȱʡ�����ֽ���
            Select Case Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��������")))
                Case 0
                    If mnuEditConfirmType(0).Enabled And mnuEditConfirmType(0).Visible Then
                        Call mnuEditConfirmType_Click(0)
                    End If
                Case 1
                    If mnuEditConfirmType(1).Enabled And mnuEditConfirmType(1).Visible Then
                        Call mnuEditConfirmType_Click(1)
                    End If
                Case 2
                    If mnuEditConfirmType(2).Enabled And mnuEditConfirmType(2).Visible Then
                        Call mnuEditConfirmType_Click(2)
                    End If
            End Select
        Case "View"
            mnuEdit_View_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Family"
           Call mnuEdit_FamilyAdd_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "OutKeep"
            mnuEditOutKeep_Click
        Case "InKeep"
            mnuEditInKeep_Click
        Case "Confirm0"
            Call mnuEditConfirmType_Click(0)
        Case "Confirm1"
            Call mnuEditConfirmType_Click(1)
        Case "Confirm2"
            Call mnuEditConfirmType_Click(2)
        Case "FamilyAdd"
            Call mnuEdit_FamilyAdd_Click
        Case "FamilyView"
            Call mnuEdit_FamilyView_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub tbsType_Click()
    mnuViewInBed.Enabled = tbsType.SelectedItem.Index = 1
    cbr.Bands(2).Visible = tbsType.SelectedItem.Index = 2
    If mnuEdit_Surety.Visible Then mnuEdit_Surety.Enabled = tbsType.SelectedItem.Index = 1
    
    Call InitUnits
    Call ShowPatis(mstrFilter)
    mshPati_Click
End Sub

Private Sub tvwDist_s_NodeClick(ByVal Node As MSComctlLib.Node)
    '��ͬ������ٴ���
    If tvwDist_s.Tag = Node.Key Then Exit Sub
    tvwDist_s.Tag = Node.Key
    
    Call ShowPatis(mstrFilter)
    mshPati_Click
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
    objOut.Title.Text = "��Ժ�����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objRow.Add "���ţ�" & tvwDist_s.SelectedItem.Text
    objRow.Add "ʱ�䣺" & Format(frmHosRegFilter.dtp��ԺB.Value, "yyyy-MM-dd") & " �� " & Format(frmHosRegFilter.dtp��ԺE.Value, "yyyy-MM-dd")
    objOut.UnderAppRows.Add objRow
    
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
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled Then
        mnuEdit_Del_Click
    ElseIf KeyCode = vbKeyReturn And mnuEdit_View.Enabled Then
        mnuEdit_View_Click
    End If
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub InitNode()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngUnitID As Long
    Dim blnByDept As Boolean
    
    On Error GoTo errHandle
    blnByDept = mnuViewByDept(1).Checked
    
    '����վ��ѡ��
    strSQL = "Select Distinct վ��, c.����" & vbNewLine & _
            " From (Select Distinct " & IIf(blnByDept, "��Ժ����id", "��Ժ����id") & " ID" & vbNewLine & _
            "       From ������ҳ" & vbNewLine & _
            "       Where ��Ժ���� Between [1] And [2] And " & IIf(blnByDept, "��Ժ����id", "��Ժ����id") & " Is Not Null) A, ���ű� B, zlnodelist C" & vbNewLine & _
            " Where A.ID = B.ID And B.վ��=c.��� " & vbNewLine & _
            " Order By վ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, frmHosRegFilter.dtp��ԺB, frmHosRegFilter.dtp��ԺE)
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
    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node, i As Integer, lngUnitID As Long
    Dim strPreKey  As String, blnByDept As Boolean
    Dim strNodeNo As String
    Dim strDeptIDs As String
      
    strPreKey = ""
    If Not tvwDist_s.SelectedItem Is Nothing Then strPreKey = tvwDist_s.SelectedItem.Key
    blnByDept = mnuViewByDept(1).Checked
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
        
    tvwDist_s.Nodes.Clear
    Set objNode = tvwDist_s.Nodes.Add(, , "Root", IIf(blnByDept, "���п���", "���в���"), 1)
    objNode.Expanded = True
    If objNode.Key = strPreKey Then objNode.Selected = True
    
    If tbsType.SelectedItem.Index = 2 And InStr(mstrPrivs, ";ȫԺԤԼ;") = 0 Then
        strDeptIDs = GetDeptOrUnitByUser()
    End If
    
    Set rsTmp = GetInDept(blnByDept, frmHosRegFilter.dtp��ԺB, frmHosRegFilter.dtp��ԺE, strNodeNo, strDeptIDs)
    If Not rsTmp.EOF Then
        If blnByDept Then
            lngUnitID = UserInfo.����ID
        Else
            lngUnitID = Get����ID(UserInfo.����ID)
        End If
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvwDist_s.Nodes.Add("Root", tvwChild, IIf(blnByDept, "D", "U") & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, 1)
            If objNode.Key = strPreKey Then objNode.Selected = True
            If rsTmp!ID = lngUnitID And tvwDist_s.SelectedItem Is Nothing Then objNode.Selected = True
            
            objNode.Expanded = True
            rsTmp.MoveNext
        Next
    End If
    If tvwDist_s.SelectedItem Is Nothing Then
        tvwDist_s.Nodes(IIf(tvwDist_s.Nodes.Count > 1, 2, 1)).Selected = True
    End If
        
    InitUnits = True
End Function

Private Sub setHeader(ByVal strHead As String)
    Dim i As Integer, j As Integer
    Dim strWidth As String, strText As String
    Dim arrText As Variant
    
    'gclsBase.GetRegister(˽��ģ��, Me.Name, strPath & "_" & TypeName(vsf(0)) & "_20101228", "")





    With mshPati
        .Redraw = False
        
        
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
            '�ָ���˳��
            '����Ƿ���Ҫ�ָ�
            strText = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.EXEName & "\" & Me.Name & "\" & TypeName(mshPati), mshPati.Name & mshPati.Tag & "����", "")
            arrText = Split(strText, ",")
            
            If strText <> "" Then
                .Cols = UBound(arrText) + 1
                For i = 0 To UBound(arrText)
                    .TextMatrix(0, i) = arrText(i)
                    .ColAlignmentFixed(i) = 4
                    For j = 0 To UBound(Split(strHead, "|"))
                        If (arrText(i) = Split(Split(strHead, "|")(j), ",")(0)) Then
                            .colAlignment(i) = Split(Split(strHead, "|")(j), ",")(1)
                            Exit For
                        End If
                    Next
                Next
            End If

            strWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.EXEName & "\" & Me.Name & "\" & TypeName(mshPati), mshPati.Name & mshPati.Tag & "���", "")
            If UBound(Split(strWidth, ",")) >= .Cols - 1 Then
                For i = 0 To .Cols - 1
                    .ColWidth(i) = Split(strWidth, ",")(i)
                Next
            End If
        End If
        
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Or .Cols = 0 Or .Cols <> UBound(Split(strHead, "|")) + 1 Then
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
                If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
                .ColAlignmentFixed(i) = 4
            Next
        End If
        
        If Not Visible Then Call RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        .ColWidth(0) = 0
        
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
End Sub

Private Sub mshPati_EnterCell()
    If mshPati.Row = 0 Or mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")) = "" Then Exit Sub
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
    
    Call SetMenu(mnuFile_Print.Enabled)
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
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
        If mshPati.TextMatrix(1, GetColNum("����ID")) = "" Then Exit Sub
        
        Set mshPati.DataSource = Nothing

        mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        
        Call ShowPatis(, True)
        mshPati_Click
    End If
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
    Dim i As Integer, blnHavePrivs As Boolean
    Dim lng���� As Long
    i = GetColNum("״̬")
    lng���� = GetColNum("��������")
    
    '����Ȩ��
    mnuEdit_Modi.Visible = True
    tbr.Buttons("Modi").Visible = True
    If InStr(mstrPrivs, "����Ǽ�") = 0 _
        And InStr(mstrPrivs, "ԤԼ�Ǽ�") = 0 _
        And InStr(mstrPrivs, "����ԤԼ") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    Else
        If tbsType.SelectedItem.Index = 1 Then
            If InStr(mstrPrivs, "����Ǽ�") = 0 Then
                mnuEdit_Modi.Visible = False
                tbr.Buttons("Modi").Visible = False
            End If
        Else
            If InStr(mstrPrivs, "ԤԼ�Ǽ�") = 0 Then
                mnuEdit_Modi.Visible = False
                tbr.Buttons("Modi").Visible = False
            End If
        End If
    End If
            
    '���ݿɲ�����
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    mnuFile_PrintMed.Enabled = blnUsed
    mnuFile_PrintWristlet.Enabled = blnUsed
    
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    If Val(mshPati.TextMatrix(mshPati.Row, i)) = 1 Then
        '�յǼǲ���
        mnuEdit_Modi.Enabled = blnUsed
        tbr.Buttons("Modi").Enabled = blnUsed
        tbr.Buttons("Del").Enabled = blnUsed
        mnuEdit_Del.Enabled = blnUsed
        
        If tbsType.SelectedItem.Index = 1 Then
            mnuEditConfirm.Enabled = False
            tbr.Buttons("Confirm").Enabled = False
            mnuEditToKeep.Enabled = blnUsed And Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��������"))) = 0
        Else
            mnuEditConfirm.Enabled = blnUsed
            tbr.Buttons("Confirm").Enabled = blnUsed
            mnuEditToKeep.Enabled = False
        End If
    Else
        '����ס����
        mnuEdit_Modi.Enabled = False
        tbr.Buttons("Modi").Enabled = False
        tbr.Buttons("Del").Enabled = False
        mnuEdit_Del.Enabled = False
        mnuEditToKeep.Enabled = False
        
        mnuEditConfirm.Enabled = False
        tbr.Buttons("Confirm").Enabled = False
    End If
    
    tbr.Buttons("View").Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
    
    'סԺ���۲���תΪסԺ����
    mnuEditToIn.Enabled = (Val(mshPati.TextMatrix(mshPati.Row, lng����)) = 2)
    
    '�շ����ʹ���
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";����;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    'tbr.Buttons("����").Visible = blnHavePrivs
    'tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    '������Ϣ
    mnuEdit_Surety.Visible = InStr(mstrPrivs, ";������Ϣ;") > 0
    '���˼���
    blnHavePrivs = InStr(";" & GetPrivFunc(glngSys, 9003) & ";", ";���˼���;") > 0
    mnuEdit_Family.Visible = blnHavePrivs
    mnuEdit_FamilyAdd.Visible = blnHavePrivs
    mnuEdit_FamilyView.Visible = blnHavePrivs
    mnuEdit_FamilyView.Enabled = blnUsed And blnHavePrivs
    
    tbr.Buttons("FamilySplit").Visible = blnHavePrivs
    tbr.Buttons("Family").Visible = blnHavePrivs
    tbr.Buttons("Family").ButtonMenus.Item("FamilyView").Enabled = blnUsed And blnHavePrivs
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������Ĳ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmHosRegFind
            If .txt����ID.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����ID")) = .txt����ID.Text
            End If
            If .txt���￨.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���￨")) = .txt���￨.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshPati.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����")) Like "*" & .txt����.Text & "*"
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

Private Sub ShowPatis(Optional ByVal strIF As String, Optional blnSort As Boolean)
'���ܣ����ݵ�ǰ�˵����Ҫ��(�Զ���������),��ȡ������Ϣ
'������strIF=" And ...."��ʽ�Ĺ�������
    Dim i As Long, strSQL As String, strDiagnoseSQL As String
    Dim strCard As String, strUnit As String
    Dim blnByDept As Boolean, lngDeptID As Long
    Dim Curdate As Date
    Dim strNodeNo As String
    Dim intDiagDays As Integer
    Dim strPerson As String, strParTable As String, strTable As String, strDiag As String
    Dim varArr As Variant
    Dim rsDiag As New ADODB.Recordset
    Dim j As Integer
    
    'by lesfeng 2010-1-11 �����Ż�
    On Error GoTo errH
    
    If blnSort = False Then PatiIdentify.Text = ""
    strPerson = ""
    strDiag = ""
    
    '��ȡվ���
    If cboNodeList.ListIndex <> -1 Then
        strNodeNo = Mid(cboNodeList.Text, 1, InStr(cboNodeList.Text, "-") - 1)
    Else
        strNodeNo = 0
    End If
    
    If Not blnSort Then
        blnByDept = mnuViewByDept(1).Checked
        
        If strIF = "" Then
            '���ó�ʼ����(��������Ժ)
            'strIF = " And B.��Ժ���� Between trunc(Sysdate,'mm') And Sysdate"
'            strIF = " AND B.��Ժ���� Between trunc(Sysdate-7) and Sysdate "
            Curdate = zlDatabase.Currentdate
            strIF = ""
            strIF = strIF & " And (B.��Ժ����  Between [1] And [2]) "
            mcllFilterA.Remove "��Ժ����"
            mcllFilterA.Add Array(Format(DateAdd("d", -7, Curdate), "yyyy-mm-dd") & " 00:00:00", Format(Curdate, "yyyy-mm-dd") & " 23:59:59"), "��Ժ����"
        End If
        If tbsType.SelectedItem.Index = 1 Then
            strIF = strIF & " And A.��ҳID=B.��ҳID And Nvl(B.��ҳID,0)<>0"
        Else
            strIF = strIF & " And Nvl(B.��ҳID,0)=0"
        End If
        
        '���￨����ʾ
        '55849:������,2012-11-21,��ԭ��Decode�жϵķ�ʽ��Ϊ�̶���ȡ�ֶ�,
        '��ΪDecode��һ������ʹ�ó�����ָ������ȡ�ֶ����ݣ����ܵ��µ��²鲻����������߷��صļ�¼�����ʳ���E-FAIL���󣬹�����ADO��Oracle�����Ե�Bug�����ض���Decode���ӱ��ѯͬʱʹ��ʱ����֣���û����ȷ�Ĺ��ɡ�
        'strCard = "Decode(" & IIf(mblnPassShowCard, 0, 1) & ",1,A.���￨��,LPAD('*',Length(A.���￨��),'*')) as ���￨,"
        If mblnPassShowCard = True Then
            strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨,"
        Else
            strCard = "A.���￨�� as ���￨,"
        End If
        '��ǰ���������
        If Not tvwDist_s.SelectedItem Is Nothing Then  '����κο��һ���û����,��ֻ�����в���
            lngDeptID = Val(Mid(tvwDist_s.SelectedItem.Key, 2))
            If blnByDept Then
                If lngDeptID <> 0 Then strUnit = " And B.��Ժ����ID=[6]"
            Else
                If lngDeptID <> 0 Then strUnit = " And B.��Ժ����ID=[6]"
            End If
        End If
        
        mintBedLen = GetMaxBedLen(lngDeptID)
        '54179:������,2012-10-12,�޸���ȡ������ϵ�sql�������ʾ���һ��������ϣ���ǰΪ������ʷ����������ϣ�
        If tbsType.SelectedItem.Index = 1 Then
            If Not (mnuViewInBed.Checked And mnuViewInBed.Enabled) Then
                '�ȴ���ƵĲ���(״̬=1)������Ҫ����Ϊ" ",��Ȼȫ��Ϊ���벡��ʱ��������
                strSQL = _
                    "Select ��������,Decode(B.��������,1,'��������',2,'סԺ����','סԺ����') as �Ǽ�����," & _
                    " A.����ID, A.�����, B.סԺ��,B.���ۺ�," & strCard & "' ' as ����," & _
                    " NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.�ѱ�,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,Nvl(A.ҽ����,F.��Ϣֵ) as ҽ����,X.���� as ����," & _
                    " To_Char(B.��Ժ����,'YYYY-MM-DD HH24:MI:SS') as ��Ժʱ��," & _
                    " C.���� as ��Ժ����,D.���� as ��Ժ����,E.���� as ����ȼ�,A.סԺ���� as ����,B.��Ժ����," & _
                    " B.��Ժ��ʽ,B.סԺĿ��,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����," & _
                    " A.ѧ��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.����״�� as ����,A.������λ,A.��ͥ��ַ,  A.��ͥ�绰, Decode(g.�������, Null, '',g.�������) As �������, B.��ע,B.�Ǽ��� as �Ǽ�Ա,B.״̬,B.��ҳID," & _
                    " Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������,B.�Һ�ID " & _
                    " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,�շ���ĿĿ¼ E,������ҳ�ӱ� F,������ϼ�¼ G,������� X" & _
                    " Where A.����ID=B.����ID And B.״̬ = 1 And B.��Ժ����ID=C.ID(+)" & _
                    " And B.��Ժ����ID=D.ID " & IIf(lngDeptID = 0 And cboNodeList.Visible, "And (d.վ��=" & strNodeNo & " Or d.վ�� Is Null)", "") & " And B.����ȼ�ID=E.ID(+)" & _
                    " And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) And B.����ID = G.����ID(+) And B.��ҳID=G.��ҳID(+) And g.��¼��Դ(+) = 2 And g.�������(+) = 1 And g.��ϴ���(+) = 1 " & _
                    " And F.��Ϣ��(+)='ҽ����' And B.����=X.���(+)" & strUnit & strIF & _
                    IIf(chkOnly.Value = 1, "  And Exists (Select 1 From ������ҳ Where ����ID=a.����id And ��ҳID>0 And ��������=1 And ��Ժʱ�� Is Not Null And ��Ժʱ�� Is Null)", "") & _
                    " Order by ��Ժʱ�� Desc,סԺ�� Desc"
            Else
                strSQL = _
                    "Select ��������,Decode(B.��������,1,'��������',2,'סԺ����','סԺ����') as �Ǽ�����," & _
                    " A.����ID, A.�����, B.סԺ��,B.���ۺ�," & strCard & "Decode(Nvl(B.״̬,0),1,NULL,LPad(B.��Ժ����," & mintBedLen & ", ' ')) as ����," & _
                    " NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.�ѱ�,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,Nvl(A.ҽ����,F.��Ϣֵ) as ҽ����,X.���� as ����," & _
                    " To_Char(B.��Ժ����,'YYYY-MM-DD HH24:MI:SS') as ��Ժʱ��," & _
                    " C.���� as ��Ժ����,D.���� as ��Ժ����,E.���� as ����ȼ�,A.סԺ���� as ����,B.��Ժ����," & _
                    " B.��Ժ��ʽ,B.סԺĿ��,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����," & _
                    " A.ѧ��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.����״�� as ����,A.������λ,A.��ͥ��ַ, A.��ͥ�绰, Decode(g.�������, Null, '',g.�������) As �������, B.��ע,B.�Ǽ��� as �Ǽ�Ա,B.״̬,B.��ҳID," & _
                    " Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������,B.�Һ�ID" & _
                    " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,�շ���ĿĿ¼ E,������ҳ�ӱ� F,������ϼ�¼ G,������� X" & _
                    " Where A.����ID=B.����ID And B.��Ժ���� is NULL And B.��Ժ����ID=C.ID(+)" & _
                    " And B.��Ժ����ID=D.ID " & IIf(lngDeptID = 0 And cboNodeList.Visible, "And (d.վ��=" & strNodeNo & " Or d.վ�� Is Null)", "") & " And B.����ȼ�ID=E.ID(+)" & _
                    " And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) And B.����ID = G.����ID(+) And B.��ҳID=G.��ҳID(+) And g.��¼��Դ(+) = 2 And g.�������(+) = 1 And g.��ϴ���(+) = 1 " & _
                    " And F.��Ϣ��(+)='ҽ����' And B.����=X.���(+)" & strUnit & strIF & _
                    IIf(chkOnly.Value = 1, "  And Exists (Select 1 From ������ҳ Where ����ID=a.����id And ��ҳID>0 And ��������=1 And ��Ժʱ�� Is Not Null And ��Ժʱ�� Is Null)", "") & _
                    " Order by ��Ժʱ�� Desc,סԺ�� Desc"
            End If
        Else
            '��ѯԤԼ�Ǽǲ���
            strSQL = _
                    "Select ��������,Decode(B.��������,1,'��������',2,'סԺ����','סԺ����') as �Ǽ�����," & _
                    " A.����ID, A.�����, B.סԺ��,B.���ۺ�," & strCard & "LPad(B.��Ժ����," & mintBedLen & ", ' ') as ����," & _
                    " NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.�ѱ�,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,Nvl(A.ҽ����,F.��Ϣֵ) as ҽ����,X.���� as ����," & _
                    " To_Char(B.��Ժ����,'YYYY-MM-DD HH24:MI:SS') as ��Ժʱ��," & _
                    " C.���� as ��Ժ����,D.���� as ��Ժ����,E.���� as ����ȼ�,A.סԺ���� as ����,B.��Ժ����," & _
                    " B.��Ժ��ʽ,B.סԺĿ��,To_Char(A.��������,'YYYY-MM-DD HH24:MI') as ��������,A.����,A.����," & _
                    " A.ѧ��,A.ְҵ,A.���,A.���֤��,A.�ֻ���,A.����״�� as ����,A.������λ,A.��ͥ��ַ,  A.��ͥ�绰, Null As �������, B.��ע,B.�Ǽ��� as �Ǽ�Ա,B.״̬,B.��ҳID," & _
                    " Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������,B.�Һ�ID" & _
                    " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,�շ���ĿĿ¼ E,������ҳ�ӱ� F,������� X" & _
                    " Where A.����ID=B.����ID And B.״̬ = 1 And B.��Ժ����ID=C.ID(+)" & _
                    " And B.��Ժ����ID=D.ID " & IIf(lngDeptID = 0 And cboNodeList.Visible, "And (d.վ��=" & strNodeNo & " Or d.վ�� Is Null)", "") & " And B.����ȼ�ID=E.ID(+)" & _
                    " And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+) " & _
                    " And F.��Ϣ��(+)='ҽ����' And B.����=X.���(+)" & strUnit & strIF & _
                    IIf(chkOnly.Value = 1, "  And Exists (Select 1 From ������ҳ Where ����ID=a.����id And ��ҳID>0 And ��������=1 And ��Ժʱ�� Is Not Null And ��Ժʱ�� Is Null)", "") & _
                    " Order by ��Ժʱ�� Desc,סԺ�� Desc"
        End If
        If Not tvwDist_s.SelectedItem Is Nothing Then
            tvwDist_s.Tag = tvwDist_s.SelectedItem.Key
        End If
        
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����嵥,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        '����17122 by lesfeng 2010-02-02
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mcllFilterA("��Ժ����")(0)), CDate(mcllFilterA("��Ժ����")(1)), _
        CLng(Val(mcllFilterA("סԺ��")(0))), CLng(Val(mcllFilterA("סԺ��")(1))), CStr(mcllFilterA("�Ǽ���")), lngDeptID, gstrLike & CStr(mcllFilterA("��������")) & "%", mcllFilterA("�����"))
'        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptID)
        If Not mrsPati.EOF And tbsType.SelectedItem.Index = 2 Then
            For i = 0 To mrsPati.RecordCount - 1
                strPerson = strPerson & "," & mrsPati!����ID
                mrsPati.MoveNext
            Next
            mrsPati.MoveFirst
            strPerson = Mid(strPerson, 2)
            intDiagDays = Val(zlDatabase.GetPara("��ϲ�������", glngSys, glngModul, "3"))
            strParTable = "Select /* +cardinality(a,10) */" & "Column_Value From Table(f_num2List([1]))"
            strTable = strParTable
            
            If Len(strPerson) >= 4000 Then
                varArr = Array()
                varArr = GetParTable(strPerson, strParTable, strTable)
            End If
            strSQL = "Select a.����id, a.�������, 1 As ���" & vbNewLine & _
                "From ������ҳ G, ������ϼ�¼ A, �������ҽ�� B, ����ҽ����¼ C, ������ĿĿ¼ D, ���˹Һż�¼ E" & vbNewLine & _
                "Where g.����id = a.����id And a.Id = b.���id And b.ҽ��id = c.Id And c.������Ŀid + 0 = d.Id And a.����id=e.����id And a.��ҳid = e.Id And a.��¼��Դ = 3 And" & vbNewLine & _
                "      e.��¼���� = 1 And e.��¼״̬ = 1 And e.�Ǽ�ʱ�� + 0 > Trunc(Sysdate-" & intDiagDays & ") And c.ҽ��״̬ In (3, 8) And d.��� = 'Z' And" & vbNewLine & _
                "      Instr(',1,11,', ',' || a.������� || ',') > 0 And Instr(',1,2,', d.��������) > 0 And g.��Ժ����id = c.ִ�п���id And" & vbNewLine & _
                "      Nvl(g.��ҳid, 0) = 0 And G.����id in (" & strTable & "A)" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select a.����id, a.�������, 2 As ���" & vbNewLine & _
                "From ������ϼ�¼ A, ���˹Һż�¼ B" & vbNewLine & _
                "Where a.����id = b.����id And a.��ҳid = b.Id And b.��¼���� = 1 And b.��¼״̬ = 1 And b.�Ǽ�ʱ�� + 0 > Trunc(Sysdate-" & intDiagDays & ") And" & vbNewLine & _
                "      Instr(',1,11,', ',' || a.������� || ',') > 0 And a.��¼��Դ = 3 And b.����id in (" & strTable & "A)"

            If Len(strPerson) >= 4000 Then
                Set rsDiag = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
            Else
                Set rsDiag = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", strPerson)
            End If
        End If
    End If
    
    mshPati.Clear
    mshPati.ClearStructure
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call setHeader(mstrHead)
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κβ���"
        Call SetMenu(False)
    Else
        Set mshPati.DataSource = mrsPati
         If tbsType.SelectedItem.Index = 2 And mshPati.Rows > 0 Then
	 If Not rsDiag Is Nothing Then
            If rsDiag.RecordCount > 0 Then
                For i = 1 To mshPati.Rows - 1
                    rsDiag.Filter = "����ID=" & Val(mshPati.TextMatrix(i, COL_����ID)) & " And ���=1"
                    If rsDiag.RecordCount > 0 Then
                        For j = 0 To rsDiag.RecordCount - 1
                            strDiag = strDiag & "," & rsDiag!�������
                            rsDiag.MoveNext
                        Next
                        mshPati.TextMatrix(i, COL_�������) = Mid(strDiag, 2)
                    Else
                        rsDiag.Filter = "����ID=" & Val(mshPati.TextMatrix(i, COL_����ID)) & " And ���=2"
                         If rsDiag.RecordCount > 0 Then
                            For j = 0 To rsDiag.RecordCount - 1
                                strDiag = strDiag & "," & rsDiag!�������
                                rsDiag.MoveNext
                            Next
                            mshPati.TextMatrix(i, COL_�������) = Mid(strDiag, 2)
                        End If
                    End If
                    strDiag = ""
                Next
            End If
        End If
	End If
        Call setHeader(mstrHead)          '�����е�enter_cell���ѵ���SetMenu(false)
        If mnuViewInBed.Checked Then Call SetInBed
        stbThis.Panels(2) = "�� " & mrsPati.RecordCount & " ������"
        Call SetMenu(True)
    End If
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub
Public Function GetParTable(ByVal strpar As String, ByVal strParTable As String, ByRef strTableOut As String) As Variant
'���ܣ����ڶ�̬�ڴ��İ󶨲��������Ĵ���
'������strPar ��������strParTable �ڴ����ʽҪ����
'���أ�һ���ַ������飬10��Ԫ��
    Dim n As Long, p As Long
    Dim varPar(0 To 9) As String
    Dim strTable As String, strThis As String
    Dim intNum As Integer '������
    
    For n = 0 To 9
        varPar(n) = ""
    Next
    
    p = InStr(strParTable, "[") + 1
    intNum = Mid(strParTable, p, 1)
    
    n = 0
    Do While True
        If Len(strpar) < 4000 Then
            p = Len(strpar) + 1
        Else
            p = InStrRev(Mid(strpar, 1, 4000), ",")
        End If
        
        strThis = Mid(strpar, 1, p - 1)
        
        If n > 9 Then
            strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "'" & strThis & "'")
        Else
            varPar(n) = strThis
            If n = 0 Then
                strTable = strParTable
            Else
                strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "[" & (n + intNum) & "]")
            End If
        End If
        
        n = n + 1
        
        strpar = Mid(strpar, p + 1)
        
        If strpar = "" Then Exit Do
    Loop
    
    strTableOut = strTable
    GetParTable = varPar
    
End Function
Private Sub SetInBed()
    Dim i As Integer, j As Integer, k As Integer
    Dim bln As Boolean
    Dim intRow As Integer, intCol As Integer
    
    intRow = mshPati.Row
    bln = mshPati.Redraw
    mshPati.Redraw = False
        
    j = GetColNum("״̬")
    k = GetColNum("����")
    For i = 1 To mshPati.Rows - 1
        '���Ų�Ϊ�յ�(������ͥ����)Ϊ��ס����
        If Val(mshPati.TextMatrix(i, j)) <> 1 Then
            mshPati.Row = i: mshPati.Col = k
            mshPati.CellBackColor = &HEBFFFF
        End If
    Next
    mshPati.Row = intRow: mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = bln
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Integer
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
'����22073 by lesfeng 2010-08-02  ��֤�Ƿ���д���Ӳ���
Private Function GetCaseHistory(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ���ȡָ�������Ƿ���ڵ��Ӳ�����¼
'˵�������ڻ�ȡ���˵��Ӳ�����¼�ļ�¼���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim int��¼�� As Integer
    
    GetCaseHistory = False
    On Error GoTo errH
    
    strSQL = "Select count(����id) As ���� From ���Ӳ�����¼ " & _
             " Where ����ID = [1] And ��ҳID = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!����) Then
            int��¼�� = rsTmp!����
            If int��¼�� > 0 Then GetCaseHistory = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


