VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClinicLists 
   BackColor       =   &H8000000C&
   Caption         =   "������Ŀ����"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12450
   Icon            =   "frmClinicLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   12450
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   10005
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Caption2        =   "����"
      Child2          =   "txtFind"
      MinHeight2      =   300
      Width2          =   1080
      Key2            =   "find"
      NewRow2         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   10620
         TabIndex        =   41
         Top             =   240
         Width           =   1740
      End
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   10
         Top             =   30
         Width           =   9810
         _ExtentX        =   17304
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
               Object.ToolTipText     =   "Ԥ����ǰ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ��ǰ��"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Class"
               Description     =   "����"
               Object.ToolTipText     =   "����ҩƷ����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "�����µ���Ŀ"
               Object.Tag             =   "����"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "add"
                     Text            =   "����"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "addcopy"
                     Text            =   "��������"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸ĵ�ǰ��Ŀ"
               Object.Tag             =   "�޸�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰ��Ŀ"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Description     =   "����"
               Object.ToolTipText     =   "����ָ����ͣ����Ŀ"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ͣ��"
               Key             =   "Stop"
               Description     =   "ͣ��"
               Object.ToolTipText     =   "ͣ��ָ����������Ŀ"
               Object.Tag             =   "ͣ��"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "���������Ŀ"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split5"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2625
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5910
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   795
      Width           =   30
   End
   Begin VB.PictureBox picClass 
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "���׷���(&2)"
         Height          =   350
         Index           =   2
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1665
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "��ҩ�䷽(&1)"
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1335
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "������Ŀ(&0)"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   4005
         Left            =   45
         TabIndex        =   5
         Tag             =   "1000"
         Top             =   2055
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   7064
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   150
      Top             =   7170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":08CA
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":0E64
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":13FE
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":1998
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":1F32
            Key             =   "���U"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":24CC
            Key             =   "���S"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":2A66
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":3000
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":359A
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":3B34
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":40CE
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":4668
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":4C02
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":519C
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":5736
            Key             =   "��ʳU"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":5CD0
            Key             =   "��ʳS"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":626A
            Key             =   "��ѪU"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":6804
            Key             =   "��ѪS"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":6D9E
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":7338
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":78D2
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":7E6C
            Key             =   "����S"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":8406
            Key             =   "��ҩU"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":89A0
            Key             =   "��ҩS"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":8F3A
            Key             =   "��ҩU"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":94D4
            Key             =   "��ҩS"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":9A6E
            Key             =   "����U"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":A008
            Key             =   "����S"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4125
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   7276
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8355
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicLists.frx":A5A2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16880
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
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
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
            Picture         =   "frmClinicLists.frx":AE34
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B04E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B268
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B482
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B69C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":BAD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":BCEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":BF04
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C11E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C33E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
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
            Picture         =   "frmClinicLists.frx":C55E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C77E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C99E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":CBB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":CDD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":CFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D206
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D420
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D63A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D854
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":DA74
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   210
      Top             =   7245
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabContent 
      Height          =   2820
      HelpContextID   =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   5505
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   4
      TabsPerRow      =   9
      TabHeight       =   520
      WordWrap        =   0   'False
      OLEDropMode     =   1
      TabCaption(0)   =   "ִ�п���(&S)"
      TabPicture(0)   =   "frmClinicLists.frx":DC94
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "hgd����ִ��"
      Tab(0).Control(1)=   "fraSubInfo(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "�շѶ���(&C)"
      TabPicture(1)   =   "frmClinicLists.frx":DCB0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSubInfo(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "����ָ��(&L)"
      TabPicture(2)   =   "frmClinicLists.frx":DCCC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSubInfo(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "��ѡ��λ(&P)"
      TabPicture(3)   =   "frmClinicLists.frx":DCE8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraSubInfo(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "�÷�����(&U)"
      TabPicture(4)   =   "frmClinicLists.frx":DD04
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fraSubInfo(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�������(T)"
      TabPicture(5)   =   "frmClinicLists.frx":DD20
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraSubInfo(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "�䷽���(&M)"
      TabPicture(6)   =   "frmClinicLists.frx":DD3C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraSubInfo(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "���׷���(&M)"
      TabPicture(7)   =   "frmClinicLists.frx":DD58
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraSubInfo(7)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Ӧ�òο�(&R)"
      TabPicture(8)   =   "frmClinicLists.frx":DD74
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "fraSubInfo(8)"
      Tab(8).ControlCount=   1
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   8
         Left            =   -74250
         TabIndex        =   20
         Top             =   495
         Width           =   5115
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
            Height          =   2880
            Left            =   150
            TabIndex        =   35
            Top             =   195
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   5080
            _Version        =   393216
            BackColor       =   -2147483628
            Rows            =   5
            Cols            =   4
            FixedRows       =   0
            BackColorBkg    =   -2147483628
            GridColor       =   -2147483628
            GridColorFixed  =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            GridLines       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            MergeCells      =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   7
         Left            =   -73980
         TabIndex        =   18
         Top             =   300
         Width           =   8010
         Begin VSFlex8Ctl.VSFlexGrid vsScheme 
            Height          =   2625
            Left            =   165
            TabIndex        =   19
            Top             =   225
            Width           =   7035
            _cx             =   12409
            _cy             =   4630
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
            BackColorSel    =   12632256
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClinicLists.frx":DD90
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
      End
      Begin VB.Frame fraSubInfo 
         Height          =   2430
         Index           =   6
         Left            =   -74835
         TabIndex        =   17
         Top             =   360
         Width           =   6570
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRecipe 
            Height          =   1710
            Left            =   195
            TabIndex        =   33
            Top             =   315
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   3016
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   5
         Left            =   -74940
         TabIndex        =   16
         Top             =   405
         Width           =   8010
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdTabu 
            Height          =   2175
            Left            =   120
            TabIndex        =   32
            Top             =   150
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   525
         Width           =   8010
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdUsage 
            Height          =   2175
            Left            =   210
            TabIndex        =   34
            Top             =   45
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   2190
         Index           =   3
         Left            =   -74835
         TabIndex        =   14
         Top             =   450
         Width           =   7710
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdPart 
            Height          =   2175
            Left            =   195
            TabIndex        =   31
            Top             =   375
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   2265
         Index           =   2
         Left            =   -74895
         TabIndex        =   13
         Top             =   345
         Width           =   7740
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdLabs 
            Height          =   2175
            Left            =   90
            TabIndex        =   30
            Top             =   180
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         BorderStyle     =   0  'None
         Height          =   2355
         Index           =   1
         Left            =   -74895
         TabIndex        =   12
         Top             =   465
         Width           =   7845
         Begin VSFlex8Ctl.VSFlexGrid vsfExse 
            Height          =   1425
            Left            =   1095
            TabIndex        =   42
            Top             =   195
            Width           =   5955
            _cx             =   10504
            _cy             =   2514
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
            BackColorSel    =   12632256
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgd����ִ�� 
         Height          =   1125
         Left            =   -71325
         TabIndex        =   37
         Top             =   1230
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   1984
         _Version        =   393216
         FixedCols       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraSubInfo 
         Enabled         =   0   'False
         Height          =   2205
         Index           =   0
         Left            =   -74910
         TabIndex        =   21
         Top             =   345
         Width           =   7890
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "���������ڿ���(&6)"
            Height          =   195
            Index           =   6
            Left            =   1035
            TabIndex        =   40
            Top             =   1830
            Width           =   1860
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "ҽԺ��ִ��(&5)"
            Height          =   195
            Index           =   5
            Left            =   1035
            TabIndex        =   27
            Top             =   1590
            Width           =   2250
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "ָ������ִ��(&4)"
            Height          =   195
            Index           =   4
            Left            =   1035
            TabIndex        =   26
            Top             =   1350
            Width           =   2250
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "����Ա���ڿ���(&3)"
            Height          =   195
            Index           =   3
            Left            =   1035
            TabIndex        =   25
            Top             =   1110
            Width           =   2250
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "�ɲ��˲���ִ��(&2)"
            Height          =   195
            Index           =   2
            Left            =   1035
            TabIndex        =   24
            Top             =   870
            Width           =   2250
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "�ɲ��˿���ִ��(&1)"
            Height          =   195
            Index           =   1
            Left            =   1035
            TabIndex        =   23
            Top             =   630
            Width           =   2250
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "������ִ�еĶ���(&0)"
            Height          =   195
            Index           =   0
            Left            =   1035
            TabIndex        =   22
            Top             =   390
            Value           =   -1  'True
            Width           =   2250
         End
         Begin VB.Label lblExcute 
            AutoSize        =   -1  'True
            Caption         =   "ִ�п��ң�"
            Height          =   180
            Left            =   150
            TabIndex        =   39
            Top             =   390
            Width           =   900
         End
         Begin VB.Label lblUseBill 
            AutoSize        =   -1  'True
            Caption         =   "���Ƶ��ݣ�"
            Height          =   180
            Left            =   150
            TabIndex        =   38
            Top             =   165
            Width           =   900
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            Caption         =   "���¿��ҿ����ֱ�һ����ָ������ִ��(&L)��"
            Height          =   180
            Left            =   3570
            TabIndex        =   29
            Top             =   645
            Width           =   3510
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            Caption         =   "һ��������              ִ�У�"
            Height          =   180
            Left            =   3570
            TabIndex        =   28
            Top             =   375
            Width           =   2700
         End
      End
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "�����ߴ�"
      Height          =   180
      Left            =   9330
      TabIndex        =   36
      Top             =   4500
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
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
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "��������(&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "����(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "����(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "�޸�(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "ɾ��(&E)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "��Ŀ(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAddcopy 
         Caption         =   "��������(&C)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRefer 
         Caption         =   "Ӧ�òο�(&R)..."
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditExse 
         Caption         =   "�շѶ���(&E)..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditLabs 
         Caption         =   "����ָ��(&L)..."
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditGather 
         Caption         =   "�ɼ���ʽ(&G)..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSample 
         Caption         =   "�걾����(&P)..."
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "ͣ��(&S)"
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRepellent 
         Caption         =   "�ų��ϵ(&N)"
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "��Ӧ����(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSpt4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditImport 
         Caption         =   "��Ŀ����(&I)"
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
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "��ʾ�����¼�(&H)"
      End
      Begin VB.Menu mnuViewStoped 
         Caption         =   "��ʾͣ��(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmClinicLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint��Χ As Integer '���׷����Ŀ�ʹ�ó��ϣ�1-����,2-סԺ,3-�����סԺ
Private mlngMode As Long
Private mstrPrivs As String       '�û����б�����ľ���Ȩ��
Private mbyt��ҩζ�� As Byte

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�

Private Const conTabִ�п��� As Integer = 0
Private Const conTab�շѶ��� As Integer = 1
Private Const conTab����ָ�� As Integer = 2
Private Const conTab��鲿λ As Integer = 3
Private Const conTab�÷����� As Integer = 4
Private Const conTab������� As Integer = 5
Private Const conTab�䷽��� As Integer = 6
Private Const conTab���׷��� As Integer = 7
Private Const conTabӦ�òο� As Integer = 8

Private Enum SelectKind
    SK_������Ŀ = "0"
    SK_��ҩ�䷽ = "1"
    SK_���׷��� = "2"
End Enum

Private Enum COL���׷���
    col��Ч = 0
    col���� = 1
    col���� = 2
    col������λ = 3
    col���� = 4
    col��λ = 5
    col���� = 6
    colƵ�� = 7
    col�÷� = 8
    col���� = 9
    colִ��ʱ�� = 10
    colִ�п��� = 11
    colִ������ = 12
    col��� = 13
    col��� = 14
    col��ĿID = 15
    col��� = 16
    col�շ�ϸĿID = 17
    col�걾��λ = 18
    col��鷽�� = 19
    colִ�б�� = 20 'ҩƷҽ���������� ��ȡҩ�Ͳ�ȡҩ
    colͣ�� = 21
End Enum
 
Private mblnStartPriceGrade As Boolean '�����˼۸�ȼ�
Private mstrPriceGrade As String
Private mstrPriceGradeFields As String

Private Sub InitPriceGrade()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���۸�ȼ�
    '����:���˺�
    '����:2017-07-01 21:37:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String, strTempFileds As String
    Dim i As Long
    mblnStartPriceGrade = zlGetrsPriceGrade(rsTemp)
    mstrPriceGrade = "": mstrPriceGradeFields = ""
    If mblnStartPriceGrade = False Then Exit Sub
    If rsTemp.RecordCount = 0 Then mblnStartPriceGrade = False: Exit Sub
    With rsTemp
        i = 1
        .MoveFirst
        Do While Not .EOF
            mstrPriceGrade = mstrPriceGrade & "," & !����
            strTempFileds = strTempFileds & ",sum(decode(P.�۸�ȼ�,'" & !���� & "',P.�ּ�, -1*NULL))  as   A" & i
            i = i + 1
            .MoveNext
        Loop
        .MoveFirst
    End With
    If mstrPriceGrade <> "" Then mstrPriceGrade = Mid(mstrPriceGrade, 2)
    mstrPriceGradeFields = strTempFileds
End Sub



Public Sub ShowMeWithScheme(frmMain As Object, ByVal int��Χ As Integer)
    mint��Χ = int��Χ
    
    On Error Resume Next
    Me.Show , frmMain
    Me.Caption = "���׷�������"
End Sub

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = SK_������Ŀ
        Else
            Me.cmdKind(intCount).Tag = SK_��ҩ�䷽
        End If
    Next
    
    '���õ���Ĳ˵�����
    Me.mnuEditImport.Enabled = (Index = SK_������Ŀ)
    'װ���ݲ���������
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Val(Me.tvwClass.Tag) <> Index Then
        Me.tvwClass.Tag = Index
        Call zlRefClasses
    End If
    Me.mnuClass.Enabled = (tvwClass.Tag = SK_������Ŀ And InStr(1, mstrPrivs, "������Ŀ�༭") > 0) Or _
                            (tvwClass.Tag = SK_��ҩ�䷽ And InStr(1, mstrPrivs, "��ҩ�䷽�༭") > 0) Or _
                            (tvwClass.Tag = SK_���׷��� And InStr(1, mstrPrivs, "���׷����༭") > 0)
    
    Me.mnuEditAdd.Enabled = Me.mnuClass.Enabled
    Me.mnuAddcopy.Enabled = Me.mnuClass.Enabled
    Me.mnuEditModify.Enabled = Me.mnuClass.Enabled
    Me.mnuEditDelete.Enabled = Me.mnuClass.Enabled
    Me.mnuEditLabs.Enabled = Me.mnuClass.Enabled
    Me.mnuEditGather.Enabled = Me.mnuClass.Enabled
    Me.mnuEditStart.Enabled = Me.mnuClass.Enabled
    Me.mnuEditStop.Enabled = Me.mnuClass.Enabled
    Me.mnuEditRepellent.Tag = ""
    Me.mnuEditBill.Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Class").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Add").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Modify").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Delete").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Start").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Stop").Enabled = Me.mnuClass.Enabled
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
End Sub

Private Sub Form_Load()
    '����ָ�
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    Call InitPriceGrade
    
    
    If mint��Χ = 0 Then mint��Χ = 3 'ֱ��ͨ��������Ŀ�������ʱ��ȱʡΪ3
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����", , , True)) = 1 Then
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picVBar.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", "0")
        If strTemp <> "0" Then
            Me.picHBar.Top = CLng(strTemp)
        End If
    End If
    
    Me.mnuViewStoped.Checked = (Val(zlDatabase.GetPara("��ʾͣ����Ŀ", glngSys, 1054, 0)) = 1)
    With Me.hgdRefer
        .ColWidth(0) = 0
        .ColWidth(1) = Me.TextWidth("�ո�")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
    End With
    
    '��ֱ��ͨ���˵����е�Ȩ�޿���
    If InStr(1, mstrPrivs, "������Ŀ�༭") = 0 And _
                            InStr(1, mstrPrivs, "��ҩ�䷽�༭") = 0 And _
                            InStr(1, mstrPrivs, "���׷����༭") = 0 Then
        Me.mnuClass.Enabled = False
        Me.mnuEditAdd.Enabled = False
        Me.mnuAddcopy.Enabled = False
        Me.mnuEditModify.Enabled = False
        Me.mnuEditDelete.Enabled = False
        Me.mnuEditLabs.Enabled = False
        Me.mnuEditGather.Enabled = False
        'Me.mnuEditSample.Enabled = False
        'Me.mnuEditExams.Enabled = False
        Me.mnuEditStart.Enabled = False
        Me.mnuEditStop.Enabled = False
        Me.mnuEditRepellent.Tag = ""
        Me.mnuEditBill.Enabled = False
        Me.tlbThis.Buttons("Class").Enabled = False
        Me.tlbThis.Buttons("Add").Enabled = False
        Me.tlbThis.Buttons("Modify").Enabled = False
        Me.tlbThis.Buttons("Delete").Enabled = False
        Me.tlbThis.Buttons("Start").Enabled = False
        Me.tlbThis.Buttons("Stop").Enabled = False
    Else
        Me.mnuEditRepellent.Tag = 1
    End If
    If InStr(1, mstrPrivs, "�շ�����") = 0 Then
        Me.mnuEditExse.Enabled = False
    End If
    If InStr(1, mstrPrivs, "�ο��༭") = 0 Then
'        Me.mnuEditRefer.Enabled = False
    End If
    If InStr(1, mstrPrivs, "��Ŀ����") = 0 Then
        Me.mnuEditSpt4.Visible = False
        Me.mnuEditImport.Visible = False
    End If
    
    If InStr(mstrPrivs, "����������Ŀ") = 0 And InStr(mstrPrivs, "������ҩ�䷽") = 0 And InStr(mstrPrivs, "������׷���") = 0 Then
        MsgBox "��û�й����κ���Ŀ���ݵ�Ȩ�ޣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Unload Me: Exit Sub
    Else
        If InStr(mstrPrivs, "����������Ŀ") = 0 Then
            cmdKind(0).Visible = False
        End If
        If InStr(mstrPrivs, "������ҩ�䷽") = 0 Then
            cmdKind(1).Visible = False
        End If
        If InStr(mstrPrivs, "������׷���") = 0 Then
            cmdKind(2).Visible = False
        End If
        
        If InStr(mstrPrivs, "����������Ŀ") > 0 Then
            Call cmdKind_Click(0)
        ElseIf InStr(mstrPrivs, "������ҩ�䷽") > 0 Then
            Call cmdKind_Click(1)
        ElseIf InStr(mstrPrivs, "������׷���") > 0 Then
            Call cmdKind_Click(2)
        End If
    End If
    
    '��ʼ������RIS�ӿ�
    If mblnPACSInterface Then
        Call IniRIS
    End If
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim i As Integer
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    err = 0: On Error Resume Next
    
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 2000 Then .Top = 2000
        If .Top > Me.ScaleHeight - lngStatus - 2500 Then .Top = Me.ScaleHeight - lngStatus - 2500
    End With
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = lngTools
        .Height = Me.picHBar.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Me.tabContent
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = Me.picHBar.Top + Me.picHBar.Height
        .Height = Me.ScaleHeight - lngStatus - .Top + 15
        .Width = Me.ScaleWidth - .Left + 15
    End With
    
    For intCount = 0 To Me.tabContent.Tabs - 1
        With Me.fraSubInfo(intCount)
            .Left = 90
            .Top = 325
            .Width = Me.tabContent.Width - .Left * 2
            .Height = Me.tabContent.Height - .Top - 90
        End With
    Next
    With Me.hgd����ִ��
'        .Visible = Me.fraSubInfo(0).Visible
        .Left = Me.fraSubInfo(0).Left + Me.lbl����ִ��.Left
        .Width = Me.fraSubInfo(0).Left + Me.fraSubInfo(0).Width - .Left - 100
        .Top = Me.fraSubInfo(0).Top + Me.lbl����ִ��.Top + Me.lbl����ִ��.Height + 45
        .Height = Me.fraSubInfo(0).Top + Me.fraSubInfo(0).Height - .Top - 100
    End With
    With vsfExse
        .Left = 30: .Top = 30
        .Width = fraSubInfo(1).Width - 60
        .Height = fraSubInfo(1).Height - 60
    End With
    
    With Me.hgdLabs
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(2).Width: .Height = Me.fraSubInfo(2).Height - .Top
    End With
    With Me.hgdPart
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(3).Width: .Height = Me.fraSubInfo(3).Height - .Top
    End With
    With Me.hgdUsage
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(4).Width: .Height = Me.fraSubInfo(4).Height - .Top
    End With
    With Me.hgdTabu
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(5).Width: .Height = Me.fraSubInfo(5).Height - .Top
    End With
    With Me.hgdRecipe
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(6).Width: .Height = Me.fraSubInfo(6).Height - .Top
    End With
    With Me.vsScheme '���׷���(ByZT)
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(7).Width: .Height = Me.fraSubInfo(7).Height - .Top
    End With

    With Me.hgdRefer
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(6).Width: .Height = Me.fraSubInfo(6).Height - .Top
        .Redraw = False
        .ColWidth(0) = 0
        .ColWidth(1) = Me.TextWidth("�ո�")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
        Call zlGrdRowHeight
        .Redraw = True
    End With
    clbThis.Bands(1).Width = Me.Width - 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\�ָ�", "����", Me.picHBar.Top)
    
    Call zlDatabase.SetPara("��ʾͣ����Ŀ", IIf(Me.mnuViewStoped.Checked, 1, 0), glngSys, 1054)
    
    mint��Χ = 3 'ֱ��ͨ��������Ŀ�������ʱ��ȱʡΪ3
    
    If Not gobjRIS Is Nothing Then
        Set gobjRIS = Nothing
    End If
End Sub

Private Sub mnuAddcopy_Click()
    Dim blnOk As Boolean
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾ��Ŀ��", vbExclamation, gstrSysName: Exit Sub
    If Val(Me.tvwClass.Tag) = 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then
            MsgBox "��ѡ����Ҫ���Ƶ���Ŀ��", vbInformation, gstrSysName
        Else
            blnOk = frmClinicItem.ShowMe(Me, 3, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
        End If
    End If
    If blnOk Then Call zlRefRecords
End Sub

Private Sub tlbThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If tvwClass.Tag <> SK_������Ŀ Then
        Button.ButtonMenus(2).Visible = False
    Else
        Button.ButtonMenus(2).Visible = True
    End If
End Sub

Private Sub vsfExse_DblClick()
    Dim i As Integer
    Dim strIDS As String
    
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    For i = Me.lvwItems.SelectedItem.Index + 1 To lvwItems.ListItems.Count
        strIDS = strIDS & Mid(Me.lvwItems.ListItems(i).Key, 2) & ","
    Next
    
    Call frmClinicExse.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2), strIDS)
End Sub

Private Sub hgdLabs_DblClick()
'    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
'    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
'    Call frmClinicLabs.ShowME(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub hgdPart_DblClick()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicPart.ShowMe(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub hgdRecipe_DblClick()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmMediRecipe.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub hgdRefer_DblClick()
    If Me.mnuEditRefer.Enabled = False Then Exit Sub
End Sub

Private Sub mnuEditImport_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,���ܵ�����Ŀ��", vbExclamation, gstrSysName: Exit Sub
    With frmClinicLoad
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefRecords
End Sub

'Private Sub mnuEditSample_Click()
'    '2007-04-17 ȥ���걾���չ���
''    If Me.lvwItems.ListItems.Count > 0 Then
''        Call frmClinicVerifySample.ShowMe(Me, Mid(Me.lvwItems.SelectedItem.Key, 2))
''    End If
'End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����id����Ŀ=��Ŀid�����=�����������
    Dim lng����id As Long
    Dim lng��Ŀid As Long
    Dim str��� As String
    
    If Not Me.tvwClass.SelectedItem Is Nothing Then
        lng����id = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng��Ŀid = Mid(Me.lvwItems.SelectedItem.Key, 2)
        str��� = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_���").Index - 1)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIf(lng����id = 0, "", lng����id), _
        "��Ŀ=" & IIf(lng��Ŀid = 0, "", lng��Ŀid), _
        "���=" & str���)
End Sub

Private Sub mnuViewShowAll_Click()
    On Error GoTo ErrHandle
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    If tvwClass.SelectedItem Is Nothing Then
        If tvwClass.Nodes.Count > 0 Then
            MsgBox "��ѡ��һ�·��࣡", vbInformation, gstrSysName
        Else
            MsgBox "���κη������ʾ��", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    Call zlRefRecords
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tabCharge_Click(PreviousTab As Integer)
    vsfExse.ZOrder 0
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim strSQLTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If txtFind.Text = "" Then Exit Sub
    If zlCommFun.IsCharChinese(txtFind.Text) Then
        strSQLTmp = " And Upper(Nvl(b.����, a.����)) Like [1]"
    ElseIf IsNumeric(txtFind.Text) Then
        strSQLTmp = " And a.���� Like [2]"
    Else
        strSQLTmp = " And (Upper(Nvl(b.����, a.����)) Like [1] Or b.���� Like [3])"
    End If
    
    On Error GoTo ErrHandle
    strSql = "Select Distinct a.Id, a.���, Nvl(b.����, a.����) As ����, a.����, b.����, c.���� As ����, a.����id, a.����ʱ��" & vbNewLine & _
            "From (Select ID, ���, ����id, ����, ����, ����ʱ��" & vbNewLine & _
            "       From ������ĿĿ¼" & vbNewLine & _
            "       Where ��� Not In ('4', '5', '6', '7') And ��� >= 'A' And" & vbNewLine & _
            "             (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)) A," & vbNewLine & _
            "     (Select Distinct a.������Ŀid, a.����, a.���� As ƴ����, b.���� As �����, a.���� || '/' || b.���� As ����" & vbNewLine & _
            "       From ������Ŀ���� A, ������Ŀ���� B" & vbNewLine & _
            "       Where a.������Ŀid = b.������Ŀid And a.���� = 1 And b.���� = 2" & _
            IIf(zlCommFun.IsCharChinese(txtFind.Text), " And Upper(a.����) Like [1] ", "") & _
            " And a.���� = 1 And b.���� = 1) B," & vbNewLine & _
            "     ���Ʒ���Ŀ¼ C" & vbNewLine & _
            "Where a.����id = c.Id(+) And a.Id = b.������Ŀid(+) And c.���� Is Not Null And C.���� In (4,5,6) " & _
            strSQLTmp
    
    vRect = zlControl.GetControlRect(txtFind.hwnd)
    If vRect.Left + 7000 > Screen.Width Then vRect.Left = Screen.Width - 7000
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "�շ�ϸĿѡ��", False, "", "", False, False, True, _
                        vRect.Left, vRect.Top, txtFind.Height, blnCancel, False, True, IIf(gstrMatch = "", "", "%") & txtFind.Text & "%", txtFind.Text & "%", IIf(gstrMatch = "", "", "%") & UCase(txtFind.Text) & "%")
    If blnCancel = True Then Exit Sub
    If Not rsTmp Is Nothing Then
        Call FindLocate(rsTmp)
    Else
        MsgBox "û���ҵ��������ҵ��շ���Ŀ��", vbInformation, Me.Caption
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindLocate(ByVal rsTmp As Recordset)
    Dim strkey As String
    Dim strItemKey As String
    
    '81291--�����ǻ������з�����в��ң���˲���Ҫ�жϵ�ǰ�������Ƿ�������
'    If lvwItems.SelectedItem Is Nothing Then Exit Sub

    On Error Resume Next
    With lvwItems.SelectedItem
        strkey = "_" & IIf(IsNull(rsTmp("����ID")), "", rsTmp("����ID"))
        strItemKey = "_" & rsTmp("id")
        If .SubItems(3) <> "δ����" Then
            Me.tvwClass.Nodes(strkey).Selected = True
            Me.tvwClass.Nodes(strkey).EnsureVisible
            Me.tvwClass_NodeClick Me.tvwClass.SelectedItem
            err.Clear
            Me.lvwItems.ListItems(strItemKey).Selected = True
            Me.lvwItems.ListItems(strItemKey).EnsureVisible
            If err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                err.Clear
                Exit Sub
            End If
            Me.lvwItems_ItemClick Me.lvwItems.SelectedItem
        Else
            Me.tvwClass.Nodes("Root").Selected = True
            Me.tvwClass.Nodes(strkey).EnsureVisible
            Me.tvwClass_NodeClick Me.tvwClass.SelectedItem
            err.Clear
            Me.lvwItems.ListItems(strItemKey).Selected = True
            Me.lvwItems.ListItems(strItemKey).EnsureVisible
            If err.Number = 35601 Then
                MsgBox "���ҵ���������¼�����ѱ�ɾ����ͣ�ã���ˢ���б�", vbInformation, gstrSysName
                err.Clear
                Exit Sub
            End If
            Me.lvwItems_ItemClick Me.lvwItems.SelectedItem
        End If
    End With
    err.Clear
End Sub

 

Private Sub vsScheme_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsScheme.FixedRows And NewCol >= vsScheme.FixedCols Then
        If NewRow <> OldRow Then
            vsScheme.ForeColorSel = vsScheme.CellForeColor
        End If
    End If
End Sub

Private Sub mnuEditGather_Click()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwItems.SelectedItem.Tag <> "C" Then Exit Sub
    Call frmLabsUsage.ShowMe(Me, Not Me.lvwItems.SelectedItem.Icon = "����S", Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub vsScheme_DblClick()
    '���ĳ��׷���(ByZT)
    If Val(Me.tvwClass.Tag) <> 2 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicScheme.ShowMe(Me, mstrPrivs, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint��Χ)
End Sub

Private Sub hgdUsage_DblClick()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmMediUsage.ShowMe(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        Call frmClinicItem.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 1
        Call frmMediRecipe.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 2
        '���ĳ��׷���(ByZT)
        Call frmClinicScheme.ShowMe(Me, mstrPrivs, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint��Χ)
    End Select
End Sub

Public Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Long, j As Long
    Dim iRow As Integer, iCol As Integer
    Dim lngForeColor As Long
    Dim strTmp As String
    Dim intIndex As Integer
    Dim lngCol As Long
    Dim lngRow As Long
    '------------------------------------------------
    '������ϸ��Ϣ��ʾ��
    Call zlClearDetail
    
    err = 0: On Error GoTo ErrHand
    '------------------------------------------------
    'ִ�п�����ʾ
    If Val(Me.tvwClass.Tag) = 0 Then
        Me.lblUseBill.Caption = "���Ƶ��ݣ�"
        gstrSql = "Select A.Ӧ�ó���,B.����" & _
                " From �����ļ��б� B,��������Ӧ�� A" & _
                " Where B.ID=A.�����ļ�id And A.������Ŀid=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            Do While Not .EOF
                Select Case !Ӧ�ó���
                Case 1
                    Me.lblUseBill.Caption = Me.lblUseBill.Caption & "�������" & !���� & "��"
                Case 2
                    Me.lblUseBill.Caption = Me.lblUseBill.Caption & "סԺ����" & !���� & "��"
                Case 4
                    Me.lblUseBill.Caption = Me.lblUseBill.Caption & "������" & !���� & "��"
                End Select
                
                .MoveNext
            Loop
        End With
        
        gstrSql = "select ִ�п��� from ������ĿĿ¼ where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        

        Me.optִ�в���(0).Value = False
        If rsTemp.RecordCount > 0 Then Me.optִ�в���(IIf(IsNull(rsTemp!ִ�п���), 0, rsTemp!ִ�п���)).Value = True
        For i = 0 To optִ�в���.Count - 1
            optִ�в���(i).Enabled = optִ�в���(i).Value
        Next
        
        gstrSql = "select R.������Դ,E.ID,E.����" & _
                " from ����ִ�п��� R,���ű� E" & _
                " where R.ִ�п���ID=E.ID and R.������Դ in (1,2) and R.��������id is null and R.������ĿID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
            
        With rsTemp
            strTemp = ""
            Do While Not .EOF
                If !������Դ = 1 Then strTemp = strTemp & "������" & !���� & "ִ�У�"
                If !������Դ = 2 Then strTemp = strTemp & "סԺ��" & !���� & "ִ�У�"
                .MoveNext
            Loop
        End With
        
        Me.lbl����ִ��.Caption = ""
        If strTemp <> "" Then Me.lbl����ִ��.Caption = "һ��" & strTemp
        
        gstrSql = "select K.���� as ������������,E.���� as ִ�в�������" & _
                " from ����ִ�п��� R,���ű� K,���ű� E" & _
                " where R.��������ID=K.ID(+) and R.ִ�п���ID=E.ID and nvl(R.������Դ,0)=0 and R.������ĿID=[1] " & _
                " order by e.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        Me.hgd����ִ��.Redraw = False
        i = 0
         With rsTemp
            Do While Not .EOF
'                If Me.hgd����ִ��.Rows - 1 < .AbsolutePosition Then Me.hgd����ִ��.Rows = Me.hgd����ִ��.Rows + 1
'                Me.hgd����ִ��.TextMatrix(.AbsolutePosition, 0) = !������������
'                Me.hgd����ִ��.TextMatrix(.AbsolutePosition, 1) = !ִ�в�������
                
                If strTmp <> !ִ�в������� Then
                    i = i + 1
                    Me.hgd����ִ��.Rows = i + 1
                    Me.hgd����ִ��.TextMatrix(i, 1) = IIf(IsNull(!������������), "�����в��ţ�", !������������)
                    Me.hgd����ִ��.TextMatrix(i, 0) = !ִ�в�������
                Else
                    Me.hgd����ִ��.TextMatrix(i, 1) = Me.hgd����ִ��.TextMatrix(i, 1) & "," & !������������
                End If
                
                strTmp = !ִ�в�������
                
                .MoveNext
            Loop
            Me.hgd����ִ��.Redraw = True
        End With
    End If
    
    '------------------------------------------------
    '�շѶ�����ʾ
    If Val(Me.tvwClass.Tag) = 0 Then
         
        
        
        gstrSql = "" & _
            " Select i.Id, r.��鲿λ, r.��鷽��, r.��������, '[' || i.���� || ']' || i.���� As ����, i.���, i.���㵥λ," & vbNewLine & _
            "       I.�Ƿ���,sum(decode(P.�۸�ȼ�,NULL,p.�ּ�,0)) as ȱʡ�۸�" & mstrPriceGradeFields & _
            "       ,Nvl(r.�շ�����, 0) As ����, Nvl(r.���ж���, 0) As �̶�," & vbNewLine & _
            "       Nvl(r.������Ŀ, 0) As ����, Nvl(i.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ��, Nvl(r.�շѷ�ʽ, 0) As �շѷ�ʽ, r.������Դ," & vbNewLine & _
            "       b.���� As ���ÿ��ұ���, b.���� As ���ÿ�������" & vbNewLine & _
            " From �����շѹ�ϵ R, �շ���ĿĿ¼ I, �շѼ�Ŀ P, ���ű� B" & vbNewLine & _
            " Where r.�շ���Ŀid = i.Id And i.Id = p.�շ�ϸĿid(+) And r.���ÿ���id = b.Id(+)  " & vbNewLine & _
            "       And p.ִ������ <= Sysdate And (p.��ֹ���� Is Null Or p.��ֹ���� >= Sysdate)  " & vbNewLine & _
            "      And r.������Ŀid = [1]" & vbNewLine & _
            "Group By i.Id, r.��鲿λ, r.��鷽��, r.��������, i.����, i.����, i.���, i.���㵥λ, i.�Ƿ���, r.�շ�����, r.���ж���, r.������Ŀ, i.����ʱ��, r.�շѷ�ʽ," & vbNewLine & _
            "         r.������Դ, b.����, b.����" & vbNewLine & _
            "order by r.������Դ,B.����,nvl(R.������Ŀ,0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)), gstrPriceClass)
        With vsfExse
            .Redraw = flexRDNone
            Do While Not rsTemp.EOF
                intIndex = Val(NVL(rsTemp!������Դ))
                i = .Rows - 1: .Rows = .Rows + 1
               .TextMatrix(i, .ColIndex("ѡ��")) = i
               .TextMatrix(i, .ColIndex("��λ")) = NVL(rsTemp!��鲿λ)
               .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!��鷽��)
               .TextMatrix(i, .ColIndex("��Ŀ��")) = NVL(rsTemp!����)
               .TextMatrix(i, .ColIndex("���")) = NVL(rsTemp!���)
               .TextMatrix(i, .ColIndex("��λ")) = NVL(rsTemp!���㵥λ)
               .TextMatrix(i, .ColIndex("�۸�")) = IIf(Val(NVL(rsTemp!�Ƿ���)) = 1, "���", Val(NVL(rsTemp!ȱʡ�۸�)))
               .TextMatrix(i, .ColIndex("����")) = FormatEx(Format(rsTemp!����, "0.00000"), 5)
               .TextMatrix(i, .ColIndex("�̶�")) = IIf(rsTemp!�̶� = 0, "", "��")
               .TextMatrix(i, .ColIndex("����")) = IIf(rsTemp!���� = 0, "", "��")
               .TextMatrix(i, .ColIndex("����")) = IIf(0 + rsTemp!�������� = 1, "��", "")
               .TextMatrix(i, .ColIndex("״̬")) = IIf(Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01", "ͣ��", "")
                
                Select Case rsTemp!�շѷ�ʽ
                Case 0
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "0-������ȡ"
                Case 1
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "1-�����Թܷ���"
                Case 2
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "2-һ�η���ֻ��ȡһ��"
                Case 3
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "3-����ֻ��ȡһ��"
                Case 4
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "4-����δִ����ȡһ��"
                Case 5
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "5-����ֻ��ȡһ�Σ��ų�������Ŀ"
                Case 6
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "6-����δִ����ȡһ�Σ��ų�������Ŀ"
                Case 7
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "7-ÿ���״β���ȡ"
                Case 9
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "9-�Զ���"
                Case Else
                   .TextMatrix(i, .ColIndex("�շѷ�ʽ")) = "0-������ȡ"
                End Select
                
                Select Case Val(NVL(rsTemp!������Դ))
                Case 0: .TextMatrix(i, .ColIndex("���ó���")) = "���п���"
                Case 1: .TextMatrix(i, .ColIndex("���ó���")) = "�������"
                Case 2: .TextMatrix(i, .ColIndex("���ó���")) = "סԺ����"
                Case 3: .TextMatrix(i, .ColIndex("���ó���")) = "������"
                End Select
                
                If Trim(NVL(rsTemp!���ÿ�������)) <> "" Then
                   .TextMatrix(i, .ColIndex("���ÿ���")) = "" & rsTemp!���ÿ������� & "(" & rsTemp!���ÿ��ұ��� & ")"
                End If
                '���ؼ۸�ȼ�����Ӧ�ļ۸�
                For intCol = 0 To .Cols - 1
                    If Left(CStr(.colData(intCol)), 1) = "A" Then
                        If Val(NVL(rsTemp!�Ƿ���)) = 1 Then
                            .TextMatrix(i, intCol) = "���"
                        Else
                            If Val(NVL(rsTemp.Fields(CStr(.colData(intCol))))) = 0 Then
                                .TextMatrix(i, intCol) = .TextMatrix(i, .ColIndex("�۸�"))
                            Else
                                .TextMatrix(i, intCol) = Val(NVL(rsTemp.Fields(CStr(.colData(intCol)))))
                            End If
                        End If
                    End If
                Next
                If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    lngForeColor = &HFF&
                Else
                    lngForeColor = &H0&
                End If
                iRow = .Row: iCol = .Col
               .Row = i
               .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = lngForeColor
               .Row = iRow: .Col = iCol
                rsTemp.MoveNext
            Loop
            .Redraw = flexRDBuffered
            If .Rows > 2 Then .Rows = .Rows - 1
            .MergeCells = flexMergeFree
            .MergeCol(.ColIndex("���ó���")) = True
        End With
        
    End If
     
    '------------------------------------------------
    '��鲿λ��ʾ
    If Val(Me.tvwClass.Tag) = 0 And Item.Tag = "D" Then
        gstrSql = "select ID from ������ĿĿ¼ I where I.ID=[1] and I.�����Ŀ=1 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            If .EOF Then
                Me.tabContent.TabVisible(conTab��鲿λ) = False
            Else
                Me.tabContent.TabVisible(conTab��鲿λ) = True
            End If
        End With
    Else
        Me.tabContent.TabVisible(conTab��鲿λ) = False
    End If
    If Me.tabContent.TabVisible(conTab��鲿λ) = True Then
        Me.hgdPart.Redraw = False
        gstrSql = "select I.ID,I.���� as ����,I.�걾��λ" & _
                " from ������Ŀ��� R,������ĿĿ¼ I" & _
                " where R.������ĿID=I.ID and R.�������ID=[1] " & _
                " order by R.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            Do While Not .EOF
                If Me.hgdPart.Rows - 1 < .AbsolutePosition Then Me.hgdPart.Rows = Me.hgdPart.Rows + 1
                Me.hgdPart.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                Me.hgdPart.TextMatrix(.AbsolutePosition, 1) = !����
                Me.hgdPart.TextMatrix(.AbsolutePosition, 2) = !�걾��λ
                .MoveNext
            Loop
        End With
        Me.hgdPart.Redraw = True
    End If
    '------------------------------------------------
    
    '------------------------------------------------
    '�䷽�����ʾ
    If Val(Me.tvwClass.Tag) = 1 Then
        Me.hgdRecipe.Redraw = False
        gstrSql = "Select b.���, b.������ĿId As ҩ��id, b.�շ�ϸĿId As ���id, a.����, c.���, a.���㵥λ, b.��������, b.ҽ������ " & vbNewLine & _
            "From ������ĿĿ¼ A, ������Ŀ��� B, �շ���ĿĿ¼ C " & vbNewLine & _
            "Where a.Id = b.������Ŀid And b.�շ�ϸĿid = c.Id(+) And b.�������id = [1] " & vbNewLine & _
            "Order By b.��� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            Do While Not .EOF
                If Me.hgdRecipe.Rows - 1 < ((.AbsolutePosition - 1) \ mbyt��ҩζ��) + 1 Then Me.hgdRecipe.Rows = Me.hgdRecipe.Rows + 1
                intCount = (.AbsolutePosition - 1) Mod mbyt��ҩζ��
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt��ҩζ�� + 1, intCount * 6 + 2) = !���� & IIf(IsNull(!���), "", "(" & !��� & ")")
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt��ҩζ�� + 1, intCount * 6 + 3) = IIf(IsNull(!��������), 0, !��������)
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt��ҩζ�� + 1, intCount * 6 + 4) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt��ҩζ�� + 1, intCount * 6 + 5) = IIf(IsNull(!ҽ������), "", !ҽ������)
                .MoveNext
            Loop
        End With
        
        For lngRow = 1 To hgdRecipe.Rows - 1
            hgdRecipe.Row = lngRow
            For lngCol = 0 To hgdRecipe.Cols - 1
                hgdRecipe.Col = lngCol
                
                If lngCol < 6 Or (lngCol > 12 And lngCol < 20) Then
                    hgdRecipe.CellBackColor = &H8000000F
                End If
            Next
        Next
        
        gstrSql = "select I.���� ,R.����,P.���� as Ƶ��,R.�Ƴ�" & _
                " from �����÷����� R,������ĿĿ¼ I,����Ƶ����Ŀ P" & _
                " where R.�÷�ID=I.ID and R.Ƶ��=P.����(+) and R.��ĿID=[1] " & _
                " order by R.���� desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
            
        With rsTemp
            strTemp = ""
            Do While Not .EOF
                If .AbsolutePosition = 1 Then strTemp = strTemp & Space(3) & IIf(IsNull(!Ƶ��), "", !Ƶ��)
                strTemp = strTemp & Space(3) & !����
                .MoveNext
            Loop
            With Me.hgdRecipe
                .Rows = .Rows + 2: .MergeRow(.Rows - 1) = True
                For intCount = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, intCount) = Trim(strTemp)
                Next
            End With
        End With
        Me.hgdRecipe.Redraw = True
    End If
    
    '------------------------------------------------
    '���׷�����ʾ(ByZT)
    If Val(Me.tvwClass.Tag) = 2 Then
        Call ShowScheme(Val(Mid(Item.Key, 2)))
    End If
    
    '------------------------------------------------
    '���ò˵��͹������Ľ�ֹ��
    If Item.ForeColor = &HFF& Then
        '�Ѿ���ֹ����Ŀ����ɾ��
        Me.mnuEditDelete.Enabled = False
        Me.mnuEditModify.Enabled = False
        '��鲿λ
        'Me.mnuEditExams.Enabled = False
        '����ָ��
        Me.mnuEditLabs.Enabled = False
        
        Me.mnuEditGather.Enabled = False
        '�걾����
        'Me.mnuEditSample.Enabled = False
        '�շѶ���
        Me.mnuEditExse.Enabled = False
        '�ų��ϵ
        Me.mnuEditRepellent.Enabled = False
        'Ӧ�òο�
'        Me.mnuEditRefer.Enabled = False
        '��Ӧ����
        Me.mnuEditBill.Enabled = False
        
        '�����ٽ�ֹ,ֻ������
        Me.mnuEditStart.Enabled = (tvwClass.Tag = SK_������Ŀ And InStr(1, mstrPrivs, "������Ŀ�༭") > 0) Or _
                            (tvwClass.Tag = SK_��ҩ�䷽ And InStr(1, mstrPrivs, "��ҩ�䷽�༭") > 0) Or _
                            (tvwClass.Tag = SK_���׷��� And InStr(1, mstrPrivs, "���׷����༭") > 0)
        Me.mnuEditStop.Enabled = False
    Else
        '����ɾ�����޸�
        Me.mnuEditDelete.Enabled = (tvwClass.Tag = SK_������Ŀ And InStr(1, mstrPrivs, "������Ŀ�༭") > 0) Or _
                            (tvwClass.Tag = SK_��ҩ�䷽ And InStr(1, mstrPrivs, "��ҩ�䷽�༭") > 0) Or _
                            (tvwClass.Tag = SK_���׷��� And InStr(1, mstrPrivs, "���׷����༭") > 0)
        Me.mnuEditModify.Enabled = Me.mnuEditDelete.Enabled
        
        '�շѶ���
        Me.mnuEditExse.Enabled = (InStr(1, mstrPrivs, "�շ�����") > 0)
        '�ų��ϵ
        Me.mnuEditRepellent.Enabled = True
        'Ӧ�òο�
'        Me.mnuEditRefer.Enabled = (InStr(1, mstrPrivs, "�ο��༭") > 0)
        '��Ӧ����
        Me.mnuEditBill.Enabled = Me.mnuEditDelete.Enabled
        'ֻ��ͣ��
        Me.mnuEditStart.Enabled = False
        Me.mnuEditStop.Enabled = Me.mnuEditDelete.Enabled
        '�������ֱ��жϽ�ֹ
        If Val(Me.tvwClass.Tag) = 0 And lvwItems.SelectedItem.Tag = "C" Then
            'Me.mnuEditExams.Enabled = False
            Me.mnuEditLabs.Enabled = Me.mnuEditDelete.Enabled
            Me.mnuEditGather.Enabled = Me.mnuEditDelete.Enabled
            'Me.mnuEditSample.Enabled = (InStr(1, mstrPrivs, "��Ŀ�༭") > 0)
        ElseIf Val(Me.tvwClass.Tag) = 0 And lvwItems.SelectedItem.Tag = "D" Then
            'Me.mnuEditExams.Enabled = (InStr(1, mstrPrivs, "��Ŀ�༭") > 0)
            Me.mnuEditLabs.Enabled = False
            Me.mnuEditGather.Enabled = False
            'Me.mnuEditSample.Enabled = False
        Else
            'Me.mnuEditExams.Enabled = False
            Me.mnuEditLabs.Enabled = False
            Me.mnuEditGather.Enabled = False
            'Me.mnuEditSample.Enabled = False
        End If
    End If
    
    Me.tlbThis.Buttons("Start").Enabled = Me.mnuEditStart.Enabled
    Me.tlbThis.Buttons("Stop").Enabled = Me.mnuEditStop.Enabled
    Me.tlbThis.Buttons("Delete").Enabled = Me.mnuEditDelete.Enabled
    Me.tlbThis.Buttons("Modify").Enabled = Me.mnuEditModify.Enabled
    Call zlGrdRowHeight
    Me.hgdRefer.Redraw = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Not lvwItems.SelectedItem Is Nothing Then
            Call PopupMenu(Me.mnuEdit, 2)
        End If
    End If
End Sub

Private Sub mnuClassAdd_Click()
    With frmClinicClass
        strTemp = Switch(Me.tvwClass.Tag = "0", 5, _
                       Me.tvwClass.Tag = "1", 4, _
                       Me.tvwClass.Tag = "2", 6)
        .lblKind.Tag = strTemp
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "����"
        
        If Me.tvwClass.SelectedItem Is Nothing Then
            If .ShowMe(1, Me, "(��)", 0, 1, True) Then
                Call zlRefClasses
            End If
        Else
            If .ShowMe(1, Me, Me.tvwClass.SelectedItem.Text, Mid(Me.tvwClass.SelectedItem.Key, 2), 1, True) Then
                Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
            End If
        End If
    End With
End Sub

Private Sub mnuClassDel_Click()
    err = 0: On Error GoTo ErrHand
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("���ɾ���÷��ࡰ" & Me.tvwClass.SelectedItem.Text & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "zl_���Ʒ���Ŀ¼_delete(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Dim strParentKey As String
    If Me.tvwClass.SelectedItem.Next Is Nothing Then
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            Call zlRefClasses
        Else
            strParentKey = Me.tvwClass.SelectedItem.Parent.Key
            Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
            If Me.tvwClass.Nodes(strParentKey).Children = 0 Then
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Key, 2))
            Else
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Child.Key, 2))
            End If
        End If
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Next.Key, 2))
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With frmClinicClass
        strTemp = Switch(Me.tvwClass.Tag = "0", 5, _
                       Me.tvwClass.Tag = "1", 4, _
                       Me.tvwClass.Tag = "2", 6)
        .lblKind.Tag = strTemp
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(��)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
        .txtSymbol = Me.tvwClass.SelectedItem.Tag
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            If .ShowMe(1, Me, "(��)", 0, 2, True) Then Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
        Else
            If .ShowMe(1, Me, Me.tvwClass.SelectedItem.Parent.Text, Mid(Me.tvwClass.SelectedItem.Parent.Key, 2), 2, True) Then Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim blnOk As Boolean
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾ��Ŀ��", vbExclamation, gstrSysName: Exit Sub
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        If Me.lvwItems.SelectedItem Is Nothing Then
            blnOk = frmClinicItem.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), 0)
        Else
            blnOk = frmClinicItem.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
        End If
    Case 1
        If Me.lvwItems.SelectedItem Is Nothing Then
            blnOk = frmMediRecipe.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), 0)
        Else
            blnOk = frmMediRecipe.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
        End If
    Case 2 '�������׷���(ByZT)
        If Me.lvwItems.SelectedItem Is Nothing Then
            blnOk = frmClinicScheme.ShowMe(Me, mstrPrivs, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), 0, mint��Χ)
        Else
            blnOk = frmClinicScheme.ShowMe(Me, mstrPrivs, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint��Χ)
        End If
    End Select
    If blnOk Then Call zlRefRecords
End Sub

Private Sub mnuEditBill_Click()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicBill.ShowMe(Me, Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditDelete_Click()
    Dim lngVItemID As Long '����������ID
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim blnRisTrans As Boolean
    
'    If Val(Me.tvwClass.Tag) >= 1 And Val(Me.tvwClass.Tag) <= 1 Then Exit Sub
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("���ɾ����" & .SelectedItem.Text & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "Select �����Ŀ,������ĿID From ������ĿĿ¼ A,���鱨����Ŀ B Where A.ID=B.������ĿID And A.ID=[1] "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(.SelectedItem.Key, 2)))
        lngVItemID = 0
        If rsTmp.RecordCount = 1 Then If rsTmp(0) = 0 Then lngVItemID = rsTmp(1)
        
        err = 0: On Error GoTo ErrHand
                        
        '����RIS�ӿڣ����ò���������顱����Ŀ���ӿڲ�����Ч��ǰ����
        If mblnPACSInterface = True And .SelectedItem.Tag = "D" Then
            If Not gobjRIS Is Nothing Then
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItem, RISBaseItemOper.Delete, Val(Mid(.SelectedItem.Key, 2))) <> 1 Then
                    '����ʱ��ʾ�ӿڴ�����Ϣ
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                    End If
                    
                    Exit Sub
                End If
                blnRisTrans = True
            Else
               '�ӿڲ�����Чʱ��ֹ����ʾ
                MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                
                Exit Sub
            End If
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        
        gstrSql = "zl_������Ŀ_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        If lngVItemID > 0 Then
            gstrSql = "zl_������Ŀ_DELETE(" & lngVItemID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        blnRisTrans = False
        
        Call .ListItems.Remove(.SelectedItem.Key)
        If .SelectedItem Is Nothing Then
            Call zlClearDetail
        Else
            Call lvwItems_ItemClick(.SelectedItem)
        End If
    End With
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    
    Call ErrCenter
    Call SaveErrLog
    
    'Ris�ӿں�HIS��ͬ��ʱ��д������־
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HISɾ��������Ŀ����RIS�ӿں�HIS���ݲ�ͬ��������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmClinicLists��mnuEditDelete_Click", "HISɾ��������Ŀ����RIS�ӿں�HIS���ݲ�ͬ��", "������ĿID=" & Val(Mid(lvwItems.SelectedItem.Key, 2)), 0)
    End If

End Sub

'Private Sub mnuEditExams_Click()
'    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
'    If Me.lvwItems.SelectedItem Is Nothing Then
'        Call frmClinicPart.ShowMe(Me, True)
'    ElseIf Me.lvwItems.SelectedItem.Tag <> "D" Then
'        Call frmClinicPart.ShowMe(Me, True)
'    Else
'        If Me.lvwItems.SelectedItem.Icon = "����S" Then MsgBox "ͣ����Ŀ���������ò�λ��ϣ�", vbExclamation, gstrSysName
'        Call frmClinicPart.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2))
'    End If
'    If Not Me.lvwItems.SelectedItem Is Nothing Then
'        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
'    End If
'End Sub

Private Sub mnuEditExse_Click()
    Dim i As Integer
    Dim strIDS As String   '���浱ǰѡ����֮�������id��
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call frmClinicExse.ShowMe(Me, True)
    Else
        If Me.lvwItems.SelectedItem.Icon = "����S" Then MsgBox "ͣ����Ŀ�����������շѶ��գ�", vbExclamation, gstrSysName
        For i = Me.lvwItems.SelectedItem.Index + 1 To lvwItems.ListItems.Count
            strIDS = strIDS & Mid(Me.lvwItems.ListItems(i).Key, 2) & ","
        Next
        
        Call frmClinicExse.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2), strIDS)
        Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub


Private Sub mnuEditLabs_Click()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call frmClinicLabs.ShowMe(Me, True)
    ElseIf Me.lvwItems.SelectedItem.Tag <> "C" Then
        Call frmClinicLabs.ShowMe(Me, True)
    Else
        If Me.lvwItems.SelectedItem.Icon = "����S" Then MsgBox "ͣ����Ŀ���������ü���ָ�꣡", vbExclamation, gstrSysName
        Call frmClinicLabs.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2))
        Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub mnuEditModify_Click()
    Dim blnOk As Boolean
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "��δ���÷���,������ɾ��Ŀ��", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        If Me.lvwItems.SelectedItem.Icon = "����S" Then MsgBox "���ܶ�ͣ����Ŀ�����޸ģ�", vbExclamation, gstrSysName
        blnOk = frmClinicItem.ShowMe(Me, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 1
        If Me.lvwItems.SelectedItem.Icon = "����S" Then MsgBox "���ܶ�ͣ���䷽�����޸ģ�", vbExclamation, gstrSysName
        blnOk = frmMediRecipe.ShowMe(Me, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 2
        If Me.lvwItems.SelectedItem.Icon = "����S" Then MsgBox "���ܶ�ͣ�÷��������޸ģ�", vbExclamation, gstrSysName
        '�޸ĳ��׷���(ByZT)
        blnOk = frmClinicScheme.ShowMe(Me, mstrPrivs, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint��Χ)
    End Select
    If blnOk Then Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub mnuEditRepellent_Click()
    If Val(Me.mnuEditRepellent.Tag) = 0 Then
        Call frmClinicTabu.ShowMe(Me, False)
    Else
        Call frmClinicTabu.ShowMe(Me, True)
    End If
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub mnuEditStart_Click()
    Dim iSubItemIndex As Integer
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If Right(.SelectedItem.Icon, 1) = "U" Then Exit Sub
        strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon))
        
        If MsgBox("����������á�" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "zl_������Ŀ_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        If Val(Me.tvwClass.Tag) = 0 Then
            strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon) - 1)
            .SelectedItem.Icon = strTemp & "U": .SelectedItem.SmallIcon = strTemp & "U"
        Else
            .SelectedItem.Icon = "����U": .SelectedItem.SmallIcon = "����U"
        End If
            
        '�ָ�������Ŀ��ʾ��ɫ����ͮ��
        .SelectedItem.ForeColor = .ForeColor
        For iSubItemIndex = 1 To .ColumnHeaders.Count - 1
            .SelectedItem.ListSubItems(iSubItemIndex).ForeColor = .ForeColor
        Next
    End With
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim iSubItemIndex As Integer
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If Right(.SelectedItem.Icon, 1) = "S" Then Exit Sub
        strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon))
        
        If MsgBox("���Ҫͣ�á�" & .SelectedItem.Text & "����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
        gstrSql = "zl_������Ŀ_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        If Me.mnuViewStoped.Checked = True Then
            If Val(Me.tvwClass.Tag) = 0 Then
                strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon) - 1)
                .SelectedItem.Icon = strTemp & "S": .SelectedItem.SmallIcon = strTemp & "S"
            Else
                .SelectedItem.Icon = "����S": .SelectedItem.SmallIcon = "����S"
            End If
            
            '��ͣ����Ŀ��ʾΪ��ɫ����ͮ��
            .SelectedItem.ForeColor = &HFF&
            For iSubItemIndex = 1 To .ColumnHeaders.Count - 1
                .SelectedItem.ListSubItems(iSubItemIndex).ForeColor = &HFF&
            Next
        Else
            Call .ListItems.Remove(.SelectedItem.Key)
        End If
    End With
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePara_Click()
    Call frmClinicPara.ShowMe(Me, mstrPrivs)
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuViewFind_Click()
    frmClinicFind.Show , Me
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStoped_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.mnuViewStoped.Checked = Not Me.mnuViewStoped.Checked
    Call zlRefRecords
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picClass_Resize()
    Dim intCount As Integer
    Dim lngTop As Long, lngButtom As Long
    Dim lngNVALL As Long, lngNVBottom As Long
    
    err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If Not cmdKind(intCount).Visible Then
            lngNVALL = lngNVALL + 1
            If Val(Me.cmdKind(intCount).Tag) = 1 Then
                lngNVBottom = lngNVBottom + 1
            End If
        End If
    Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + lngTop
            lngTop = lngTop + IIf(cmdKind(intCount).Visible, 285, 0)
            Me.tvwClass.Top = Me.picClass.ScaleTop + lngTop
        Else
            If lngButtom = 0 Then
                lngButtom = 285 * (Me.cmdKind.UBound - intCount + 1 - lngNVBottom)
            End If
            If cmdKind(intCount).Visible Then
                Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - lngButtom
                lngButtom = lngButtom - 285
            End If
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1 - lngNVALL) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHBar.Top = Me.picHBar.Top + y
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + x
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub tabContent_Click(PreviousTab As Integer)
    For intCount = 0 To Me.tabContent.Tabs - 1
        If intCount = Me.tabContent.Tab Then
            Me.fraSubInfo(intCount).Visible = True
        Else
            Me.fraSubInfo(intCount).Visible = False
        End If
    Next
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        Select Case tvwClass.Tag
            Case SK_������Ŀ
                If InStr(1, mstrPrivs, "������Ŀ�༭") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_��ҩ�䷽
                If InStr(1, mstrPrivs, "��ҩ�䷽�༭") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_���׷���
                If InStr(1, mstrPrivs, "���׷����༭") > 0 Then Call PopupMenu(Me.mnuClass, 2)
        End Select
    Case "Add"
        Call mnuEditAdd_Click
    Case "Modify"
        Call mnuEditModify_Click
    Case "Delete"
        Call mnuEditDelete_Click
    Case "Start"
        Call mnuEditStart_Click
    Case "Stop"
        Call mnuEditStop_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Key = "add" Then
        Call mnuEditAdd_Click
    Else
        Call mnuAddcopy_Click
    End If
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Select Case tvwClass.Tag
            Case SK_������Ŀ
                If InStr(1, mstrPrivs, "������Ŀ�༭") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_��ҩ�䷽
                If InStr(1, mstrPrivs, "��ҩ�䷽�༭") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_���׷���
                If InStr(1, mstrPrivs, "���׷����༭") > 0 Then Call PopupMenu(Me.mnuClass, 2)
        End Select
    End If
End Sub

Public Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwItems.Tag = Node.Key Then Exit Sub
    Me.lvwItems.Tag = Node.Key
    Call zlRefRecords
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    '---------------------------------------------
    '��д���Ʒ�����Ŀ(�˴�ΪҩƷ����)�����ղ�ͬ���͵�������
    '---------------------------------------------
    
    'Ȩ�޿���
    
    '������ʾ����
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        Me.mnuEditAdd.Visible = True: Me.mnuEditModify.Visible = True: Me.mnuEditDelete.Visible = True: Me.mnuEditSpt1.Visible = True
'        Me.mnuEditExse.Visible = True: Me.mnuEditLabs.Visible = True: Me.mnuEditExams.Visible = True
        Me.mnuEditExse.Visible = True: Me.mnuEditLabs.Visible = False: Me.mnuEditGather.Visible = True ': Me.mnuEditSample.Visible = True
        Me.mnuEditSpt2.Visible = True: Me.mnuEditStart.Visible = True: Me.mnuEditStop.Visible = True
        Me.mnuEditSpt3.Visible = True: Me.mnuEditRepellent.Visible = True: Me.mnuEditBill.Visible = True
        Me.mnuAddcopy.Visible = True
        
        Me.tlbThis.Buttons("Split2").Visible = True
        Me.tlbThis.Buttons("Add").Visible = True: Me.tlbThis.Buttons("Modify").Visible = True: Me.tlbThis.Buttons("Delete").Visible = True
        Me.tlbThis.Buttons("Split3").Visible = True
        Me.tlbThis.Buttons("Start").Visible = True: Me.tlbThis.Buttons("Stop").Visible = True
    
    Case 1, 2
        Me.mnuEditAdd.Visible = True: Me.mnuEditModify.Visible = True: Me.mnuEditDelete.Visible = True: Me.mnuEditSpt1.Visible = False
        Me.mnuEditExse.Visible = False: Me.mnuEditLabs.Visible = False: Me.mnuEditGather.Visible = False ': Me.mnuEditSample.Visible = False
        Me.mnuEditSpt2.Visible = True: Me.mnuEditStart.Visible = True: Me.mnuEditStop.Visible = True
        Me.mnuEditSpt3.Visible = False: Me.mnuEditRepellent.Visible = False: Me.mnuEditBill.Visible = False
        
        Me.tlbThis.Buttons("Split2").Visible = True
        Me.tlbThis.Buttons("Add").Visible = True: Me.tlbThis.Buttons("Modify").Visible = True: Me.tlbThis.Buttons("Delete").Visible = True
        Me.tlbThis.Buttons("Split3").Visible = True
        Me.tlbThis.Buttons("Start").Visible = True: Me.tlbThis.Buttons("Stop").Visible = True
        Me.mnuAddcopy.Visible = False
    End Select
    
    Me.lvwItems.ListItems.Clear
    With Me.tabContent
        .TabVisible(conTabִ�п���) = False
        .TabVisible(conTab�շѶ���) = False
        .TabVisible(conTab����ָ��) = False
        .TabVisible(conTab��鲿λ) = False
        .TabVisible(conTab�÷�����) = False
        .TabVisible(conTab�������) = False
        .TabVisible(conTab�䷽���) = False
        .TabVisible(conTab���׷���) = False
        .TabVisible(conTabӦ�òο�) = False
    End With
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_����", "����", 2500
            .Add , "_����", "����", 1200
            .Add , "_�걾��λ", "�걾��λ", 900
            .Add , "_���㵥λ", "���㵥λ", 900
            .Add , "_���", "���", 600
            .Add , "_��������", "��������", 1200
            .Add , "_ִ��Ƶ��", "ִ��Ƶ��", 900
            .Add , "_���㷽ʽ", "���㷽ʽ", 900
            .Add , "_�������", "�������", 900
            .Add , "_�������", "�������", 1200
            '.Add , "_վ��", "վ��", IIf(gstrNodeNo = "-", 0, 1000)
            .Add , "_Ժ��", "Ժ��", 600
        End With
        With Me.tabContent
            .TabVisible(conTabִ�п���) = True
            .TabVisible(conTab�շѶ���) = True
'            .TabVisible(conTabӦ�òο�) = True
            .Tab = conTabִ�п���: Call tabContent_Click(conTabִ�п���)
        End With
    Case 1
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_����", "����", 2000
            .Add , "_����", "����", 1200
            .Add , "_˵��", "˵��", 3000
            '.Add , "_վ��", "վ��", IIf(gstrNodeNo = "-", 0, 1000)
            .Add , "_Ժ��", "Ժ��", 600
        End With
        With Me.tabContent
            .TabVisible(conTab�䷽���) = True
            .Tab = conTab�䷽���: Call tabContent_Click(conTab�䷽���)
        End With
    Case 2
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_����", "����", 2500
            .Add , "_����", "����", 1200
            .Add , "_˵��", "˵��", 3000
            '.Add , "_վ��", "վ��", IIf(gstrNodeNo = "-", 0, 1000)
            .Add , "_������", "������", "1200"
            .Add , "_����ʱ��", "����ʱ��", "2000"
            .Add , "_Ժ��", "Ժ��", 600
            
        End With
        With Me.tabContent
            .TabVisible(conTab���׷���) = True
            .Tab = conTab���׷���: Call tabContent_Click(conTab���׷���)
        End With
    End Select
    With Me.lvwItems
        .ColumnHeaders("_����").Position = 1
        .SortKey = .ColumnHeaders("_����").Index - 1: .SortOrder = lvwAscending
    End With
    Call RestoreListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    
    '��д����
    err = 0: On Error GoTo ErrHand
    
    strTemp = Switch(Me.tvwClass.Tag = "0", 5, _
                   Me.tvwClass.Tag = "1", 4, _
                   Me.tvwClass.Tag = "2", 6)
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = [1] " & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp)
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Visible = True
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
        Else
            Me.tvwClass.Nodes(1).Selected = True
        End If
        Call zlRefRecords
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    Dim iSubItemIndex As Integer
    '---------------------------------------------
    '��д��Ŀ�б�
    '---------------------------------------------
    err = 0: On Error GoTo ErrHand
    
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        If mnuViewShowAll.Checked = True Then
'            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,I.���㵥λ,I.��� as �����,K.���� as ���,I.��������,I.ִ��Ƶ��,I.���㷽ʽ,I.�������," & _
'                    "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ',4,'���','��ֱ��Ӧ���ڲ���') as �������," & _
'                    "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ�� " & _
'                    " from ������ĿĿ¼ I,������Ŀ��� K, " & _
'                    " (Select ID, ���� From ���Ʒ���Ŀ¼ Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id" & _
'                    " Union ALL Select ID, ���� From ���Ʒ���Ŀ¼ Where ID=[1]) B " & _
'                    " where I.���=K.���� And I.����id = B.ID And (I.վ�� = '" & gstrNodeNo & "' Or I.վ�� is Null) "
            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,I.���㵥λ,I.��� as �����,K.���� as ���,I.��������,I.ִ��Ƶ��,I.���㷽ʽ,I.�������," & _
                    "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ',4,'���','��ֱ��Ӧ���ڲ���') as �������," & _
                    "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ�� " & _
                    " from ������ĿĿ¼ I,������Ŀ��� K, " & _
                    " (Select ID, ���� From ���Ʒ���Ŀ¼ Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id" & _
                    " Union ALL Select ID, ���� From ���Ʒ���Ŀ¼ Where ID=[1]) B " & _
                    " where I.���=K.���� And I.����id = B.ID "
        Else
'            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,I.���㵥λ,I.��� as �����,K.���� as ���,I.��������,I.ִ��Ƶ��,I.���㷽ʽ,I.�������," & _
'                    "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ',4,'���','��ֱ��Ӧ���ڲ���') as �������," & _
'                    "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ�� " & _
'                    " from ������ĿĿ¼ I,������Ŀ��� K" & _
'                    " where I.���=K.���� and I.����ID=[1] And (I.վ�� = '" & gstrNodeNo & "' Or I.վ�� is Null)"
            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,I.���㵥λ,I.��� as �����,K.���� as ���,I.��������,I.ִ��Ƶ��,I.���㷽ʽ,I.�������," & _
                    "        decode(I.�������,1,'����',2,'סԺ',3,'�����סԺ',4,'���','��ֱ��Ӧ���ڲ���') as �������," & _
                    "        nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ�� " & _
                    " from ������ĿĿ¼ I,������Ŀ��� K" & _
                    " where I.���=K.���� and I.����ID=[1] "
        End If
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        gstrSql = gstrSql & " order by I.����"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
        
        With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
                Select Case !�����
                Case "C"
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "D"
                    objItem.Icon = "���" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "E"
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "F"
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "G"
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "H"
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "I"
                    objItem.Icon = "��ʳ" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "K"
                    objItem.Icon = "��Ѫ" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "L"
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case Else
                    objItem.Icon = "����" & IIf(Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                End Select
                
                objItem.SmallIcon = objItem.Icon
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�걾��λ").Index - 1) = IIf(IsNull(!�걾��λ), "", !�걾��λ)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_���").Index - 1) = !���
                Select Case IIf(IsNull(!ִ��Ƶ��), 0, !ִ��Ƶ��)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ִ��Ƶ��").Index - 1) = "��ѡƵ��"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ִ��Ƶ��").Index - 1) = "һ����"
                Case 2
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_ִ��Ƶ��").Index - 1) = "������"
                End Select
                Select Case IIf(IsNull(!���㷽ʽ), 0, !���㷽ʽ)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_���㷽ʽ").Index - 1) = "��ȷ��"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_���㷽ʽ").Index - 1) = "����"
                Case 2
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_���㷽ʽ").Index - 1) = "��ʱ"
                Case 3
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_���㷽ʽ").Index - 1) = "�ƴ�"
                End Select
                Select Case IIf(IsNull(!�������), 0, !�������)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_�������").Index - 1) = "��������"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_�������").Index - 1) = "ȡ������"
                End Select
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_�������").Index - 1) = !�������
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_Ժ��").Index - 1) = IIf(IsNull(!վ��), "", !վ��)
                Select Case !�����
                Case "E"
                    intCount = Val(IIf(IsNull(!��������), 0, !��������))
                    strTemp = Switch(intCount = 0, "��ͨ", _
                                    intCount = 1, "��������", _
                                    intCount = 2, "��ҩ����(��ҩ)", _
                                    intCount = 3, "��ҩ�巨", _
                                    intCount = 4, "��ҩ��(��)��", _
                                    intCount = 5, "��������", _
                                    intCount = 6, "�ɼ�����", _
                                    intCount = 7, "��Ѫ����", _
                                    intCount = 8, "��Ѫ;��", _
                                    intCount = 9, "��Ѫ�ɼ�")
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = strTemp
                Case "H"
                    If IIf(IsNull(!��������), "0", !��������) = "1" Then
                        objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = "����ȼ�"
                    Else
                        objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = "������"
                    End If
                Case "Z"
                    intCount = Val(IIf(IsNull(!��������), 0, !��������))
                    strTemp = Switch(intCount = 0, "��ͨ", _
                                    intCount = 1, "����", _
                                    intCount = 2, "סԺ", _
                                    intCount = 3, "ת��", _
                                    intCount = 4, "����", _
                                    intCount = 5, "��Ժ", _
                                    intCount = 6, "תԺ", _
                                    intCount = 7, "����", _
                                    intCount = 8, "����", _
                                    intCount = 9, "����", _
                                    intCount = 10, "��Σ", _
                                    intCount = 11, "����", _
                                    intCount = 12, "��¼�����", _
                                    intCount = 14, "��ǰ")
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = strTemp
                Case Else
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_��������").Index - 1) = IIf(IsNull(!��������), "", !��������)
                End Select
                objItem.Tag = !�����
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                
                '��ͣ����Ŀ��ʾΪ��ɫ����ͮ��
                If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = &HFF&
                    For iSubItemIndex = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(iSubItemIndex).ForeColor = &HFF&
                    Next
                End If
                
                .MoveNext
            Loop
        End With
    Case 1, 2 '��ҩ�䷽�����׷���
        '���׷���Ȩ�޷�Χ����
        gstrSql = ""
        If Val(Me.tvwClass.Tag) = 2 Then
            If InStr(mstrPrivs, "ȫԺ���׷���") > 0 Then
                '��ȫԺ���׷���Ȩ��ʱ��������
            ElseIf InStr(mstrPrivs, "���Ƴ��׷���") > 0 Then
                'ֻ�б��Ƴ��׷���Ȩ��ʱ�����ڱ����ڻ����ѵ�
                gstrSql = " And (I.��ԱID=[2] Or Exists(Select 1 From �������ÿ��� X,������Ա Y Where X.����ID=Y.����ID And X.��ĿID=I.ID And Y.��ԱID=[2]))"
            Else
                '��û����ֻ�ܿ����ѵ�
                gstrSql = " And I.��ԱID=[2]"
            End If
        End If
        '-------------------------------
        If mnuViewShowAll.Checked = True Then
'            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ�� " & _
'                    " from ������ĿĿ¼ I," & _
'                    " (Select ID, ���� From ���Ʒ���Ŀ¼ Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id" & _
'                    " Union ALL Select ID, ���� From ���Ʒ���Ŀ¼ Where ID=[1]) B " & _
'                    " where I.����id = B.ID And (I.վ�� = '" & gstrNodeNo & "' Or I.վ�� is Null) " & gstrSql
            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ��,I.������,I.����ʱ�� " & _
                    " from ������ĿĿ¼ I," & _
                    " (Select ID, ���� From ���Ʒ���Ŀ¼ Start With �ϼ�id = [1] Connect By Prior ID = �ϼ�id" & _
                    " Union ALL Select ID, ���� From ���Ʒ���Ŀ¼ Where ID=[1]) B " & _
                    " where I.����id = B.ID " & gstrSql
        Else
'            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ�� " & _
'                    " from ������ĿĿ¼ I where (I.վ�� = '" & gstrNodeNo & "' Or I.վ�� is Null) And I.����ID=[1] " & gstrSql
            gstrSql = "select I.ID,I.����,I.����,I.�걾��λ,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,I.վ��,I.������,I.����ʱ�� " & _
                    " from ������ĿĿ¼ I where I.����ID=[1] " & gstrSql
        End If
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        gstrSql = gstrSql & " order by I.����"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)), UserInfo.ID)
        
        With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
                If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                    objItem.Icon = "����U": objItem.SmallIcon = "����U"
                Else
                    objItem.Icon = "����S": objItem.SmallIcon = "����S"
                End If
                
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_˵��").Index - 1) = IIf(IsNull(!�걾��λ), "", !�걾��λ)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_Ժ��").Index - 1) = IIf(IsNull(!վ��), "", !վ��)
                
                If Val(Me.tvwClass.Tag) = 2 Then
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_������").Index - 1) = IIf(IsNull(!������), "", !������)
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_����ʱ��").Index - 1) = IIf(IsNull(!����ʱ��), "", Format(!����ʱ��, "YYYY-MM-DD"))
                End If
                
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                
                '��ͣ����Ŀ��ʾΪ��ɫ����ͮ��
                If Format(!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = &HFF&
                    For iSubItemIndex = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(iSubItemIndex).ForeColor = &HFF&
                    Next
                End If
                
                .MoveNext
            Loop
        End With
    End Select

    If Me.lvwItems.ListItems.Count > 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then Me.lvwItems.ListItems(1).Selected = True
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.lvwItems.ListItems.Count & "����Ŀ"
    Else
        Call zlClearDetail
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsfExseGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���շѶ�������
    '����:���˺�
    '����:2017-07-01 21:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, str��� As String
    Dim i As Integer, varGrade As Variant
    On Error GoTo ErrHandle
    
    Set objItem = Me.lvwItems.SelectedItem
    If Not objItem Is Nothing Then
        str��� = objItem.SubItems(Me.lvwItems.ColumnHeaders("_���").Index - 1)
    Else
        str��� = ""
    End If
    
    varGrade = Split(mstrPriceGrade, ",")
    With vsfExse
        .Redraw = flexRDNone
        '.FixedRows = IIf(mblnStartPriceGrade, 2, 1)
        .Rows = .FixedRows + 1
        i = UBound(varGrade)
        i = IIf(i < 0, 0, i + IIf(mstrPriceGrade <> "", 1, 0))
        .Cols = 15 + i: .FixedCols = 1
        
        .TextMatrix(0, 0) = "":
        .TextMatrix(0, 1) = "��λ":
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "��Ŀ��":
        .TextMatrix(0, 4) = "���":
        .TextMatrix(0, 5) = "��λ":
        .TextMatrix(0, 6) = "�۸�"
        .TextMatrix(0, 7) = "����":
        .TextMatrix(0, 8) = "�̶�":
        .TextMatrix(0, 9) = "����"
        .TextMatrix(0, 10) = "����":
        .TextMatrix(0, 11) = "״̬":
        .TextMatrix(0, 12) = "�շѷ�ʽ"
        .TextMatrix(0, 13) = "���ó���"
        .TextMatrix(0, 14) = "���ÿ���"
        For i = 0 To UBound(varGrade)
            .TextMatrix(0, 15 + i) = varGrade(i)
            .colData(15 + i) = "A" & i + 1
            .ColAlignment(15 + i) = flexAlignRightCenter
        Next
        For i = 0 To .Cols - 1
            .ColKey(i) = IIf(i = 0, "ѡ��", .TextMatrix(0, i))
            .TextMatrix(.FixedRows, i) = ""
            .FixedAlignment(i) = 4
        Next
        .ColWidth(.ColIndex("ѡ��")) = 250
        .ColWidth(.ColIndex("��λ")) = 900
        .ColWidth(.ColIndex("����")) = 1200
        .ColWidth(.ColIndex("��Ŀ��")) = 3000
        .ColWidth(.ColIndex("���")) = 1000
        .ColWidth(.ColIndex("��λ")) = 800
        .ColWidth(.ColIndex("�۸�")) = 1200
        .ColWidth(.ColIndex("����")) = 1200
        .ColWidth(.ColIndex("�̶�")) = 600
        .ColWidth(.ColIndex("����")) = 600
        .ColWidth(.ColIndex("����")) = 600
        .ColWidth(.ColIndex("״̬")) = 0
        .ColWidth(.ColIndex("�շѷ�ʽ")) = 3000
        .ColWidth(.ColIndex("���ó���")) = 850
        .ColWidth(.ColIndex("���ÿ���")) = 1800
        If str��� <> "���" Then
            .ColWidth(.ColIndex("��λ")) = 0: .ColWidth(.ColIndex("����")) = 0: .ColWidth(.ColIndex("����")) = 0
        End If
        .ColAlignment(.ColIndex("�շѷ�ʽ")) = flexAlignLeftCenter
        .Redraw = flexRDBuffered
    End With
    
    

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub zlClearDetail()
    '---------------------------------------------
    '���������ϸ��Ϣ��ʾ����
    '---------------------------------------------
    Dim objItem As ListItem, str��� As String
    Dim i As Integer
    If Val(Me.tvwClass.Tag) = 0 Then
        'ִ�п�����ʾ
        Me.lblUseBill.Caption = "���Ƶ��ݣ�"
        Me.optִ�в���(0).Value = False 'True
        Me.lbl����ִ��.Caption = ""
        With Me.hgd����ִ��
            .Rows = .FixedRows + 1: .Cols = 2
            .TextMatrix(0, 0) = "ִ�п���": .TextMatrix(0, 1) = "���˿���"
            .ColWidth(0) = 1800: .ColWidth(1) = 6000
            For intCount = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCount) = "": .ColAlignmentFixed(intCount) = 4
            Next
        End With
    
        '�շѶ�����ʾ
        If Val(Me.tvwClass.Tag) = 0 Then Call InitVsfExseGrid
    
        '����ָ����ʾ
'        If Val(Me.tvwClass.Tag) = 0 Then
'            With Me.hgdLabs
'                .Rows = .FixedRows + 1: .Cols = 5: .FixedCols = 1
'                .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "����걾": .TextMatrix(0, 2) = "������Ŀ": .TextMatrix(0, 3) = "����": .TextMatrix(0, 4) = "��λ"
'                .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 3800: .ColWidth(3) = 600: .ColWidth(4) = 1000
'                For intCount = 0 To .Cols - 1
'                    .TextMatrix(.FixedRows, intCount) = "": .ColAlignmentFixed(intCount) = 4
'                Next
'            End With
'        End If
    
        '��鲿λ��ʾ
        With Me.hgdPart
            .Rows = .FixedRows + 1: .Cols = 3: .FixedCols = 1
            .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "�����Ŀ": .TextMatrix(0, 2) = "��鲿λ"
            .ColWidth(0) = 250: .ColWidth(1) = 3500: .ColWidth(2) = 2000
            For intCount = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCount) = "": .ColAlignmentFixed(intCount) = 4
            Next
        End With
    End If
    
    If Val(Me.tvwClass.Tag) = 1 Then
        mbyt��ҩζ�� = zlDatabase.GetPara(213, glngSys)
        '�䷽�����ʾ
        With Me.hgdRecipe
            .Rows = .FixedRows + 1: .Cols = mbyt��ҩζ�� * 6: .RowHeight(0) = 0
            .GridColor = &H80000005: .BackColorBkg = &H80000005
            .MergeCells = flexMergeFree: .MergeRow(0) = True
            For intCount = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCount) = ""
                If (intCount Mod 6) = 0 Then .ColWidth(intCount) = 150
                If (intCount Mod 6) = 1 Then .ColWidth(intCount) = 0
                If (intCount Mod 6) = 2 Then .ColWidth(intCount) = 1500
                If (intCount Mod 6) = 3 Then .ColWidth(intCount) = 500
                If (intCount Mod 6) = 4 Then .ColWidth(intCount) = 200
                If (intCount Mod 6) = 5 Then .ColWidth(intCount) = 800
            Next
        End With
    End If
    
    If Val(Me.tvwClass.Tag) = 2 Then
        '���׷�����ʾ(ByZT)
        With vsScheme
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End With
    End If
    
    '���Ʋο���ʾ
    With Me.hgdRefer
        .Rows = 1: .ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 1
        For intCount = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCount) = ""
        Next
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        objPrint.Title.Text = "������Ŀ�嵥"
    Case 1
        objPrint.Title.Text = "��ҩ�䷽�嵥"
    Case 2
        objPrint.Title.Text = "�������Ʒ����嵥"
    End Select
    objPrint.UnderAppItems.Add "���ࣺ" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '���ݵ������ݵ�������������и߶ȣ��Ա�֤���ݵ�������ʾ
    '---------------------------------------------
    Dim intRow As Integer, lngColWidth As Long
    With Me.hgdRefer
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, 1) = "" Then
                lngColWidth = .ColWidth(2)
            Else
                lngColWidth = .ColWidth(1) + .ColWidth(2)
            End If
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(intRow, 2)
            .RowHeight(intRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Public Sub zlLocateItem(lngClassId As Long, lngItemId As Long)
    '---------------------------------------------
    '��λ��ָ������ϲο���Ŀ���ڲ���ʱʹ��
    '---------------------------------------------
    On Error Resume Next
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lngClassId)
    Me.tvwClass.Nodes("_" & lngClassId).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & lngItemId)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub vsScheme_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsScheme.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsScheme.TextMatrix(vsScheme.FixedRows - 1, Col) & "A")
        If vsScheme.ColWidth(Col) < lngW Then
            vsScheme.ColWidth(Col) = lngW
        ElseIf vsScheme.ColWidth(Col) > vsScheme.Width * 0.5 Then
            vsScheme.ColWidth(Col) = vsScheme.Width * 0.5
        End If
    End If
End Sub

Private Sub vsScheme_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsScheme
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col����: lngRight = col�÷�
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) And Me.ActiveControl Is vsScheme Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsScheme
        If .TextMatrix(lngRow, col���) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col���)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���)) = Val(.TextMatrix(lngRow, col���)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Function ShowScheme(ByVal lng����ID As Long) As Boolean
'���ܣ���ȡ����ʾ���ݿ��еĳ��׷�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim str��ҩ As String, str�巨 As String
    Dim str���� As String, Str�걾 As String
    Dim i As Long, j As Long

    On Error GoTo errH

    strSql = "Select A.���,A.������,A.��Ч,A.������ĿID,A.ҽ������,A.����," & _
             " A.��������,A.ִ��Ƶ��,A.ҽ������,Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���," & _
             " A.ִ������,A.ִ�б��,A.ʱ�䷽��,Nvl(B.���,'*') as ���,Nvl(D.����||Decode(D.���,NULL,NULL,' '||D.���),B.����) as ����," & _
             " B.���㵥λ,A.�걾��λ,A.��鷽��,A.�ܸ�����,D.���㵥λ as ������λ,D.ID as �շ�ϸĿID," & _
             " Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ��" & _
             " From ������Ŀ��� A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D" & _
             " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=C.ID(+)" & _
             " And A.�շ�ϸĿid=D.ID(+) And A.�������ID=[1] " & _
             " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)

    With vsScheme
        .Redraw = flexRDNone
        .Rows = .FixedRows    '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col��Ч) = IIf(NVL(rsTmp!��Ч, 0) = 0, "����", "��ʱ")
                .TextMatrix(i, col����) = NVL(rsTmp!ҽ������, NVL(rsTmp!����))
                .TextMatrix(i, col�걾��λ) = NVL(rsTmp!�걾��λ)    '����걾
                .TextMatrix(i, col��鷽��) = NVL(rsTmp!��鷽��)
                .TextMatrix(i, col����) = FormatEx(NVL(rsTmp!��������), 4)
                If Not IsNull(rsTmp!��������) Then
                    If rsTmp!��� = "4" Then
                        .TextMatrix(i, col��λ) = NVL(rsTmp!������λ)
                    Else
                        .TextMatrix(i, col��λ) = NVL(rsTmp!���㵥λ)
                    End If
                End If
                If .TextMatrix(i, col��Ч) = "��ʱ" Then
                    If Not IsNull(rsTmp!�ܸ�����) Then
                        .TextMatrix(i, col����) = FormatEx(NVL(rsTmp!�ܸ�����), 4)
                        If Not IsNull(rsTmp!������λ) Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!������λ)
                        ElseIf InStr(",4,5,6,7,", rsTmp!���) = 0 Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!���㵥λ)
                        End If
                    End If
                End If
                .TextMatrix(i, col����) = NVL(rsTmp!����)
                .TextMatrix(i, colƵ��) = NVL(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, col����) = NVL(rsTmp!ҽ������)
                .TextMatrix(i, colִ��ʱ��) = NVL(rsTmp!ʱ�䷽��)
                .TextMatrix(i, colִ�п���) = NVL(rsTmp!ִ�п���)
                .Cell(flexcpData, i, colִ������) = NVL(rsTmp!ִ������, 0)
                .TextMatrix(i, col���) = rsTmp!���
                .TextMatrix(i, col���) = NVL(rsTmp!������)
                .TextMatrix(i, col��ĿID) = NVL(rsTmp!������Ŀid)
                .TextMatrix(i, col�շ�ϸĿID) = NVL(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, col���) = rsTmp!���
                .TextMatrix(i, colִ�б��) = NVL(rsTmp!ִ�б��)
                .TextMatrix(i, colͣ��) = IIf(Format(rsTmp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01", "��", "")
                If Format(rsTmp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF&
                End If
                rsTmp.MoveNext
            Next

            '�ٴ���һЩ�����е�����,��������ݵ���ʾ
            For i = 1 To .Rows - 1
                '��ҩ;��
                If .TextMatrix(i, col���) = "E" And Val(.TextMatrix(i, col���)) = 0 _
                   And Val(.TextMatrix(i - 1, col���)) = Val(.TextMatrix(i, col���)) _
                   And InStr(",5,6,", .TextMatrix(i - 1, col���)) > 0 Then
                    .RowHidden(i) = True
                    '��ʾ��ҩ;��
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .TextMatrix(j, col�÷�) = .TextMatrix(i, col����)

                            '��ʾ��ҩ��ִ������
                            If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                                .TextMatrix(j, colִ������) = IIf(Val(.TextMatrix(j, colִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
                            ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                                .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                            Else
                                .TextMatrix(j, colִ������) = IIf(Val(.TextMatrix(j, colִ�б��)) = 1, "��ȡҩ", "����")
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If

                '��Ѫ;��
                If .TextMatrix(i, col���) = "E" And .TextMatrix(i - 1, col���) = "K" _
                   And Val(.TextMatrix(i, col���)) = Val(.TextMatrix(i - 1, col���)) Then
                    .RowHidden(i) = True
                    .TextMatrix(i - 1, col�÷�) = .TextMatrix(i, col����)
                    .TextMatrix(i - 1, col����) = .TextMatrix(i - 1, col����) & "(" & .TextMatrix(i, col����) & ")"
                End If

                '��ҩ�䷽�ͼ������
                If .TextMatrix(i, col���) = "E" And Val(.TextMatrix(i, col���)) = 0 _
                   And Val(.TextMatrix(i - 1, col���)) = Val(.TextMatrix(i, col���)) _
                   And InStr(",7,E,C,", .TextMatrix(i - 1, col���)) > 0 Then

                    str��ҩ = "": str�巨 = "": Str�걾 = "": strTmp = ""
                    j = .FindRow(CStr(Val(.TextMatrix(i, col���))), , col���)

                    '��ҩ�������ִ�п���
                    .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)

                    '��ʾ��ҩ�䷽ִ������:��ҩƷΪ׼�ж�
                    If .TextMatrix(i - 1, col���) <> "C" Then
                        If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                            .TextMatrix(i, colִ������) = IIf(Val(.TextMatrix(i, colִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
                        ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                            .TextMatrix(i, colִ������) = "��Ժ��ҩ"
                        Else
                            .TextMatrix(i, colִ������) = IIf(Val(.TextMatrix(i, colִ�б��)) = 1, "��ȡҩ", "����")
                        End If
                    End If

                    For j = j To i - 1
                        .RowHidden(j) = j <> i
                        If .TextMatrix(j, col���) = "7" Then
                            str��ҩ = str��ҩ & "," & RTrim(.TextMatrix(j, col����) & _
                                                        " " & .TextMatrix(j, col����) & .TextMatrix(j, col��λ) & _
                                                        " " & .TextMatrix(j, col����))
                        ElseIf .TextMatrix(j, col���) = "C" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col����)
                            Str�걾 = .TextMatrix(j, col�걾��λ)    'ȡ��һ��������Ŀ�ı걾
                        ElseIf .TextMatrix(j, col���) = "E" And Val(.TextMatrix(j, col���)) <> 0 Then
                            str�巨 = .TextMatrix(j, col����) & .TextMatrix(j, col�걾��λ)
                        End If
                    Next

                    .TextMatrix(i, col�÷�) = .TextMatrix(i, col����)    '��ʾ��ҩ�÷������ɼ�����

                    If .TextMatrix(i - 1, col���) = "C" Then
                        .TextMatrix(i, col����) = Mid(strTmp, 2) & IIf(Str�걾 <> "", "(" & Str�걾 & ")", "")
                    Else
                        .TextMatrix(i, col����) = "��ҩ�䷽," & .TextMatrix(i, colƵ��) & "," & _
                                                str�巨 & "," & .TextMatrix(i, col����) & ":" & Mid(str��ҩ, 2)
                        .TextMatrix(i, col������λ) = "��"
                    End If
                End If

                '������
                If .TextMatrix(i, col���) = "D" And Val(.TextMatrix(i, col���)) = 0 Then
                    Str�걾 = "": str�巨 = "": strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col�걾��λ) <> "" _
                               And Val(.TextMatrix(j, col��ĿID)) = Val(.TextMatrix(i, col��ĿID)) Then    '��ͬ����ĿID�����·�ʽ
                                If .TextMatrix(j, col�걾��λ) <> strTmp And strTmp <> "" Then
                                    Str�걾 = Str�걾 & "," & strTmp & IIf(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                                    str�巨 = ""
                                End If
                                If .TextMatrix(j, col��鷽��) <> "" Then
                                    str�巨 = str�巨 & "," & .TextMatrix(j, col��鷽��)
                                End If

                                strTmp = .TextMatrix(j, col�걾��λ)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        Str�걾 = Str�걾 & "," & strTmp & IIf(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                    End If
                    If Str�걾 <> "" Then    '��ǰ�ļ�鷽ʽʱ����ʾ��ϸҽ������
                        .TextMatrix(i, col����) = .TextMatrix(i, col����) & ":" & Mid(Str�걾, 2)
                    End If
                End If

                '������Ŀ
                If .TextMatrix(i, col���) = "F" And Val(.TextMatrix(i, col���)) = 0 Then
                    strTmp = "": str���� = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col���) = "F" Then
                                strTmp = strTmp & "," & .TextMatrix(j, col����)
                            ElseIf .TextMatrix(j, col���) = "G" Then
                                str���� = .TextMatrix(j, col����)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Or str���� <> "" Then
                        If str���� <> "" Then
                            .TextMatrix(i, col����) = "�� " & str���� & " ���� " & .TextMatrix(i, col����)
                        Else
                            .TextMatrix(i, col����) = "�� " & .TextMatrix(i, col����)
                        End If
                        If strTmp <> "" Then
                            .TextMatrix(i, col����) = .TextMatrix(i, col����) & " �� " & Mid(strTmp, 2)
                        End If
                    End If
                End If
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col����
        .Redraw = flexRDDirect
    End With
    ShowScheme = True
    Exit Function
errH:
    vsScheme.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub
