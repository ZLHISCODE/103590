VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDrugQuery 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "ҩƷ����ѯ"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11325
   Icon            =   "frmDrugQuery.frx":0000
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7110
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4920
      ScaleHeight     =   255
      ScaleWidth      =   4935
      TabIndex        =   16
      Top             =   5760
      Width           =   4935
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   3000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "��Ч��"
         Height          =   180
         Left            =   2325
         TabIndex        =   24
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "���"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   23
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "ͣ��"
         Height          =   180
         Index           =   2
         Left            =   900
         TabIndex        =   22
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "���Σ�"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   21
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "�������������Ԥ��"
         Height          =   180
         Index           =   0
         Left            =   3285
         TabIndex        =   20
         Top             =   30
         Width           =   1620
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1095
      Left            =   9360
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   1931
      Appearance      =   0
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDrugQuery.frx":0982
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
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
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.CheckBox Chk���� 
      Appearance      =   0  'Flat
      Caption         =   "ȫѡ"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   70
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5655
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1470
      ScaleHeight     =   255
      ScaleWidth      =   2415
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6810
      Width           =   2415
      Begin VB.TextBox txtҩƷ��Ϣ 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   780
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lblҩƷ��Ϣ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ҩƷ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView lst����_S 
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   5910
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1270
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6744
      Width           =   11328
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugQuery.frx":09D0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14182
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin MSComctlLib.TreeView tvwSection_S 
      Height          =   4350
      Left            =   60
      TabIndex        =   1
      Top             =   1275
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   7673
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglvw"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   2985
      Top             =   2205
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
            Picture         =   "frmDrugQuery.frx":1264
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":2F6E
            Key             =   "child"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":4C78
            Key             =   "clock"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVLine_S 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   2940
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5460
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   1305
      Width           =   45
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   11325
      _CBHeight       =   1125
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   2730
      NewRow1         =   0   'False
      Caption2        =   "�ⷿ"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   6780
      NewRow2         =   -1  'True
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   585
         TabIndex        =   6
         Text            =   "cboStock"
         Top             =   780
         Width           =   10650
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               ImageIndex      =   3
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ϸ"
               Key             =   "��ϸ"
               Object.ToolTipText     =   "ҩƷ��ϸ��"
               Object.Tag             =   "��ϸ"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "ҩƷ����"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   1425
      Top             =   780
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
            Picture         =   "frmDrugQuery.frx":B4DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":B6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":B912
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":BB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":BD48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":BF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C17C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C398
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C5B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C7D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C9EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   690
      Top             =   810
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
            Picture         =   "frmDrugQuery.frx":D2C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D4E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D916
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":DB32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":DD4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":DF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E182
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E39E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E5BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E7D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1245
      Left            =   4080
      TabIndex        =   12
      Top             =   1440
      Width           =   5055
      _cx             =   8916
      _cy             =   2196
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   15724527
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugQuery.frx":F0AE
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
      VirtualData     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
      Height          =   1245
      Left            =   3960
      TabIndex        =   13
      Top             =   4080
      Width           =   5055
      _cx             =   8916
      _cy             =   2196
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
      MouseIcon       =   "frmDrugQuery.frx":F4AD
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   15724527
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugQuery.frx":F4C9
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
      VirtualData     =   0   'False
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
   Begin VB.Label lbl����_S 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ҩƷ����"
      Height          =   285
      Left            =   15
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   5625
      Width           =   2865
   End
   Begin VB.Label lbl����_S 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "�������"
      Height          =   180
      Left            =   3600
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   3240
      Width           =   6585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileBatch 
         Caption         =   "������ӡ��ϸ��(&B)"
      End
      Begin VB.Menu mnuViewLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
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
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "С����"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "������"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewForeColor 
         Caption         =   "ǰ��ɫ(&C)"
      End
      Begin VB.Menu mnuViewBackColor 
         Caption         =   "����ɫ(&B)"
      End
      Begin VB.Menu mnuviewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuviewLineNoVerify 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewNoVerify 
         Caption         =   "δ�󵥾ݲ�ѯ(&N)"
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
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuOpen 
         Caption         =   "��(&O)"
      End
      Begin VB.Menu mnuPopuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "С����"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "������"
         Index           =   1
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "������"
         Index           =   2
      End
   End
   Begin VB.Menu mnuReportBill 
      Caption         =   "����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuBill 
         Caption         =   "����(&D)"
      End
   End
End
Attribute VB_Name = "frmDrugQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Dim intFont As Integer
'Dim WithEvents DataRecordSet  As adodb.Recordset
Dim appData As Collection
Public intChoose���� As Byte         '0-�ۼ۵�λ;1-���ﵥλ;2-ҩ�ⵥλ;3-סԺ��λ
Dim bln����� As Boolean
Dim bln����ͣ��ҩƷ As Boolean
Dim intMonths As Integer
Dim BlnHourse As Boolean 'Ϊ���ʾ�ⷿ
Public BlnDO As Boolean
Dim Bln����ҩ As Boolean '��ʾ�Ƿ���в�ѯ����ҩ��Ȩ��
Dim Bln�г�ҩ As Boolean '��ʾ�Ƿ���в�ѯ�г�ҩ��Ȩ��
Dim Bln�в�ҩ As Boolean '��ʾ�Ƿ���в�ѯ�в�ҩ��Ȩ��
Dim Str���� As String
Dim StrSort As String    '��ʾҩƷ���
Dim mstrPrivs As String
Private mlngMode As Long
Private mblnViewCost As Boolean       '�鿴�ɱ��� true-����鿴 false-������鿴

Private mblnRefresh As Boolean                  '�Ƿ�����ˢ��
Private mstrUnShow_List As String               '��������ʾ���У�ҩƷ�б�
Private mstrUnShow_Batch As String              '��������ʾ���У������б�

Private LngCardRow As Long
Private LngPhysicRow As Long
Private StrCardSortBy As String                 '������
'Modified By ���� 2003-12-10 ���������� ѡ����ɫ��Ϊ��ɫ����ɫ���е������滻���Ϊ��ɫ
Private Const glng��ɫ As Long = &H80000005
Private Const glng��ɫ As Long = &H80000008
Private Const glng��ɫ As Long = &HFFCECE
Private Const glng��ɫ As Long = &H8000000F
Private Const glng��ɫ As Long = &HC0C0C0
Private Const glng��ɫ As Long = &HC0           'ͣ��

Private mStr�ɱ��� As String
Private mStr���� As String
Private mStr���� As String
Private mStr��� As String
Private mStrMax��� As String

Private mblnExportState As Boolean          '�������״̬

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

'----------------------
'���ű���ı�������
Public WithEvents ObjReport As zl9Report.clsReport
Attribute ObjReport.VB_VarHelpID = -1
Private lngCurReport As Long
Private CurSheet As Object
Dim strNoS As String
'-----------------------

Private Type Type_SQLCondition
    strͨ���� As String
    str���� As String
    str���� As String
    str���� As String
    str��� As String
    str���� As String
    strҩƷ��Ϣ As String
    lngҩƷ���� As Long
    lng�ⷿID As Long
End Type

Private SQLCondition As Type_SQLCondition

Private Enum IniListType
    AllList = 0
    MainList = 1
    BatchList = 2
End Enum

Private Sub SetSortCode()
    '����ҩƷ���뷵�ظ�ʽ�����������
    '�����п��ܺ���"-"���ţ��������б�����"-"ǰ��༸λ��"-"����༸λ�����б��붼�����λ�����и�ʽ������
    Dim lngRow As Long
    Dim intǰ׺ As Integer
    Dim int��׺ As Integer
    Dim str����ǰ׺ As String
    Dim str�����׺ As String
    Dim blnLine As Boolean
    
    With vsfList
        For lngRow = 1 To vsfList.rows - 1
            If InStr(1, .TextMatrix(lngRow, .ColIndex("����")), "-") > 0 Then
                blnLine = True
                If Len(Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") - 1)) > intǰ׺ Then
                    intǰ׺ = Len(Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") - 1))
                End If
                
                If Len(Mid(.TextMatrix(lngRow, .ColIndex("����")), InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") + 1)) > int��׺ Then
                    int��׺ = Len(Mid(.TextMatrix(lngRow, .ColIndex("����")), InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") + 1))
                End If
            Else
                If Len(.TextMatrix(lngRow, .ColIndex("����"))) > intǰ׺ Then
                    intǰ׺ = Len(.TextMatrix(lngRow, .ColIndex("����")))
                End If
            End If
        Next
        
        For lngRow = 1 To .rows - 1
            If blnLine = False Then
                .TextMatrix(lngRow, .ColIndex("�������")) = Format(.TextMatrix(lngRow, .ColIndex("����")), String(intǰ׺, "0"))
            Else
                If InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") > 0 Then
                    str����ǰ׺ = Mid(.TextMatrix(lngRow, .ColIndex("����")), 1, InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") - 1)
                    str�����׺ = Mid(.TextMatrix(lngRow, .ColIndex("����")), InStr(.TextMatrix(lngRow, .ColIndex("����")), "-") + 1)
                    
                    str����ǰ׺ = Format(str����ǰ׺, String(intǰ׺, "0"))
                    str�����׺ = Format(str�����׺, String(int��׺, "0"))
                Else
                    str����ǰ׺ = Format(.TextMatrix(lngRow, .ColIndex("����")), String(intǰ׺, "0"))
                    str�����׺ = String(int��׺, "0")
                End If
                
                .TextMatrix(lngRow, .ColIndex("�������")) = str����ǰ׺ & "-" & str�����׺
            End If
        Next
    End With
End Sub

Private Sub SetDrugDigit(ByVal intUnit As Integer)
    Dim strUnit As String
    Dim intDrugUnit As Integer
    
    Const conInt���㾫�� As Integer = 0
    
    Const conIntҩƷ As Integer = 1
    
    'ҩƷ����ѯ�������õĵ�λ��˳����ܺ�����ģ�����ò�һ��
    Const conint�ۼ۵�λ As Integer = 1
    Const conint���ﵥλ As Integer = 2
    Const conintסԺ��λ As Integer = 4
    Const conintҩ�ⵥλ As Integer = 3
        
    Const conInt�ɱ��� As Integer = 1
    Const conInt�ۼ� As Integer = 2
    Const conInt���� As Integer = 3
    Const conInt��� As Integer = 4
    
    intDrugUnit = intUnit
    
    Select Case intDrugUnit
        Case conint�ۼ۵�λ            '�ۼ۵�λ����Ҫ���Ƽ���
            intDrugUnit = 1
        Case conint���ﵥλ
            intDrugUnit = 2
        Case conintסԺ��λ
            intDrugUnit = 3
        Case conintҩ�ⵥλ
            intDrugUnit = 4
    End Select

    '�ֱ�ȡҩƷ�ɱ��ۡ��ۼۡ�����������С��λ��
    mintCostDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt�ɱ���, intDrugUnit)
    mintPriceDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt�ۼ�, intDrugUnit)
    mintNumberDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt����, intDrugUnit)
    mintMoneyDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt���)
    
    mStr�ɱ��� = "####0." & String(mintCostDigit, "0") & ";-####0." & String(mintCostDigit, "0") & "; ;"
    mStr���� = "####0." & String(mintPriceDigit, "0") & ";-####0." & String(mintPriceDigit, "0") & "; ;"
    mStr���� = "####0." & String(mintNumberDigit, "0") & ";-####0." & String(mintNumberDigit, "0") & "; ;"
    mStr��� = "####0." & String(mintMoneyDigit, "0") & ";-####0." & String(mintMoneyDigit, "0") & "; ;"
    
    mStrMax��� = "####0." & String(5, "0") & ";-####0." & String(5, "0") & "; ;"
End Sub
Private Sub openҩƷ����()
    Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_1", Me, "�ⷿ=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)))
End Sub

Private Sub openҩƷ��ϸ��()
'    If DataRecordSet Is Nothing Then Exit Sub
'    If Not (DataRecordSet.State = 1) Then Exit Sub
'    If DataRecordSet.RecordCount = 0 Then Exit Sub
    
    If vsfList.Row = 0 Then Exit Sub
    If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����")) = "" Then Exit Sub
    
    If cboStock.ItemData(cboStock.ListIndex) = 0 Then
        Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_2", Me, "ҩƷ=" & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����")) & "|" & Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))), "�ⷿ=���пⷿ|is not null", "��λ=" & Choose(intChoose����, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(intChoose����, 1, 3, 2, 4))    ' , "��ʼ����=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "��������=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    Else
        Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_2", Me, "ҩƷ=" & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����")) & "|" & Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ҩƷID"))), "�ⷿ=" & cboStock.Text & "|=  " & cboStock.ItemData(cboStock.ListIndex), "��λ=" & Choose(intChoose����, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(intChoose����, 1, 3, 2, 4)) ' , "��ʼ����=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "��������=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    End If
End Sub
Private Sub openҩƷ��ϸ��()
    Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_3", Me, "�ⷿ=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)), "��λ=" & Choose(intChoose����, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(intChoose����, 1, 3, 2, 4))
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "H,I,J,K,L,M,N"

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), str��������, IIf(zlStr.IsHavePrivs(mstrPrivs, "���пⷿ"), False, True)) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub


Private Sub cbrThis_Resize()
    Form_Resize
End Sub

Private Sub cboStock_Click()
    If Me.tvwSection_S.Nodes.count = 0 Then Exit Sub
    Me.tvwSection_S.Tag = ""
    
'    If cboStock.Text = "���пⷿ" Then
'        vsfList.ColHidden(vsfList.ColIndex("�������")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("�������")) = False
'    Else
'        vsfList.ColHidden(vsfList.ColIndex("�������")) = False
'        vsfBatch.ColHidden(vsfBatch.ColIndex("�������")) = True
'    End If
'
'    If bln����� = False Then
'        vsfList.ColHidden(vsfList.ColIndex("�������")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("�������")) = True
'    End If
    
    If Val(cboStock.Tag) <> cboStock.ItemData(cboStock.ListIndex) Then
        If IIf(Val(cboStock.Tag) = 0, 0, 1) <> IIf(cboStock.ItemData(cboStock.ListIndex) = 0, 0, 1) Then
            SaveVsFlexState vsfList, App, Me, IIf(Val(cboStock.Tag) = 0, "���пⷿ", "")
            SaveVsFlexState vsfBatch, App, Me, IIf(Val(cboStock.Tag) = 0, "���пⷿ", "")
            
            RestoreVsFlexState vsfList, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "���пⷿ", "")
            RestoreVsFlexState vsfBatch, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "���пⷿ", "")
        End If
        
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
        
        ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
    End If
End Sub

Private Sub Chk����_Click()
    Dim lstItem As ListItem
    
    For Each lstItem In Me.lst����_S.ListItems
        lstItem.Checked = (Chk����.Value = 1)
    Next
    
    DoEvents
    
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub Form_Activate()
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub Form_Load()
    intChoose���� = Val(zlDatabase.GetPara("��λ", glngSys, 1309, 3))
    Call SetDrugDigit(intChoose����)
    
    bln����� = (zlDatabase.GetPara("�Ƿ���ʾ�޿��ҩƷ", glngSys, 1309) = 1)
    intMonths = Val(zlDatabase.GetPara("Ч�ڱ�������", glngSys, 1309, 3))
    bln����ͣ��ҩƷ = (zlDatabase.GetPara("�Ƿ���ʾͣ��ҩƷ", glngSys, 1309) = 1)
    intFont = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ����ѯ", "����", 0)

    mlngMode = glngModul
    mstrPrivs = gstrprivs
    gstrStockSearchPrivs = gstrprivs 'ר����Կ���ѯ��Ȩ��
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")

    Call mnuViewFontSize_Click(intFont)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '���ر���
    Set ObjReport = New zl9Report.clsReport
    Call ����Ȩ��
    
    Call SetParent(picFind.hWnd, stbThis.hWnd)
    picFind.Top = 80
    picFind.Left = stbThis.Panels(1).Width + 160
    
    If Not ReFreshTreeView() Then Unload Me: Exit Sub
    
    RestoreWinState Me, App.ProductName, Me.Caption
    RestoreVsFlexState vsfList, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "���пⷿ", "")
    RestoreVsFlexState vsfBatch, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "���пⷿ", "")

    Call SetFormat(IniListType.AllList)
    
    stbThis.Panels(2).Picture = picColor
    
'    Set vsfBatch.Icons = imgList.Icons '���ù�����ͼ��ؼ�
End Sub

Private Sub Form_Resize()
    Dim intTop As Integer, intButton As Integer
    If Me.WindowState = 1 Then Exit Sub
    intTop = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    intButton = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    Me.picVLine_S.Top = intTop + Me.ScaleTop
    Me.picVLine_S.Height = Me.ScaleHeight - Me.tvwSection_S.Top - intButton
    If Me.picVLine_S.Left < 500 Then Me.picVLine_S.Left = 500
    If Me.picVLine_S.Left > Me.ScaleWidth - 500 Then Me.picVLine_S.Left = Me.ScaleWidth - 500
    
    Me.tvwSection_S.Left = Me.ScaleLeft
    Me.tvwSection_S.Width = Me.picVLine_S.Left - Me.tvwSection_S.Left
    Me.tvwSection_S.Top = Me.ScaleTop + intTop
    Me.tvwSection_S.Height = Me.lbl����_S.Top - Me.tvwSection_S.Top
    
    If Me.ScaleWidth - Me.picVLine_S.Left - Me.picVLine_S.Width < 500 Then
        Me.Width = Me.picVLine_S.Left + Me.picVLine_S.Width + 500
    End If
    If Me.ScaleHeight - Me.lbl����_S.Top - Me.lbl����_S.Height < 500 Then
        Me.Height = Me.lbl����_S.Top + Me.lbl����_S.Height + 2000
    End If
    If Me.ScaleHeight - Me.lbl����_S.Top - Me.lbl����_S.Height < 500 Then
        Me.Height = Me.lbl����_S.Top + Me.lbl����_S.Height + 2000
    End If
    Me.lbl����_S.Left = Me.tvwSection_S.Left
    Me.lbl����_S.Width = Me.tvwSection_S.Width
    Me.Chk����.Left = Me.lbl����_S.Left + 55
    Me.Chk����.Top = Me.lbl����_S.Top + 30
    With lst����_S
        .Top = Me.lbl����_S.Top + Me.lbl����_S.Height
        .Height = Me.ScaleHeight - .Top - intButton
        .Width = Me.lbl����_S.Width
        .Left = Me.lbl����_S.Left
    End With

    Me.lbl����_S.Left = Me.picVLine_S.Left + Me.picVLine_S.Width - 20
    Me.lbl����_S.Width = Me.ScaleWidth - Me.lbl����_S.Left
    With Me.vsfBatch
        .Left = Me.lbl����_S.Left
        .Width = Me.lbl����_S.Width
    End With
    
    Me.vsfList.Left = Me.lbl����_S.Left
    Me.vsfList.Width = Me.lbl����_S.Width
        
    If Me.vsfBatch.Visible Then
        With Me.vsfBatch
            .Top = Me.lbl����_S.Top + Me.lbl����_S.Height
            .Height = Me.ScaleHeight - .Top - intButton
        End With
        Me.vsfList.Top = intTop + 50
        Me.vsfList.Height = Me.lbl����_S.Top - Me.vsfList.Top
    Else
        Me.vsfList.Top = intTop + 50
        Me.vsfList.Height = Me.ScaleHeight - Me.vsfList.Top - intButton
    End If

'    If cboStock.Text = "���пⷿ" Then
'        vsfList.ColHidden(vsfList.ColIndex("�������")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("�������")) = False
'    Else
'        vsfList.ColHidden(vsfList.ColIndex("�������")) = False
'        vsfBatch.ColHidden(vsfBatch.ColIndex("�������")) = True
'    End If
'    If bln����� = False Then
'        vsfList.ColHidden(vsfList.ColIndex("�������")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("�������")) = True
'    End If
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    SaveWinState Me, App.ProductName, Me.Caption
    SaveVsFlexState vsfList, App, Me, IIf(Val(cboStock.Tag) = 0, "���пⷿ", "")
    SaveVsFlexState vsfBatch, App, Me, IIf(Val(cboStock.Tag) = 0, "���пⷿ", "")
End Sub

Private Sub lbl����_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.lbl����_S.Top = Me.lbl����_S.Top + y
        If Me.lbl����_S.Top < 5000 Then Me.lbl����_S.Top = 5000
        If Me.Height - Me.lbl����_S.Top < 2000 Then Me.lbl����_S.Top = Me.Height - 2000
        Form_Resize
    End If
End Sub

Private Sub lbl����_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.lbl����_S.Top = Me.lbl����_S.Top + y
        If Me.lbl����_S.Top < 2000 Then Me.lbl����_S.Top = 2000
        If Me.Height - Me.lbl����_S.Top < 2000 Then Me.lbl����_S.Top = Me.Height - 2000
        Form_Resize
    End If
End Sub

Private Sub lst����_S_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub SaveVsFlexState(objGrid As VSFlexGrid, objApp As App, objForm As Form, Optional strType As String)
    '����VSFlexGrid�ؼ�����״̬���������м�ֵ����״̬��0-����;1-��ʾ�����п��ж��뷽ʽ
    '��ʽ������1,�м�ֵ1,��״̬1,�п�1,�ж��䷽ʽ1|����2,�м�ֵ2,��״̬2,�п�2,�ж��䷽ʽ2������
    'objApp�����̶���
    'objForm�������ڶ���
    'strType������ҵ�������ͬһ�������ж����ʾ��ʽ���Զ�����ʾ��ʽ����
    Dim strText As String
    Dim i As Integer
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        Exit Sub
    End If
    
    With objGrid
        For i = 0 To .Cols - 1
            strText = IIf(strText = "", "", strText & "|") & .TextMatrix(0, i) & "," & .ColKey(i) & "," & IIf(.colHidden(i) = True, 0, 1) & "," & .ColWidth(i) & "," & .ColAlignment(i)
        Next
    End With
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & objApp.ProductName & "\" & objForm.Name & objForm.Caption & "\" & TypeName(objGrid), objGrid.Name & strType, strText
End Sub

Private Sub RestoreVsFlexState(objGrid As VSFlexGrid, objApp As App, objForm As Form, Optional strType As String)
    '�ָ�VSFlexGrid�ؼ�����״̬��ͬʱ�ָ���˳�򣩣��������м�ֵ����״̬��0-����;1-��ʾ�����п��ж��뷽ʽ
    '��ʽ������1,�м�ֵ1,��״̬1,�п�1,�ж��䷽ʽ1|����2,�м�ֵ2,��״̬2,�п�2,�ж��䷽ʽ2������
    'objApp�����̶���
    'objForm�������ڶ���
    'strType������ҵ�������ͬһ�������ж����ʾ��ʽ���Զ�����ʾ��ʽ����
    Dim strText As String
    Dim i As Integer
    Dim intCols As Integer
    Dim arrText
    
    Dim strName As String
    Dim strkey As String
    Dim blnHidden As Boolean
    Dim dblWidth As Double
    Dim intAlignment As Integer
    
    Dim blnFindKey As Boolean
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        Exit Sub
    End If
    
    strText = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & objApp.ProductName & "\" & objForm.Name & objForm.Caption & "\" & TypeName(objGrid), objGrid.Name & strType, "")
    
    'ע���ֵΪ�գ�������
    If strText = "" Then Exit Sub
    
    arrText = Array()
    arrText = Split(strText, "|")
    
    '��������ȣ��˳�
    If UBound(arrText) + 1 <> objGrid.Cols Then Exit Sub
    
    'û�ҵ��м�ֵ��������
    For i = 0 To UBound(arrText)
        strkey = Split(arrText(i), ",")(1)
        blnFindKey = False
        For intCols = 0 To objGrid.Cols - 1
            If strkey = objGrid.ColKey(intCols) Then
                blnFindKey = True
                Exit For
            End If
        Next
        If blnFindKey = False Then
            Exit Sub
        End If
    Next
    
    '�ָ���״̬
    With objGrid
        .Clear
        For i = 0 To UBound(arrText)
            .TextMatrix(0, i) = Split(arrText(i), ",")(0)
            .ColKey(i) = Split(arrText(i), ",")(1)
            .colHidden(i) = IIf(Val(Split(arrText(i), ",")(2)) = 0, True, False)
            .ColWidth(i) = Val(Split(arrText(i), ",")(3))
            .ColAlignment(i) = Val(Split(arrText(i), ",")(4))

            .FixedAlignment(i) = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub mnuEXCEL_Click()
    grdPrint 1
End Sub

Private Sub mnuFileBatch_Click()
    With Frm������ӡ��ϸ��
        .Show 1, Me
    End With
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    BlnDO = False
    frmDrugQueryParaSet.In_Ȩ�� = mstrPrivs
    frmDrugQueryParaSet.Show 1, Me
    
    If Not BlnDO Then Exit Sub
    
    intChoose���� = Val(zlDatabase.GetPara("��λ", glngSys, 1309, 3))
    Call SetDrugDigit(intChoose����)
    
    bln����� = (zlDatabase.GetPara("�Ƿ���ʾ�޿��ҩƷ", glngSys, 1309) = 1)
    intMonths = Val(zlDatabase.GetPara("Ч�ڱ�������", glngSys, 1309))
    bln����ͣ��ҩƷ = (zlDatabase.GetPara("�Ƿ���ʾͣ��ҩƷ", glngSys, 1309) = 1)
    
    If Me.tvwSection_S.Nodes.count = 0 Then Exit Sub
    Me.tvwSection_S.Tag = ""
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub mnuFilePrint_Click()
    grdPrint 3
End Sub

Private Sub mnuFilePrintSet_Click()
     zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
  grdPrint 0
End Sub
Private Sub grdPrint(blnIsPreview As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������
    '     blnIsPreview: 0��ʾԤ�� 1��ʾ�����EXCEL ������ʾ��ӡ
    '���أ�
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    objPrint.Title.Text = "ҩƷ����ѯ"
    Set objRow = New zlTabAppRow
    objRow.Add "�ⷿ��" & Me.cboStock.Text
    objRow.Add "ҩƷ��;��" & Me.tvwSection_S.SelectedItem.Text
    objRow.Add "��ֹ���ڣ�" & Format(Sys.Currentdate, "yyyy��MM��DD��")
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:MM")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = vsfList
    
    mblnExportState = True
    
    If blnIsPreview = 0 Then
         zlPrintOrView1Grd objPrint, 2
    Else
      If blnIsPreview = 1 Then
            zlPrintOrView1Grd objPrint, 3
      Else
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
      End If
    End If
    Set objPrint = Nothing
    mblnExportState = False
End Sub

Private Sub mnuHelpAbout_Click()
   ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ������ⷿ=�ⷿid������=����id��ҩƷ=ҩƷid
    Dim strReportName As String
    
    strReportName = Split(mnuReportItem(Index).Tag, ",")(1)
    
    Select Case strReportName
        Case "ZL1_INSIDE_1309_2"
            Call openҩƷ��ϸ��
        Case "ZL1_INSIDE_1309_3"
            Call openҩƷ��ϸ��
        Case "ZL1_INSIDE_1309_1"
            Call openҩƷ����
        Case Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "ҩƷ=", _
                "�ⷿ=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
                "����=" & IIf(SQLCondition.lngҩƷ���� = 0, "", SQLCondition.lngҩƷ����))
    End Select
End Sub

Private Sub mnuViewFind_Click()
    Dim strFind As String
    Me.tvwSection_S.Tag = ""
    strFind = Frm������.GetSearch(Me, _
         SQLCondition.strͨ����, _
         SQLCondition.str����, _
         SQLCondition.str����, _
         SQLCondition.str����, _
         SQLCondition.str���, _
         SQLCondition.str����)
    
    If strFind = "" Then Exit Sub
    If Not ReFreshDrugData(cboStock.ItemData(cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2)), strFind, False) Then Exit Sub
    Me.tvwSection_S.Tag = "T"
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True

    Select Case Index
    Case 0
        Me.vsfList.Font.Size = 9
        Me.tvwSection_S.Font.Size = 9
        vsfBatch.Font.Size = 9
     Case 1
        Me.vsfList.Font.Size = 11
        Me.tvwSection_S.Font.Size = 11
        vsfBatch.Font.Size = 11
    Case 2
        Me.vsfList.Font.Size = 15
        Me.tvwSection_S.Font.Size = 15
        vsfBatch.Font.Size = 15
    End Select
    intFont = Index
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ����ѯ", "����", intFont
    
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.vsfList.ForeColor)
    Me.vsfList.ForeColor = lngForeColor
    
End Sub

Private Sub mnuViewBackColor_Click()
    Dim lngBackColor As Long
    lngBackColor = zlGetColor(Me.vsfList.BackColor)
    Me.vsfList.BackColor = lngBackColor
End Sub


Private Sub mnuViewNoVerify_Click()
    frmQueryUnVerify.ShowCard Me, cboStock, intChoose����, mintNumberDigit
End Sub

Private Sub mnuViewRefresh_Click()
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub



Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.mnuViewToolbarText.Enabled = Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub
Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize
End Sub

Private Sub vsfBatch_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '������ѡ���б�
    Call InitColSelList(IniListType.BatchList, vsfBatch)
End Sub

Private Sub vsfBatch_Click()
    vsfBatch.BackColorSel = glngRowByFocus
'    vsfBatch.GridColorFixed = &H80000008
'    vsfBatch.GridColor = &H80000008
    
    vsfList.BackColorSel = glngRowByNotFocus
'    vsfList.GridColorFixed = &H80000010
'    vsfList.GridColor = &H80000010
End Sub

Private Sub vsfBatch_EnterCell()
    If mblnRefresh = True Then Exit Sub
    If vsfBatch.Row = 0 Then Exit Sub

    With vsfBatch
        '��ǰ������
        .Redraw = flexRDNone

        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfBatch_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 2 Then '��ѡ����
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
        
        If vsfBatch.MouseRow <> 0 Then Exit Sub
        
        InitColSelList IniListType.BatchList, vsfBatch
        
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfBatch.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfBatch.colHidden(.RowData(i)) Or vsfBatch.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = vsfBatch.Top + vsfBatch.RowHeight(0)
                .Width = 1750
'                If .Top + .Height > Me.ScaleHeight - vsfBatch.Top Then
'                    .Height = Me.ScaleHeight - .Top - vsfBatch.Top
'                    .Width = 1750
'                Else
'                    .Width = 1470
'                End If
                
                .Left = vsfBatch.Left + x
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    lngCol = vsfColSel.RowData(Row)
    If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
        If Val(vsfColSel.Tag) = IniListType.MainList Then
            vsfList.colHidden(lngCol) = False
        Else
            vsfBatch.colHidden(lngCol) = False
        End If
    Else
        If Val(vsfColSel.Tag) = IniListType.MainList Then
            vsfList.colHidden(lngCol) = True
        Else
            vsfBatch.colHidden(lngCol) = True
        End If
    End If
End Sub

Private Sub vsfColSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfColSel
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub


Private Sub vsfColSel_LostFocus()
    vsfColSel.Visible = False
End Sub

Private Sub vsfColSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsfColSel.Cell(flexcpForeColor, Row, 1) = vsfColSel.BackColorFixed Then Cancel = True
End Sub


Private Sub InitColSelList(ByVal intListType As Integer, ByVal objGrid As VSFlexGrid)
    Dim i As Integer
    Dim strUnShow As String
    
    If intListType = IniListType.MainList Then
        strUnShow = mstrUnShow_List
    ElseIf intListType = IniListType.BatchList Then
        strUnShow = mstrUnShow_Batch
    End If
        
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 0 To objGrid.Cols - 1
            '���ڲ�������ʾ�б���в��ܼ�����ѡ���б�
            If InStr(1, ";" & strUnShow & ";", ";" & objGrid.ColKey(i) & ";") = 0 Then
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 1) = objGrid.TextMatrix(0, i)
                .RowData(.rows - 1) = i
                
                '�п�Ϊ�ջ������ص�������Ϊ����ѡ
                If Not (objGrid.ColWidth(i) = 0 Or objGrid.colHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
'                'ָ����������Ϊ������������
'                If IsInString(mstrUnallowSetColHide, objGrid.ColKey(i), ";") = True Then
'                    .Cell(flexcpForeColor, .Rows - 1, 1) = .BackColorFixed
'                End If
            End If
        Next
    End With
End Sub

Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    '������ѡ���б�
    Call InitColSelList(IniListType.MainList, vsfList)
End Sub

Private Sub vsfList_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfList
        If Col = .ColIndex("����") Then
            .Col = .ColIndex("�������")
            .Sort = Order
        End If
    End With
End Sub


Private Sub vsfList_Click()
    vsfList.BackColorSel = glngRowByFocus
'    vsfList.GridColorFixed = &H80000008
'    vsfList.GridColor = &H80000008
    
    vsfBatch.BackColorSel = glngRowByNotFocus
    vsfBatch.ForeColorSel = IIf(Val(vsfBatch.RowData(vsfBatch.Row)) = 0, glng��ɫ, glng����)
'    vsfBatch.GridColorFixed = &H80000010
'    vsfBatch.GridColor = &H80000010
End Sub

Private Sub vsfList_DblClick()
'    If DataRecordSet Is Nothing Then Exit Sub
'    If Not (DataRecordSet.State = 1) Then Exit Sub
'    If DataRecordSet.RecordCount = 0 Then Exit Sub
    
    If vsfList.MouseRow = 0 Or vsfList.MouseRow = -1 Then Exit Sub
    
    If vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("����")) = "" Then Exit Sub
    
    If cboStock.ItemData(cboStock.ListIndex) = 0 Then
        Call ObjReport.ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "ҩƷ=" & vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("����")) & "|" & Val(vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("ҩƷID"))), "�ⷿ=���пⷿ|is not null", "��λ=" & Choose(intChoose����, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(intChoose����, 1, 3, 2, 4))    ' , "��ʼ����=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "��������=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    Else
        Call ObjReport.ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "ҩƷ=" & vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("����")) & "|" & Val(vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("ҩƷID"))), "�ⷿ=" & cboStock.Text & "|=  " & cboStock.ItemData(cboStock.ListIndex), "��λ=" & Choose(intChoose����, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(intChoose����, 1, 3, 2, 4))  ' , "��ʼ����=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "��������=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    End If
End Sub

Private Sub vsfList_EnterCell()
    On Error Resume Next
    
    If mblnExportState = True Then Exit Sub
    If mblnRefresh = True Then Exit Sub
    If vsfList.Row = 0 Then Exit Sub
    If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("����")) = "" Then
        RefreshBatch Me.cboStock.ItemData(Me.cboStock.ListIndex), 0
        Me.vsfBatch.Visible = False
        Me.lbl����_S.Visible = False
        Me.vsfBatch.rows = 2
        Me.vsfBatch.Redraw = flexRDDirect
        Call Form_Resize
        Exit Sub
    End If
    
    With vsfList
        '��ǰ������
        .Redraw = flexRDNone
        
        '����ҩƷ��ͣ��ҩƷ��ǰ��ɫ
        .ForeColorSel = IIf(Trim(.TextMatrix(.Row, vsfList.ColIndex("����ʱ��"))) = "", glng��ɫ, glng��ɫ)
        
        .Redraw = flexRDDirect
    
        '��ȡ������Ϣ
        RefreshBatch Me.cboStock.ItemData(Me.cboStock.ListIndex), Val(.TextMatrix(.Row, .ColIndex("ҩƷID")))
        
        If Me.tvwSection_S.Tag <> "T" Then Exit Sub
        
        Err = 0
       
        Me.tvwSection_S.Nodes("_" & vsfList.TextMatrix(.Row, .ColIndex("��;����ID"))).Selected = True
        Me.tvwSection_S.Nodes("_" & vsfList.TextMatrix(.Row, .ColIndex("��;����ID"))).Expanded = True
    End With
End Sub

Private Sub picVLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVLine_S.Left = Me.picVLine_S.Left + x
        Form_Resize
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "Ԥ��"
            mnuFilePrintView_Click
        Case "��ӡ"
            grdPrint 3
        Case "����"
            Call openҩƷ����
        Case "��ϸ"
            Call openҩƷ��ϸ��
        Case "����"
            mnuViewFind_Click
        Case "ˢ��"
            mnuViewRefresh_Click
        Case "����"
             PopupMenu mnuViewFont
        Case "ǰ��ɫ"
            mnuViewForeColor_Click
        Case "����ɫ" '
            mnuViewBackColor_Click
        Case "����"
            mnuHelpTitle_Click
        Case "�˳�"
           mnufileexit_Click
        End Select
    End With
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewToolbar
    End If
End Sub

Private Sub tvwSection_S_GotFocus()
    If Me.tvwSection_S.Tag = "T" Then Me.tvwSection_S.Tag = "F"
End Sub

Private Sub tvwSection_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.tvwSection_S.Tag = "T" Then Exit Sub
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Function ReFreshTreeView() As Boolean
    '-------------------------------------------------------------------------
    '--����:���»�ȡ�����ͽṹ����
    '--����:
    '--����:������ݿ�򿪳ɹ�,��True,���򷵻�False
    '-------------------------------------------------------------------------
    Dim objNode As Node
    Dim RecDept As New ADODB.Recordset
    Dim RecDrug As New ADODB.Recordset
    Dim Str���� As String
    Dim i As Integer
    Dim RsTreeRecordset As ADODB.Recordset
    
    ReFreshTreeView = False
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select distinct a.ID,(a.���� || '-' || a.����) As ���� From ���ű� a,��������˵�� b,�������ʷ��� C " & _
              "Where (a.վ�� = [2] Or a.վ�� is Null) And a.id=b.����id And b.��������=c.���� And (c.���� in ('H','I','J','K','L','M','N'))" & _
              IIf(zlStr.IsHavePrivs(mstrPrivs, "���пⷿ"), "", " And A.id In (Select ����ID From ������Ա Where ��ԱID=[1])") & _
              "  and (to_char(a.����ʱ��,'yyyy-mm-dd')='3000-01-01' or a.����ʱ�� is null) " & _
              "Order By A.���� || '-' || A.���� "
    Set RecDept = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���пⷿ", UserInfo.�û�ID, gstrNodeNo)
    
    With RecDept
        If .RecordCount = 0 Then
            MsgBox "ҩ����ϵδ������Ȩ�޲��㣬����ִ�б�����!", vbInformation, gstrSysName
            Exit Function
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "���пⷿ") Then
            Me.cboStock.Clear
            Me.cboStock.AddItem "���пⷿ"
            Me.cboStock.ItemData(Me.cboStock.NewIndex) = 0
            Me.cboStock.ListIndex = Me.cboStock.NewIndex
        End If
        Do While Not .EOF
            Me.cboStock.AddItem .Fields("����").Value
            Me.cboStock.ItemData(Me.cboStock.NewIndex) = .Fields("ID").Value
            .MoveNext
        Loop
        
        gstrSQL = "Select Distinct ����id From ������Ա Where ȱʡ = 1 And ��Աid = [1]"
        Set RecDept = zlDatabase.OpenSQLRecord(gstrSQL, "ȡȱʡ����", UserInfo.�û�ID)
        
        Me.cboStock.ListIndex = 0
        
        '��λ��ȱʡ����
        If Not RecDept.EOF Then
            For i = 0 To Me.cboStock.ListCount - 1
                If Me.cboStock.ItemData(i) = RecDept!����ID Then
                    Me.cboStock.ListIndex = i
                    Exit For
                End If
            Next
        End If
        
        Me.cboStock.Tag = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    End With
    
    Str���� = ""
    StrSort = ""
    If zlStr.IsHavePrivs(mstrPrivs, "����ҩ") Then
        Bln����ҩ = True
        Str���� = "1"
        StrSort = ",'5'"
    Else
        Bln����ҩ = False
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "�г�ҩ") Then
        Bln�г�ҩ = True
        If Str���� = "" Then
            Str���� = "2"
        Else
            Str���� = Str���� & ",2"
        End If
        StrSort = StrSort & ",'6'"
    Else
        Bln�г�ҩ = False
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "�в�ҩ") Then
        Bln�в�ҩ = True
        If Str���� = "" Then
            Str���� = "3"
        Else
            Str���� = Str���� & ",3"
        End If
        StrSort = StrSort & ",'7'"
    Else
        Bln�в�ҩ = False
    End If
    
    If Str���� = "" Then
        MsgBox "�Բ��𣬱������һ������ҩƷ���ʵ�Ȩ�ޣ�����ϵͳ����Ա��ϵ��", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select a.id,a.�ϼ�id,a.����,Decode(a.����,1,'5',2,'6','7') As ���� " & _
              "From ���Ʒ���Ŀ¼ a " & _
              "Where a.���� in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) " & _
              "Start with a.�ϼ�id is null " & _
              "Connect by prior a.id=a.�ϼ�id " & _
              "Order by level,a.id"
    Set RsTreeRecordset = zlDatabase.OpenSQLRecord(gstrSQL, "ҩƷ��;����", Str����)
    
    With RsTreeRecordset
        If .RecordCount = 0 Then
            MsgBox "ҩƷ��;��ϵδ����������ִ�б�����!", vbInformation, gstrSysName
            Exit Function
        End If
        Me.tvwSection_S.Nodes.Clear
        
        If Bln����ҩ = True Then
            Me.tvwSection_S.Nodes.Add , , "R" & "5", "����ҩ", "child"
        End If
        
        If Bln�в�ҩ = True Then
            Me.tvwSection_S.Nodes.Add , , "R" & "7", "�в�ҩ", "child"
        End If
        
        If Bln�г�ҩ = True Then
            Me.tvwSection_S.Nodes.Add , , "R" & "6", "�г�ҩ", "child"
        End If
        
        Do While Not .EOF
            If IsNull(.Fields("�ϼ�id").Value) Then
                Set objNode = Me.tvwSection_S.Nodes.Add("R" & !����, 4, "_" & .Fields("id").Value, .Fields("����").Value, "child")
            Else
                Set objNode = Me.tvwSection_S.Nodes.Add("_" & .Fields("�ϼ�id").Value, 4, "_" & .Fields("id").Value, .Fields("����").Value, "child")
            End If
            .MoveNext
         Loop
         
         If tvwSection_S.Nodes(1).Children <> 0 Then
            tvwSection_S.Nodes(1).Child.Selected = True
         Else
            tvwSection_S.Nodes(1).Selected = True
         End If
    End With
    With RecDrug
        gstrSQL = " Select ����,���� From ҩƷ���� "
        Call zlDatabase.OpenRecordset(RecDrug, gstrSQL, "ҩƷ����")
        
        If .RecordCount = 0 Then
            MsgBox "ҩƷ����δ����,����ִ�г���!", vbInformation, gstrSysName
            Exit Function
        End If
        Me.lst����_S.ListItems.Clear
        Do While Not .EOF
            Me.lst����_S.ListItems.Add , "K" & !����, !����
            Me.lst����_S.ListItems("K" & !����).Checked = True
            .MoveNext
        Loop
        Me.lst����_S.ListItems(1).Selected = True
    End With
    ReFreshTreeView = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Unload Me
End Function

Private Function ReFreshDrugData(ByVal lngDeptId As Long, _
    Optional ByVal lngUseId As Long = 0, Optional ByVal strFind As String = "", Optional ByVal Click As Boolean = True) As Boolean
    '-------------------------------------------------------------------------
    '--����:���»�ȡ��ҩƷ�����
    '--����:
    '       lngDeptId:ҩƷ��id
    '       lngUseId:��;idֵ
    '       strFind:���ڿ��ٲ��ң�������롢���ƻ���룩
    '       ClickΪ���ʾ���ѡ��,Ϊ�ٱ�ʾ����
    '--����:
    '-------------------------------------------------------------------------
    Dim strOrder As String, strSql As String, str�շ���ĿĿ¼ As String
    Dim str���� As String
    Dim lstItem As ListItem
    Dim blnAllCheck As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSqlҩ�� As String
    
    blnAllCheck = True
    str���� = ""
    strOrder = ""
    If strFind = "" Then
        str�շ���ĿĿ¼ = " �շ���ĿĿ¼ "
    Else
        str�շ���ĿĿ¼ = "(Select distinct A.ID, A.����, A.����, A.���, A.����, A.�Ƿ���, A.����ʱ��, A.���, A.���㵥λ " & _
                 " From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                 " Where A.ID=B.�շ�ϸĿID " & strFind & ")"
    End If
    
    Call FS.ShowFlash("���ڲ�������,���Ժ� ...", Me)
    DoEvents
    For Each lstItem In Me.lst����_S.ListItems
        If lstItem.Checked Then
            str���� = str���� & "," & lstItem
        Else
            blnAllCheck = False
        End If
    Next
    If str���� <> "" Then
        If blnAllCheck = True Then
            str���� = ""
        Else
            str���� = Mid(str����, 2) & ",����"
        End If
    Else
        str���� = "С��"
    End If
    
    If lngDeptId = 0 Then
        Select Case intChoose����
        Case 1
            gstrSQL = ",A.���㵥λ as ��λ,'' as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) as ��ǰ�ۼ�,nvl(M.����ϵ��,0) as ϵ��1,Sum(B.��������) As ��������,Sum(B.ʵ������) As ʵ������,Sum(B.ʵ�ʽ��) As ʵ�ʽ��,Sum(B.ʵ�ʲ��) As ʵ�ʲ��,sum(b.ƽ���ɱ���) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,1 as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��, '' �ⷿ��λ "
            strOrder = " Group by M.ҩƷID,X.����id,A.����,M.����ҩ��,M.��ʶ��,A.����,L.����,A.���,decode(b.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),b.�ϴβ���),m.ԭ����,M.ҩ�����,A.���㵥λ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) ,nvl(M.����ϵ��,0),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����"
        Case 2
            gstrSQL = ",M.���ﵥλ as ��λ,'' as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) * Nvl(m.�����װ, 0) as ��ǰ�ۼ�,nvl(M.�����װ,0) as ϵ��1,Sum(B.��������/Decode(M.�����װ,0,1,null,1,M.�����װ)) as ��������,Sum(B.ʵ������/Decode(M.�����װ,0,1,null,1,M.�����װ)) as ʵ������,Sum(B.ʵ�ʽ��) As ʵ�ʽ��,Sum(B.ʵ�ʲ��) As ʵ�ʲ��,sum(b.ƽ���ɱ���)*Decode(M.�����װ,0,1,null,1,M.�����װ) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.�����װ,0,1,null,1,M.�����װ) as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��, '' �ⷿ��λ "
            strOrder = " Group by M.ҩƷID,X.����id,A.����,M.����ҩ��,M.��ʶ��,A.����,L.����,A.���,decode(b.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),b.�ϴβ���),m.ԭ����,M.ҩ�����,M.���ﵥλ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) * Nvl(m.�����װ, 0),nvl(M.�����װ,0),Decode(M.�����װ,0,1,null,1,M.�����װ),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����"
        Case 3
            gstrSQL = ",M.ҩ�ⵥλ as ��λ,'' as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) * Nvl(m.ҩ���װ, 0) as ��ǰ�ۼ�,nvl(M.ҩ���װ,0) as ϵ��1,Sum(B.��������/Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ)) as ��������, Sum(B.ʵ������/Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ)) as ʵ������,Sum(B.ʵ�ʽ��) As ʵ�ʽ��,Sum(B.ʵ�ʲ��) As ʵ�ʲ��,sum(b.ƽ���ɱ���)*Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ) as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��,'' �ⷿ��λ "
            strOrder = " Group by M.ҩƷID,X.����id,A.����,M.����ҩ��,M.��ʶ��,A.����,L.����,A.���,decode(b.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),b.�ϴβ���),m.ԭ����,M.ҩ�����,M.ҩ�ⵥλ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) * Nvl(m.ҩ���װ, 0),nvl(M.ҩ���װ,0),Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����"
        Case 4
            gstrSQL = ",M.סԺ��λ as ��λ,'' as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) * Nvl(m.סԺ��װ, 0) as ��ǰ�ۼ�,nvl(M.סԺ��װ,0) as ϵ��1,Sum(B.��������/Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ)) as ��������, Sum(B.ʵ������/Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ)) as ʵ������,Sum(B.ʵ�ʽ��) As ʵ�ʽ��,Sum(B.ʵ�ʲ��) As ʵ�ʲ��,sum(b.ƽ���ɱ���)*Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ) as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��,'' �ⷿ��λ "
            strOrder = " Group by M.ҩƷID,X.����id,A.����,M.����ҩ��,M.��ʶ��,A.����,L.����,A.���,decode(b.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),b.�ϴβ���),m.ԭ����,M.ҩ�����,M.סԺ��λ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(b.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), b.���ۼ�)) * Nvl(m.סԺ��װ, 0),nvl(M.סԺ��װ,0),Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����"
        End Select
    Else
        Select Case intChoose����
        Case 1
            gstrSQL = ",A.���㵥λ as ��λ,Nvl(Avg(S.�ϴβɹ���), m.�ɱ���) as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) As ��ǰ�ۼ�,nvl(M.����ϵ��,0) as ϵ��1,Sum(S.��������) as ��������, Sum(S.ʵ������) as ʵ������,sum(i.����) as ����,Sum(S.ʵ�ʽ��) as ʵ�ʽ��,Sum(S.ʵ�ʲ��) as ʵ�ʲ��,sum(s.ƽ���ɱ���) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,1 as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��, C.�ⷿ��λ "
            strOrder = " Group by M.ҩƷID,A.����,M.����ҩ��,M.��ʶ��,X.����id,A.����,L.����,A.���,decode(s.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),s.�ϴβ���),Nvl(m.ԭ����, s.ԭ����),nvl(M.���Ч��,0),s.���Ч��,s.����,M.ҩ�����,A.���㵥λ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)),nvl(M.����ϵ��,0),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����, C.�ⷿ��λ, m.�ɱ���"
        Case 2
            gstrSQL = ",M.���ﵥλ as ��λ,Nvl(Avg(S.�ϴβɹ���*nvl(M.�����װ,0)), m.�ɱ���*nvl(M.�����װ,0)) as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) * Nvl(m.�����װ, 0) As ��ǰ�ۼ�,nvl(M.�����װ,0) as ϵ��1,Sum(S.�������� /Decode(M.�����װ,0,1,null,1,M.�����װ)) as ��������, Sum(S.ʵ������ /Decode(M.�����װ,0,1,null,1,M.�����װ)) as ʵ������,sum(i.����/Decode(M.�����װ,0,1,null,1,M.�����װ)) as ����,Sum(S.ʵ�ʽ��) as ʵ�ʽ��,Sum(S.ʵ�ʲ��) as ʵ�ʲ��,sum(s.ƽ���ɱ���)*Decode(M.�����װ,0,1,null,1,M.�����װ) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.�����װ,0,1,null,1,M.�����װ) as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��, C.�ⷿ��λ "
            strOrder = " Group by M.ҩƷID,A.����,M.����ҩ��,M.��ʶ��,X.����id,A.����,L.����,A.���,decode(s.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),s.�ϴβ���),Nvl(m.ԭ����, s.ԭ����),nvl(M.���Ч��,0),s.���Ч��,s.����,M.ҩ�����,M.���ﵥλ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) * Nvl(m.�����װ, 0),nvl(M.�����װ,0),Decode(M.�����װ,0,1,null,1,M.�����װ),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����, C.�ⷿ��λ, m.�ɱ���"
        Case 3
            gstrSQL = ",M.ҩ�ⵥλ as ��λ,Nvl(Avg(S.�ϴβɹ���*nvl(M.ҩ���װ,0)), m.�ɱ���*nvl(M.ҩ���װ,0)) as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) * Nvl(m.ҩ���װ, 0) As ��ǰ�ۼ�,nvl(M.ҩ���װ,0) as ϵ��1,Sum(S.�������� /Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ)) as ��������,Sum(S.ʵ������ /Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ)) as ʵ������,sum(i.����/Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ)) as ����,Sum(S.ʵ�ʽ��) as ʵ�ʽ��,Sum(S.ʵ�ʲ��) as ʵ�ʲ��,sum(s.ƽ���ɱ���)*Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ) as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��, C.�ⷿ��λ "
            strOrder = " Group by M.ҩƷID,A.����,M.����ҩ��,M.��ʶ��,X.����id,A.����,L.����,A.���,decode(s.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),s.�ϴβ���),Nvl(m.ԭ����, s.ԭ����),nvl(M.���Ч��,0),s.���Ч��,s.����,M.ҩ�����,M.ҩ�ⵥλ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) * Nvl(m.ҩ���װ, 0),nvl(M.ҩ���װ,0),Decode(M.ҩ���װ,0,1,null,1,M.ҩ���װ),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����, C.�ⷿ��λ, m.�ɱ���"
        Case 4
            gstrSQL = ",M.סԺ��λ as ��λ,Nvl(Avg(S.�ϴβɹ���*nvl(M.סԺ��װ,0)), m.�ɱ���*nvl(M.סԺ��װ,0)) as �ϴβɹ���,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) * Nvl(m.סԺ��װ, 0) As ��ǰ�ۼ�,nvl(M.סԺ��װ,0) as ϵ��1,Sum(S.�������� /Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ)) as ��������, Sum(S.ʵ������ /Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ)) as ʵ������,sum(i.����/Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ)) as ����,Sum(S.ʵ�ʽ��) as ʵ�ʽ��,Sum(S.ʵ�ʲ��) as ʵ�ʲ��,sum(s.ƽ���ɱ���)*Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ) ƽ���ɱ���,Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')) ����ʱ��,Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ) as ����, Nvl(A.�Ƿ���, 0) ���, G.���� As �ϴι�Ӧ��, C.�ⷿ��λ "
            strOrder = " Group by M.ҩƷID,A.����,M.����ҩ��,M.��ʶ��,X.����id,A.����,L.����,A.���,decode(s.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),s.�ϴβ���),Nvl(m.ԭ����, s.ԭ����),nvl(M.���Ч��,0),s.���Ч��,s.����,M.ҩ�����,M.סԺ��λ,Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(p.�ּ�, 0), Decode(Nvl(s.���ۼ�, 0), 0, Nvl(p.�ּ�, 0), s.���ۼ�)) * Nvl(m.סԺ��װ, 0),nvl(M.סԺ��װ,0),Decode(M.סԺ��װ,0,1,null,1,M.סԺ��װ),Decode(To_Char(A.����ʱ��,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.����ʱ��,'yyyy-MM-dd')),Nvl(A.�Ƿ���, 0),G.����, C.�ⷿ��λ, m.�ɱ���"
        End Select
    End If
    
    On Error GoTo ErrHand:

    If lngDeptId = 0 Then
        strSql = "SELECT Distinct M.ҩƷID,X.����ID As ��;����ID,A.����,M.����ҩ��,M.��ʶ�� As ҩ����,A.���� As ͨ����,L.���� As ��Ʒ��,A.���,decode(b.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),b.�ϴβ���) AS ����,m.ԭ���� as ԭ����, NULL AS Ч��,DECODE(M.ҩ�����,1,'��','��') AS ҩ����� " & gstrSQL & _
                " FROM ҩƷ��� M,�շѼ�Ŀ P," & str�շ���ĿĿ¼ & " A "
                
        strSql = strSql & " ,(Select a.ҩƷid, Avg(a.�ϴβɹ���) As �ϴγɱ���,Sum(a.ʵ������ * a.ƽ���ɱ���) / Decode(Sum(Nvl(a.ʵ������, 0)), 0, 1, Sum(Nvl(a.ʵ������, 0))) as ƽ���ɱ���," & _
                " Sum(a.ʵ������ * a.���ۼ�) / Decode(Sum(Nvl(a.ʵ������, 0)), 0, 1, Sum(Nvl(a.ʵ������, 0))) as ���ۼ�, '' �ϴβ���, Max(nvl(a.����,0)) As ����, Sum(a.��������) As ��������," & _
                " Sum(a.ʵ������) As ʵ������, Sum(a.ʵ�ʽ��) As ʵ�ʽ��, Sum(a.ʵ�ʲ��) As ʵ�ʲ�� " & _
                " From ҩƷ��� A, ҩƷ��� B, ������ĿĿ¼ C, �շ���ĿĿ¼ D " & _
                " Where a.ҩƷid = b.ҩƷid And b.ҩ��id = c.Id And b.ҩƷid = d.Id And ���� = 1 and (Nvl(a.��������,0)<>0 or Nvl(a.ʵ������,0)<>0 or Nvl(a.ʵ�ʽ��,0)<>0 or Nvl(a.ʵ�ʲ��,0)<>0) "
        If Click Then
            strSql = strSql & (IIf(lngUseId = 0, " AND D.���=[10]", _
                 " AND C.����ID IN ( SELECT ID FROM ���Ʒ���Ŀ¼ Q Where Q.���� In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=�ϼ�ID)"))
        End If
        strSql = strSql & " Group By a.ҩƷid) B "
        
        strSql = strSql & " ,������ĿĿ¼ X,ҩƷ���� T,�շ���Ŀ���� L, ��Ӧ�� G " & _
                " WHERE M.ҩ��ID=X.ID And X.ID=T.ҩ��ID And Nvl(M.�ϴι�Ӧ��id, 0) = G.ID(+) " & _
                " AND M.ҩƷID=P.�շ�ϸĿID AND SYSDATE BETWEEN P.ִ������ AND NVL(P.��ֹ����,SYSDATE) " & _
                GetPriceClassString("P") & _
                IIf(bln����ͣ��ҩƷ, "", " AND (TO_CHAR(A.����ʱ��, 'YYYY-MM-DD') = '3000-01-01' OR A.����ʱ�� IS NULL) ") & _
                " AND A.ID=M.ҩƷID AND M.ҩƷID=B.ҩƷID(+)  " & _
                " And M.ҩƷID=L.�շ�ϸĿID(+) And L.����(+)=3 And L.����(+)=1"
        If Not Click Then
            strSql = strSql & " And A.��� in (" & Mid(StrSort, 2) & ")"
        Else
            strSql = strSql & (IIf(lngUseId = 0, " AND A.���=[10]", _
                 " AND X.����ID IN ( SELECT ID FROM ���Ʒ���Ŀ¼ Q Where Q.���� In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=�ϼ�ID)"))
        End If
        
        If str���� <> "" Then
'            StrSql = StrSql & " And T.ҩƷ����=E.Column_Value "
            strSql = strSql & " And Instr(',' || [11] || ',' , T.ҩƷ����)>0 "
        End If

        strSql = strSql + strOrder
    Else
        strSql = "SELECT M.ҩƷID,X.����ID As ��;����ID,A.����,M.����ҩ��,M.��ʶ�� As ҩ����,A.���� As ͨ����,L.���� As ��Ʒ��,A.���,decode(s.�ϴβ���,null,decode(m.�ϴβ���,null,a.����,m.�ϴβ���),s.�ϴβ���) AS ����,Nvl(m.ԭ����, s.ԭ����) as ԭ����,NVL(M.���Ч��,0) AS Ч��,s.����,s.���Ч��,DECODE(M.ҩ�����,1,'��','��') AS ҩ����� " & gstrSQL & _
                " FROM ҩƷ��� M,�շѼ�Ŀ P,(select nvl(����,0) as ����,ҩƷid from ҩƷ�����޶� where �ⷿid=[9]) I ," & str�շ���ĿĿ¼ & _
                " A,������ĿĿ¼ X,ҩƷ���� T "
        strSql = strSql & " ,(Select b.ҩƷid, b.�ϴβɹ���, b.ƽ���ɱ���,b.���ۼ�,a.�ϴβ���,a.ԭ����, b.��������,b.ʵ������,b.ʵ�ʽ��, b.ʵ�ʲ��, A.�ϴι�Ӧ��id,Decode(Sign(Add_Months(Sysdate, " & intMonths & ") - Ч��), -1, 0, 1) ����,Ч�� as ���Ч�� " & _
                "  From ҩƷ��� a,(SELECT a.ҩƷID,avg(a.�ϴβɹ���) AS �ϴβɹ���,Sum(a.ʵ������ * a.ƽ���ɱ���) / Decode(Sum(Nvl(a.ʵ������, 0)), 0, 1, Sum(Nvl(a.ʵ������, 0))) ƽ���ɱ���," & _
                " Sum(a.ʵ������ * a.���ۼ�) / Decode(Sum(Nvl(a.ʵ������, 0)), 0, 1, Sum(Nvl(a.ʵ������, 0))) ���ۼ�, " & _
                " Max(nvl(a.����,0)) AS ����,SUM(a.��������) AS ��������,SUM(a.ʵ������) AS ʵ������,SUM(a.ʵ�ʽ��) AS ʵ�ʽ��,SUM(a.ʵ�ʲ��) AS ʵ�ʲ�� " & _
                "  FROM ҩƷ��� A, ҩƷ��� B, ������ĿĿ¼ C, �շ���ĿĿ¼ D " & _
                " WHERE a.ҩƷid = b.ҩƷid And b.ҩ��id = c.Id And b.ҩƷid = d.Id And a.�ⷿID=[9] AND a.����=1 " & _
                " and (Nvl(a.��������,0)<>0 or Nvl(a.ʵ������,0)<>0 or Nvl(a.ʵ�ʽ��,0)<>0 or Nvl(a.ʵ�ʲ��,0)<>0)  "
        If Click Then
            strSql = strSql & (IIf(lngUseId = 0, " AND D.���=[10]", _
                 " AND C.����ID IN ( SELECT ID FROM ���Ʒ���Ŀ¼ Q Where Q.���� In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=�ϼ�ID)"))
        End If
        strSql = strSql & "  GROUP BY a.ҩƷID) b Where a.�ⷿID=[9] and a.ҩƷid=b.ҩƷid And a.���� = 1 And nvl(a.����,0) = b.����) S, "
        
        strSql = strSql & " �շ���Ŀ���� L, ��Ӧ�� G, (Select Distinct �շ�ϸĿid, ִ�п���id From �շ�ִ�п��� Where ִ�п���id=[9]) K, " & _
                " (Select �ⷿid, ҩƷid, �ⷿ��λ From ҩƷ�����޶� Where �ⷿid = [9] And �ⷿ��λ Is Not Null) C "
        
        strSql = strSql & " WHERE  i.ҩƷid(+)=m.ҩƷid and M.ҩƷID =A.ID And M.ҩ��ID =X.ID And X.ID=T.ҩ��ID And Nvl(S.�ϴι�Ӧ��id, 0) = G.ID(+) " & _
                " And M.ҩƷID=L.�շ�ϸĿID(+) And L.����(+)=3 AND L.����(+)=1 And M.ҩƷid = C.ҩƷid(+) " & _
                " AND M.ҩƷID=S.ҩƷID(+) And M.ҩƷID = K.�շ�ϸĿid " & _
                IIf(bln����ͣ��ҩƷ, "", " AND (TO_CHAR(A.����ʱ��, 'YYYY-MM-DD') = '3000-01-01' OR A.����ʱ�� IS NULL)  ") & _
                "       And M.ҩƷID+0=P.�շ�ϸĿID AND SYSDATE BETWEEN P.ִ������ AND NVL(P.��ֹ����,SYSDATE) " & _
                GetPriceClassString("P")
        
        If Not Click Then
            strSql = strSql & " And A.��� in (" & Mid(StrSort, 2) & ")"
        Else
            strSql = strSql & (IIf(lngUseId = 0, " AND A.���=[10]", _
                 " AND X.����ID IN ( SELECT ID FROM ���Ʒ���Ŀ¼ Q Where Q.���� In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=�ϼ�ID)"))
        End If
        
        If str���� <> "" Then
            strSql = strSql & " And T.ҩƷ���� in (select * from Table(Cast(f_Str2list([11]) As zlTools.t_Strlist))) "
        End If
        
        strSql = strSql + strOrder
    End If
    gstrSQL = "Select * From (" & strSql & ")" & IIf(bln�����, " Where NVL(ʵ������,0)<>0 ", "")
    gstrSQL = gstrSQL & " Order By ����"
    
    SQLCondition.lngҩƷ���� = lngUseId
    SQLCondition.lng�ⷿID = lngDeptId
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.strͨ����, _
            SQLCondition.str����, _
            SQLCondition.str����, _
            SQLCondition.str����, _
            SQLCondition.str���, _
            SQLCondition.str����, _
            SQLCondition.strҩƷ��Ϣ, _
            SQLCondition.lngҩƷ����, _
            SQLCondition.lng�ⷿID, _
            Mid(tvwSection_S.SelectedItem.Key, 2), _
            str����)
    
    With rsData
        If .RecordCount = 0 Then
            mnuExcel.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFilePrintView.Enabled = False
            mnuViewFind.Enabled = False
'            mnuViewList.Enabled = False
            tbrThis.Buttons.Item(1).Enabled = False
            tbrThis.Buttons.Item(2).Enabled = False
            tbrThis.Buttons.Item(6).Enabled = False
            tbrThis.Buttons.Item(7).Enabled = False
        Else
            mnuExcel.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFilePrintView.Enabled = True
            mnuViewFind.Enabled = True
'            mnuViewList.Enabled = True
            tbrThis.Buttons.Item(1).Enabled = True
            tbrThis.Buttons.Item(2).Enabled = True
            tbrThis.Buttons.Item(6).Enabled = True
            tbrThis.Buttons.Item(7).Enabled = True
        End If
    End With
        
    Call FS.StopFlash
    Call SetFormat(IniListType.MainList)
    If Not rsData.EOF Then DataBound rsData
    
    With vsfList
        .Row = 1
        Call vsfList_EnterCell
    End With
    
    ReFreshDrugData = (rsData.RecordCount <> 0)
    If ReFreshDrugData Then Me.vsfList.SetFocus
    Exit Function
ErrHand:
    Call FS.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
End Function

Private Sub RefreshBatch(lng�ⷿID As Long, lngҩƷid As Long)
    '-------------------------------------------------------------------------
    '--����:���»�ȡ��ҩƷ���������
    '--����:
    '       lng�ⷿId:ҩƷ��id
    '       lngҩƷId:��;idֵ
    '--����:
    '-------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intRow As Long
    Dim intCol As Long
    Dim lngColor As Long
    
    Dim int���� As Integer
    Dim intҩ�� As Integer
    Dim lng���� As Long
    Dim strTemp As String
    Dim dbl��װϵ�� As Double
    Dim Dbl���� As Double
    
    On Error GoTo ErrHand
    
    mblnRefresh = True
            
    Me.vsfBatch.Redraw = flexRDNone
    Me.vsfBatch.rows = 1

    gstrSQL = "Select 1 From ��������˵�� Where ����id=[1] And �������� IN ('��ҩ��','��ҩ��','��ҩ��')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ҩ���ж�]", lng�ⷿID)
    
    If rsTemp.EOF Then
        intҩ�� = 0
    Else
        intҩ�� = 1
    End If
    
    If lngҩƷid = 0 Then Exit Sub
    
    dbl��װϵ�� = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("��װ"))
    
    gstrSQL = " Select Decode(nvl(ҩ�����,0),1,Decode(Nvl(ҩ������,0),1,2,1),0) As ���� " & _
              " From ҩƷ��� Where ҩƷid=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lngҩƷid)
        
    '���ҩ�������ҩ��������int����=2������ҩ�������int����=1������������int����=0��
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        int���� = rsTemp!����
        If lng�ⷿID = 0 Or (intҩ�� = 1 And int���� = 2) Or (intҩ�� = 0 And int���� <> 0) Then
            '�����пⷿ ���� ��ҩ���ҿ�����������ʾ�ֿⷿ�������
            If lng�ⷿID = 0 Then
                gstrSQL = "Select �ⷿ, ����, ƽ���ɱ���, ʧЧ��, ����, ����, ԭ����,�ϴβɹ���, Sum(��������) As ��������, Sum(ʵ������) As ʵ������, Sum(ʵ�ʽ��) As ʵ�ʽ��, Sum(ʵ�ʲ��) As ʵ�ʲ��," & _
                          "     �������� , NO, ��ҩ��λ, ��Ӧ��, �ⷿ��λ, ����, �ۼ�, �Ƿ���, �ּ� " & _
                        " From ( " & _
                        " SELECT (D.���� || '-' || D.����) AS �ⷿ,s.�ϴ����� AS ����,avg(s.ƽ���ɱ���) ƽ���ɱ���,s.Ч�� AS ʧЧ��,DECODE(Nvl(SIGN(ADD_MONTHS(SYSDATE," & intMonths & ")-S.Ч��),-1),-1,0,1) ����,NULL AS ����,Null As ԭ����,NULL AS �ϴβɹ���," & _
                        "        SUM(S.��������)/" & dbl��װϵ�� & " AS ��������,SUM(S.ʵ������)/" & dbl��װϵ�� & " AS ʵ������,SUM(S.ʵ�ʽ��) AS ʵ�ʽ��,SUM(S.ʵ�ʲ��) AS ʵ�ʲ��," & _
                        "        NULL AS ��������,NULL AS NO,NULL AS ��ҩ��λ,NULL As ��Ӧ��, C.�ⷿ��λ,����/" & dbl��װϵ�� & " as ����, decode(nvl(s.���ۼ�,0),0,decode(sum(s.ʵ������),0,0,Sum(s.ʵ�ʽ��) / Sum(s.ʵ������)),s.���ۼ�)*" & dbl��װϵ�� & " As �ۼ�,nvl(a.�Ƿ���,0) as �Ƿ���,b.�ּ�*" & dbl��װϵ�� & "as �ּ� " & _
                        " FROM ҩƷ��� S,���ű� D, (Select Distinct �շ�ϸĿid, ִ�п���id From �շ�ִ�п���) K, ҩƷ�����޶� C,�շ���ĿĿ¼ A,�շѼ�Ŀ B " & _
                        " WHERE S.�ⷿID=D.ID AND S.����=1 AND S.ҩƷID=[1] And S.�ⷿid = C.�ⷿid(+) And S.ҩƷid = C.ҩƷid(+) " & _
                        "       And K.ִ�п���id(+) = S.�ⷿID And K.�շ�ϸĿid(+) = S.ҩƷID AND s.ҩƷid=a.id and a.Id = b.�շ�ϸĿid And Sysdate Between ִ������ And ��ֹ���� " & _
                        GetPriceClassString("B") & " AND (Nvl(S.ʵ������,0)<>0 OR Nvl(S.ʵ�ʽ��,0)<>0 OR Nvl(S.ʵ�ʲ��,0)<>0) " & _
                        " GROUP BY D.���� || '-' || D.����,s.�ϴ�����, C.�ⷿ��λ,����,�Ƿ���,�ּ�,����,���ۼ�,s.Ч��)" & _
                        " Group By ����,�ⷿ, ƽ���ɱ���, ʧЧ��, ����, ����, �ϴβɹ���, ��������, NO, ��ҩ��λ, ��Ӧ��, �ⷿ��λ, ����, �ۼ�, �Ƿ���, �ּ�" & _
                        " order by �ⷿ"
            Else
               gstrSQL = "SELECT (D.���� || '-' || D.����) AS �ⷿ,S.�ϴ����� As ����,s.ƽ���ɱ���, S.Ч�� ʧЧ��, S.�ϴβ��� As ����,Decode(s.ԭ����, Null, a.ԭ����, s.ԭ����) As ԭ����,DECODE(Nvl(SIGN(ADD_MONTHS(SYSDATE," & intMonths & ")-S.Ч��),-1),-1,0,1) ����," & _
                        "        S.��������/" & dbl��װϵ�� & " AS ��������,S.ʵ������/" & dbl��װϵ�� & " AS ʵ������,S.ʵ�ʽ��,S.ʵ�ʲ��," & _
                        "        S.�ϴβɹ���*" & dbl��װϵ�� & " AS �ϴβɹ���, G.���� As ��Ӧ��,'' �ⷿ��λ, decode(nvl(s.���ۼ�,0),0,decode(s.ʵ������,0,0,s.ʵ�ʽ�� / s.ʵ������),s.���ۼ�)*" & dbl��װϵ�� & " As �ۼ�,nvl(a.�Ƿ���,0) as �Ƿ���, b.�ּ�*" & dbl��װϵ�� & "as �ּ�" & _
                        " FROM ҩƷ��� S,���ű� D,ҩƷ��� A, ��Ӧ�� G, (Select Distinct �շ�ϸĿid, ִ�п���id From �շ�ִ�п���) K,�շ���ĿĿ¼ A, �շѼ�Ŀ B " & _
                        " WHERE S.�ⷿID=D.ID AND A.ҩƷID=S.ҩƷID" & _
                        "       AND S.ҩƷID=[1] AND S.����=1 AND S.�ⷿID=[2] And Nvl(S.�ϴι�Ӧ��id, 0) = G.ID(+) " & _
                        "       And K.ִ�п���id(+) = S.�ⷿID And K.�շ�ϸĿid(+) = S.ҩƷID and s.ҩƷid=a.id and a.Id = b.�շ�ϸĿid And Sysdate Between ִ������ And ��ֹ���� " & _
                        GetPriceClassString("B") & " AND (Nvl(S.ʵ������,0)<>0 OR Nvl(S.ʵ�ʽ��,0)<>0 OR Nvl(S.ʵ�ʲ��,0)<>0) " & _
                        " ORDER BY D.���� || '-' || D.����,S.�ϴ�����"
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩƷid, lng�ⷿID)
            
            Call SetFormat(IniListType.BatchList)
            
            Me.vsfBatch.rows = 2
            With rsTemp
                Do While Not .EOF
                    If lng�ⷿID = 0 Then
                        Dbl���� = 0
                        
                        If !�ⷿ <> vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("�ⷿ")) And vsfBatch.rows - 2 <> 0 Then
                            For intRow = 1 To vsfBatch.rows - 1
                                If vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("�ⷿ")) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("�ⷿ")) Then
                                    Dbl���� = Dbl���� + Val(vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("ʵ������")))
                                End If
                            Next
                            
                            vsfBatch.MergeCells = flexMergeRestrictRows
                            vsfBatch.MergeRow(vsfBatch.rows - 1) = True
                            
                            For intCol = 0 To vsfBatch.Cols - 1
                                vsfBatch.TextMatrix(vsfBatch.rows - 1, intCol) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("�ⷿ")) & "ʵ������Ϊ��" & Dbl���� & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("��λ"))
                            Next
                            
                            Dbl���� = 0
                            vsfBatch.rows = vsfBatch.rows + 1
                        End If
                    End If
                    
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ⷿ")) = !�ⷿ
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("����")) = IIf(IsNull(!����), "", !����)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("Ч��")) = Format(!ʧЧ��, "yyyy��MM��dd��")
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("Ч��")) <> "" And cboStock.Text <> "���пⷿ" Then
                        '����Ϊ��Ч��
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("Ч��")) = Format(DateAdd("D", -1, Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("Ч��"))), "yyyy-mm-dd")
                    End If
                    
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("����")) = IIf(IsNull(!����), "", !����)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("ԭ����")) = IIf(IsNull(!ԭ����), "", !ԭ����)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("��������")) = Format(!��������, mStr����)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("ʵ������")) = Format(!ʵ������, mStr����)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("ʵ�ʽ��")) = Format(!ʵ�ʽ��, mStr���)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("ʵ�ʲ��")) = Format(!ʵ�ʲ��, mStr���)
                    If !�Ƿ��� = 0 Then
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ۼ�")) = Format(!�ּ�, mStr����)
                    Else
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ۼ�")) = Format(!�ۼ�, mStr����)
                    End If
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ϴβɹ���")) = Format(!�ϴβɹ���, mStr�ɱ���)
                    If !ʵ������ <> 0 Then
'                        Me.vsfBatch.TextMatrix(vsfbatch.rows-1, vsfBatch.ColIndex("�ɱ���")) = Format((!ʵ�ʽ�� - !ʵ�ʲ��) / !ʵ������, mStr����)
'                        Me.vsfBatch.TextMatrix(vsfbatch.rows-1, vsfBatch.ColIndex("�ɱ����")) = Format(!ʵ�ʽ�� - !ʵ�ʲ��, mStr���)
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ɱ���")) = Format(!ƽ���ɱ��� * dbl��װϵ��, mStr�ɱ���)
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ɱ����")) = Format(!ƽ���ɱ��� * dbl��װϵ�� * !ʵ������, mStr���)
                    End If
          
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("��Ӧ��")) = IIf(IsNull(!��Ӧ��), "", !��Ӧ��)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("�ⷿ��λ")) = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
                    
                    With vsfBatch
                        If Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("ʵ������")) <> "" And lng�ⷿID = 0 Then
                            lng���� = IIf(IsNull(rsTemp!����), 0, rsTemp!����)
                            If lng���� = 0 Then
                                strTemp = "������"
                            ElseIf Val(.TextMatrix(vsfBatch.rows - 1, .ColIndex("ʵ������"))) > lng���� Then
                                strTemp = "����"
                            ElseIf Val(.TextMatrix(vsfBatch.rows - 1, .ColIndex("ʵ������"))) = lng���� Then
                                strTemp = "��ƽ"
                            ElseIf Val(.TextMatrix(vsfBatch.rows - 1, .ColIndex("ʵ������"))) < lng���� Then
                                strTemp = "����"
                            End If
                            .TextMatrix(vsfBatch.rows - 1, .ColIndex("�������")) = strTemp
                        End If
                    End With
                    '���ݼ�¼״̬�Ĳ�ͬ��������ɫ
                    lngColor = IIf(!���� = 0, glng����, glng����)
                    Me.vsfBatch.Cell(flexcpForeColor, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, vsfBatch.Cols - 1) = lngColor
                    'ʧЧ��ҩƷ����ǰ����ʱ��ͼ��
                    If !���� = 1 And IsNull(!ʧЧ��) = False Then
                        Me.vsfBatch.Cell(flexcpPicture, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, 0) = imglvw.ListImages(3).Picture
                    End If
                    
                    Me.vsfBatch.RowData(vsfBatch.rows - 1) = 0 ' CStr(!����)
                    
                    'ʵ�������������Ϊ0������������Ϊ0�ı�ʾ��Ԥ�������������ݣ���ɫ������ʾ
                    If zlCommFun.NVL(!ʵ������, 0) = 0 And zlCommFun.NVL(!ʵ�ʽ��, 0) = 0 And zlCommFun.NVL(!ʵ�ʲ��, 0) = 0 Then
                        Me.vsfBatch.Cell(flexcpForeColor, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, vsfBatch.Cols - 1) = vbRed
                    End If
                    
                    '���������ú�ɫ������ʾ
                    '1.��������<0�ı�ʾ�ǿ�������Ԥ����ռ��(��������ⵥ�����̿����)
                    '2.ʵ������<=0�����ֿ�����û�н��п���飬������������������
                    If zlCommFun.NVL(!��������, 0) < 0 Or zlCommFun.NVL(!ʵ������, 0) <= 0 Then
                        Me.vsfBatch.Cell(flexcpForeColor, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, vsfBatch.Cols - 1) = vbRed
                    End If
                    
                    .MoveNext
                    vsfBatch.rows = vsfBatch.rows + 1
                Loop
                
                If lng�ⷿID = 0 Then
                    Dbl���� = 0
                    vsfBatch.MergeCells = flexMergeRestrictRows
                    vsfBatch.MergeRow(vsfBatch.rows - 1) = True
                    
                    For intRow = 1 To vsfBatch.rows - 1
                        If vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("�ⷿ")) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("�ⷿ")) Then
                            Dbl���� = Dbl���� + Val(vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("ʵ������")))
                        End If
                    Next
                    For intCol = 0 To vsfBatch.Cols - 1
                        vsfBatch.TextMatrix(vsfBatch.rows - 1, intCol) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("�ⷿ")) & "ʵ������Ϊ��" & Dbl���� & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("��λ"))
                    Next
                Else
                    vsfBatch.RemoveItem vsfBatch.rows - 1
                End If
            End With
        End If
    End If
    
    If Me.vsfBatch.rows = 1 Then
        Me.vsfBatch.Visible = False
        Me.lbl����_S.Visible = False
        Me.vsfBatch.rows = 2
    Else
        Me.vsfBatch.Visible = True
        Me.lbl����_S.Visible = True
    End If
'    Me.vsfBatch.FixedRows = 1
'    Me.vsfBatch.Row = 1
    Me.vsfBatch.Redraw = flexRDDirect
    Call Form_Resize
    
    mblnRefresh = False
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    mblnRefresh = False
    Exit Sub
End Sub
Private Sub mnuBill_Click()
    Dim strNo As String
    Dim byt���� As Integer
    Dim byt��¼״̬ As Integer
    
    Select Case Mid(strNoS, 4)
        Case "_INSIDE_1309_1"  '����
            strNo = Mid(Trim(CurSheet.TextMatrix(CurSheet.Row, 3)), 3)
            byt���� = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
            byt��¼״̬ = Val(CurSheet.TextMatrix(CurSheet.Row, 11))
        Case "_INSIDE_1309_2"  '��ϸ��
            strNo = Trim(CurSheet.TextMatrix(CurSheet.Row, 3))
            byt���� = Val(CurSheet.TextMatrix(CurSheet.Row, 2))
            byt��¼״̬ = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
        Case "_INSIDE_1309_3"  '��ϸ��
        
    End Select
    
    If strNo = "" Or byt���� = 0 Or byt��¼״̬ = 99 Then Exit Sub
    If byt���� = 0 Then Exit Sub
    ShowBill Me, strNo, byt��¼״̬, byt����
End Sub

Private Sub ObjReport_ReportActive(ByVal strNo As String, Form As Object)
    lngCurReport = Form.hWnd
    strNoS = strNo
End Sub

Private Sub ObjReport_SheetDblClick(ByVal strNo As String, Sheet As Object, frmParent As Object)
    lngCurReport = frmParent.hWnd
    strNoS = strNo
    Set CurSheet = Sheet
    If Mid(UCase(strNo), 4) = "_INSIDE_1309_3" Then Exit Sub
    mnuBill_Click
End Sub

Private Sub ObjReport_SheetMouseDown(ByVal strNo As String, Button As Integer, Shift As Integer, x As Single, y As Single, Sheet As Object, frmParent As Object)
    lngCurReport = frmParent.hWnd
    strNoS = strNo
    Set CurSheet = Sheet
    If Mid(UCase(strNo), 4) <> "_INSIDE_1309_3" Then
        If Button = 2 Then PopupMenu mnuReportBill, 2
    End If
End Sub

Private Sub SetMenu(ByVal intState As Integer)
    If intState = 0 Then mnuReportBill.Visible = False: Exit Sub
End Sub
Private Sub ShowBill(frmObject As Object, strNo As String, int��¼״̬ As Integer, int���� As Integer, Optional bln���� As Boolean = False)
    '--------------------------------------------------------------------------------------
    '����:��ʾָ������
    '����:
    '       frmObject:����
    '           strNo:���ݺ�
    '     int��¼״̬:����״̬(mod(��¼״̬,3)=1-������¼;mod(��¼״̬,3)=2-������¼;mod(��¼״̬,3)=0-�Ѿ������ļ�¼)
    '         int����:�������( �ⷿ:1-�⹺��ⵥ;2-�������;3-�ƿⵥ;4-����;5-��������;6-�̴�;7-������;
    '                           ����:1-����;2-����;3-���ϵ�;4-Ȩ�����)
    '--------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Select Case int����
        Case 1
            frmPurchaseCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 2
            frmSelfMakeCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 3
            frmAccordDrugCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 4
            frmOtherInputCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 5
            frmDiffPriceAdjustCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 6
            frmTransferCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 7
            frmDrawCard.ShowCard frmObject, strNo, 4, False, int��¼״̬
        Case 11
            frmOtherOutputCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 12
            frmCheckCard.ShowCard frmObject, strNo, 4, int��¼״̬
        Case 13
            Dim rsTemp As New ADODB.Recordset
            gstrSQL = "Select id,����,NO,nvl(�۸�id,0) as �۸�id " & _
                      "From ҩƷ�շ���¼ " & _
                      "Where No=[1] and ����ID is null And ����=[2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�۸��¼ID]", strNo, int����)
                  
            If rsTemp.EOF Or rsTemp.BOF Then Exit Sub
              
            gstrUserName = UserInfo.�û�����
            With frmAdjust
                .lngBillId = rsTemp!�۸�id
                .lngMediId = 1
                .intUnit = intChoose����
                .Show 1, frmObject
            End With
        Case Else
            
            With Frm����See
                .int��¼״̬ = int��¼״̬
                .byt���� = int����
                .strNo = strNo
                .mstrPrivs = mstrPrivs
                .Show 1, frmObject
            End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ����Ȩ��()
    If Not zlStr.IsHavePrivs(mstrPrivs, "ҩƷ��ϸ��") Then
        tbrThis.Buttons("��ϸ").Visible = False
        mnuFileBatch.Visible = False
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "ҩƷ����") Then
        tbrThis.Buttons("����").Visible = False
    End If
End Sub
Private Sub SetFormat(ByVal intType As Integer)
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln��ҩ���� As Boolean
    Dim int���� As Integer
    
    On Error GoTo errHandle
    
    If Val(cboStock.ItemData(cboStock.ListIndex)) < 0 Then Exit Sub
    
    gstrSQL = "select a.���� from ���Ʒ���Ŀ¼ a where a.id=[1]"
    If Left(Me.tvwSection_S.SelectedItem.Key, 1) <> "R" Then
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�ѡ��ҩƷ�����", Mid(Me.tvwSection_S.SelectedItem.Key, 2))
        int���� = rsDetail!����
    End If
    
    If int���� = 3 Or (Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R" And Right(Me.tvwSection_S.SelectedItem.Key, 1) = "7") Then bln��ҩ���� = True
    If bln��ҩ���� Then
        vsfList.ColWidth(vsfList.ColIndex("ԭ����")) = 1500
        vsfBatch.ColWidth(vsfBatch.ColIndex("ԭ����")) = 1500
    Else
        vsfList.ColWidth(vsfList.ColIndex("ԭ����")) = 0
        vsfBatch.ColWidth(vsfBatch.ColIndex("ԭ����")) = 0
    End If
    
    If intType = IniListType.AllList Or intType = IniListType.MainList Then
        With vsfList
            .rows = 1
            .rows = 2

            If gintҩƷ������ʾ = 2 Then
                '��ʾ��Ʒ����
                .ColWidth(.ColIndex("��Ʒ��")) = IIf(.ColWidth(.ColIndex("��Ʒ��")) = 0, 2000, .ColWidth(.ColIndex("��Ʒ��")))
            Else
                '��������ʾ��Ʒ����
                .ColWidth(.ColIndex("��Ʒ��")) = 0
            End If
            
            .ColWidth(.ColIndex("�ϴβɹ���")) = IIf(mblnViewCost, IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("�ϴβɹ���")) = 0, 1000, .ColWidth(.ColIndex("�ϴβɹ���")))), 0)
            .ColWidth(.ColIndex("ƽ���ɱ���")) = IIf(mblnViewCost, IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("ƽ���ɱ���")) = 0, 1000, .ColWidth(.ColIndex("ƽ���ɱ���")))), 0)
            .ColWidth(.ColIndex("�����")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("�����")) = 0, 1000, .ColWidth(.ColIndex("�����"))), 0)
            .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("�ɱ����")) = 0, 1000, .ColWidth(.ColIndex("�ɱ����"))), 0)
            .ColWidth(.ColIndex("�ϴι�Ӧ��")) = IIf(zlStr.IsHavePrivs(mstrPrivs, "��Ӧ�̲�ѯ"), IIf(.ColWidth(.ColIndex("�ϴι�Ӧ��")) = 0, 2500, .ColWidth(.ColIndex("�ϴι�Ӧ��"))), 0)
            .ColWidth(.ColIndex("�ⷿ��λ")) = IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("�ⷿ��λ")) = 0, 1500, .ColWidth(.ColIndex("�ⷿ��λ"))))
            .ColWidth(.ColIndex("�������")) = IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("�������")) = 0, 1500, .ColWidth(.ColIndex("�������"))))
            .Row = 1
            
            mstrUnShow_List = "ҩƷID;��;����ID;Ч��;����ϵ��;����ʱ��;��װ"
            If .ColWidth(.ColIndex("��Ʒ��")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";��Ʒ��"
            If .ColWidth(.ColIndex("�ϴβɹ���")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";�ϴβɹ���"
            If .ColWidth(.ColIndex("ƽ���ɱ���")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";ƽ���ɱ���"
            If .ColWidth(.ColIndex("�����")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";�����"
            If .ColWidth(.ColIndex("�ɱ����")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";�ɱ����"
            If .ColWidth(.ColIndex("�ϴι�Ӧ��")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";�ϴι�Ӧ��"
            If .ColWidth(.ColIndex("�ⷿ��λ")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";�ⷿ��λ"
            If .ColWidth(.ColIndex("�������")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";�������"
            If .ColWidth(.ColIndex("ԭ����")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";ԭ����"
        End With
    End If
    
    If intType = IniListType.AllList Or intType = IniListType.BatchList Then
        With vsfBatch
            .rows = 1
            .rows = 2
            
            .TextMatrix(0, .ColIndex("Ч��")) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
            
            If Val(cboStock.ItemData(cboStock.ListIndex)) = 0 Then
                .ColWidth(.ColIndex("�ⷿ")) = IIf(.ColWidth(.ColIndex("�ⷿ")) = 0, 1500, .ColWidth(.ColIndex("�ⷿ")))
                .ColWidth(.ColIndex("����")) = IIf(.ColWidth(.ColIndex("����")) = 0, 1500, .ColWidth(.ColIndex("����")))
                .ColWidth(.ColIndex("Ч��")) = IIf(.ColWidth(.ColIndex("Ч��")) = 0, 1500, .ColWidth(.ColIndex("Ч��")))
                .ColWidth(.ColIndex("����")) = 0
                .ColWidth(.ColIndex("�ϴβɹ���")) = 0
                .ColWidth(.ColIndex("��Ӧ��")) = 0
                .ColWidth(.ColIndex("�ⷿ��λ")) = IIf(.ColWidth(.ColIndex("�ⷿ��λ")) = 0, 1500, .ColWidth(.ColIndex("�ⷿ��λ")))
                .ColWidth(.ColIndex("�������")) = IIf(.ColWidth(.ColIndex("�������")) = 0, 1500, .ColWidth(.ColIndex("�������")))
                .ColWidth(.ColIndex("ԭ����")) = 0
            Else
                .ColWidth(.ColIndex("�ⷿ")) = 0
                .ColWidth(.ColIndex("����")) = IIf(.ColWidth(.ColIndex("����")) = 0, 1500, .ColWidth(.ColIndex("����")))
                .ColWidth(.ColIndex("Ч��")) = IIf(.ColWidth(.ColIndex("Ч��")) = 0, 1500, .ColWidth(.ColIndex("Ч��")))
                .ColWidth(.ColIndex("����")) = IIf(.ColWidth(.ColIndex("����")) = 0, 1500, .ColWidth(.ColIndex("����")))
                .ColWidth(.ColIndex("�ϴβɹ���")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("�ϴβɹ���")) = 0, 1500, .ColWidth(.ColIndex("�ϴβɹ���"))), 0)
                .ColWidth(.ColIndex("��Ӧ��")) = IIf(zlStr.IsHavePrivs(mstrPrivs, "��Ӧ�̲�ѯ"), IIf(.ColWidth(.ColIndex("��Ӧ��")) = 0, 2500, .ColWidth(.ColIndex("��Ӧ��"))), 0)
                .ColWidth(.ColIndex("�ⷿ��λ")) = 0
                .ColWidth(.ColIndex("�������")) = 0
                If bln��ҩ���� Then
                    .ColWidth(.ColIndex("ԭ����")) = IIf(.ColWidth(.ColIndex("ԭ����")) = 0, 1500, .ColWidth(.ColIndex("ԭ����")))
                End If
            End If

            .ColWidth(.ColIndex("ʵ�ʲ��")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("ʵ�ʲ��")) = 0, 1500, .ColWidth(.ColIndex("ʵ�ʲ��"))), 0)
            .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("�ɱ���")) = 0, 1500, .ColWidth(.ColIndex("�ɱ���"))), 0)
            .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("�ɱ����")) = 0, 1500, .ColWidth(.ColIndex("�ɱ����"))), 0)
            
            mstrUnShow_Batch = "������������"
            If .ColWidth(.ColIndex("�ⷿ")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";�ⷿ"
            If .ColWidth(.ColIndex("����")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";����"
            If .ColWidth(.ColIndex("Ч��")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";Ч��"
            If .ColWidth(.ColIndex("����")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";����"
            If .ColWidth(.ColIndex("�ϴβɹ���")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";�ϴβɹ���"
            If .ColWidth(.ColIndex("��Ӧ��")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";��Ӧ��"
            If .ColWidth(.ColIndex("�ⷿ��λ")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";�ⷿ��λ"
            If .ColWidth(.ColIndex("ʵ�ʲ��")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";ʵ�ʲ��"
            If .ColWidth(.ColIndex("�ɱ���")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";�ɱ���"
            If .ColWidth(.ColIndex("�ɱ����")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";�ɱ����"
            If .ColWidth(.ColIndex("�������")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";�������"
            If .ColWidth(.ColIndex("ԭ����")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";ԭ����"
        End With
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataBound(ByVal rsData As ADODB.Recordset)
    Dim lngColor As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngPrice As Single
    Dim lng���� As Long
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    If rsData.EOF Then Exit Sub
    
    On Error GoTo ErrHand
    
    mblnRefresh = True
    
    With vsfList
        .Redraw = flexRDNone
        Do While Not rsData.EOF
            If rsData.AbsolutePosition > .rows - 1 Then .rows = .rows + 1
            .Row = rsData.AbsolutePosition
            '�������
            .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsData!ҩƷid
            .TextMatrix(.Row, .ColIndex("��;����ID")) = rsData!��;����id
            .TextMatrix(.Row, .ColIndex("����")) = rsData!����
            .TextMatrix(.Row, .ColIndex("��ʶ��")) = zlCommFun.NVL(rsData!ҩ����, "")
            
            If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                .TextMatrix(.Row, .ColIndex("����")) = rsData!ͨ����
            Else
                .TextMatrix(.Row, .ColIndex("����")) = IIf(IsNull(rsData!��Ʒ��), rsData!ͨ����, rsData!��Ʒ��)
            End If
            
            .TextMatrix(.Row, .ColIndex("��Ʒ��")) = IIf(IsNull(rsData!��Ʒ��), "", rsData!��Ʒ��)
            .TextMatrix(.Row, .ColIndex("����ҩ��")) = IIf(IsNull(rsData!����ҩ��), "", rsData!����ҩ��)
            .TextMatrix(.Row, .ColIndex("���")) = zlCommFun.NVL(rsData!���, "")
            .TextMatrix(.Row, .ColIndex("����")) = zlCommFun.NVL(rsData!����, "")
            .TextMatrix(.Row, .ColIndex("ԭ����")) = zlCommFun.NVL(rsData!ԭ����, "")
            .TextMatrix(.Row, .ColIndex("Ч��")) = zlCommFun.NVL(rsData!Ч��, "")
            .TextMatrix(.Row, .ColIndex("ҩ�����")) = zlCommFun.NVL(rsData!ҩ�����, "��")
            .TextMatrix(.Row, .ColIndex("��λ")) = zlCommFun.NVL(rsData!��λ, "")
            .TextMatrix(.Row, .ColIndex("ƽ���ɱ���")) = Format(rsData!ƽ���ɱ���, mStr�ɱ���)
            .TextMatrix(.Row, .ColIndex("�ϴβɹ���")) = Format(rsData!�ϴβɹ���, mStr�ɱ���)
            
            lngPrice = rsData!��ǰ�ۼ�
            
            .TextMatrix(.Row, .ColIndex("��ǰ�ۼ�")) = Format(lngPrice, mStr����)
            .TextMatrix(.Row, .ColIndex("����ϵ��")) = Format(rsData!ϵ��1, mStr����)
            .TextMatrix(.Row, .ColIndex("��������")) = Format(rsData!��������, mStr����)
            .TextMatrix(.Row, .ColIndex("�������")) = Format(rsData!ʵ������, mStr����)
            .TextMatrix(.Row, .ColIndex("�����")) = Format(rsData!ʵ�ʽ��, mStr���)
            .TextMatrix(.Row, .ColIndex("�����")) = Format(rsData!ʵ�ʲ��, mStr���)
            .TextMatrix(.Row, .ColIndex("�ɱ����")) = Format(rsData!ƽ���ɱ��� * rsData!ʵ������, mStr���)
            .TextMatrix(.Row, .ColIndex("����ʱ��")) = zlCommFun.NVL(rsData!����ʱ��, "")
            .TextMatrix(.Row, .ColIndex("��װ")) = zlCommFun.NVL(rsData!����, 1)
            .TextMatrix(.Row, .ColIndex("�ϴι�Ӧ��")) = zlCommFun.NVL(rsData!�ϴι�Ӧ��, "")
            .TextMatrix(.Row, .ColIndex("�ⷿ��λ")) = zlCommFun.NVL(rsData!�ⷿ��λ, "")
            If .TextMatrix(.Row, .ColIndex("�������")) <> "" And cboStock.Tag <> 0 Then
                lng���� = IIf(IsNull(rsData!����), 0, rsData!����)
                If lng���� = 0 Then
                    strTemp = "������"
                ElseIf Val(.TextMatrix(.Row, .ColIndex("�������"))) > lng���� Then
                    strTemp = "����"
                ElseIf Val(.TextMatrix(.Row, .ColIndex("�������"))) = lng���� Then
                    strTemp = "��ƽ"
                ElseIf Val(.TextMatrix(.Row, .ColIndex("�������"))) < lng���� Then
                    strTemp = "����"
                End If
                .TextMatrix(.Row, .ColIndex("�������")) = strTemp
            End If
            '��ɫ
            If bln����ͣ��ҩƷ Then
                lngColor = IIf(Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) = "", glng��ɫ, glng��ɫ)
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = lngColor
            End If
            If cboStock.ItemData(Me.cboStock.ListIndex) > 0 Then
                '����Ƿ�������Ч�ڹ��ˣ�ֻ��ѡ�����ĳ���ⷿʱ�Ŵ���
                If rsData!���� = 1 And IsNull(rsData!���Ч��) = False Then
                    .Cell(flexcpPicture, .Row, 0, .Row, 0) = imglvw.ListImages(3).Picture
                End If
            End If
            rsData.MoveNext
        Loop
        
        '��д�������
        Call SetSortCode
                
        .Redraw = flexRDDirect
    End With
    
    mblnRefresh = False
    Exit Sub
ErrHand:
    If ErrCenter() Then Resume
    Call SaveErrLog
    mblnRefresh = False
End Sub

'Modified By ���� 2003-12-10 ����������
Private Sub txtҩƷ��Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txtҩƷ��Ϣ)
End Sub

Private Sub txtҩƷ��Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String, StrBit As String
    Dim strInput As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    txtҩƷ��Ϣ.Text = Replace(txtҩƷ��Ϣ.Text, "'", "")
    strInput = Trim(UCase(txtҩƷ��Ϣ.Text))
    
    If strInput = "" Then Exit Sub
    
    StrBit = GetSetting(appName:="ZLSOFT", Section:="����ģ��\����", Key:="����ƥ��", Default:="0")
    StrBit = IIf(StrBit = "0", "%", "")
    
    strFind = " And (A.���� Like [7] OR B.���� Like [7] OR B.���� LIKE [7])"
    
    If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
        If Mid(gtype_UserSysParms.P44_����ƥ��, 1, 1) = "1" Then strFind = " And (A.���� Like [7] Or B.���� Like [7] And B.����=3)"
    ElseIf zlStr.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
        If Mid(gtype_UserSysParms.P44_����ƥ��, 2, 1) = "1" Then strFind = " And B.���� Like [7] "
    ElseIf zlStr.IsCharChinese(strInput) Then
        strFind = " And B.���� Like [7] "
    End If
    
    SQLCondition.strҩƷ��Ϣ = StrBit & strInput & "%"
        
    If Not ReFreshDrugData(cboStock.ItemData(cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2)), strFind, False) Then Exit Sub
    Me.tvwSection_S.Tag = "T"
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 2 Then '��ѡ����
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
        
        If vsfList.MouseRow <> 0 Then Exit Sub
        
        InitColSelList IniListType.MainList, vsfList
        
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.colHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = vsfList.Top + vsfList.RowHeight(0)
                If .Top + .Height > Me.ScaleHeight - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = vsfList.Left + x
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub



