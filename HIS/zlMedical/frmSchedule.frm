VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSchedule 
   Caption         =   "���ԤԼ����"
   ClientHeight    =   7140
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11310
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSchedule.frx":1CFA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14870
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11310
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "��첿��"
      Child2          =   "cboDept"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   465
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
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
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ԤԼ"
               Key             =   "ԤԼ"
               Object.ToolTipText     =   "ԤԼ"
               Object.Tag             =   "ԤԼ"
               ImageIndex      =   3
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȷ��"
               Key             =   "ȷ��"
               Object.ToolTipText     =   "ȷ��"
               Object.Tag             =   "ȷ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "ȡ��"
               Object.ToolTipText     =   "ȡ��"
               Object.Tag             =   "ȡ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8760
      Top             =   900
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
            Picture         =   "frmSchedule.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":29CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":2BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":2E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":301C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3236
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3450
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":366A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":388A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3AAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8040
      Top             =   900
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
            Picture         =   "frmSchedule.frx":3CC4
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":3EE4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4104
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4456
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4670
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":49C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":4DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5010
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5230
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5450
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1770
      Left            =   150
      TabIndex        =   3
      Top             =   900
      Width           =   2775
      _cx             =   4895
      _cy             =   3122
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   0
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   4560
      Top             =   2055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5962
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":5CFC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":6296
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":B300
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":B89A
            Key             =   "ȡ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":BE34
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":C3CE
            Key             =   "�¿�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":C968
            Key             =   "���"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":CF02
            Key             =   "up"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D0C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   5475
      TabIndex        =   5
      Top             =   1980
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ilsGrid 
      Left            =   6315
      Top             =   3165
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
            Picture         =   "frmSchedule.frx":D286
            Key             =   "T����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D620
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":D9BA
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":DD54
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E0EE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E488
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchedule.frx":E64A
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPerson 
      Height          =   1950
      Left            =   390
      TabIndex        =   6
      Top             =   3495
      Width           =   2985
      _cx             =   5265
      _cy             =   3440
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   900
         X2              =   900
         Y1              =   810
         Y2              =   2025
      End
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   75
         X2              =   1860
         Y1              =   945
         Y2              =   945
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   1950
      Left            =   3615
      TabIndex        =   7
      Top             =   3495
      Width           =   2505
      _cx             =   4419
      _cy             =   3440
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   75
         X2              =   1860
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   900
         X2              =   900
         Y1              =   810
         Y2              =   2025
      End
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   3450
      MousePointer    =   9  'Size W E
      Top             =   2835
      Width           =   45
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   255
      MousePointer    =   7  'Size N S
      Top             =   3135
      Width           =   3690
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "���ɵǼǱ��(&N)"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "ԤԼ(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "���Ӹ���ԤԼ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAddGroup 
         Caption         =   "��������ԤԼ(&N)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�ԤԼ(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��ԤԼ(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "ȷ��ԤԼ(&O)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "ȡ��ԤԼ(&C)"
      End
   End
   Begin VB.Menu mnuAddition 
      Caption         =   "����(&A)"
      Begin VB.Menu mnuAdditionItems 
         Caption         =   "�����Ŀ(&I)"
      End
      Begin VB.Menu mnuAdditionPerItems 
         Caption         =   "��Ա��Ŀ(&H)"
      End
      Begin VB.Menu mnuAddition_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdditionPersons 
         Caption         =   "�ܼ���Ա(&P)"
      End
      Begin VB.Menu mnuAddition_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdditionGroupMember 
         Caption         =   "��Ա����(&C)"
      End
      Begin VB.Menu mnuAdditionPhoto 
         Caption         =   "��Ƭ�ɼ�(&S)"
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
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mstrPrivs As String

Private mblnAllowChange As Boolean
Private mstrCondition As String
Private mlngSvrDept As Long                             '�����ϴε������첿��
Private mstrSvrGoup As String                           '�����ϴε����������
Private mlngSvrKey As Long                              '�����ϴε�������ԤԼ

Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

Private Enum mCol
    i���� = 1
    i��鲿λ
    i�ɼ���ʽ
    i����걾
    i�����۸�
    i���۸�
    iִ�п���
    i���㷽ʽ
End Enum

'�������Զ�����̻���************************************************************************************************

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    Call InitSysPara
    
    mlngSvrDept = 0
    mblnAllowChange = True
    mstrSvrGoup = ""
    mlngSvrKey = 0
    imgY_S.Width = 60
    
    mstrCondition = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "'" & Format(DateAdd("d", 7, CDate(zlDatabase.Currentdate)), "yyyy-MM-dd")
    mstrCondition = mstrCondition & "''''''"
    
    strVsf = ",255,4,1,1,[����];,255,4,1,1,[״̬];No,810,1,1,1,;ԤԼ��,750,1,1,1,;ԤԼʱ��,990,1,1,1,;����,2400,1,1,1,;����,450,1,1,1,;Ӧ�ս��,900,7,1,1,;ʵ�ս��,900,7,1,1,"
    strVsf = strVsf & ";�������,1800,1,1,1,;��ϵ�绰,900,1,1,1,;��ϵ��ַ,1800,1,1,1,;����˵��,1500,1,1,1,;���״̬,0,1,1,1,;��Լ��λid,0,1,1,1,"
    
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "0")) = 1 Then
    
        'ʹ�ø��Ի�����
                        
    End If
    
    Call CreateVsf(vsf, strVsf)
'    vsf.Cols = vsf.Cols + 1
'    vsf.ColWidth(vsf.Cols - 1) = 15
    
    Set vsf.Cell(flexcpPicture, 0, 0) = ils13.ListImages("״̬").Picture
    Set vsf.Cell(flexcpPicture, 0, 1) = ils13.ListImages("״̬").Picture
    
    vsf.ColFormat(GetCol(vsf, "Ӧ�ս��")) = gstrDec
    vsf.ColFormat(GetCol(vsf, "ʵ�ս��")) = gstrDec
    
    strVsf = "����,900,1,1,1,;�����,900,7,1,1,;�Ա�,810,1,1,1,;����,600,1,1,1,;����id,0,1,1,1,"
    Call CreateVsf(vsfPerson, strVsf)
    With vsfPerson
'        .Cols = .Cols + 1
'        .ColWidth(vsfPerson.Cols - 1) = 15
        .MergeCells = flexMergeFree
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarComplete
    End With
        
    strVsf = ",255,4,1,1,[����];����,2400,1,1,1,;��鲿λ,1200,1,1,1,;�ɼ���ʽ,900,1,1,1,;����걾,900,1,1,1,;�����۸�,900,7,1,1,;���۸�,900,7,1,1,;ִ�п���,1200,1,1,1,;���㷽ʽ,810,1,1,1,"
    Call CreateVsf(vsfItem, strVsf)
'    vsfItem.Cols = vsfItem.Cols + 1
'    vsfItem.ColWidth(vsfItem.Cols - 1) = 15

    vsfItem.ColFormat(GetCol(vsfItem, "���۸�")) = gstrDec
    vsfItem.ColFormat(GetCol(vsfItem, "�����۸�")) = gstrDec
    Set vsfItem.Cell(flexcpPicture, 0, 0) = ilsGrid.ListImages("T����").Picture
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitActivate() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ�����ݣ������ڴ����Activate�¼�
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = GetPublicSQL(SQL.��첿���嵥, IIf(InStr(gstrPrivs, "���п���") > 0, "����", ""))
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
    
    If rs.BOF Then
        ShowSimpleMsg "û��������ʵĲ��ţ����ڲ��Ź��������ã�"
        Exit Function
    End If
    
    '�����ݵ��ؼ���
    Call AddComboData(cboDept, rs)
    
    '��ʼѡ�����ݴ���
    zlControl.CboLocate cboDept, UserInfo.����ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    InitActivate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� Ӧ��Ȩ�޴���
    '������ strPrivilege                    Ȩ��
    '------------------------------------------------------------------------------------------------------------------
        
    '�������
    'strPrivilege = "���п���;���ԤԼ;ȷ��ԤԼ;ȡ��ԤԼ"
    
    '�����С�ԤԼ��Ȩ��ʱ
    If InStr(strPrivilege, "���ԤԼ") = 0 Then
        mnuEditAdd.Visible = False
        mnuEditAddGroup.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
                
        mnuAddition.Visible = False
    End If
    
    If InStr(strPrivilege, "ȷ��ԤԼ") = 0 Then mnuEditCheck.Visible = False
    
    If InStr(strPrivilege, "ȡ��ԤԼ") = 0 Then
        If InStr(strPrivilege, "ȷ��ԤԼ") = 0 And InStr(strPrivilege, "���ԤԼ") = 0 Then
            mnuEdit.Visible = False
        Else
            mnuEditCancel.Visible = False
        End If
    End If
    
    mnuEdit_1.Visible = mnuEditAdd.Visible And (mnuEditCheck.Visible Or mnuEditCancel.Visible)
    
    tbrThis.Buttons("ԤԼ").Visible = mnuEdit.Visible And mnuEditAdd.Visible
    tbrThis.Buttons("�޸�").Visible = mnuEdit.Visible And mnuEditModify.Visible
    tbrThis.Buttons("ɾ��").Visible = mnuEdit.Visible And mnuEditDelete.Visible
    tbrThis.Buttons("ȷ��").Visible = mnuEdit.Visible And mnuEditCheck.Visible
    tbrThis.Buttons("ȡ��").Visible = mnuEdit.Visible And mnuEditCancel.Visible
    
    tbrThis.Buttons("Split_2").Visible = tbrThis.Buttons("ԤԼ").Visible
    tbrThis.Buttons("Split_3").Visible = tbrThis.Buttons("ȷ��").Visible Or tbrThis.Buttons("ȡ��").Visible
    
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���������ܲ˵��Ŀ���״̬
    '------------------------------------------------------------------------------------------------------------------
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    
    mnuEditCheck.Enabled = True
    mnuEditCancel.Enabled = True

    mnuAdditionGroupMember.Enabled = True
    mnuAdditionItems.Enabled = True
    mnuAdditionPerItems.Enabled = True
    mnuAdditionPersons.Enabled = True
            
    If Val(vsf.RowData(1)) = 0 Then
                
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditCheck.Enabled = False
        mnuEditCancel.Enabled = False
        
        mnuAdditionGroupMember.Enabled = False
        mnuAdditionItems.Enabled = False
        mnuAdditionPerItems.Enabled = False
        mnuAdditionPersons.Enabled = False
            
    Else
        Select Case Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "���״̬")))
        Case 1, 3         '�¿�ԤԼ
        
            mnuEditCancel.Enabled = False
            
        Case 2          'ȷ��ԤԼ
        
            mnuEditCheck.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            
        Case 4, 5
        
            mnuEditCheck.Enabled = False
            mnuEditCancel.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
            
            mnuAdditionGroupMember.Enabled = False
            mnuAdditionItems.Enabled = False
            mnuAdditionPerItems.Enabled = False
            mnuAdditionPersons.Enabled = False
            
        End Select
        
        If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[����]")) <> "" Then
           
           '�Ǹ���ԤԼ
           mnuAdditionGroupMember.Enabled = False
           mnuAdditionItems.Enabled = False
           mnuAdditionPersons.Enabled = False
           
        Else
            
            If Val(vsfPerson.TextMatrix(vsfPerson.Row, GetCol(vsfPerson, "�����"))) = 0 Then
                mnuAdditionPerItems.Enabled = False
            End If
            
        End If
        
    End If
    
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("�޸�").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("ɾ��").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("ȷ��").Enabled = mnuEditCheck.Enabled
    tbrThis.Buttons("ȡ��").Enabled = mnuEditCancel.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��״̬����ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------------------
    
    If Val(vsf.RowData(1)) <= 0 Then
        stbThis.Panels(2).Text = "��" & cboDept.Text & "����û�����ԤԼ��"
    Else
        stbThis.Panels(2).Text = "��" & cboDept.Text & "���¹��� " & vsf.Rows - 1 & "�����ԤԼ��"
    End If
    
End Sub

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";���ԤԼ;") > 0 Then
        Call ResetVsf(vsf)
        
'        Call AppendSapceRows(vsf, lnX, lnY)
    End If
    
    If InStr(strMenuItem, ";�����Ŀ;") > 0 Then
        Call ResetVsf(vsfItem)
        
        Call AppendSapceRows(vsfItem, lnX1, lnY1)
    End If
        
End Function

Public Function EditRefresh(ByVal strMenuItem As String, ByVal strPara As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����༭���ݴ�����ã��ӿں���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey2 As Long
    Dim lngSvrKey3 As Long
    Dim varPara As Variant
        
    On Error GoTo errHand

    '���������������Ŀ
    varPara = Split(strPara, "'")
    
    Select Case strMenuItem
    Case "���ԤԼ"
        Call ClearData("���ԤԼ;�����Ŀ")
        
        Call RefreshData("���ԤԼ")
        
        '�ָ����ԤԼ
        Call RestoreRow(vsf, Val(varPara(0)))
        
    Case "�����Ա"
        
    Case Else
        Call ClearData("�����Ŀ")
    End Select
    
    mblnAllowChange = True
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��/װ������
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    Dim varCondition As Variant
    Dim strCondition As String
    Dim lngLoop As Long
    
    Dim intState As Integer
    Dim intGroup As Integer
    
    On Error GoTo errHand
       
    Select Case strMenuItem
    Case "���ԤԼ"
                                
        varCondition = Split(mstrCondition, "'")
        strCondition = " AND A.���ʱ�� BETWEEN [2] AND [3] "
        
        If Trim(varCondition(2)) <> "" Then strCondition = strCondition & " AND A.��ϵ�� LIKE [4] "
        If Trim(varCondition(3)) <> "" Then strCondition = strCondition & " AND A.����=[5] "
        
        If Val(varCondition(5)) > 0 Then strCondition = strCondition & " AND A.��Լ��λID=[6] "
        
        strCondition = strCondition & " AND A.���״̬<=[7]"
        
        If Val(varCondition(6)) > 0 Then
            intState = 3
        Else
            intState = 1
        End If

        If Val(varCondition(7)) = 1 Then
            intGroup = 1
            strCondition = strCondition & " AND NVL(A.�Ƿ�����,0)=[8]"
        ElseIf Val(varCondition(7)) = 2 Then
            intGroup = 0
            strCondition = strCondition & " AND NVL(A.�Ƿ�����,0)=[8]"
        End If
                                        
        gstrSQL = GetPublicSQL(SQL.���ԤԼ����, strCondition)
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngSvrDept, CDate(varCondition(0)), CDate(varCondition(1)) + 1 - 1 / 24 / 60 / 60, "%" & CStr(varCondition(2)) & "%", CStr(varCondition(3)), Val(varCondition(5)), intState, intGroup)

        If rs.BOF = False Then Call LoadGrid(vsf, rs, , , ils13)
        
    Case "�����Ա"
        
        vsfPerson.RowHidden(1) = False
        
        gstrSQL = "Select '' As ���ʱ��,0 AS ����,0 As ID,0 As ����id,������� As ����,0 AS �����,'' AS �Ա�,'' AS ����,������� " & _
            "From ������ Where �Ǽ�id=[1]"
        
        gstrSQL = "Select * From (" & gstrSQL & " Union All " & _
            "Select TO_CHAR(B.���ʱ��,'yyyy-mm-dd') As ���ʱ��,B.����,A.����id AS ID,A.����id,A.����,A.�����,A.�Ա�,A.����,B.������� " & _
            "from ������Ϣ A,�����Ա���� B  " & _
            "WHERE A.����ID=B.����ID AND B.�Ǽ�id=[1]) Order By �������,����� "
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)))
        
        If rs.BOF = False Then
        
            Call FillGrid(vsfPerson, rs)
            
            With vsfPerson
                For lngLoop = 1 To .Rows - 1
                    If Val(.TextMatrix(lngLoop, GetCol(vsfPerson, "�����"))) = 0 Then
                        .MergeRow(lngLoop) = True
                        .Cell(flexcpText, lngLoop, 0, lngLoop, .Cols - 2) = .TextMatrix(lngLoop, 0)
                        .Cell(flexcpFontBold, lngLoop, 0, lngLoop, .Cols - 2) = True
                        .RowOutlineLevel(lngLoop) = 1
                        .IsSubtotal(lngLoop) = True
                    End If
                Next
                
                If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[����]")) <> "" Then
                    .RowHidden(1) = True
                    .Row = 2
                Else
                    .RowHidden(1) = False
                End If
            End With
        End If
        
    Case "��Ա��Ŀ"
        
        If vsfPerson.IsSubtotal(vsfPerson.Row) Then
            gstrSQL = GetPublicSQL(SQL.�����Ŀ�嵥)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), vsfPerson.TextMatrix(vsfPerson.Row, 0))
        Else
            gstrSQL = GetPublicSQL(SQL.��Ա�����Ŀ)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), Val(vsfPerson.RowData(vsfPerson.Row)))
        End If
        
        Dim sglSum(0 To 1) As Single
        
        If rs.BOF = False Then
            Call LoadGrid(vsfItem, rs, , , ilsGrid)
            
            If vsfPerson.IsSubtotal(vsfPerson.Row) = False Then
                
                '��������ܶ�,900,7,1,1,;���۸�,
                For lngLoop = 1 To vsfItem.Rows - 1
                    sglSum(0) = sglSum(0) + Val(vsfItem.TextMatrix(lngLoop, mCol.i�����۸�))
                    sglSum(1) = sglSum(1) + Val(vsfItem.TextMatrix(lngLoop, mCol.i���۸�))
                Next
                
                vsfItem.Rows = vsfItem.Rows + 1
                vsfItem.TextMatrix(vsfItem.Rows - 1, mCol.i�����۸�) = " " & Format(sglSum(0), "0.00")
                vsfItem.TextMatrix(vsfItem.Rows - 1, mCol.i���۸�) = Format(sglSum(1), "0.00")
                vsfItem.MergeCells = flexMergeFree
                vsfItem.MergeRow(vsfItem.Rows - 1) = True
                vsfItem.Cell(flexcpText, vsfItem.Rows - 1, 0, vsfItem.Rows - 1, mCol.i�����۸� - 1) = "�ϼƣ�"
                vsfItem.Cell(flexcpForeColor, vsfItem.Rows - 1, 0, vsfItem.Rows - 1, vsfItem.Cols - 1) = COLOR.��ɫ
            End If
        End If
        
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItem(ByRef lngKey As Long, ByVal intFoot As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����༭���ݴ�����ã��ӿں���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    Dim lngLoop As Long
    
    
    On Error GoTo errHand
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey Then
            Exit For
        End If
    Next
    
    If lngLoop < vsf.Rows And lngLoop > 0 Then
        
        lngIndex = lngLoop
        lngIndex = lngLoop + intFoot
        
        If Val(vsf.RowData(lngIndex)) > 0 Then
            lngKey = Val(vsf.RowData(lngIndex))
            GetItem = True
        End If
    End If
    
    Exit Function
    
errHand:
    
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݱ༭/����
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim lngTmp As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim rsItems As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strGroup As String
    Dim strPrompt As String
    Dim intCount2 As Integer
    Dim lng����� As Long
    Dim bytNew As Byte
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    lngKey = Val(vsf.RowData(vsf.Row))
    
    '��һ������
    Select Case strMenuItem
    Case "����ԤԼ"             '����ԤԼ---------------------------------------------------------------------------------
            
        If Not frmScheduleEdit.ShowEdit(Me, 0, mlngSvrDept) Then Exit Function
        
    Case "����ԤԼ"             '����ԤԼ---------------------------------------------------------------------------------
        
        If Not frmScheduleEdit.ShowEdit(Me, 0, mlngSvrDept, True) Then Exit Function
        
    Case "�޸ĸ���ԤԼ"         '�޸ĸ���ԤԼ---------------------------------------------------------------------------------
        
        If lngKey = 0 Then Exit Function
        If Not frmScheduleEdit.ShowEdit(Me, lngKey, mlngSvrDept) Then Exit Function
        
    Case "�޸�����ԤԼ"         '�޸�����ԤԼ---------------------------------------------------------------------------------
                
        If lngKey = 0 Then Exit Function
        If Not frmScheduleEdit.ShowEdit(Me, lngKey, mlngSvrDept, True) Then Exit Function
        
    Case "ɾ��ԤԼ"             'ɾ��ԤԼ---------------------------------------------------------------------------------

        If lngKey = 0 Then Exit Function
        
        If MsgBox("�����Ҫɾ����ǰ���ԤԼ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                
        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_DELETE(" & lngKey & ")"
        
    Case "ȷ��ԤԼ"             'ȷ��ԤԼ---------------------------------------------------------------------------------
        If lngKey = 0 Then Exit Function
        
        If MsgBox("�����Ҫȷ�ϵ�ǰ���ԤԼ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        Select Case CheckAllowMedical(lngKey)
        Case 1
            strPrompt = "��ǰ��컹û������������壬������"
            If MsgBox(strPrompt, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 2
            strPrompt = "��ǰ��컹û�����������Ա��������"
            If MsgBox(strPrompt, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 3
            strPrompt = "��ǰ���������Ŀ��������ÿ���������������Ŀ����������"
            If MsgBox(strPrompt, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Case 4
            strPrompt = "����û�з���������Ա�����Ƚ�����Ա��𻮷֣�"
            ShowSimpleMsg strPrompt
            Exit Function
        End Select
                
        
        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_STATE(" & lngKey & ",2)"
        
    Case "ȡ��ԤԼ"             'ȡ��ԤԼ---------------------------------------------------------------------------------
        If lngKey = 0 Then Exit Function
        
        If MsgBox("�����Ҫȡ����ǰ���ԤԼȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_STATE(" & lngKey & ",1)"
        
    Case "�����Ŀ"             '�������Ŀѡ����---------------------------------------------------------------------------------
                
        If lngKey = 0 Then Exit Function
        
        Call MedicalItemsRecord(rsItems)
        
        '��ȡ�����Ŀ
        gstrSQL = GetPublicSQL(SQL.���������Ŀ)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)))
        
        If WriteItems(rs, rsItems, 2) = False Then Exit Function
        
        If Not frmItemsEdit.ShowEdit(Me, Val(vsf.RowData(vsf.Row)), rsItems, mlngSvrDept, IIf(vsf.TextMatrix(vsf.Row, GetCol(vsf, "[����]")) <> "", False, True)) Then Exit Function

        '�����Ѿ�ɾ���������Ŀ
        Call FilterRecord(rsItems, "ɾ��='1'")
        Call DeleteMedicalItems(strSQL, rsItems, vsf.TextMatrix(vsf.Row, GetCol(vsf, "No")), lngKey, 0)

        '��������ӵ������Ŀ
        Call FilterRecord(rsItems, "�¼�<>'1'")
        Call InsertMedicalItems(strSQL, rsItems, lngKey, 0)

        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_�������(" & lngKey & ")"
        
    Case "��Ա��Ŀ"
        
        '�༭�����Ա�ĸ��������Ŀ
        
        If lngKey = 0 Then Exit Function
        If Val(vsfPerson.RowData(vsfPerson.Row)) = 0 Then Exit Function
        
        Call MedicalItemsRecord(rsItems)

        gstrSQL = GetPublicSQL(SQL.��Ա�����Ŀ)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), Val(vsfPerson.RowData(vsfPerson.Row)))
        Call WriteItems(rs, rsItems, 2)

        If Not frmItemsEdit.ShowEdit(Me, lngKey, rsItems, mlngSvrDept, False, 1, Val(vsfPerson.RowData(vsfPerson.Row))) Then Exit Function

        '�����Ѿ�ɾ���������Ŀ
        Call FilterRecord(rsItems, "ɾ��='1'")
        Call DeleteMedicalItems(strSQL, rsItems, vsf.TextMatrix(vsf.Row, GetCol(vsf, "No")), lngKey, Val(vsfPerson.TextMatrix(vsfPerson.Row, GetCol(vsfPerson, "����id"))))

        '��������ӵ������Ŀ
        Call FilterRecord(rsItems, "�¼�<>'1'")
        Call InsertMedicalItems(strSQL, rsItems, lngKey, Val(vsfPerson.TextMatrix(vsfPerson.Row, GetCol(vsfPerson, "����id"))))

    Case "�ܼ���Ա"
        
        If lngKey = 0 Then Exit Function
        
        Dim lng����id As Long
                
        Call MedicalItemsRecord(rsItems, 2)
        
        gstrSQL = GetPublicSQL(SQL.�����Ա����)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If WriteItems(rs, rsItems, , 2) = False Then Exit Function
        
        lngTmp = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "��Լ��λid")))
        If lngTmp = 0 Then Exit Function
        
        If Not frmPersonEdit.ShowEdit(Me, Val(vsf.RowData(vsf.Row)), rsItems, IIf(vsf.TextMatrix(vsf.Row, GetCol(vsf, "[����]")) <> "", False, True), , lngTmp) Then Exit Function
        
        Dim intCount As Integer
        Dim intCount1 As Integer
        
        intCount = -1
        
        strSQL(ReDimArray(strSQL)) = "zl_�����Ա����_Delete(" & lngKey & ")"
        
        rsItems.Filter = ""
        Do While Not rsItems.EOF
            
            '����������
            If rsItems("��������") <> "" Then
                
                If CheckStrValid(rsItems("��������"), CHECKFORMAT.����) = False Then
                    ShowSimpleMsg rsItems("����").Value & "�ĳ���������Ч��"
                    Exit Function
                End If
            End If
            
            lng����id = rsItems("����ID").Value
            bytNew = 0
            If lng����id = 0 Then
                intCount = intCount + 1
                lng����id = GetNextNo(1)
                bytNew = 1
            End If
            
            intCount1 = intCount1 + 1
            
            If zlCommFun.NVL(rsItems("�����").Value, 0) < 1 Then
                lng����� = GetNextNo(3)
                
                intCount2 = intCount2 + 1
            Else
                lng����� = zlCommFun.NVL(rsItems("�����").Value, 0)
            End If
            
            strSQL(ReDimArray(strSQL)) = "ZL_�����Ա����_INSERT(" & lngKey & "," & _
                                                                IIf(lng����id = 0, "NULL", lng����id) & ",'" & _
                                                                rsItems("���").Value & "','" & _
                                                                rsItems("����").Value & "','" & _
                                                                rsItems("���֤").Value & "','" & _
                                                                rsItems("�Ա�").Value & "'," & _
                                                                IIf(rsItems("��������").Value = "", "NULL", "TO_DATE('" & rsItems("��������").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                                rsItems("����״��").Value & "','" & _
                                                                rsItems("����").Value & "','" & _
                                                                rsItems("����").Value & "','" & _
                                                                rsItems("ѧ��").Value & "','" & _
                                                                rsItems("ְҵ").Value & "','" & _
                                                                rsItems("��ϵ������").Value & "','" & _
                                                                rsItems("��ϵ�˵绰").Value & "','" & _
                                                                rsItems("�����ʼ�").Value & "','" & _
                                                                rsItems("��ϵ�˵�ַ").Value & "','" & _
                                                                rsItems("������λ").Value & "','" & _
                                                                rsItems("����").Value & "'," & _
                                                                lng����� & ",'" & _
                                                                rsItems("IC����").Value & "','" & _
                                                                rsItems("������").Value & "'," & _
                                                                rsItems("���￨��").Value & "'," & _
                                                                "1," & _
                                                                IIf(intCount1 = rsItems.RecordCount, "1", "0") & ",0," & bytNew & _
                                                                ",Null)"
            
            rsItems.MoveNext
        Loop
        
    Case "�������"                 '��������Ϣ�༭����(��Լ��λ)------------------------------------------------------------
        
        If lngKey = 0 Then Exit Function
        If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[����]")) <> "" Then Exit Function
        lngTmp = Val(vsf.TextMatrix(vsf.Row, GetCol(vsf, "��Լ��λid")))
        If lngTmp = 0 Then Exit Function
                                        
    Case "�����Ա"
        
        If lngKey = 0 Then Exit Function
        If Not frmPatientGroupEdit.ShowEdit(Me, lngKey) Then Exit Function
                            
    Case "��Ƭ�ɼ�"
        
        If lngKey > 0 Then
            Call frmPersonPhoto.ShowEdit(Me, lngKey, 0)
            Exit Function
        End If
        
    End Select
    
    '�ڶ�������
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    Select Case strMenuItem
    Case "ɾ��ԤԼ"
    
        'ɾ����
        If vsf.Rows = 2 Then
            Call ResetVsf(vsf)
        Else
            vsf.RemoveItem vsf.Row
        End If
        
        Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
'        Call AppendSapceRows(vsf, lnX, lnY)
        
        MenuClick = True
        
        Exit Function
        
    End Select
    
    Call mnuViewRefresh_Click
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Sub PrintData(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��ӡ����
    '������ bytMode                         ��ӡ��ʽ��1-��ӡ��2-Ԥ����3-�����Excel��
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
'    If cboGroup.ListCount < 1 Then Exit Sub
    
    If UserInfo.���� = "" Then Call GetUserInfo
    
    Call CopyGrid(vsf, vsfPrint, 2)
    objPrint.Title.Text = "���ԤԼ�嵥"
    
    Set objRow = New zlTabAppRow
    objRow.Add "��첿�ţ�" & cboDept.Text
    objRow.Add ""
    
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Body = vsfPrint

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
        
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub cboDept_Click()
    
    If mblnStartUp Then Exit Sub
    If mlngSvrDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngSvrDept = cboDept.ItemData(cboDept.ListIndex)
    
    Call mnuViewRefresh_Click
    
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
        
    If InitActivate = False Then
        mblnStartUp = False
        Unload Me
        Exit Sub
    End If
    
    mblnStartUp = False
    
    Call cboDept_Click
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitLoad
    
    Call ApplyPrivilege(gstrPrivs)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1500 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1500
    End If
    
    With vsf
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height
        .Width = vsf.Width
    End With

    With vsfPerson
        .Left = 0
        .Top = imgX_S.Top + imgX_S.Height
        .Width = imgY_S.Left - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
    With imgY_S
        .Top = vsfPerson.Top
        .Height = vsfPerson.Height
    End With
    
    With vsfItem
        .Left = imgY_S.Left + imgY_S.Width
        .Top = vsfPerson.Top
        .Width = Me.ScaleWidth - .Left - 30
        .Height = vsfPerson.Height
    End With
    
'    Call AppendSapceRows(vsf, lnX, lnY)
'    Call AppendSapceRows(vsfPerson, lnX1, lnY1)
'    Call AppendSapceRows(vsfItem, lnX2, lnY2)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnStartUp Then
        Cancel = True
        Exit Sub
    End If
    'ʹ�ø��Ի�����

    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1500 Then imgX_S.Top = Me.Height - imgX_S.Height - 1500
    
            
    Form_Resize
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 3000 Then imgY_S.Left = 3000
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub

Private Sub mnuAdditionPerItems_Click()
    If vsf.TextMatrix(vsf.Row, GetCol(vsf, "[����]")) <> "" Then
        Call MenuClick("�����Ŀ")
    Else
        Call MenuClick("��Ա��Ŀ")
    End If
End Sub

Private Sub mnuAdditionPhoto_Click()
    Call MenuClick("��Ƭ�ɼ�")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("����ԤԼ")
End Sub

Private Sub mnuEditAddGroup_Click()
    Call MenuClick("����ԤԼ")
End Sub

Private Sub mnuEditCancel_Click()
    Call MenuClick("ȡ��ԤԼ")
End Sub

Private Sub mnuEditCheck_Click()
    Call MenuClick("ȷ��ԤԼ")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("ɾ��ԤԼ")
End Sub

Private Sub mnuAdditionGroup_Click()
    Call MenuClick("�������")
End Sub

Private Sub mnuAdditionGroupMember_Click()
    Call MenuClick("�����Ա")
End Sub

Private Sub mnuAdditionItems_Click()
    Call MenuClick("�����Ŀ")
End Sub

Private Sub mnuEditModify_Click()
    If vsf.TextMatrix(vsf.Row, 0) <> "" Then
        Call MenuClick("�޸ĸ���ԤԼ")
    Else
        Call MenuClick("�޸�����ԤԼ")
    End If
End Sub

Private Sub mnuAdditionPersons_Click()
    Call MenuClick("�ܼ���Ա")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileNew_Click()
    frmScheduleExcel.ShowEdit Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    Dim lngSvrKey As Long
    Dim lngSvrKey2 As Long
    Dim lngSvrKey3 As Long
                
    '�������ԤԼ�������������Ŀ
    lngSvrKey = Val(vsf.RowData(vsf.Row))
    
    mblnAllowChange = False
    
    Call ResetVsf(vsf)
    Call ResetVsf(vsfPerson)
    Call ResetVsf(vsfItem)
        
    Call RefreshData("���ԤԼ")
    
    '�ָ����ԤԼ
    Call RestoreRow(vsf, lngSvrKey)
    
    Call RefreshData("�����Ա")
    
    mblnAllowChange = True
    
    Call vsfPerson_AfterRowColChange(0, 0, vsfPerson.Row, vsfPerson.Col)
    
'    Call AppendSapceRows(vsfPerson, lnX1, lnY1)
'    Call AppendSapceRows(vsf, lnX, lnY)
    
        
    Call AdjustEnableState
    Call RefreshStateInfo
End Sub

Private Sub mnuViewSearch_Click()
    If frmScheduleFilter.ShowFilter(Me, mstrCondition) Then
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditAdd.Visible Then mobjPopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditAddGroup.Visible Then mobjPopMenu.Add 2, mnuEditAddGroup.Caption, , , mnuEditAddGroup.Enabled
    Case 2
        
        If mnuAddition.Visible = False Then Exit Sub

        If mnuAdditionPersons.Visible Then mobjPopMenu.Add 1, mnuAdditionPersons.Caption, , , mnuAdditionPersons.Enabled
        
        mobjPopMenu.Add 2, "-", , 2, True
        
        If mnuAdditionGroupMember.Visible Then mobjPopMenu.Add 3, mnuAdditionGroupMember.Caption, , , mnuAdditionGroupMember.Enabled
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditAddGroup_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuAdditionPersons_Click
        Case 3
            Call mnuAdditionGroupMember_Click
        End Select
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "ԤԼ"
        
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "�޸�"
        Call mnuEditModify_Click
    Case "ɾ��"
        Call mnuEditDelete_Click
    Case "ȷ��"
        Call mnuEditCheck_Click
    Case "ȡ��"
        Call mnuEditCancel_Click
    Case "����"
        Call mnuViewSearch_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "����"
        Call mnuEditAdd_Click
    Case "����"
        Call mnuEditAddGroup_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnAllowChange = False Then Exit Sub
        
    If mlngSvrKey = Val(vsf.RowData(NewRow)) Then Exit Sub
    mlngSvrKey = Val(vsf.RowData(NewRow))
    
    Call ResetVsf(vsfPerson)
    Call ResetVsf(vsfItem)
        
    Call RefreshData("�����Ա")
    
    Call vsfPerson_AfterRowColChange(0, 0, vsfPerson.Row, vsfPerson.Col)

    
    Call AdjustEnableState
    
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 1)
End Sub

Private Sub vsf_DblClick()
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then
        Call mnuEditModify_Click
    End If
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.����
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vsf_DblClick
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.�ǽ���
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsf.hWnd, X, Y)
        If mnuEdit.Visible Then Me.PopupMenu mnuEdit
    End If
End Sub


Private Sub vsfItem_GotFocus()
    vsfItem.BackColorSel = COLOR.����
End Sub

Private Sub vsfItem_LostFocus()
    vsfItem.BackColorSel = COLOR.�ǽ���
End Sub

Private Sub vsfPerson_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If mblnAllowChange = False Then Exit Sub
    
    Call ResetVsf(vsfItem)
    Call RefreshData("��Ա��Ŀ")
    
    Call AdjustEnableState
    
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub vsfPerson_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsfPerson.hWnd, X, Y)
        If mnuAddition.Visible Then Me.PopupMenu mnuAddition
    End If
End Sub
