VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBalance 
   Caption         =   "����������"
   ClientHeight    =   6885
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "frmBalance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TabStrip tbs 
      Height          =   360
      Left            =   3510
      TabIndex        =   8
      Top             =   2505
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1.������ϸ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2.���㷽ʽ"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6525
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBalance.frx":1CFA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13097
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1244
      BandCount       =   2
      _CBWidth        =   10305
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "��첿��"
      Child2          =   "cboDept"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   2100
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   8115
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9255
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":29CE
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":2BEA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":2E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3024
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3244
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3464
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8535
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3BDE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":3DFE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":401E
            Key             =   "Parent"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":423A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":445A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":4674
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":4894
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":4AB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1380
      Index           =   0
      Left            =   3420
      TabIndex        =   3
      Top             =   855
      Width           =   2790
      _cx             =   4921
      _cy             =   2434
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
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX0 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY0 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1845
      Index           =   1
      Left            =   4590
      TabIndex        =   4
      Top             =   3750
      Width           =   3795
      _cx             =   6694
      _cy             =   3254
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
      GridColor       =   12698049
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1845
      Index           =   2
      Left            =   3810
      TabIndex        =   5
      Top             =   2610
      Width           =   3720
      _cx             =   6562
      _cy             =   3254
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
      GridColor       =   12698049
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8910
      Top             =   3105
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
            Picture         =   "frmBalance.frx":522E
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":55C8
            Key             =   "��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalance.frx":5B62
            Key             =   "��"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   9480
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   4305
      MousePointer    =   7  'Size N S
      Top             =   2355
      Width           =   4845
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
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBalanceBill 
         Caption         =   "��ӡƱ��(&O)"
      End
      Begin VB.Menu mnuFileBalanceDetail 
         Caption         =   "������ϸ(C)"
         Begin VB.Menu mnuFileBalanceDetaiPrintView 
            Caption         =   "Ԥ��(&1)"
         End
         Begin VB.Menu mnuFileBalanceDetaiPrint 
            Caption         =   "��ӡ(&2)"
         End
         Begin VB.Menu mnuFileBalanceDetaiExcel 
            Caption         =   "�����Excel(&3)"
         End
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditBalance 
         Caption         =   "������(&B)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditBalanceCancel 
         Caption         =   "��������(&M)"
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
      Begin VB.Menu mnuViewList 
         Caption         =   "δ�������嵥(&L)"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
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
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                                  '����������־
Private mlngSvrKey(0 To 2)  As Long                             '���ڱ����������ѡ�е��йؼ���
Private mlngDept As Long
Private mblnNoAllowChange As Boolean
Private WithEvents mobjPopMenu As clsPopMenu                    '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mstrCondition As String
Private mbytKind As Byte
Private Type TYPE_USR_CELL
    Row As Integer
    Col As Integer
End Type
Private mblnDataMoved As Boolean

Private musrSavePos As TYPE_USR_CELL

'�������Զ�����̻���************************************************************************************************
Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����ؼ��е�����
    '����:  strMenuItem             Ҫ����ķ�Χ
    '����;  True                    ����ɹ�
    '       False                   ���ʧ��
    '------------------------------------------------------------------------------------------------------------------
    
    Select Case strMenuItem
    Case "������"
        Call ResetVsf(vsf(0))
        Call InheritAppendSpaceRows(0)
    Case "���㵥��"
        Call ResetVsf(vsf(1))
        Call InheritAppendSpaceRows(1)
    Case "���㷽ʽ"
        Call ResetVsf(vsf(2))
        Call InheritAppendSpaceRows(2)
    End Select
        
End Function

Private Sub InheritAppendSpaceRows(ByVal intIndex As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���������
    '����:  intIndex                Ҫ������еı��ؼ�������
    '------------------------------------------------------------------------------------------------------------------
    Select Case intIndex
    Case 0
        Call AppendRows(vsf(intIndex), lnX0, lnY0)
    Case 1
        Call AppendRows(vsf(intIndex), lnX1, lnY1)
    Case 2
        Call AppendRows(vsf(intIndex), lnX2, lnY2)
    End Select
End Sub

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ�����ݣ������ڴ����Load�¼�
    '����:  True                    �ɹ�
    '       False                   ����
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mbytKind = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 3))
    
    mlngDept = 0
    mstrCondition = Format(DateAdd("d", -7, CDate(zlDatabase.Currentdate)), "yyyy-MM-dd") & "'" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mstrCondition = mstrCondition & "''''''''"
        
    strVsf = ",255,4,1,1,[״̬];���ݺ�,900,1,1,1,;Ʊ�ݺ�,900,1,1,1,;��������,1800,1,1,1,;������,900,7,1,1,;������,810,1,1,1,;����ʱ��,1670,1,1,1,;��ϵ��,810,1,1,1,;��ϵ�绰,1200,1,1,1,;��ϵ��ַ,1800,1,1,1,"
    Call CreateVsf(vsf(0), strVsf)
    vsf(0).Cols = vsf(0).Cols + 1
    vsf(0).ColWidth(vsf(0).Cols - 1) = 15
    Set vsf(0).Cell(flexcpPicture, 0, 0) = ils13.ListImages("״̬").Picture
    
    strVsf = "����,810,1,1,1,;���ݺ�,900,1,1,1,;��Ŀ,2400,1,1,1,;��Ŀ,750,1,1,1,;������,900,7,1,1,;��������,1080,1,1,1,;����ʱ��,1670,1,1,1,"
    Call CreateVsf(vsf(1), strVsf)
    vsf(1).Cols = vsf(1).Cols + 1
    vsf(1).ColWidth(vsf(1).Cols - 1) = 15
        
    strVsf = "���ݺ�,900,1,1,1,;���,810,7,1,1,;���㷽ʽ,900,1,1,1,;�������,900,1,1,1,"
    Call CreateVsf(vsf(2), strVsf)
    vsf(2).Cols = vsf(2).Cols + 1
    vsf(2).ColWidth(vsf(2).Cols - 1) = 15
    
    'Ʊ���ϸ����
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    strTmp = zlDatabase.GetPara(24, glngSys, , "00000")
    If strTmp <> "" Then
        gblnBill���� = (Mid(strTmp, 3, 1) = "1")
        gblnStrictCtrl = (Mid(strTmp, 3, 1) = "1")
    End If
               
    glng����ID = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0)
    glngShareUseID = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0)
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���ݱ༭/����
    '����:  strMenuItem             ��������
    '����:  True                    �ɹ�
    '       False                   ʧ��/ȡ��/����
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strNo As String
    
    On Error GoTo errHand
        
    ReDim Preserve strSQL(1 To 1)
    
    lngKey = Val(vsf(0).RowData(vsf(0).Row))
        
    '��һ������
    Select Case strMenuItem
    Case "������"
        
        If Not frmBalanceEdit.ShowEdit(Me, 0) Then Exit Function
                
    
    Case "��������"
        
        If lngKey = 0 Then Exit Function
        strNo = vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "���ݺ�"))
        If strNo = "" Then Exit Function
        
        '����ת������
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "�˽��ʵ����Ѿ�ת���������ٲ�����"
            Exit Function
        End If
        
        If MsgBox("���Ҫ���ϵ�ǰ���㵥����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "zl_�������¼_Cancel(" & Val(vsf(0).RowData(vsf(0).Row)) & ")"
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "��������", "������"
        Call mnuViewRefresh_Click
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
End Function


Private Sub PrintData(ByVal bytMode As Byte)
    '--------------------------------------------------------------------------------------------------------
    '���ܣ� ��ӡ����
    '������ bytMode                         ��ӡ��ʽ��1-��ӡ��2-Ԥ����3-�����Excel��
    '--------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
                
    mblnNoAllowChange = True
    
    musrSavePos.Row = vsf(0).Row
    musrSavePos.Col = vsf(0).Col
    
    If UserInfo.���� = "" Then Call GetUserInfo

    objPrint.Title = "��" & zlCommFun.GetNeedName(cboDept.Text) & "�����������㵥"
    
    Call CopyGrid(vsf(0), vsfPrint, 1)
    
    Set objPrint.Body = vsfPrint

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)
    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
    
    On Error Resume Next
    vsf(0).Row = musrSavePos.Row
    vsf(0).Col = musrSavePos.Col
    vsf(0).ShowCell vsf(0).Row, vsf(0).Col
    On Error GoTo 0
    
    mblnNoAllowChange = False
End Sub

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� Ӧ��Ȩ�޴���
    '������ strPrivilege                    Ȩ��
    '------------------------------------------------------------------------------------------------------------------
    
'    strPrivilege = "������;��������;�����ش�"
    
    '�����С������㡱�͡��������ϡ�Ȩ��ʱ
    If InStr(strPrivilege, "������") = 0 And InStr(strPrivilege, "��������") = 0 Then
        mnuEdit.Visible = False
    Else
        '�����С������㡱Ȩ��ʱ
        If InStr(strPrivilege, "������") = 0 Then
            mnuEditBalance.Visible = False
        End If
        
        '�����С��������ϡ�Ȩ��ʱ
        If InStr(strPrivilege, "��������") = 0 Then
            mnuEditBalanceCancel.Visible = False
        End If
        
    End If
    
    '�����С������ش�Ȩ��ʱ
    If InStr(strPrivilege, "�����ش�") = 0 Then
        mnuFileBalanceBill.Visible = False
        mnuFileBalanceDetail.Visible = False
        mnuFile_2.Visible = False
    End If
            
    '��������
    tbrThis.Buttons("����").Visible = mnuEdit.Visible And mnuEditBalance.Visible
    tbrThis.Buttons("����").Visible = mnuEdit.Visible And mnuEditBalanceCancel.Visible
    tbrThis.Buttons("Split_2").Visible = tbrThis.Buttons("����").Visible Or tbrThis.Buttons("����").Visible
    
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���������ܲ˵��Ŀ���״̬
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuFileBalanceBill.Enabled = True
    mnuFileBalanceDetail.Enabled = True
    
    mnuEditBalanceCancel.Enabled = True
    
    If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
        
        mnuFileBalanceBill.Enabled = False
        mnuFileBalanceDetail.Enabled = False
        
        mnuEditBalanceCancel.Enabled = False
    Else
        Select Case vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "[״̬]"))
        Case "��"

        Case "��"
            mnuEditBalanceCancel.Enabled = False
            mnuFileBalanceBill.Enabled = False
        End Select
    End If
    
    If Val(vsf(1).RowData(vsf(1).Row)) = 0 Then
        mnuFileBalanceDetail.Enabled = False
    End If
    
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("����").Enabled = mnuEditBalance.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditBalanceCancel.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��״̬����ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------------------
    If Val(vsf(0).RowData(1)) = 0 Then
        stbThis.Panels(2).Text = "û�н��㵥�ݡ�"
    Else
        stbThis.Panels(2).Text = "���� " & vsf(0).Rows - 1 & " �Ž��㵥�ݡ�"
    End If
    
End Sub

Private Function GetQueryCondition(ByVal strCondition As String) As String
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim strResult As String
     
   
    '�����Ǹ��������������ɵ��������
    '�洢��ʽ:��ʼʱ��'����ʱ��'��ʼ���ݺ�'�������ݺ�'��ʼƱ�ݺ�'����Ʊ�ݺ�'������'�������'�������id'����'����ȷ��
    
    If strCondition = "" Then Exit Function
        
    varTmp = Split(strCondition, "'")
    
    strResult = " AND C.�շ�ʱ�� BETWEEN TO_DATE('" & Format(varTmp(0), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(1), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
    
    '���ݺ�
    If Trim(varTmp(2)) <> "" And Trim(varTmp(3)) <> "" Then
        strResult = strResult & " AND C.No BETWEEN '" & Trim(varTmp(2)) & "' AND '" & Trim(varTmp(3)) & "'"
    ElseIf Trim(varTmp(2)) <> "" Then
        strResult = strResult & " AND C.No='" & Trim(varTmp(2)) & "'"
    ElseIf Trim(varTmp(3)) <> "" Then
        strResult = strResult & " AND C.No='" & Trim(varTmp(3)) & "'"
    End If
    
    'ʵ��Ʊ��
    If Trim(varTmp(4)) <> "" And Trim(varTmp(5)) <> "" Then
        strResult = strResult & " AND C.ʵ��Ʊ�� BETWEEN '" & Trim(varTmp(4)) & "' AND '" & Trim(varTmp(5)) & "'"
    ElseIf Trim(varTmp(4)) <> "" Then
        strResult = strResult & " AND C.ʵ��Ʊ��='" & Trim(varTmp(4)) & "'"
    ElseIf Trim(varTmp(5)) <> "" Then
        strResult = strResult & " AND C.ʵ��Ʊ��='" & Trim(varTmp(5)) & "'"
    End If
    
    '������
    If Trim(varTmp(6)) <> "" Then strResult = strResult & " AND C.����Ա����='" & Trim(varTmp(6)) & "'"
    
    '��������
    If Val(varTmp(8)) > 0 Then strResult = strResult & " AND A.��Լ��λid=" & Val(varTmp(8))
        
    '��¼״̬
    If Val(varTmp(9)) = 0 Then strResult = strResult & " AND C.��¼״̬=1"
    
    GetQueryCondition = strResult
    
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��/װ������
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strTmp As String
    Dim blnDataMoved As Boolean
    
    On Error GoTo errHand
    
    Call InitSysPara
    
    Select Case strMenuItem
    Case "������"
        
        Call ClearGrid(vsf(0))
        gstrSQL = "SELECT A.ID,A.���ݺ�,A.Ʊ�ݺ�,A.������,A.������,TO_CHAR(A.����ʱ��,'yyyy-mm-dd hh24:mi') AS ����ʱ��," & _
                        "DECODE(A.��¼״̬,1,'��','��') AS ״̬," & _
                        "DECODE(A.��¼״̬,1,'0','192') AS ǰ��ɫ," & _
                        "B.���� AS ��������,B.��ϵ��,B.�绰 AS ��ϵ�绰,B.��ַ AS ��ϵ��ַ FROM " & _
                        "( " & _
                        "SELECT C.ID,C.��¼״̬,A.��Լ��λid," & _
                               "C.NO AS ���ݺ�, " & _
                               "A.������, C.ʵ��Ʊ�� AS Ʊ�ݺ�," & _
                               "C.����Ա���� AS ������, " & _
                               "C.�շ�ʱ�� AS ����ʱ�� " & _
                        "FROM �������¼ A, " & _
                             "���˽��ʼ�¼ C " & _
                        "Where C.ID=A.����id " & _
                              "AND A.���㲿��id+0=" & mlngDept & " " & _
                              "AND A.��¼״̬ IN (1,2) " & GetQueryCondition(mstrCondition) & " " & _
                        ") A, " & _
                        "��Լ��λ B " & _
                        "WHERE A.��Լ��λID=B.ID "
                        
        '����ת������
        '--------------------------------------------------------------------------------------------------------------
        blnDataMoved = zlDatabase.DateMoved(Format(Split(mstrCondition, "'")(0), "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
        If blnDataMoved Then
            strTmp = gstrSQL
            strTmp = Replace(strTmp, "�������¼", "H�������¼")
            strTmp = Replace(strTmp, "���˽��ʼ�¼", "H���˽��ʼ�¼")
            gstrSQL = "Select * From (" & gstrSQL & " Union All " & strTmp & ") a "
        End If
        gstrSQL = gstrSQL & " Order By a.���ݺ� Desc"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then
            Call LoadGrid(vsf(0), rs, Array("", "", "", "", gstrDec, gstrDec), , ils13)
        End If
        Call InheritAppendSpaceRows(0)
        
    Case "���㵥��"
        
        lngKey = Val(vsf(0).RowData(vsf(0).Row))
        If lngKey = 0 Then Exit Function
        
        Call ClearGrid(vsf(1))
        gstrSQL = "SELECT A.ID,A.����,A.NO AS ���ݺ�,B.���� AS ��������,C.���� AS ��Ŀ,A.�վݷ�Ŀ AS ��Ŀ,A.���ʽ�� AS ������,TO_CHAR(A.����ʱ��,'yyyy-mm-dd hh24:mi') AS ����ʱ�� " & _
                    "FROM ���˷��ü�¼ A, " & _
                         "���ű� B, " & _
                         "�շ���ĿĿ¼ C " & _
                    "WHERE A.����id = [1] " & _
                          "AND A.��������ID=B.ID " & _
                          "AND C.ID=A.�շ�ϸĿID "
        
        '����ת������
        '--------------------------------------------------------------------------------------------------------------
        mblnDataMoved = zlDatabase.NOMoved("���˽��ʼ�¼", vsf(0).TextMatrix(vsf(0).Row, 1))
        If mblnDataMoved Then
            gstrSQL = Replace(gstrSQL, "���˷��ü�¼", "H���˷��ü�¼")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf(1), rs, Array("", "", "", "", gstrDec), , ils13)
        End If
        Call InheritAppendSpaceRows(1)
        
    Case "���㷽ʽ"
        
        lngKey = Val(vsf(0).RowData(vsf(0).Row))
        If lngKey = 0 Then Exit Function
        
        Call ClearGrid(vsf(2))
        
        gstrSQL = "SELECT A.ID,A.NO AS ���ݺ�, A.��Ԥ�� AS ���,A.���㷽ʽ,A.������� " & _
                    "FROM ����Ԥ����¼ A " & _
                    "WHERE A.����ID=[1]"
        
        '����ת������
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            gstrSQL = Replace(gstrSQL, "����Ԥ����¼", "H����Ԥ����¼")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf(2), rs, Array("", "0.00##"), , ils13)
        End If
        Call InheritAppendSpaceRows(2)
        
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitActive() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ�����ݣ������ڴ����Active�¼�
    '����:  True        �ɹ�
    '       False       ����
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
    zlControl.CboLocate cboDept, UserInfo.����ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
        
    InitActive = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub PrintDetail(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ������б�
    '------------------------------------------------------------------------------------------------------------------
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    
    Dim strNo As String
    
    strNo = vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "���ݺ�"))
    If strNo = "" Then Exit Sub
    
    Call CopyGrid(vsf(1), vsfPrint)
    
    '��ͷ
    objOut.Title.Text = "���������㵥��ϸ"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����

        objRow.Add "���ݺţ�" & vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "���ݺ�"))
        'objRow.Add "���ʷ�Χ��" & mshList.TextMatrix(mshList.Row, GetColNum("��ʼ����")) & " �� " & mshList.TextMatrix(mshList.Row, GetColNum("��������"))
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        'objRow.Add "סԺ�ţ�" & mshList.TextMatrix(mshList.Row, GetColNum("סԺ��"))
        'objRow.Add "������" & mshList.TextMatrix(mshList.Row, GetColNum("����"))
        objOut.UnderAppRows.Add objRow

    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����

    Set objOut.Body = vsfPrint
    
    If bytMode = 1 Then bytMode = zlPrintAsk(objOut)
    
    Me.Refresh
    
    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objOut, bytMode)
    
'    bytR = zlPrintAsk(objOut)
'    Me.Refresh
'    If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR

End Sub

Private Sub mnuFileBalanceDetaiExcel_Click()
    Call PrintDetail(3)
End Sub

Private Sub mnuFileBalanceDetaiPrint_Click()
            
    Call PrintDetail(1)
End Sub

Private Sub mnuFileBalanceDetaiPrintView_Click()
    Call PrintDetail(2)
End Sub

Private Sub mnuViewList_Click()
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1862", Me, "��첿��id=" & mlngDept, 0)
    
End Sub

Private Sub mnuViewSearch_Click()
    If frmBalanceFilter.ShowFilter(Me, mstrCondition) Then
        Call mnuViewRefresh_Click
    End If
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    If mnuEdit.Visible Then
        If mnuEditBalance.Visible Then mobjPopMenu.Add 1, mnuEditBalance.Caption, , , mnuEditBalance.Enabled
        If mnuEditBalanceCancel.Visible Then mobjPopMenu.Add 2, mnuEditBalanceCancel.Caption, , , mnuEditBalanceCancel.Enabled

    End If

End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case Key
    Case 1
        Call mnuEditBalance_Click
    Case 2
        Call mnuEditBalanceCancel_Click
    End Select
End Sub

Private Sub cboDept_Click()
    If mblnStartUp Then Exit Sub
    If mlngDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngDept = cboDept.ItemData(cboDept.ListIndex)
    Call mnuViewRefresh_Click
    
End Sub


Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub

    If InitActive = False Then
        Unload Me
        Exit Sub
    End If
    DoEvents
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
    
    If Me.WindowState = 1 Then Exit Sub
    
    '�����������
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
    End If

    With vsf(0)
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = 0
        .Width = vsf(0).Width
    End With
    
    With tbs
        .Left = vsf(0).Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = vsf(0).Width
    End With
    
    With vsf(1)
        .Left = vsf(0).Left
        .Top = tbs.Top + tbs.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With vsf(2)
        .Left = vsf(0).Left
        .Top = vsf(1).Top
        .Width = vsf(1).Width
        .Height = vsf(1).Height
    End With
    
    Call InheritAppendSpaceRows(0)
    Call InheritAppendSpaceRows(1)
    Call InheritAppendSpaceRows(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000

    Call Form_Resize
End Sub


Private Sub mnuFileBalanceBill_Click()
    
    '���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
    
    Dim strNo As String
    Dim lng����ID As Long
    
    lng����ID = Val(vsf(0).RowData(vsf(0).Row))
    If lng����ID = 0 Then Exit Sub
    
    strNo = vsf(0).TextMatrix(vsf(0).Row, GetCol(vsf(0), "���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����ش�Ʊ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If RePrintBalance(strNo, Me, lng����ID, mbytKind) Then
        
        Call mnuViewRefresh_Click
        
    End If
End Sub

Private Sub mnuEditBalanceCancel_Click()
    Call MenuClick("��������")
End Sub

Private Sub mnuEditBalance_Click()
    Call MenuClick("������")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePara_Click()
    If frmSetExpence.ShowParameter(Me) Then
        
        '���¶�ȡ����
        
        glng����ID = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0)
        glngShareUseID = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 0)
        
        mbytKind = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "���ý���Ʊ������", 3))
        
    End If
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
                
    '����
    musrSavePos.Row = vsf(0).Row
    musrSavePos.Col = vsf(0).Col
    
    mblnNoAllowChange = True
    
    Call ClearData("������")
    Call ClearData("���㵥��")
    Call ClearData("���㷽ʽ")
    
    Call RefreshData("������")
    
    '�ָ����ԤԼ
    
    On Error Resume Next
    vsf(0).Row = musrSavePos.Row
    vsf(0).Col = musrSavePos.Col
    vsf(0).ShowCell vsf(0).Row, vsf(0).Col
    Call SelectRow(vsf(0), 0, vsf(0).Row)
    On Error GoTo 0
    
    Call RefreshData("���㵥��")
    Call RefreshData("���㷽ʽ")
    
    mblnNoAllowChange = False
    
    Call AdjustEnableState
    Call RefreshStateInfo
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

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    
    Case "����"
        Call mnuEditBalance_Click
    
    Case "����"
        Call mnuEditBalanceCancel_Click
    Case "����"
        Call mnuViewSearch_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tbs_Click()
    
    vsf(1).Visible = False
    vsf(2).Visible = False
    
    vsf(tbs.SelectedItem.Index).Visible = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    If mblnNoAllowChange Then Exit Sub
    
    If Index = 0 Then
        Call SelectRow(vsf(Index), OldRow, NewRow)
        
        Call RefreshData("���㵥��")
        Call RefreshData("���㷽ʽ")
        
        Call AdjustEnableState
    End If
    
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Index = 0 And Col = 0 Then Cancel = True
End Sub

Private Sub vsf_GotFocus(Index As Integer)
    vsf(Index).BackColorSel = COLOR.����
    If Index = 0 Then Call SelectRow(vsf(Index), 1, vsf(Index).Row)
End Sub

Private Sub vsf_LostFocus(Index As Integer)
    vsf(Index).BackColorSel = COLOR.�ǽ���
    If Index = 0 Then Call SelectRow(vsf(Index), 1, vsf(Index).Row)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If Index <> 0 Then Exit Sub
    
    Call SendLMouseButton(vsf(Index).hWnd, X, Y)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenuByCursor
    
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


