VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDiagnoseAdvice 
   Caption         =   "�����Ͻ���"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmDiagnoseAdvice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1935
      Left            =   3105
      TabIndex        =   6
      Top             =   3240
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   3413
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDiagnoseAdvice.frx":1CFA
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   315
      Left            =   3150
      TabIndex        =   5
      Top             =   2460
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1.�ο�����"
            Key             =   "�ο�����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2.��������"
            Key             =   "��������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&3.������Ŀ"
            Key             =   "������Ŀ"
            Object.ToolTipText     =   "���ݵ������Ŀ"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagnoseAdvice.frx":1D97
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
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
   Begin MSComctlLib.TreeView tvw 
      Height          =   1770
      Left            =   330
      TabIndex        =   4
      Top             =   945
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3122
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1035
      Top             =   4695
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
            Picture         =   "frmDiagnoseAdvice.frx":262B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":2A7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":2D97
            Key             =   "class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1485
      Left            =   3135
      TabIndex        =   3
      Top             =   870
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   2619
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   405
      Top             =   4695
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
            Picture         =   "frmDiagnoseAdvice.frx":3331
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3783
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10455
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
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
               ImageKey        =   "Class"
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�б�"
               Key             =   "�б�"
               Object.ToolTipText     =   "�б�"
               Object.Tag             =   "�б�"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Text            =   "��ͼ��(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Text            =   "Сͼ��(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8325
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3A9D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3CBD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3EDD
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":40F9
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4315
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":452F
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":474F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":496F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4B8F
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4DAF
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7515
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4FCF
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":51EF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":540F
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":562B
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5847
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5B99
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5DB9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5FD9
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":61F9
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":6419
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1995
      Left            =   4035
      TabIndex        =   7
      Top             =   3630
      Width           =   4275
      _cx             =   7541
      _cy             =   3519
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   1965
         X2              =   1965
         Y1              =   1590
         Y2              =   2805
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   1140
         X2              =   2925
         Y1              =   1725
         Y2              =   1725
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   1995
      Left            =   6435
      TabIndex        =   8
      Top             =   3060
      Width           =   4275
      _cx             =   7541
      _cy             =   3519
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   1140
         X2              =   2925
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   1965
         X2              =   1965
         Y1              =   1590
         Y2              =   2805
      End
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   4860
      MousePointer    =   7  'Size N S
      Top             =   2385
      Width           =   5115
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   2670
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   45
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditClass 
         Caption         =   "��Ϸ���(&C)"
         Begin VB.Menu mnuEditClassAdd 
            Caption         =   "���ӷ���(&A)"
         End
         Begin VB.Menu mnuEditClassModify 
            Caption         =   "�޸ķ���(&M)"
         End
         Begin VB.Menu mnuEditClassDelete 
            Caption         =   "ɾ������(&D)"
         End
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "�������(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸����(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ�����(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdvice 
         Caption         =   "�ο�����(S)"
      End
      Begin VB.Menu mnuEditCondition 
         Caption         =   "��������(&P)"
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
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
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
Attribute VB_Name = "frmDiagnoseAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mstrVsf As String                               '����б���
Private mstrKey As String                               '������ǰ��ѡ��
Private Const mstrLvw As String = "����,2100,0,1;����,900,0,0;����,900,0,0;�Ƿ񼲲�,900,0,0;��Ͻ���,2100,0,0;������,900,0,0;�������,900,0,0;��������,1500,0,0"
Private mlngLoop As Long
Private WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

'�������Զ�����̻���************************************************************************************************
Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrKey = ""
    lvw.Tag = "�ɱ仯��"
    
    If lvw.ListItems.Count = 0 Then zlControl.LvwSelectColumns lvw, mstrLvw, True

    strVsf = "����,900,1,1,1,;�Ա�,600,1,1,1,;��ʼ����,900,1,1,1,;��������,900,1,1,1,;��Ŀ,2100,1,1,1,;����,900,1,1,1,;��Ŀֵ,900,1,1,1,"
    Call CreateVsf(vsf, strVsf)
    
    vsf.MergeCol(0) = True
    
    
    strVsf = "����,2400,1,1,1,;����,1200,1,1,1,;���,900,1,1,1,"
    Call CreateVsf(vsfItem, strVsf)
    vsfItem.Cols = vsfItem.Cols + 1
    vsfItem.ColWidth(vsfItem.Cols - 1) = 15
    
    
    InitLoad = True
    
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
    'strPrivilege = "����;��ɾ��"
    
    '�����С���ɾ�ġ�Ȩ��ʱ
    If InStr(strPrivilege, "��ɾ��") = 0 Then
        mnuEdit.Visible = False
    End If
    
    tbrThis.Buttons("����").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("����").Visible = mnuEdit.Visible
    tbrThis.Buttons("�޸�").Visible = mnuEdit.Visible
    tbrThis.Buttons("ɾ��").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("Split_2").Visible = mnuEdit.Visible
    tbrThis.Buttons("Split_3").Visible = mnuEdit.Visible

End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���������ܲ˵��Ŀ���״̬
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditClassAdd.Enabled = True
    mnuEditClassModify.Enabled = True
    mnuEditClassDelete.Enabled = True
    
    mnuEditClass.Enabled = True
    
    mnuEditAdd.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditAdvice.Enabled = True
    mnuEditCondition.Enabled = True
    
    If lvw.ListItems.Count = 0 Then
                
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditAdvice.Enabled = False
        mnuEditCondition.Enabled = False
    End If
    
    If Not (tvw.SelectedItem Is Nothing) Then
        If tvw.SelectedItem.Key = "K0" Then
            mnuEditClassModify.Enabled = False
            mnuEditClassDelete.Enabled = False
        End If
    End If
    
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("����").Enabled = mnuEditClassModify.Enabled Or mnuEditClassAdd.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("�޸�").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("ɾ��").Enabled = mnuEditDelete.Enabled
        
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��״̬����ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------------------
    
    stbThis.Panels(2).Text = "���� " & lvw.ListItems.Count & " �������ϣ�"
    
End Sub

Public Function GetItem(ByRef lngKey As Long, ByVal intFoot As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����༭���ݴ�����ã��ӿں���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    Set objItem = lvw.ListItems("K" & lngKey)
    If Not (objItem Is Nothing) Then
        
        lngIndex = objItem.Index
        lngIndex = lngIndex + intFoot
        
        Set objItem = Nothing
        Set objItem = lvw.ListItems(lngIndex)
        
        If Not (objItem Is Nothing) Then lngKey = Val(Mid(objItem.Key, 2))
            
        GetItem = True
    Else
        GetItem = False
    End If
    
    Exit Function
    
errHand:
    
End Function

Public Function EditRefresh(ByVal strMenuItem As String, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����༭���ݴ�����ã��ӿں���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    On Error GoTo errHand

    Select Case strMenuItem
    Case "�����Ϸ���"
        
        Call ClearData("�����Ϸ���;������Ŀ¼;�����ϸ���")
        
        Call RefreshData("�����Ϸ���")
        
        On Error Resume Next
        tvw.Nodes("K" & lngKey).Selected = True
        tvw.Nodes("K" & lngKey).EnsureVisible
        On Error GoTo 0
        
        If tvw.Nodes.Count > 0 Then
            If tvw.SelectedItem Is Nothing Then
                tvw.Nodes(1).Selected = True
                tvw.Nodes(1).EnsureVisible
            End If
        End If
        
        Call RefreshData("������Ŀ¼")
        
        Call tbs_Click
        
        
    Case "������Ŀ¼"
    
        Call ClearData("������Ŀ¼;�����ϸ���")
        
        Call RefreshData("������Ŀ¼")
        
        '�ָ��������
        Call zlControl.LvwRestoreItem(lvw, "K" & lngKey)
            
        Call tbs_Click
   
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";�����Ϸ���;") > 0 Then
        tvw.Nodes.Clear
    End If
    
    If InStr(strMenuItem, ";������Ŀ¼;") > 0 Then
        lvw.ListItems.Clear
    End If
    
    If InStr(strMenuItem, ";�����ϸ���;") > 0 Then
        rtb.Text = ""
        Call ResetVsf(vsf)
        Call ResetVsf(vsfItem)
        Call AppendRows(vsf, lnX, lnY)
        Call AppendRows(vsfItem, lnX1, lnY1)
    End If
        
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��/װ������
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    Dim objNode As Node
    Dim rsPrice As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "�����Ϸ���"
        
        gstrSQL = GetPublicSQL(SQL.�����Ϸ���)
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If rs.BOF = False Then Call FillTreeData(tvw, rs)
        
    Case "������Ŀ¼"
        
        If Val(Mid(tvw.SelectedItem.Key, 2)) = 0 Then
            
            gstrSQL = " SELECT E.��� AS ID," & _
                              "E.����," & _
                              "E.����," & _
                              "E.����," & _
                              "E.��Ͻ���,Decode(E.������,Null,'',Trim(To_Char(E.������,'99999'))||'��') As ������,Decode(E.�������,Null,'',Trim(To_Char(E.�������,'99999'))||'��') as �������," & _
                              "Decode(E.�Ƿ񼲲�,1,'��','') As �Ƿ񼲲�," & _
                              "1 as ͼ��, "
                                         
            gstrSQL = gstrSQL & _
                              "F.���� AS �������� " & _
                         "FROM �����Ͻ��� E, �����Ͻ��� F " & _
                        "WHERE E.ĩ�� = 1 AND E.�ϼ���� = F.���(+)"
        Else
            gstrSQL = " SELECT E.��� AS ID," & _
                      "E.����," & _
                      "E.����," & _
                      "E.����," & _
                        "E.��Ͻ���,Decode(E.������,Null,'',Trim(To_Char(E.������,'99999'))||'��') as ������,Decode(E.�������,Null,'',Trim(To_Char(E.�������,'99999'))||'��') as �������," & _
                      "Decode(E.�Ƿ񼲲�,1,'��','') As �Ƿ񼲲�," & _
                      "1 as ͼ��,"
                                         
                gstrSQL = gstrSQL & _
                      "F.���� AS �������� " & _
                 "FROM �����Ͻ��� E, �����Ͻ��� F " & _
                "WHERE E.ĩ�� = 1 AND E.�ϼ���� = F.��� AND E.�ϼ����=[1]"
            
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(tvw.SelectedItem.Key, 2)))
        If rs.BOF = False Then Call FillLvw(lvw, rs)
        
    Case "������Ŀ"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = "SELECT B.ID,B.����,b.����,c.���� As ��� FROM ���������� A,������ĿĿ¼ B,������Ŀ��� c WHERE A.������Ŀid=B.ID And A.������=[1] And c.����=b.���"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call FillGrid(vsfItem, rs)
        
            Call AppendRows(vsfItem, lnX1, lnY1)
        End If
    Case "�����Ͻ���"
    
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = "select �ο����� from �����Ͻ��� WHERE ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        
        If rs.BOF = False Then rtb.Text = zlCommFun.NVL(rs("�ο�����").Value)
        
    Case "����������"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        Call ResetVsf(vsf)
        Call AppendRows(vsf, lnX, lnY)
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = "SELECT b.ID,a.������ As ����,b.������ As ��Ŀ,a.��ϵʽ As ����,a.����ֵ As ��Ŀֵ,a.�Ա�,a.��ʼ����,a.�������� from ���������� A,����������Ŀ B WHERE A.��Ŀid=B.ID AND A.������=[1] ORDER BY A.������"
                   
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf, rs)
            Call AppendRows(vsf, lnX, lnY)
        End If
        
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ݱ༭/����
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
                
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    Select Case strMenuItem
    Case "���ӷ���"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmDiagnoseAdviceClass.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
        
    Case "�޸ķ���"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
                        
        If Not frmDiagnoseAdviceClass.ShowEdit(Me, Val(Mid(tvw.SelectedItem.Key, 2)), Val(Mid(tvw.SelectedItem.Parent.Key, 2))) Then Exit Function
        
        
    Case "ɾ������"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
        
        If MsgBox("�����Ҫɾ����" & tvw.SelectedItem.Text & "�����ࣿ" & vbCrLf & "ɾ������ͬʱҲɾ����Ӧ�������Ͻ��顣", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(tvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_�����Ͻ���_DELETE(" & lngKey & ")"
        
    Case "�������"
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmDiagnoseAdviceEdit.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "�޸����"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmDiagnoseAdviceEdit.ShowEdit(Me, lngKey, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "ɾ�����"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        If MsgBox("�����Ҫɾ����" & lvw.SelectedItem.Text & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_�����Ͻ���_DELETE(" & lngKey & ")"
        
    Case "�ο�����"
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmDiagnoseAdviceContent.ShowEdit(Me, lngKey) Then Exit Function
    
    Case "��������"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmDiagnoseAdviceEvaluate.ShowEdit(Me, lngKey) Then Exit Function
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "ɾ������"
        
        If Not (tvw.SelectedItem Is Nothing) Then tvw.Nodes.Remove tvw.SelectedItem.Index
        
        Call ClearData("�����Ͻ���")
        Call RefreshData("�����Ͻ���")
        If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("�����ϸ���")
                
        
    Case "ɾ�����"
    
        'ɾ����
        lngLoop = lvw.SelectedItem.Index
        lvw.ListItems.Remove lngLoop
        Call NextLvwPos(lvw, lngLoop)
        
    Case "�ο�����", "��������"
        Call tbs_Click
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
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��ӡ����
    '������ bytMode                         ��ӡ��ʽ��1-��ӡ��2-Ԥ����3-�����Excel��
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrintLvw
        
    If tvw.SelectedItem Is Nothing Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If UserInfo.���� = "" Then Call GetUserInfo

    objPrint.Title.Text = "�����Ͻ����嵥"
    
    objPrint.UnderAppItems.Add "���ࣺ" & tvw.SelectedItem.Text
    objPrint.UnderAppItems.Add ""
        
    Set objPrint.Body.objData = lvw

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrViewLvw(objPrint, bytMode)
        
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
    
    Call mnuViewIcon_Click(lvw.View)
    Call mnuViewRefresh_Click
    
    mblnStartUp = False
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
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then
        imgY_S.Left = Me.ScaleWidth - 1000
    End If
    
    With tvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With imgY_S
        .Top = tvw.Top
        .Height = tvw.Height
    End With
    
    With lvw
        .Left = imgY_S.Left + imgY_S.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = lvw.Left
        .Width = lvw.Width
    End With
    
    With tbs
        .Left = imgX_S.Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = imgX_S.Width
    End With
    
    With rtb
        .Left = lvw.Left
        .Top = tbs.Top + tbs.Height + 30
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
    With vsf
        .Left = rtb.Left
        .Top = rtb.Top
        .Width = rtb.Width
        .Height = rtb.Height
    End With

    With vsfItem
        .Left = rtb.Left
        .Top = rtb.Top
        .Width = rtb.Width
        .Height = rtb.Height
    End With
    
    Call AppendRows(vsf, lnX, lnY)
    Call AppendRows(vsfItem, lnX1, lnY1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = mblnStartUp
    If Cancel Then Exit Sub
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������", mstrVsf)
                
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000
    
            
    Form_Resize
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 1500 Then imgY_S.Left = 1500
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        
    Call zlControl.LvwSortColumn(lvw, ColumnHeader.Index)
    
End Sub

Private Sub lvw_DblClick()
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    Call ClearData("�����ϸ���")
    Call tbs_Click
    
    '�ָ�
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    mbytPopMenu = 2
    Set mobjPopMenu = New clsPopMenu
    mobjPopMenu.ShowPopupMenuByCursor

End Sub

Private Sub mnuEditAdvice_Click()
    Call MenuClick("�ο�����")
End Sub

Private Sub mnuEditClassAdd_Click()
    Call MenuClick("���ӷ���")
End Sub

Private Sub mnuEditClassDelete_Click()
    Call MenuClick("ɾ������")
End Sub

Private Sub mnuEditClassModify_Click()
    Call MenuClick("�޸ķ���")
End Sub

Private Sub mnuEditCondition_Click()
    Call MenuClick("��������")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("ɾ�����")
End Sub

Private Sub mnuEditModify_Click()
    Call MenuClick("�޸����")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("�������")
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
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

Private Sub mnuViewIcon_Click(Index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    Dim strKeyClass As String
        
    '����������ͷ��ࡢ�������
    If Not (tvw.SelectedItem Is Nothing) Then strKeyClass = tvw.SelectedItem.Key
    strKey = zlControl.LvwSaveItem(lvw)
            
    Call ClearData("�����Ϸ���;������Ŀ¼;�����ϸ���")
    
    Call RefreshData("�����Ϸ���")
    
    '�ָ�ˢ��ǰѡ���������ͷ���
    
    tvw.Nodes(1).Selected = True
    tvw.Nodes(1).Expanded = True
    
    On Error Resume Next
    tvw.Nodes(strKeyClass).Selected = True
    tvw.Nodes(strKeyClass).EnsureVisible
    On Error GoTo 0
    
    If Not (tvw.SelectedItem Is Nothing) Then
        Call RefreshData("������Ŀ¼")
        
        '�ָ�ˢ��ǰѡ����������
        Call zlControl.LvwRestoreItem(lvw, strKey)
        
        Call tbs_Click
        
    End If
    
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

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditClassAdd.Visible Then mobjPopMenu.Add 1, mnuEditClassAdd.Caption, , , mnuEditClassAdd.Enabled
        If mnuEditClassModify.Visible Then mobjPopMenu.Add 2, mnuEditClassModify.Caption, , , mnuEditClassModify.Enabled
        If mnuEditClassDelete.Visible Then mobjPopMenu.Add 3, mnuEditClassDelete.Caption, , , mnuEditClassDelete.Enabled
        
    Case 2
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditAdd.Visible Then mobjPopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditModify.Visible Then mobjPopMenu.Add 2, mnuEditModify.Caption, , , mnuEditModify.Enabled
        If mnuEditDelete.Visible Then mobjPopMenu.Add 3, mnuEditDelete.Caption, , , mnuEditDelete.Enabled
            
        mobjPopMenu.Add 4, "-", , 2, True
        
        If mnuEditAdvice.Visible Then mobjPopMenu.Add 5, mnuEditAdvice.Caption, , , mnuEditAdvice.Enabled
        If mnuEditCondition.Visible Then mobjPopMenu.Add 6, mnuEditCondition.Caption, , , mnuEditCondition.Enabled
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditClassAdd_Click
        Case 2
            Call mnuEditClassModify_Click
        Case 3
            Call mnuEditClassDelete_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
        Case 5
            Call mnuEditAdvice_Click
        Case 6
            Call mnuEditCondition_Click
        End Select
    End Select
End Sub


Private Sub rtb_DblClick()
    If mnuEdit.Visible And mnuEditAdvice.Visible And mnuEditAdvice.Enabled Then
        Call mnuEditAdvice_Click
    End If
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call rtb_DblClick
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        
        Call mnuFilePrint_Click
        
    Case "����"
                
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "����"
        Call mnuEditAdd_Click
    Case "�޸�"
        Call mnuEditModify_Click
    Case "ɾ��"
        Call mnuEditDelete_Click
    Case "�б�"
        Call mnuViewIcon_Click(IIf(lvw.View = 3, 0, lvw.View + 1))
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
    Case "Large"
        Call mnuViewIcon_Click(0)
    Case "Small"
        Call mnuViewIcon_Click(1)
    Case "List"
        Call mnuViewIcon_Click(2)
    Case "Detail"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tbs_Click()

    If lvw.SelectedItem Is Nothing Then Exit Sub
    
'    rtb.Visible = True
'    vsf.Visible = True
'    vsf.Visible = True
    Select Case tbs.SelectedItem.Key
    Case "�ο�����"
        rtb.Visible = True
        vsf.Visible = False
        vsfItem.Visible = False
        
        Call RefreshData("�����Ͻ���")

        
    Case "��������"
        
        rtb.Visible = False
        vsf.Visible = True
        vsfItem.Visible = False
        Call RefreshData("����������")

    
        
    Case "������Ŀ"
        
        rtb.Visible = False
        vsf.Visible = False
        vsfItem.Visible = True
        
        Call RefreshData("������Ŀ")

        
    End Select
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        mobjPopMenu.ShowPopupMenuByCursor
        
    End If
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call ClearData("������Ŀ¼;�����ϸ���")
    
    Call RefreshData("������Ŀ¼")
    Call tbs_Click
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_DblClick()
    If mnuEdit.Visible And mnuEditAdvice.Visible And mnuEditCondition.Enabled Then
        Call mnuEditCondition_Click
    End If
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.����
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick
    End If
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.�ǽ���
End Sub

Private Sub vsfItem_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsfItem, lnX1, lnY1)
End Sub

Private Sub vsfItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsfItem, lnX1, lnY1)
End Sub

Private Sub vsfItem_GotFocus()
    vsfItem.BackColorSel = COLOR.����
End Sub

Private Sub vsfItem_LostFocus()
    vsfItem.BackColorSel = COLOR.�ǽ���
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

