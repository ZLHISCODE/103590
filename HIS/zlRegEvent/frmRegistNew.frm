VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRegistNew 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҺŹ���"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9120
   Icon            =   "frmRegistNew.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9120
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9000
         _ExtentX        =   15875
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
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Һ�"
               Key             =   "Add"
               Description     =   "�Һ�"
               Object.ToolTipText     =   "����ҺŴ���"
               Object.Tag             =   "�Һ�"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˺�"
               Key             =   "Del"
               Description     =   "�˺�"
               Object.ToolTipText     =   "�Ե�ǰѡ�е����˺�"
               Object.Tag             =   "�˺�"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DelBook"
                     Object.Tag             =   "�˲�����"
                     Text            =   "�˲�����"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DelExtra"
                     Object.Tag             =   "�˸��ӷ�"
                     Text            =   "�˸��ӷ�"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fun_1"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ԤԼ"
               Key             =   "ԤԼ"
               Description     =   "ԤԼ"
               Object.ToolTipText     =   "ԤԼ�Һ�"
               Object.Tag             =   "ԤԼ"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����ԤԼ"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "ȡ��"
               Description     =   "ȡ��"
               Object.ToolTipText     =   "ȡ��ԤԼ"
               Object.Tag             =   "ȡ��"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fun_2"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "���¶����������ļ�¼"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ����ǰ�б������������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�շ�����"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��չ"
               Key             =   "Extra"
               ImageIndex      =   14
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExtraItem"
                     Object.Tag             =   "����"
                     Text            =   "����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRegistNew.frx":014A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11007
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
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7455
      Top             =   390
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
            Picture         =   "frmRegistNew.frx":09DE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":0BF8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":0E12
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":150C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":1C06
            Key             =   "ԤԼ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":2300
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":29FA
            Key             =   "ȡ��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":30F4
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":37EE
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":3A08
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":3C22
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":3E3C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":4056
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":D9ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6870
      Top             =   390
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
            Picture         =   "frmRegistNew.frx":E167
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":E381
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":E59B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":EC95
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":F38F
            Key             =   "ԤԼ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":FA89
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":10183
            Key             =   "ȡ��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":1087D
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":10F77
            Key             =   "Go"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":11191
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":113AB
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":115C5
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":117DF
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegistNew.frx":11ED9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsThis 
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   9015
      _cx             =   15901
      _cy             =   7858
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
   Begin MSComctlLib.TabStrip tbsType 
      Height          =   720
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   1270
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2290
      TabFixedHeight  =   564
      HotTracking     =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�Һ��嵥(&1)"
            Key             =   "�Һ�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ԤԼ�嵥(&2)"
            Key             =   "ԤԼ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ԤԼ������(&3)"
            Key             =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "�ֽ�㳮(&E)"
         Shortcut        =   {F9}
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
         Caption         =   "���˹Һ�(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "�����˺�"
         Begin VB.Menu mnuEdit_Del 
            Caption         =   "�����˺�(&D)"
            Shortcut        =   {DEL}
         End
         Begin VB.Menu mnuEdit_DelBook 
            Caption         =   "�˲�����(&B)"
         End
         Begin VB.Menu mnuEdit_DelExtra 
            Caption         =   "�˸��ӷ�(&E)"
         End
      End
      Begin VB.Menu mnuEdit_21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_BindPatNum 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_BatchChangeNum 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuEdit_22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Bespeak 
         Caption         =   "ԤԼ�Һ�(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEdit_Incept 
         Caption         =   "����ԤԼ(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEdit_CancelAuditing 
         Caption         =   "�˺����(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEdit_Cancel 
         Caption         =   "ȡ��ԤԼ(&C)"
      End
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "���ԤԼ(&R)"
      End
      Begin VB.Menu mnuEdit_Defer 
         Caption         =   "ԤԼ����(&F)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "���ĵ���(&V)"
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "�ش�Ʊ��(&P)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "����Ʊ��(&S)"
      End
      Begin VB.Menu mnuEdit_Print_Slip 
         Caption         =   "��ӡ�Һ�ƾ��(&I)"
      End
      Begin VB.Menu mnuEdit_Print_Case 
         Caption         =   "��ӡ������ǩ(&Q)"
      End
      Begin VB.Menu mnuEdit_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Extra 
         Caption         =   "��չ"
         Begin VB.Menu mnuEdit_ExtraItem 
            Caption         =   "����"
            Index           =   0
         End
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
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
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
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Checked         =   -1  'True
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewReFlash 
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
Attribute VB_Name = "frmRegistNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    DeptID As Long
    Patientid As Long
    Doctor As String
    Operator As String
    FeeType As String   '�ѱ�
    ItemType As String '����
    PatiName As String
End Type
Private SQLCondition As Type_SQLCondition
Private mrsList As ADODB.Recordset  '�����б�
Private mstrVsType As String
Private mstrFilter As String
Private mbytCancel As Byte
Private mstr���ӷ� As String, mstr������ĿID As String
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNOMoved As Boolean
Private mstrColWidth As String
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

'�����һ��������������
Private Enum AcitonType
    t_��ͨ
    t_ʱ��
End Enum
'ģ�����
Private Type Ty_ModulePara
    lngN��ȡ��ԤԼ          As Long    'ԤԼN���ڲ���ȡ��ԤԼ
    bln�˺����             As Boolean '��N����ȡ��ԤԼ �Ƿ���Ҫͨ�����
    blnReuseRegNo           As Boolean '�����������Һ�
End Type
Private mTy_Para     As Ty_ModulePara
Private mactionType  As AcitonType
'�˿���ش���
Private mstrPassWord As String
Private mcolCardPayMode As Collection
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    If InStr(mstrPrivs, ";LED������;") = 0 Then gblnLED = False
End Sub

Private Sub mnuEdit_BatchChangeNum_Click()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ź���
    '����:����
    '����:2011-08-24 10:42:19
    '�����:45507
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim datNow As Date, strMsgResult As String
    Err.Clear
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "ϵͳ����" & Format(gdatRegistTime, "yyyy-mm-dd") & _
                                            "�����°������Ű�ģʽ�Һ�,�����ԤԼʱ��ѡ��ģʽ:" & vbCrLf & _
                                            "�ƻ��Ű�ԤԼ(��):" & Format(datNow, "yyyy-mm-dd hh:mm:ss") & "��" & Format(gdatRegistTime - 1, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                                            "������Ű�ԤԼ(��):" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & "���Ժ�", "�ƻ��Ű�ԤԼ,������Ű�ԤԼ,ȡ��", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "ȡ��" Then Exit Sub
        If strMsgResult = "�ƻ��Ű�ԤԼ" Then
            frmBatchChangeNum.Show 1
        End If
        If strMsgResult = "������Ű�ԤԼ" Then
            frmBatchChangeNumNew.Show 1
        End If
    Else
        frmBatchChangeNumNew.Show 1
    End If
End Sub

Private Sub mnuEdit_Bespeak_Click()
    On Error Resume Next
    Dim datNow As Date, strMsgResult As String
    Err.Clear
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "ϵͳ����" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & _
                                            "�����°������Ű�ģʽ�Һ�,�����ԤԼʱ��ѡ��ģʽ:" & vbCrLf & _
                                            "�ƻ��Ű�ԤԼ(��):" & Format(datNow, "yyyy-mm-dd hh:mm:ss") & "��" & Format(gdatRegistTime - 1, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                                            "������Ű�ԤԼ(��):" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & "���Ժ�", "�ƻ��Ű�ԤԼ,������Ű�ԤԼ,ȡ��", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "ȡ��" Then Exit Sub
        If strMsgResult = "�ƻ��Ű�ԤԼ" Then
            If gbln������� Then
                frmRegistEditSimple.mlngModul = mlngModul
                frmRegistEditSimple.mstrPrivs = mstrPrivs
                frmRegistEditSimple.mbytMode = 1
                frmRegistEditSimple.mbytInState = 0
                Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
                frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
            Else
                frmRegistEdit.mlngModul = mlngModul
                frmRegistEdit.mstrPrivs = mstrPrivs
                frmRegistEdit.mbytMode = 1
                frmRegistEdit.mbytInState = 0
                Set frmRegistEdit.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
                frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
            End If
        End If
        If strMsgResult = "������Ű�ԤԼ" Then
            frmRegistEditNew.mlngModul = mlngModul
            frmRegistEditNew.mstrPrivs = mstrPrivs
            frmRegistEditNew.mbytMode = 1
            frmRegistEditNew.mbytInState = 0
            Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
            frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
    Else
        frmRegistEditNew.mlngModul = mlngModul
        frmRegistEditNew.mstrPrivs = mstrPrivs
        frmRegistEditNew.mbytMode = 1
        frmRegistEditNew.mbytInState = 0
        Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If gblnOk And tbsType.SelectedItem.Key <> "�Һ�" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub mnuEdit_BindPatNum_Click()
  ' ������� �󶨴��� ��������ŵİ�
  frmBindPatientNo.Show 1, Me
End Sub

Private Sub mnuEdit_Del_Click()
    Call DeleteRegist
End Sub

Private Sub mnuEdit_DelBook_Click()
    Dim strNO As String
    Dim str�Һ�ʱ�� As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))

    If strNO = "" Then
        MsgBox "��ǰû�м�¼�����˲����ѣ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬����������˲����Ѳ�����", vbInformation, gstrSysName
        Exit Sub
    End If

    If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then
        '�жϵ�ǰ��Ա�Ե����Ƿ��в���Ȩ��,ʱ������,������Һŵ���Ч����
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
                              CDate(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("�Һ�ʱ��"))), "�˺�") Then Exit Sub
    End If

    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If

    If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs, False, 1) = False Then Exit Sub
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub LoadPlugInMnu()
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim blnHave As Boolean, blnTool As Boolean
    Dim strTemp As String
    Dim intToolCounter As Integer
    
    If CreatePlugInOK(mlngModul) Then
        blnHave = True
    End If
    
    mnuEdit_Extra.Visible = blnHave
    tbr.Buttons("Extra").Visible = blnHave
    
    If blnHave Then
        blnTool = False
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, mlngModul, 3)
        Call zlPlugInErrH(Err, "GetFuncNames")
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then
            mnuEdit_Extra.Visible = False
            tbr.Buttons("Extra").Visible = False
            Exit Sub
        End If
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        intToolCounter = 0
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuEdit_ExtraItem(i)
            End If
            mnuEdit_ExtraItem(i).Caption = Replace(CStr(arrTmp(i)), "InTool:", "")
            mnuEdit_ExtraItem(i).Tag = Replace(CStr(arrTmp(i)), "InTool:", "")
            
            If InStr(CStr(arrTmp(i)), "InTool:") > 0 Then
                strTemp = Split(CStr(arrTmp(i)), ":")(1)
                blnTool = True
                If intToolCounter <> 0 Then
                    tbr.Buttons("Extra").ButtonMenus.Add tbr.Buttons("Extra").ButtonMenus.Count + 1, strTemp, strTemp
                    intToolCounter = intToolCounter + 1
                End If
                tbr.Buttons("Extra").ButtonMenus(tbr.Buttons("Extra").ButtonMenus.Count).Text = strTemp
                tbr.Buttons("Extra").ButtonMenus(tbr.Buttons("Extra").ButtonMenus.Count).Tag = strTemp
            End If
        Next
        tbr.Buttons("Extra").Visible = blnTool
    End If
End Sub

Private Sub mnuEdit_ExtraItem_Click(index As Integer)
    Call ExcPlugInFun(mnuEdit_ExtraItem(index).Tag)
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lngPatiID As Long
    Dim strNO As String
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    
    If strNO = "" Or strNO = "���ݺ�" Then
        MsgBox "δѡ���κε��ݣ�����ִ�д˲�����", vbExclamation, gstrSysName: Exit Sub
    End If
        
    If CreatePlugInOK(mlngModul) Then
        lngPatiID = Val(Me.vsThis.TextMatrix(vsThis.Row, getColNum("����ID")))
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, mlngModul, strFunName, lngPatiID, strNO, 0, "", 3)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    With ButtonMenu
        Select Case .Key
            Case "DelBook"
                mnuEdit_DelBook_Click
            Case "DelExtra"
                mnuEdit_DelExtra_Click
            Case Else
                Call ExcPlugInFun(.Tag)
        End Select
    End With
End Sub


Private Sub mnuEdit_DelExtra_Click()
    Dim strNO As String, str�Һ�ʱ�� As String
    '�˸��ӷ�
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))

    If strNO = "" Then
        MsgBox "��ǰû�м�¼������" & mstr���ӷ� & "��", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬�����������" & mstr���ӷ� & "������", vbInformation, gstrSysName
        Exit Sub
    End If

    If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then
        '�жϵ�ǰ��Ա�Ե����Ƿ��в���Ȩ��,ʱ������,������Һŵ���Ч����
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
                              CDate(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("�Һ�ʱ��"))), "�˺�") Then Exit Sub
    End If

    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If

    If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs, False, 2) = False Then Exit Sub
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub CancelOldRegist()
    Dim strSQL As String, strNO As String
    Dim Datsys As Date
    Dim datTmp As Date
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�йҺ�ԤԼ����ȡ����", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
        CDate(vsThis.TextMatrix(vsThis.Row, getColNum("�Ǽ�ʱ��"))), "ȡ��ԤԼ") Then Exit Sub
    If mbytCancel <> 1 Then
        If vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) <> "1" Then
            MsgBox "��ǰ�Һ�ԤԼ�Ѿ�ȡ����", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "��ǰ�Һ�ԤԼ�Ѿ����ա�", vbExclamation, gstrSysName
        Exit Sub
    End If
    If tbsType.SelectedItem.Key <> "�Һ�" And mTy_Para.bln�˺���� And mTy_Para.lngN��ȡ��ԤԼ > 0 Then
        '�˺���� �����շ�ԤԼ��ԤԼ���յ��˺�
        If vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("�˺������")) = "" Then
                '�Ƿ�ԤԼ�жϷŵ����� ����Ӱ������
                 Datsys = zlDatabase.Currentdate
                 datTmp = DateAdd("d", -1 * mTy_Para.lngN��ȡ��ԤԼ, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��"))))
                   'ԤԼʱ��-K >datSys
                   If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                           MsgBox "���ݺ�Ϊ" & strNO & "���շ�ԤԼ����û�о����˺����!���ܽ����˺�!", vbInformation, Me.Caption
                           Exit Sub
                   End If
            End If
        End If
        
    End If
    If gbln������� Then
        frmRegistEditSimple.mlngModul = mlngModul
        frmRegistEditSimple.mstrPrivs = mstrPrivs
        frmRegistEditSimple.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
        frmRegistEditSimple.mblnNOMoved = mblnNOMoved
        frmRegistEditSimple.mbytMode = 3
        frmRegistEditSimple.mbytInState = 1
        Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        frmRegistEdit.mlngModul = mlngModul
        frmRegistEdit.mstrPrivs = mstrPrivs
        frmRegistEdit.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
        frmRegistEdit.mblnNOMoved = mblnNOMoved
        frmRegistEdit.mbytMode = 3
        frmRegistEdit.mbytInState = 1
        Set frmRegistEdit.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If mnuViewRefeshOptionItem(1).Checked And gblnOk Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked And gblnOk Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuEdit_Cancel_Click()
    Dim strSQL As String, strNO As String
    Dim Datsys As Date
    Dim datTmp As Date
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�йҺ�ԤԼ����ȡ����", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If IsNewModeRegist(strNO) = False Then
        Call CancelOldRegist
        Exit Sub
    End If
    
    If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
        CDate(vsThis.TextMatrix(vsThis.Row, getColNum("�Ǽ�ʱ��"))), "ȡ��ԤԼ") Then Exit Sub
    If mbytCancel <> 1 Then
        If vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) <> "1" Then
            MsgBox "��ǰ�Һ�ԤԼ�Ѿ�ȡ����", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "��ǰ�Һ�ԤԼ�Ѿ����ա�", vbExclamation, gstrSysName
        Exit Sub
    End If
    If tbsType.SelectedItem.Key <> "�Һ�" And mTy_Para.bln�˺���� And mTy_Para.lngN��ȡ��ԤԼ > 0 Then
        '�˺���� �����շ�ԤԼ��ԤԼ���յ��˺�
        If vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("�˺������")) = "" Then
                '�Ƿ�ԤԼ�жϷŵ����� ����Ӱ������
                 Datsys = zlDatabase.Currentdate
                 datTmp = DateAdd("d", -1 * mTy_Para.lngN��ȡ��ԤԼ, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��"))))
                   'ԤԼʱ��-K >datSys
                   If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                           MsgBox "���ݺ�Ϊ" & strNO & "���շ�ԤԼ����û�о����˺����!���ܽ����˺�!", vbInformation, Me.Caption
                           Exit Sub
                   End If
            End If
        End If
        
    End If
    
    frmRegistEditNew.mlngModul = mlngModul
    frmRegistEditNew.mstrPrivs = mstrPrivs
    frmRegistEditNew.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    frmRegistEditNew.mblnNOMoved = mblnNOMoved
    frmRegistEditNew.mbytMode = 3
    frmRegistEditNew.mbytInState = 1
    Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
    frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If mnuViewRefeshOptionItem(1).Checked And gblnOk Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked And gblnOk Then
        mnuViewReFlash_Click
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

 

Private Sub mnuEdit_CancelAuditing_Click()
    '�˺����
    Dim strSQL As String, strNO As String
    If vsThis.Rows <= 1 Then Exit Sub
    If InStr(1, mstrPrivs, ";�˺����;") = 0 Then
        MsgBox "��û�ж�ԤԼ�Ž����˺���˵�Ȩ�ޡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�йҺ�ԤԼ���Խ����˺���ˡ�", vbExclamation, gstrSysName
        Exit Sub
    End If

    If mbytCancel <> 1 Then
        If vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) <> "1" Then
            MsgBox "��ǰ�Һ�ԤԼ�Ѿ�ȡ����", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    Select Case tbsType.SelectedItem.Key
    Case "ԤԼ", "����":
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
            CDate(vsThis.TextMatrix(vsThis.Row, getColNum("�Ǽ�ʱ��"))), "ȡ��ԤԼ") Then Exit Sub
        
    Case "�Һ�":
        If vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")) = "" Then Exit Sub
        If vsThis.TextMatrix(vsThis.Row, getColNum("�˺������")) <> "" Then Exit Sub
    Case Else:
        Exit Sub
    End Select
    
   
     If MsgBox("ȷʵҪ������[" & strNO & "]����ȡ���˺������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
           '  Zl_����ԤԼ�Һ�_Cancelauditing
   strSQL = "Zl_����ԤԼ�Һ�_Cancelauditing("
           '  No_In       ���˹Һż�¼.NO%Type,
   strSQL = strSQL & "'" & strNO & "',"
           '  ����Ա_In   ���˹Һż�¼.�˺������%Type,
   strSQL = strSQL & "'" & UserInfo.���� & "',"
           '  ���ʱ��_In ���˹Һż�¼.�˺����ʱ��%Type
    strSQL = strSQL & "Sysdate)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    mnuViewReFlash_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Clear_Click()
    If MsgBox("�ò����������� " & gintԤԼ���� & " ���ڵǼǣ���ԤԼʱ���ѹ��ڵ�ԤԼ��¼��Ҫ������" & _
        vbCrLf & vbCrLf & "˵����Ϊ��֤��Ч�����Щ���ڵ�ԤԼ��¼������Ҫ����ִ�иù��ܡ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure("zl_����ԤԼ�Һ�_Clear", Me.Caption)
    On Error GoTo 0
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("���������ִ����ϡ��嵥���ݿ����Ѹ��ģ�Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        MsgBox "���������ִ����ϡ�", vbInformation, gstrSysName
        mnuViewReFlash_Click
    Else
        MsgBox "���������ִ����ϡ�", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Defer_Click()
    Dim datNow As Date, strMsgResult As String
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        strMsgResult = zlCommFun.ShowMsgbox(gstrSysName, "ϵͳ����" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & _
                                            "�����°������Ű�ģʽ�Һ�,�����ԤԼʱ��ѡ��ģʽ:" & vbCrLf & _
                                            "�ƻ��Ű�ԤԼ(��):" & Format(datNow, "yyyy-mm-dd hh:mm:ss") & "��" & Format(gdatRegistTime - 1, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
                                            "������Ű�ԤԼ(��):" & Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") & "���Ժ�", "�ƻ��Ű�ԤԼ,������Ű�ԤԼ,ȡ��", Me, vbQuestion)
        If strMsgResult = "" Or strMsgResult = "ȡ��" Then Exit Sub
        If strMsgResult = "�ƻ��Ű�ԤԼ" Then
            frmBookingDefer.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
        If strMsgResult = "������Ű�ԤԼ" Then
            frmBookingDeferNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
    Else
        frmBookingDeferNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
End Sub

Private Sub InceptOldRegist()
    Dim strNO As String
    Dim datTime As Date
    Dim datThis As Date
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�н��յ�ԤԼ�Һš�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "��ǰ�����Ѿ������ա�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    datTime = CDate(vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")))
    datThis = zlDatabase.Currentdate
    If Format(datTime, "YYYY-MM-DD") > Format(datThis, "YYYY-MM-DD") Then
        If MsgBox("��ǰ���յļ�¼���ǵ����ԤԼ��¼���Ƿ�������գ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    If gbln������� Then
        frmRegistEditSimple.mlngModul = mlngModul
        frmRegistEditSimple.mstrPrivs = mstrPrivs
        frmRegistEditSimple.mbytMode = 2
        frmRegistEditSimple.mbytInState = 0
        frmRegistEditSimple.mstrNoIn = strNO
        Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        frmRegistEdit.mlngModul = mlngModul
        frmRegistEdit.mstrPrivs = mstrPrivs
        frmRegistEdit.mbytMode = 2
        frmRegistEdit.mbytInState = 0
        frmRegistEdit.mstrNoIn = strNO
        Set frmRegistEdit.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If gblnOk And tbsType.SelectedItem.Key <> "�Һ�" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Incept_Click()
    Dim strNO As String
    Dim datTime As Date
    Dim datThis As Date
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO = "" Then
        MsgBox "��ǰû�н��յ�ԤԼ�Һš�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If IsNewModeRegist(strNO) = False Then
        Call InceptOldRegist
        Exit Sub
    End If
    
    If CheckRegistAppointment(strNO) = False Then
        MsgBox "��ǰ�����Ѿ������ա�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    datTime = CDate(vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")))
    datThis = zlDatabase.Currentdate
    If Format(datTime, "YYYY-MM-DD") > Format(datThis, "YYYY-MM-DD") Then
        If MsgBox("��ǰ���յļ�¼���ǵ����ԤԼ��¼���Ƿ�������գ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    frmRegistEditNew.mlngModul = mlngModul
    frmRegistEditNew.mstrPrivs = mstrPrivs
    frmRegistEditNew.mbytMode = 2
    frmRegistEditNew.mbytInState = 0
    frmRegistEditNew.mstrNoIn = strNO
    Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
    frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If gblnOk And tbsType.SelectedItem.Key <> "�Һ�" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub
 

Private Sub mnuEdit_Print_Case_Click()
    Dim lng����ID As Long
    lng����ID = Val(Me.vsThis.TextMatrix(vsThis.Row, getColNum("����ID")))
    If lng����ID <> 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_2", Me, "����ID=" & lng����ID, 2)
    Else
        MsgBox "�ùҺŵ���صĲ���û�н������˵���!", vbInformation
    End If
End Sub

Private Sub mnuEdit_Print_Slip_Click()
    Dim strNO As String
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO <> "" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
        '72704:л��,2014-07-23,д��ƾ����ӡ��¼
        gstrSQL = "Zl_ƾ����ӡ��¼_Update(4,'" & strNO & "',1,'" & UserInfo.���� & "')"
        zlDatabase.ExecuteProcedure gstrSQL, ""
    Else
        MsgBox "��ǰû�йҺŻ���ռ�¼��", vbExclamation, gstrSysName
    End If
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    frmLocalPara.mlngModul = mlngModul
    frmLocalPara.mstrPrivs = mstrPrivs
    frmLocalPara.Show 1, Me
    If gblnOk Then InitPara
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub
 

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(index As Integer)
    Dim strNO As String
    
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If strNO <> "" Then
        With vsThis
            Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me, _
                "NO=" & strNO, "Ʊ�ݺ�=" & .TextMatrix(.Row, getColNum("����Ʊ��")), _
                "�ű�=" & .TextMatrix(.Row, getColNum("�ű�")), "ҽ��=" & .TextMatrix(.Row, getColNum("ҽ��")), _
                "�����=" & .TextMatrix(.Row, getColNum("�����")))
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(index).Tag, ",")(0), Split(mnuReportItem(index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    '��λ���˷�Χ
    frmRegistFilter.mlngModule = mlngModul
    frmRegistFilter.bytType = tbsType.SelectedItem.index - 1
    
    '�б���ʾ��ʽ,���շ���ʱ����ʾ ���ǰ��յǼ�ʱ����ʾ��
    'frmRegistFilter.mblnFilterType = mTy_Para.bln�Һ��б����
    frmRegistFilter.Show 1, Me
    If gblnOk Then
        mstrFilter = frmRegistFilter.mstrFilter
        mbytCancel = IIf(frmRegistFilter.optRegistRecord(0).Value = True, 1, IIf(frmRegistFilter.optRegistRecord(1).Value = True, 2, 3))
        
        With SQLCondition
            .DateB = frmRegistFilter.dtpBegin.Value
            .DateE = frmRegistFilter.dtpEnd.Value
            .NOB = frmRegistFilter.txtNOBegin.Text
            .NOE = frmRegistFilter.txtNOEnd.Text
            .FactB = frmRegistFilter.txtFactBegin.Text
            .FactE = frmRegistFilter.txtFactEnd.Text
            If frmRegistFilter.cbo����.ListIndex > 0 Then .DeptID = frmRegistFilter.cbo����.ItemData(frmRegistFilter.cbo����.ListIndex)
            .PatiName = gstrLike & frmRegistFilter.txtPatient.Text & "%"
            .Patientid = frmRegistFilter.mlngPrePatient
            .Doctor = frmRegistFilter.txtҽ��.Text & "%"
            If frmRegistFilter.cbo����Ա.ListIndex > 0 Then .Operator = NeedName(frmRegistFilter.cbo����Ա.Text)
            If frmRegistFilter.cbo�ѱ�.ListIndex > 0 Then .FeeType = NeedName(frmRegistFilter.cbo�ѱ�.Text)
            If frmRegistFilter.cbo����.ListIndex > 0 Then .ItemType = frmRegistFilter.cbo����.Text
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub vsThis_AfterMoveColumn(ByVal Col As Long, Position As Long)
     zl_vsGrid_Para_Save mlngModul, vsThis, Me.Caption, Me.tbsType.SelectedItem.Key, False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngWidth As Long
    Dim i As Integer
    If mstrColWidth <> "" Then
        lngWidth = vsThis.ColWidth(Col)
        For i = 0 To UBound(Split(mstrColWidth, "|"))
            vsThis.ColWidth(i) = Split(mstrColWidth, "|")(i)
        Next i
        vsThis.ColWidth(Col) = lngWidth
    End If
    zl_vsGrid_Para_Save mlngModul, vsThis, Me.Caption, Me.tbsType.SelectedItem.Key, False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsThis_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Integer
    mstrColWidth = ""
    For i = 0 To vsThis.Cols - 1
        mstrColWidth = mstrColWidth & "|" & vsThis.ColWidth(i)
    Next i
    If mstrColWidth <> "" Then mstrColWidth = Mid(mstrColWidth, 2)
End Sub

Private Sub vsThis_DblClick()
    Dim lngCols As Long
    Dim lngRow As Long
    If vsThis.MouseRow <= 0 Then Exit Sub
    If vsThis.Row <= 0 Then Exit Sub
     lngCols = getColNum("��¼״̬")
     lngRow = vsThis.Row
    
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub vsThis_EnterCell()
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    If vsThis.MouseRow <= 0 Then Exit Sub
    If Mid(stbThis.Panels(2).Text, 1, 2) = "ժҪ" Then stbThis.Panels(2).Text = ""
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If vsThis.Row = 0 Or strNO = "" Then Exit Sub
    
    mlngGo = vsThis.Row
    mlngCurRow = vsThis.Row: mlngTopRow = vsThis.TopRow
    
    If tbsType.SelectedItem.Key <> "�Һ�" Then
        stbThis.Panels(2).Text = "ժҪ:" & vsThis.TextMatrix(vsThis.Row, getColNum("ժҪ"))
    End If
    Call SetMenuEnable
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsThis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then
        Call mnuEdit_Del_Click
        Exit Sub
    End If
    If (tbsType.SelectedItem.Key = "ԤԼ" Or tbsType.SelectedItem.Key = "����") And KeyCode = vbKeyDelete And mnuEdit_Cancel.Enabled = True And mnuEdit_Cancel.Visible Then Call mnuEdit_Cancel_Click
End Sub

Private Sub vsThis_RowColChange()
    Call SetMenuEnable
End Sub
Private Sub SetMenuEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò˵���Enable����
    '����:���˺�
    '����:2013-11-05 16:03:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��¼״̬ As Long, lng����ID As Long
    Dim bln���� As Boolean, blnδ���� As Boolean
    Dim blnEnabled As Boolean, bln���� As Boolean, bln���� As Boolean
    Dim strStatus As String, rsStatus As ADODB.Recordset
    
    With vsThis
      If .Row <= 0 Then
            blnEnabled = False
            mnuEdit_Cancel.Enabled = blnEnabled
            tbr.Buttons("ȡ��").Enabled = blnEnabled
            mnuEdit_Del.Enabled = blnEnabled
            mnuEdit_DelBook.Enabled = blnEnabled
            mnuEdit_DelExtra.Enabled = blnEnabled
            tbr.Buttons("Del").Enabled = blnEnabled
             mnuEdit_CancelAuditing.Enabled = blnEnabled
             Exit Sub
      End If
      
      lng��¼״̬ = 0
      If getColNum("��¼״̬") <> -1 Then
            lng��¼״̬ = Val(.TextMatrix(.Row, getColNum("��¼״̬")))
      End If
      
      If getColNum("��¼״̬") <> -1 Then
        lng����ID = Val(.TextMatrix(.Row, getColNum("����ID")))
      End If
      
      If getColNum("���ݺ�") <> -1 Then
            strStatus = "Select Sum(����) As ����, Sum(����) As ����, Sum(δ����) As δ����" & vbNewLine & _
                        "From (Select Decode(Nvl(Max(a.Id), 0), 0, 0, 1) As ����, 0 As ����, 0 As δ����" & vbNewLine & _
                        "       From ������ü�¼ A" & vbNewLine & _
                        "       Where a.No(+) = [1] And a.��¼����(+) = 4 And a.��¼״̬(+) = 1 And" & vbNewLine & _
                        "             Instr(',' || [2] || ',' , ',' || a.�շ�ϸĿid(+) || ',') > 0" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select 0 As ����, Decode(Nvl(Max(b.Id), 0), 0, 0, 1) As ����, 0 As δ����" & vbNewLine & _
                        "       From ������ü�¼ B" & vbNewLine & _
                        "       Where b.No(+) = [1] And b.��¼����(+) = 4 And b.��¼״̬(+) = 1 And b.���ӱ�־(+) = 1" & vbNewLine & _
                        "       Union All" & vbNewLine & _
                        "       Select 0 As ����, 0 As ����, Decode(Nvl(Max(c.Id), 0), 0, 0, 1) As δ����" & vbNewLine & _
                        "       From ������ü�¼ C" & vbNewLine & _
                        "       Where c.No(+) = [1] And c.��¼����(+) = 4 And c.��¼״̬(+) = 1)"
            Set rsStatus = zlDatabase.OpenSQLRecord(strStatus, Me.Caption, .TextMatrix(.Row, getColNum("���ݺ�")), mstr������ĿID)
      End If
      
      mnuEdit_CancelAuditing.Enabled = lng��¼״̬ = 1 And .TextMatrix(.Row, getColNum("�˺������")) = ""
      
      Select Case tbsType.SelectedItem.Key
      Case "�Һ�"
            If getColNum("���ʷ���") <> -1 Then
                bln���� = Val(.TextMatrix(.Row, getColNum("���ʷ���"))) = 1
            Else
                bln���� = False
            End If
            
            If Not rsStatus.EOF Then
                blnδ���� = Val(Nvl(rsStatus!δ����)) = 1
                bln���� = Val(Nvl(rsStatus!����)) = 1
                bln���� = Val(Nvl(rsStatus!����)) = 1
            Else
                blnδ���� = False
                bln���� = False
                bln���� = False
            End If
            
            '������ش�
            mnuEdit_Print.Enabled = blnδ���� And Not bln����
            blnEnabled = Trim(vsThis.TextMatrix(vsThis.Row, getColNum("����Ʊ��"))) = ""
            mnuEdit_Print_Supplemental.Enabled = blnδ���� And Not bln���� And blnEnabled
            
            mnuEdit_Cancel.Enabled = False
            tbr.Buttons("ȡ��").Enabled = False
            
            blnEnabled = lng��¼״̬ = 1
            mnuEdit_Del.Enabled = blnEnabled
            tbr.Buttons("Del").Enabled = blnEnabled
            
            blnEnabled = mnuEdit_CancelAuditing.Enabled And .TextMatrix(.Row, getColNum("ԤԼʱ��")) <> ""
            mnuEdit_CancelAuditing.Enabled = blnEnabled
            
            If lng��¼״̬ <> 2 Then
                mnuEdit_DelBook.Enabled = .TextMatrix(.Row, getColNum("����")) <> ""
                tbr.Buttons("Del").ButtonMenus("DelBook").Enabled = .TextMatrix(.Row, getColNum("����")) <> ""
            Else
                mnuEdit_DelBook.Enabled = bln����
                tbr.Buttons("Del").ButtonMenus("DelBook").Enabled = bln����
            End If
            
            mnuEdit_DelExtra.Enabled = bln����
            tbr.Buttons("Del").ButtonMenus("DelExtra").Enabled = bln����
      Case Else
            mnuEdit_Cancel.Enabled = lng��¼״̬ = 1
            tbr.Buttons("ȡ��").Enabled = mnuEdit_Cancel.Enabled
            mnuEdit_Del.Enabled = False
            tbr.Buttons("Del").Enabled = False
            
            mnuEdit_Print.Enabled = False
            mnuEdit_Print_Supplemental.Enabled = False
      End Select
   End With
 End Sub

Private Sub vsThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        Call vsThis_EnterCell
    End If
End Sub

Private Function getColNum(strColName As String) As Long
    Dim i As Long
    For i = 0 To vsThis.Cols - 1
        If vsThis.TextMatrix(0, i) = strColName Then
            getColNum = i
            Exit Function
        End If
    Next
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub
Private Function IsCheckCancelValied(ByVal lng�ҺŽ���ID As Long, ByVal lng���ѽ���ID As Long, _
    ByVal cllBillBalance As Collection, ByVal dbl��� As Double, Optional ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�ʱ��������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln���ѿ� As Boolean, lng�����ID As Long
    Dim str��֤����  As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim strXmlIn As String, bln�˿��鿨 As Boolean, strˢ������ As String
    strName = IIf(glngSys \ 100 = 8, "��Ա��", "ҽ�ƿ�")
    If cllBillBalance Is Nothing Then IsCheckCancelValied = True: Exit Function
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
    bln���ѿ� = Val(cllBillBalance(1)(2)) = 1
    lng�����ID = cllBillBalance(1)(0)
    If lng�����ID = 0 Then IsCheckCancelValied = True: Exit Function
    '4.3.3.2.6   zlReturnCheck:�ʻ����˽���ǰ�ļ��
    'zlPaymentCheck�ʻ��ۿ�׼��
    '������  ��������    ��/��   ��ע
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  ģ���
    'lngCardTypeID   Long    In  �����ID:ҽ�ƿ����.ID
    'strCardNo   String  IN  ����
    'strBalanceIDs:��ʽ:�շ�����( 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�)|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    'dblMoney    Double  IN  �˿���
    'strSwapNo   String  In  ������ˮ��(�˿�ʱ���)
    'strSwapMemo String  In  ����˵��(�˿�ʱ����)
    '    Boolean ��������    True:���óɹ�,False:����ʧ��
    '˵��:
    '�ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬��ˣ��ٵ��û��˽���ǰ���Ƚ������ݵĺϷ��Լ��,�Ա�������������
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
    'mcolBillBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
    Dim str���� As String, str������ˮ�� As String, str����˵�� As String, str������Ϣ As String
    Dim strXMLExpend As String
    str���� = cllBillBalance(1)(1)
    str������ˮ�� = cllBillBalance(1)(3)
    str����˵�� = cllBillBalance(1)(4)
    If lng���ѽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||5|" & lng���ѽ���ID
    If lng�ҺŽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||4|" & lng�ҺŽ���ID
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 3)
    
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng�����ID, bln���ѿ�, str����, str������Ϣ, dbl���, str������ˮ��, str����˵��, strXMLExpend) = False Then
        Exit Function
    End If
    
    strSQL = "Select �Ƿ��˿��鿨 From ҽ�ƿ���� Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����ID)
    If rsTmp.EOF Then
        bln�˿��鿨 = False
    Else
        bln�˿��鿨 = Val(Nvl(rsTmp!�Ƿ��˿��鿨)) = 1
    End If
    
    strSQL = "Select ����,�Ա�,���� From ���˹Һż�¼ Where NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    If bln�˿��鿨 Then
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If rsTmp.EOF Then
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng�����ID, bln���ѿ�, _
                "", "", "", dbl���, str����, strˢ������, _
                False, True, False, True, Nothing, False, True, strXmlIn) = False Then Exit Function
        Else
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng�����ID, bln���ѿ�, _
                Nvl(rsTmp!����), Nvl(rsTmp!�Ա�), Nvl(rsTmp!����), dbl���, str����, strˢ������, _
                False, True, False, True, Nothing, False, True, strXmlIn) = False Then Exit Function
        End If
    End If
    
    IsCheckCancelValied = True
End Function

Private Function IsCheckCancel��Ԥ��(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ������ʱʱ��鲡���Ƿ���Ԥ����δ��
    '����:��Ч,����true,���򷵻�False
    '����:������
    '����:2014-04-24
    '�����:62568
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsBill As Recordset, rsCard As Recordset
    
    strSQL = "Select Count(1) As ҽ�ƿ��� From ����ҽ�ƿ���Ϣ Where ״̬=0 And ����ID=[1]"
    Set rsCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    strSQL = _
            "Select Ԥ�����,������� From ������� Where ����=1 And ����=1 And ����ID=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Format(Nvl(rsBill!Ԥ�����, 0) - Nvl(rsBill!�������, 0), "0.00") > 0 Then
        If Val(Nvl(rsCard!ҽ�ƿ���)) = 1 Then
            MsgBox "�ò�������Ԥ������ȥҽ�ƿ����Ź������Ըÿ�����ȡ���󶨲���!", vbInformation, gstrSysName
            IsCheckCancel��Ԥ�� = False
            Exit Function
        End If
    End If
    IsCheckCancel��Ԥ�� = True
End Function

Private Function CheckRegistAppointment(ByVal strNO As String) As Boolean
    '���ԤԼ��¼�Ƿ񱻽���
    'True-ԤԼ��¼δ����;False-ԤԼ��¼�ѱ�����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 1 From ���˹Һż�¼ Where NO = [1] And ����ʱ�� Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If Not rsTmp.EOF Then
        CheckRegistAppointment = True
    Else
        CheckRegistAppointment = False
    End If
End Function

Private Sub DelOldRegist()
    Dim strSQL As String, strNO As String, str����NO As String, strCardNo As String
    Dim intInsure As Integer, lng����ID As Long, lngCard����ID As Long, msgBoxResult As String
    Dim str�Һ�ʱ�� As String, strSQLCard As String, strMessage As String
    Dim str����� As String, strAdvance As String, str�����ʻ� As String
    Dim blnEnableDel As Boolean, blnTrans As Boolean, strSQLBound As String
    Dim bytTogetherDo As Byte, bln�˷��ش� As Boolean, blnPromptClear As Boolean
    Dim rsTmp As ADODB.Recordset, rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset
    Dim objICCard As Object
    Dim cllPro As Collection, cllBillBalance As Collection, dblThreeMoney As Double
    Dim cllUpdate As Collection, cllThreeIns As Collection, strErrMsg As String
    Dim byt�˷ѷ�ʽ As Byte  '0-ȫ�� 1-ֻ�˹Һŷ� 2-ֻ�˲���
    Dim bln������ As Boolean    '�Ƿ����������
    Dim blnCardReprint As Boolean    '���˿��ش�
    Dim Datsys As Date
    Dim datTmp As Date
    Dim lngPatientID As Long
    Dim dblAdvanceMoney As Double    'Ԥ���������
    Dim strInvoice As String, lng����ID As Long, lng����ID As Long
    Dim blnVirtualPrint As Boolean
    Dim bln���� As Boolean, int���� As Integer, bln���� As Boolean
    
    Set cllPro = New Collection
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))

    If strNO = "" Then
        MsgBox "��ǰû�м�¼�����˺ţ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬����������˺Ų�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    str�Һ�ʱ�� = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("�Һ�ʱ��"))
    str����� = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("�����"))
    bln������ = Trim(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("����"))) <> ""
    lngPatientID = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("����ID")))
    
    int���� = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("����")))
    bln���� = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("���ʷ���"))) = 1
    
    
    If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then
        '�жϵ�ǰ��Ա�Ե����Ƿ��в���Ȩ��,ʱ������,������Һŵ���Ч����
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
                             CDate(str�Һ�ʱ��), "�˺�") Then Exit Sub
    End If

    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If

    If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then   '����:
        '���Һŵ��Ƿ���ִ��
        If InStr(";" & mstrPrivs & ";", ";��ҽ�����˺�;") > 0 Then
            blnEnableDel = True
        End If
        If CheckExecuted(strNO, blnEnableDel) Then
            MsgBox "�Һŵ�" & strNO & "�Ѿ���ҽ��������¹�ҽ��,�����˺ţ�", vbInformation, gstrSysName
            Exit Sub
        End If
        'ҽ��վ�ҵĺ�-�շ��ж�
        If CheckPriceHaveFee(strNO, str����NO) Then Exit Sub
        '�Ƿ���������,��δ�˷�
        If InStr(1, mstrPrivs, ";�շѺ��˺�;") = 0 Then
            If ExistFee(strNO) Then
                MsgBox strNO & "�Һŵ��Ĳ����Ѿ������˷���,�����˷Ѳ����˺�.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If str����NO = "" Then
        '�˺�,��ȡ���۵�
        strSQL = "Select NO,��¼״̬ From ������ü�¼ " & _
                " Where ��¼����=1 And No = (Select �շѵ� From ���˹Һż�¼ Where NO=[1] And ��¼����=1 and ��¼״̬=1 and  Rownum<2 )" & _
                " And ��¼״̬ IN(0,1,3) And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", strNO, "%" & strNO & "%")
        If Not rsTmp.EOF Then
            If Nvl(rsTmp!��¼״̬, 0) = 0 Then
                str����NO = Nvl(rsTmp!NO)
            End If
        End If
    End If
    
    '�˺�����ʾ��ϸ��Ϣ
    If gbln������� Then
        If frmRegistEditSimple.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
    Else
        If frmRegistEdit.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
    End If
    GoTo ReFlash    'ˢ��


    'ȥ����ҽ������ƥ����
    lng����ID = GetBill����ID(strNO, 4, lng����ID, bln����)
   If zlCheckIsAllowBackSN(strNO, bln����, bln����) = False Then Exit Sub

    'ҽ���˺ż��
    If bln���� Then
        intInsure = int����
    Else
        intInsure = ExistInsure(strNO)
    End If
    
    Dim blnStartFactUseType  As Boolean, strUseType As String
    If gblnSharedInvoice And bln���� = False Then
        '�Һ�������Ʊ��:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        End If
    End If
    
    If intInsure > 0 And bln���� = False Then
        Set rsTmp = Get���㷽ʽ("�Һ�", "3")
        If rsTmp.RecordCount > 0 Then str�����ʻ� = rsTmp!����
        strAdvance = IIf(str�����ʻ� <> "", str�����ʻ�, "�����ʻ�")
        If gclsInsure.GetCapability(support�����������, , intInsure, strAdvance) Then
            strAdvance = ""     '����̴��벻�����˵Ľ��㷽ʽ,�ձ�ʾȫ������
        End If
        '67143
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure)
        If blnVirtualPrint Then
            If zlGetInvoiceGroupUseID(lng����ID, , , strUseType) = False Then Exit Sub
            strInvoice = GetNextBill(lng����ID)
        End If
    End If
    If bln���� = False Then
        
        Call zlReadRegThreeBalance(strNO, cllBillBalance)
        
        If Not cllBillBalance Is Nothing Then
            '���������˻�֧����,��Ҫ������������ʾ
            If gbln������� Then
                If frmRegistEditSimple.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
            Else
                If frmRegistEdit.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
            End If
            GoTo ReFlash    'ˢ��
        End If
    End If
    
    blnPromptClear = True
    strSQLCard = ExistCardFee(strNO, lngCard����ID)
    If strSQLCard <> "" Then
        '��Ե�������������,����Ҫ������ʾ,���˿�����ȡ����!
        strSQL = "Select c.�����id, c.����, c.����id, d.�Ƿ�����" & vbNewLine & _
             "From ������ü�¼ A, ����ҽ�ƿ��䶯 B, ����ҽ�ƿ���Ϣ C, ҽ�ƿ���� D" & vbNewLine & _
             "Where a.��¼���� = 4 And a.No = [1] And a.��¼״̬ = 1 And b.����id = a.����id And b.�䶯ʱ�� = a.�Ǽ�ʱ�� And" & vbNewLine & _
             "      b.�����id = c.�����id And b.���� = c.���� And c.����id = a.����id And c.״̬ = 0 And c.�����id=d.id And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.RecordCount <> 0 Then
            If Val(Nvl(rsTmp!�Ƿ�����)) = 0 Then
                msgBoxResult = zlCommFun.ShowMsgbox(gstrSysName, "�ò��˹Һ�ʱ�д�����,�˺�ʱ�˿�����ȡ����?", "�˿�,ȡ����,ȡ��", Me, vbQuestion)
                If msgBoxResult = "" Or msgBoxResult = "ȡ��" Then
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                ElseIf msgBoxResult = "�˿�" Then
                    strSQLCard = "zl_ҽ�ƿ���¼_DELETE('" & strSQLCard & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                ElseIf msgBoxResult = "ȡ����" Then
                    If IsCheckCancel��Ԥ��(rsTmp!����ID) = True Then
                        strSQLBound = "Zl_ҽ�ƿ��䶯_Insert(14," & Nvl(rsTmp!����ID) & "," & Nvl(rsTmp!�����ID) & ",Null,'" & Nvl(rsTmp!����) & "','�˺�ȡ����'," & _
                                  "Null,'" & UserInfo.���� & "',Sysdate)"
                    End If
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                End If
            Else
                If MsgBox("�ò��˹Һ�ʱ������,�˺�ͬʱ�˿���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strSQLCard = "zl_ҽ�ƿ���¼_DELETE('" & strSQLCard & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                Else
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                End If
            End If
        Else
            blnPromptClear = False
            strSQLCard = ""
            lngCard����ID = 0
            bln�˷��ش� = gbln�˷��ش�
        End If
    End If
    
    If bln���� = False Then
        dblThreeMoney = zlGetRegThreeMoney(lng����ID, lngCard����ID, cllBillBalance)
    End If
    bytTogetherDo = 0
    '����Һŵ��ĵǼ�����-������Ϣ�ĵǼ������ڹҺŵ���Ч����֮��,����ʾ�Ƿ�ɾ�������   txtʱ��
    If str����� <> "" And blnPromptClear Then
        If Check�Һ�ʱ����(strNO, str�Һ�ʱ��) Then
            Select Case gbyt���������Ϣ  '35176
            Case 0  '�����
            Case 1  '���
                bytTogetherDo = 1
            Case 2  '��ʾ���
                If MsgBox("�˺ź�Ҫ�����ò�����ص��������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bytTogetherDo = 1
                End If
            End Select
        End If
    End If

    If tbsType.SelectedItem.Key = "�Һ�" And mTy_Para.bln�˺���� And mTy_Para.lngN��ȡ��ԤԼ > 0 Then
        '�˺���� �����շ�ԤԼ��ԤԼ���յ��˺�
        If vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("�˺������")) = "" Then
                '�Ƿ�ԤԼ�жϷŵ����� ����Ӱ������
                Datsys = zlDatabase.Currentdate
                datTmp = DateAdd("d", -1 * mTy_Para.lngN��ȡ��ԤԼ, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��"))))
                'ԤԼʱ��-K >datSys
                If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                    MsgBox "���ݺ�Ϊ" & strNO & "���շ�ԤԼ����û�о����˺����,���ܽ����˺�!", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End If

    End If
  
    
    If dblThreeMoney <> 0 And bln������ And bln���� = False Then
        '�����ӿ�,ͬʱ���˲���,�˺�ͬʱҲ�����˲���,��Ϊ�ӿڻ����϶�Ҫ��ȫ��,��֧�ֲ����˿�
        If MsgBox("���ݺ�Ϊ" & strNO & "�ĵ���,�Һŵ�ͬʱ�����˲���,ͬʱ�����������ӿڿ۷�,�˺�ʱ��ͬʱ�˲���,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
        '������������ӿ�,����ȫ��
        byt�˷ѷ�ʽ = 0    'ȫ��,��Ϊ�漰���ӿ�,�ӿ�һ��Ҫ��ȫ��,���ܲ�����

    ElseIf bln������ Then
        If MsgBox("���ݺ�Ϊ" & strNO & "�ĵ���,�ڹҺŵ�ͬʱ�����˲���,�Ƿ�ͬʱ�˲���?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            byt�˷ѷ�ʽ = 0    'ȫ��
        Else
            byt�˷ѷ�ʽ = 1    'ֻ�˹Һŷ�
            bln�˷��ش� = gbln�˷��ش�
        End If
    End If

    If MsgBox(strMessage & "ȷʵҪ������[" & strNO & "]�˺���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub

    If intInsure = 0 And bln���� = False Then
        Set rsOneCard1 = GetOneCardBalance(lng����ID)
        If rsOneCard1.RecordCount > 0 Then
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
                Exit Sub
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Sub
            If strCardNo <> rsOneCard1!��λ�ʺ� Then
                MsgBox "��ǰ������ۿ�Ų�һ��!���ܽ����˷�.", vbInformation, gstrSysName
                Exit Sub
            End If

            If lngCard����ID <> 0 Then
                Set rsOneCard2 = GetOneCardBalance(lngCard����ID)
            End If
        End If
        '�����������
        If IsCheckCancelValied(lng����ID, lngCard����ID, cllBillBalance, dblThreeMoney, strNO) = False Then Exit Sub
    End If

    If str����NO <> "" And bln���� = False Then
        strSQL = "zl_���ﻮ�ۼ�¼_Delete('" & str����NO & "')"
        zlAddArray cllPro, strSQL
    End If
    strSQL = "zl_���˹Һż�¼_DELETE( "
    '���ݺ�_In       ������ü�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '����Ա���_In   ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In   ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    'ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
    strSQL = strSQL & "NULL,"
    'ɾ�������_In   Number := 0,
    strSQL = strSQL & "" & bytTogetherDo & ","
    '��ԭ���˽���_In Varchar2 := Null, --ҽ����������˷ѽ��㷽ʽ,�ձ�ʾȫ������
    strSQL = strSQL & "'" & strAdvance & "',"
    '�˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲�����
    strSQL = strSQL & "" & 0 & ","
    '��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
    strSQL = strSQL & "NULL,"
    '�˺�����_In   Number := 1
    strSQL = strSQL & IIf(mTy_Para.blnReuseRegNo, 1, 0) & ")"
    
    zlAddArray cllPro, strSQL
    If strSQLCard <> "" Then zlAddArray cllPro, strSQLCard
    If strSQLBound <> "" Then zlAddArray cllPro, strSQLBound
    If gbytԤ����˷��鿨 <> 0 And bln���� = False Then
        dblAdvanceMoney = zlGetRegAdvanceMoney(lng����ID, lngCard����ID)
        If dblAdvanceMoney <> 0 Then
            If Not zlDatabase.PatiIdentify(Me, glngSys, lngPatientID, dblAdvanceMoney, mlngModul, 1, , , True, _
                , , (gbytԤ����˷��鿨 = 2)) Then Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If intInsure > 0 Then
        Dim strAdvanceTemp As String
        If bln���� Then
            strAdvanceTemp = "1|" & strNO
        End If
        If Not gclsInsure.RegistDelSwap(lng����ID, intInsure, strAdvanceTemp) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    ElseIf Not rsOneCard1 Is Nothing And bln���� = False Then
        If rsOneCard1.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(Nvl(rsOneCard1!��λ�ʺ�), rsOneCard1!ҽԺ����, "" & Nvl(rsOneCard1!�������), rsOneCard1!���) Then
                gcnOracle.RollbackTrans
                MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                Exit Sub
            End If
            If Not rsOneCard2 Is Nothing Then
                If rsOneCard2.RecordCount > 0 Then
                    If Not objICCard.ReturnSwap(rsOneCard2!��λ�ʺ�, rsOneCard2!ҽԺ����, "" & rsOneCard2!�������, rsOneCard2!���) Then
                        gcnOracle.RollbackTrans
                        MsgBox "һ��ͨ�˿��ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    If bln���� = False Then
        '��������
        '�˷�
        If CallBackBalanceInterface(cllBillBalance, lng����ID, lngCard����ID, dblThreeMoney, cllUpdate, cllThreeIns, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then
                MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
            Else
                MsgBox "���õ������ӿڽ���ʧ��,�˴��˷Ѳ���ʧ��!", vbExclamation + vbOKOnly, gstrSysName
            End If
            Exit Sub
        End If
    
        If Not cllBillBalance Is Nothing And Not cllUpdate Is Nothing Then
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        End If
    End If
    gcnOracle.CommitTrans
    If Not cllThreeIns Is Nothing And bln���� = False Then
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllThreeIns, Me.Caption
    End If
    '����ִ��
ResumeExecute:
    '����:31634
    Err = 0: On Error GoTo NotCommit:
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, True, intInsure)
    blnTrans = False
    If gblnBillPrint Then
        Err = 0: On Error Resume Next
        Call gobjBillPrint.zlEraseBill_Reg("'" & strNO & "'")
        If Err <> 0 Then
            Err = 0
        End If
        On Error GoTo errH
    End If
    '70262:������,2014-03-04,�˺���Ʊ������
'    If bln�˷��ش� And bln���� = False And (byt�˷ѷ�ʽ <> 0 Or blnCardReprint) Then Call RePrintBill(Me, strNO, lng����ID, intInsure, blnVirtualPrint, , strUseType)
    If strAdvance <> "" And bln���� = False Then
        MsgBox "ҽ����֧��[" & strAdvance & "]����,��Ϊ�ֽ�." & vbCrLf & vbCrLf & "�˿��:" & Format(GetCashMoney(strNO), "0.00") & " Ԫ.", vbInformation, gstrSysName
    End If
ReFlash:
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrOthers:
    If ErrCenter = 1 Then gcnOracle.RollbackTrans: Resume
    gcnOracle.CommitTrans
    GoTo ResumeExecute:
    Exit Sub
    '����:31634
NotCommit:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, False, intInsure)
End Sub

Private Sub DeleteRegist()
    Dim strSQL As String, strNO As String, str����NO As String, strCardNo As String
    Dim intInsure As Integer, lng����ID As Long, lngCard����ID As Long, msgBoxResult As String
    Dim str�Һ�ʱ�� As String, strSQLCard As String, strMessage As String
    Dim str����� As String, strAdvance As String, str�����ʻ� As String
    Dim blnEnableDel As Boolean, blnTrans As Boolean, strSQLBound As String
    Dim bytTogetherDo As Byte, bln�˷��ش� As Boolean, blnPromptClear As Boolean
    Dim rsTmp As ADODB.Recordset, rsOneCard1 As ADODB.Recordset, rsOneCard2 As ADODB.Recordset
    Dim objICCard As Object
    Dim cllPro As Collection, cllBillBalance As Collection, dblThreeMoney As Double
    Dim cllUpdate As Collection, cllThreeIns As Collection, strErrMsg As String
    Dim byt�˷ѷ�ʽ As Byte  '0-ȫ�� 1-ֻ�˹Һŷ� 2-ֻ�˲���
    Dim bln������ As Boolean    '�Ƿ����������
    Dim blnCardReprint As Boolean    '���˿��ش�
    Dim Datsys As Date
    Dim datTmp As Date
    Dim lngPatientID As Long
    Dim dblAdvanceMoney As Double    'Ԥ���������
    Dim strInvoice As String, lng����ID As Long, lng����ID As Long
    Dim blnVirtualPrint As Boolean
    Dim bln���� As Boolean, int���� As Integer, bln���� As Boolean
    
    Set cllPro = New Collection
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))

    If strNO = "" Then
        MsgBox "��ǰû�м�¼�����˺ţ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If IsNewModeRegist(strNO) = False Then
        Call DelOldRegist
        Exit Sub
    End If
    
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬����������˺Ų�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    str�Һ�ʱ�� = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("�Һ�ʱ��"))
    str����� = vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("�����"))
    bln������ = Trim(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("����"))) <> ""
    lngPatientID = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("����ID")))
    
    int���� = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("����")))
    bln���� = Val(vsThis.TextMatrix(vsThis.Row, vsThis.ColIndex("���ʷ���"))) = 1
    
    
    If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then
        '�жϵ�ǰ��Ա�Ե����Ƿ��в���Ȩ��,ʱ������,������Һŵ���Ч����
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
                             CDate(str�Һ�ʱ��), "�˺�") Then Exit Sub
    End If

    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If

    If InStr(1, mstrPrivs, ";ǿ���˺�;") = 0 Then   '����:
        '���Һŵ��Ƿ���ִ��
        If InStr(";" & mstrPrivs & ";", ";��ҽ�����˺�;") > 0 Then
            blnEnableDel = True
        End If
        If CheckExecuted(strNO, blnEnableDel) Then
            MsgBox "�Һŵ�" & strNO & "�Ѿ���ҽ��������¹�ҽ��,�����˺ţ�", vbInformation, gstrSysName
            Exit Sub
        End If
        'ҽ��վ�ҵĺ�-�շ��ж�
        If CheckPriceHaveFee(strNO, str����NO) Then Exit Sub
        '�Ƿ���������,��δ�˷�
        If InStr(1, mstrPrivs, ";�շѺ��˺�;") = 0 Then
            If ExistFee(strNO) Then
                MsgBox strNO & "�Һŵ��Ĳ����Ѿ������˷���,�����˷Ѳ����˺�.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
    GoTo ReFlash    'ˢ��


    'ȥ����ҽ������ƥ����
    lng����ID = GetBill����ID(strNO, 4, lng����ID, bln����)
   If zlCheckIsAllowBackSN(strNO, bln����, bln����) = False Then Exit Sub

    'ҽ���˺ż��
    If bln���� Then
        intInsure = int����
    Else
        intInsure = ExistInsure(strNO)
    End If
    
    Dim blnStartFactUseType  As Boolean, strUseType As String
    If gblnSharedInvoice And bln���� = False Then
        '�Һ�������Ʊ��:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        End If
    End If
    
    If intInsure > 0 And bln���� = False Then
        Set rsTmp = Get���㷽ʽ("�Һ�", "3")
        If rsTmp.RecordCount > 0 Then str�����ʻ� = rsTmp!����
        strAdvance = IIf(str�����ʻ� <> "", str�����ʻ�, "�����ʻ�")
        If gclsInsure.GetCapability(support�����������, , intInsure, strAdvance) Then
            strAdvance = ""     '����̴��벻�����˵Ľ��㷽ʽ,�ձ�ʾȫ������
        End If
        '67143
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure)
        If blnVirtualPrint Then
            If zlGetInvoiceGroupUseID(lng����ID, , , strUseType) = False Then Exit Sub
            strInvoice = GetNextBill(lng����ID)
        End If
    End If
    
    If bln���� = False Then
        Call zlReadRegThreeBalance(strNO, cllBillBalance)
        If Not cllBillBalance Is Nothing Then
            '���������˻�֧����,��Ҫ������������ʾ
            If frmRegistEditNew.CancelBill(Me, strNO, mlngModul, mstrPrivs) = False Then Exit Sub
            GoTo ReFlash    'ˢ��
        End If
    End If
    
    blnPromptClear = True
    strSQLCard = ExistCardFee(strNO, lngCard����ID)
    If strSQLCard <> "" Then
        '��Ե�������������,����Ҫ������ʾ,���˿�����ȡ����!
        strSQL = "Select c.�����id, c.����, c.����id, d.�Ƿ�����" & vbNewLine & _
             "From ������ü�¼ A, ����ҽ�ƿ��䶯 B, ����ҽ�ƿ���Ϣ C, ҽ�ƿ���� D" & vbNewLine & _
             "Where a.��¼���� = 4 And a.No = [1] And a.��¼״̬ = 1 And b.����id = a.����id And b.�䶯ʱ�� = a.�Ǽ�ʱ�� And" & vbNewLine & _
             "      b.�����id = c.�����id And b.���� = c.���� And c.����id = a.����id And c.״̬ = 0 And c.�����id=d.id And Rownum = 1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.RecordCount <> 0 Then
            If Val(Nvl(rsTmp!�Ƿ�����)) = 0 Then
                msgBoxResult = zlCommFun.ShowMsgbox(gstrSysName, "�ò��˹Һ�ʱ�д�����,�˺�ʱ�˿�����ȡ����?", "�˿�,ȡ����,ȡ��", Me, vbQuestion)
                If msgBoxResult = "" Or msgBoxResult = "ȡ��" Then
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                ElseIf msgBoxResult = "�˿�" Then
                    strSQLCard = "zl_ҽ�ƿ���¼_DELETE('" & strSQLCard & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                ElseIf msgBoxResult = "ȡ����" Then
                    If IsCheckCancel��Ԥ��(rsTmp!����ID) = True Then
                        strSQLBound = "Zl_ҽ�ƿ��䶯_Insert(14," & Nvl(rsTmp!����ID) & "," & Nvl(rsTmp!�����ID) & ",Null,'" & Nvl(rsTmp!����) & "','�˺�ȡ����'," & _
                                  "Null,'" & UserInfo.���� & "',Sysdate)"
                    End If
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                End If
            Else
                If MsgBox("�ò��˹Һ�ʱ������,�˺�ͬʱ�˿���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strSQLCard = "zl_ҽ�ƿ���¼_DELETE('" & strSQLCard & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
                Else
                    blnPromptClear = False
                    strSQLCard = ""
                    lngCard����ID = 0
                    bln�˷��ش� = gbln�˷��ش�
                    blnCardReprint = gbln�˷��ش�
                End If
            End If
        Else
            blnPromptClear = False
            strSQLCard = ""
            lngCard����ID = 0
            bln�˷��ش� = gbln�˷��ش�
        End If
    End If
    
    If bln���� = False Then
        dblThreeMoney = zlGetRegThreeMoney(lng����ID, lngCard����ID, cllBillBalance)
    End If
    bytTogetherDo = 0
    '����Һŵ��ĵǼ�����-������Ϣ�ĵǼ������ڹҺŵ���Ч����֮��,����ʾ�Ƿ�ɾ�������   txtʱ��
    If str����� <> "" And blnPromptClear Then
        If Check�Һ�ʱ����(strNO, str�Һ�ʱ��) Then
            Select Case gbyt���������Ϣ  '35176
            Case 0  '�����
            Case 1  '���
                bytTogetherDo = 1
            Case 2  '��ʾ���
                If MsgBox("�˺ź�Ҫ�����ò�����ص��������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    bytTogetherDo = 1
                End If
            End Select
        End If
    End If

    If tbsType.SelectedItem.Key = "�Һ�" And mTy_Para.bln�˺���� And mTy_Para.lngN��ȡ��ԤԼ > 0 Then
        '�˺���� �����շ�ԤԼ��ԤԼ���յ��˺�
        If vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��")) <> "" Then
            If vsThis.TextMatrix(vsThis.Row, getColNum("�˺������")) = "" Then
                '�Ƿ�ԤԼ�жϷŵ����� ����Ӱ������
                Datsys = zlDatabase.Currentdate
                datTmp = DateAdd("d", -1 * mTy_Para.lngN��ȡ��ԤԼ, CDate(vsThis.TextMatrix(vsThis.Row, getColNum("ԤԼʱ��"))))
                'ԤԼʱ��-K >datSys
                If Format(Datsys, "yyyy-MM-dd hh:mm:ss") > Format(datTmp, "yyyy-MM-dd hh:mm:ss") Then
                    MsgBox "���ݺ�Ϊ" & strNO & "���շ�ԤԼ����û�о����˺����,���ܽ����˺�!", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End If

    End If
  
    
    If dblThreeMoney <> 0 And bln������ And bln���� = False Then
        '�����ӿ�,ͬʱ���˲���,�˺�ͬʱҲ�����˲���,��Ϊ�ӿڻ����϶�Ҫ��ȫ��,��֧�ֲ����˿�
        If MsgBox("���ݺ�Ϊ" & strNO & "�ĵ���,�Һŵ�ͬʱ�����˲���,ͬʱ�����������ӿڿ۷�,�˺�ʱ��ͬʱ�˲���,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
        '������������ӿ�,����ȫ��
        byt�˷ѷ�ʽ = 0    'ȫ��,��Ϊ�漰���ӿ�,�ӿ�һ��Ҫ��ȫ��,���ܲ�����

    ElseIf bln������ Then
        If MsgBox("���ݺ�Ϊ" & strNO & "�ĵ���,�ڹҺŵ�ͬʱ�����˲���,�Ƿ�ͬʱ�˲���?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            byt�˷ѷ�ʽ = 0    'ȫ��
        Else
            byt�˷ѷ�ʽ = 1    'ֻ�˹Һŷ�
            bln�˷��ش� = gbln�˷��ش�
        End If
    End If

    If MsgBox(strMessage & "ȷʵҪ������[" & strNO & "]�˺���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub

    If intInsure = 0 And bln���� = False Then
        Set rsOneCard1 = GetOneCardBalance(lng����ID)
        If rsOneCard1.RecordCount > 0 Then
            On Error Resume Next
            Set objICCard = CreateObject("zlICCard.clsICCard")
            On Error GoTo 0
            If objICCard Is Nothing Then
                MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
                Exit Sub
            End If
            strCardNo = objICCard.Read_Card(Me)
            If strCardNo = "" Then Exit Sub
            If strCardNo <> rsOneCard1!��λ�ʺ� Then
                MsgBox "��ǰ������ۿ�Ų�һ��!���ܽ����˷�.", vbInformation, gstrSysName
                Exit Sub
            End If

            If lngCard����ID <> 0 Then
                Set rsOneCard2 = GetOneCardBalance(lngCard����ID)
            End If
        End If
        '�����������
        If IsCheckCancelValied(lng����ID, lngCard����ID, cllBillBalance, dblThreeMoney, strNO) = False Then Exit Sub
    End If

    If str����NO <> "" And bln���� = False Then
        strSQL = "zl_���ﻮ�ۼ�¼_Delete('" & str����NO & "')"
        zlAddArray cllPro, strSQL
    End If
    strSQL = "zl_���˹Һż�¼_����_DELETE( "
    '���ݺ�_In       ������ü�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '����Ա���_In   ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In   ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    'ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
    strSQL = strSQL & "NULL,"
    'ɾ�������_In   Number := 0,
    strSQL = strSQL & "" & bytTogetherDo & ","
    '��ԭ���˽���_In Varchar2 := Null, --ҽ����������˷ѽ��㷽ʽ,�ձ�ʾȫ������
    strSQL = strSQL & "'" & strAdvance & "',"
    '�˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲�����
    strSQL = strSQL & "" & 0 & ","
    '��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
    strSQL = strSQL & "NULL,"
    '�˺�����_In   Number := 1
    strSQL = strSQL & IIf(mTy_Para.blnReuseRegNo, 1, 0) & ")"
    
    zlAddArray cllPro, strSQL
    If strSQLCard <> "" Then zlAddArray cllPro, strSQLCard
    If strSQLBound <> "" Then zlAddArray cllPro, strSQLBound
    If gbytԤ����˷��鿨 <> 0 And bln���� = False Then
        dblAdvanceMoney = zlGetRegAdvanceMoney(lng����ID, lngCard����ID)
        If dblAdvanceMoney <> 0 Then
            If Not zlDatabase.PatiIdentify(Me, glngSys, lngPatientID, dblAdvanceMoney, mlngModul, 1, , , True, _
                , , (gbytԤ����˷��鿨 = 2)) Then Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If intInsure > 0 Then
        Dim strAdvanceTemp As String
        If bln���� Then
            strAdvanceTemp = "1|" & strNO
        End If
        If Not gclsInsure.RegistDelSwap(lng����ID, intInsure, strAdvanceTemp) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    ElseIf Not rsOneCard1 Is Nothing And bln���� = False Then
        If rsOneCard1.RecordCount > 0 Then
            If Not objICCard.ReturnSwap(Nvl(rsOneCard1!��λ�ʺ�), rsOneCard1!ҽԺ����, "" & Nvl(rsOneCard1!�������), rsOneCard1!���) Then
                gcnOracle.RollbackTrans
                MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                Exit Sub
            End If
            If Not rsOneCard2 Is Nothing Then
                If rsOneCard2.RecordCount > 0 Then
                    If Not objICCard.ReturnSwap(rsOneCard2!��λ�ʺ�, rsOneCard2!ҽԺ����, "" & rsOneCard2!�������, rsOneCard2!���) Then
                        gcnOracle.RollbackTrans
                        MsgBox "һ��ͨ�˿��ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    If bln���� = False Then
        '��������
        '�˷�
        If CallBackBalanceInterface(cllBillBalance, lng����ID, lngCard����ID, dblThreeMoney, cllUpdate, cllThreeIns, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then
                MsgBox strErrMsg, vbExclamation + vbOKOnly, gstrSysName
            Else
                MsgBox "���õ������ӿڽ���ʧ��,�˴��˷Ѳ���ʧ��!", vbExclamation + vbOKOnly, gstrSysName
            End If
            Exit Sub
        End If
    
        If Not cllBillBalance Is Nothing And Not cllUpdate Is Nothing Then
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        End If
    End If
    gcnOracle.CommitTrans
    If Not cllThreeIns Is Nothing And bln���� = False Then
        Err = 0: On Error GoTo ErrOthers:
        zlExecuteProcedureArrAy cllThreeIns, Me.Caption
    End If
    '����ִ��
ResumeExecute:
    '����:31634
    Err = 0: On Error GoTo NotCommit:
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, True, intInsure)
    blnTrans = False
    '70262:������,2014-03-04,�˺���Ʊ������
'    If bln�˷��ش� And bln���� = False And (byt�˷ѷ�ʽ <> 0 Or blnCardReprint) Then Call RePrintBill(Me, strNO, lng����ID, intInsure, blnVirtualPrint, , strUseType)
    If strAdvance <> "" And bln���� = False Then
        MsgBox "ҽ����֧��[" & strAdvance & "]����,��Ϊ�ֽ�." & vbCrLf & vbCrLf & "�˿��:" & Format(GetCashMoney(strNO), "0.00") & " Ԫ.", vbInformation, gstrSysName
    End If
ReFlash:
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
ErrOthers:
    If ErrCenter = 1 Then gcnOracle.RollbackTrans: Resume
    gcnOracle.CommitTrans
    GoTo ResumeExecute:
    Exit Sub
    '����:31634
NotCommit:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    If intInsure > 0 Then Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, False, intInsure)
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub mnuEdit_Add_Click()
    On Error Resume Next
    Dim datNow As Date
    Err.Clear
    
    datNow = zlDatabase.Currentdate
    If Format(datNow, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        If gbln������� Then
            frmRegistEditSimple.mlngModul = mlngModul
            frmRegistEditSimple.mstrPrivs = mstrPrivs
            frmRegistEditSimple.mbytMode = 0
            frmRegistEditSimple.mbytInState = 0
            Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
            frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            frmRegistEdit.mlngModul = mlngModul
            frmRegistEdit.mstrPrivs = mstrPrivs
            frmRegistEdit.mbytMode = 0
            frmRegistEdit.mbytInState = 0
            Set frmRegistEdit.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
            frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
    Else
        frmRegistEditNew.mlngModul = mlngModul
        frmRegistEditNew.mstrPrivs = mstrPrivs
        frmRegistEditNew.mbytMode = 0
        frmRegistEditNew.mbytInState = 0
        Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
    If gblnOk And tbsType.SelectedItem.Key = "�Һ�" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub ViewOldRegist()
    If vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�")) = "" Then
        MsgBox "��ǰû�м�¼���Բ��ģ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Err.Clear
    Dim blnCancel As Boolean
    Dim bytInState As Byte
    Dim bytViewState As Byte
    Dim strNO As String, rsTemp As ADODB.Recordset, strSQL As String
    bytInState = 1
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    If tbsType.SelectedItem.Key = "ԤԼ" Then
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) <> "1"
        bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬"))))
    Else
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "2"
    End If
    If tbsType.SelectedItem.Key = "�Һ�" Then
        bytViewState = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬"))
        '���ղ��������˷ѵ�ʱ������Ƿ�Ϊ�쳣�˷ѵ���
        If bytViewState = 2 Then
            If CheckBillExistReplenishData(strNO) Then
                strSQL = "Select 1 From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ = 2 And Nvl(����״̬, 0) = 1 And NO = [1] And Rownum < 2"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�Ϊ�˷��쳣����", strNO)
                If Not rsTemp.EOF Then
                    MsgBox "��ǰѡ��Һŵ������ڱ��ղ�������˷��쳣״̬���ݲ�����鿴��", vbExclamation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        If mbytCancel <> 1 And bytViewState = 1 Then
            bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬"))))
            blnCancel = bytInState <> 1
        End If
    End If
    If gbln������� Then
        frmRegistEditSimple.mlngModul = mlngModul
        frmRegistEditSimple.mstrPrivs = mstrPrivs
        frmRegistEditSimple.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
        frmRegistEditSimple.mblnNOMoved = mblnNOMoved
        frmRegistEditSimple.mbytMode = tbsType.SelectedItem.index - 1
        frmRegistEditSimple.mbytInState = bytInState
        frmRegistEditSimple.mblnViewCancel = blnCancel
        frmRegistEditSimple.mblnViewOriginal = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "3"
        Set frmRegistEditSimple.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEditSimple.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        frmRegistEdit.mlngModul = mlngModul
        frmRegistEdit.mstrPrivs = mstrPrivs
        frmRegistEdit.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
        frmRegistEdit.mblnNOMoved = mblnNOMoved
        frmRegistEdit.mbytMode = tbsType.SelectedItem.index - 1
        frmRegistEdit.mbytInState = bytInState
        frmRegistEdit.mblnViewCancel = blnCancel
        frmRegistEdit.mblnViewOriginal = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "3"
        Set frmRegistEdit.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
        frmRegistEdit.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    End If
End Sub

Private Sub mnuEdit_View_Click()
    If vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�")) = "" Then
        MsgBox "��ǰû�м�¼���Բ��ģ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Err.Clear
    Dim blnCancel As Boolean
    Dim bytInState As Byte
    Dim bytViewState As Byte
    Dim strNO As String, rsTemp As ADODB.Recordset, strSQL As String
    bytInState = 1
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    
    If IsNewModeRegist(strNO) = False Then
        Call ViewOldRegist
        Exit Sub
    End If
    
    If tbsType.SelectedItem.Key = "ԤԼ" Then
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) <> "1"
        bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬"))))
    Else
        blnCancel = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "2"
    End If
    If tbsType.SelectedItem.Key = "�Һ�" Then
        bytViewState = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬"))
        '���ղ��������˷ѵ�ʱ������Ƿ�Ϊ�쳣�˷ѵ���
        If bytViewState = 2 Then
            If CheckBillExistReplenishData(strNO) Then
                strSQL = "Select 1 From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ = 2 And Nvl(����״̬, 0) = 1 And NO = [1] And Rownum < 2"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�Ϊ�˷��쳣����", strNO)
                If Not rsTemp.EOF Then
                    MsgBox "��ǰѡ��Һŵ������ڱ��ղ�������˷��쳣״̬���ݲ�����鿴��", vbExclamation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        If mbytCancel <> 1 And bytViewState = 1 Then
            bytInState = IIf(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "1", 1, Val(vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬"))))
            blnCancel = bytInState <> 1
        End If
    End If
    frmRegistEditNew.mlngModul = mlngModul
    frmRegistEditNew.mstrPrivs = mstrPrivs
    frmRegistEditNew.mstrNoIn = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    frmRegistEditNew.mblnNOMoved = mblnNOMoved
    frmRegistEditNew.mbytMode = tbsType.SelectedItem.index - 1
    frmRegistEditNew.mbytInState = bytInState
    frmRegistEditNew.mblnViewCancel = blnCancel
    frmRegistEditNew.mblnViewOriginal = vsThis.TextMatrix(vsThis.Row, getColNum("��¼״̬")) = "3"
    Set frmRegistEditNew.mobjMsgModule = mobjMsgModule '��Ϣ����ģ��
    frmRegistEditNew.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    If Me.Enabled And Me.Visible Then Me.SetFocus
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    Call ShowBills
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

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Add"
            mnuEdit_Add_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "ԤԼ"
            mnuEdit_Bespeak_Click
        Case "����"
            mnuEdit_Incept_Click
        Case "ȡ��"
            mnuEdit_Cancel_Click
        Case "����"
            mnuFileRollingCurtain_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
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
    
    intRow = vsThis.Row
    
    '��ͷ
    objOut.Title.Text = "����Һŵ����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmRegistFilter
        If IsNull(.dtpEnd.Value) Then
            objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, "yyyy-MM-dd")
        Else
            objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, "yyyy-MM-dd HH:MM") & " �� " & Format(.dtpEnd.Value, "yyyy-MM-dd HH:MM")
        End If
        objRow.Add "���ʣ�" & IIf(.optRegistRecord(1).Value = True, "�˿��¼", "�տ��¼")
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    vsThis.Redraw = False
    Set objOut.Body = vsThis
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    vsThis.Row = intRow
    vsThis.Col = 0: vsThis.ColSel = vsThis.Cols - 1
    vsThis.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Hwnd
End Sub

Private Function IsNewModeRegist(ByVal strNO As String) As Boolean
'���ܣ��жϹҺŵ��Ƿ�Ϊ������Ű�ģʽ�Һŵ�
'������strNo = �Һŵ����ݺ�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select 1 From ���˹Һż�¼ Where NO = [1] And �����¼Id Is Null And ����ʱ�� < [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, gdatRegistTime)
    If rsTemp.EOF Then
        IsNewModeRegist = True
    Else
        IsNewModeRegist = False
    End If
End Function

Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
'������blnUsed=������ǰ�嵥����������
    
    mnuFile_Print.Enabled = blnUsed
    mnuFile_Preview.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed And tbsType.SelectedItem.Key = "�Һ�"
    '�ش�
    mnuEdit_Print.Enabled = blnUsed And tbsType.SelectedItem.Key = "�Һ�"
    '����
    mnuEdit_Print_Supplemental.Enabled = blnUsed And tbsType.SelectedItem.Key = "�Һ�"
    mnuEdit_Print_Slip.Enabled = blnUsed And tbsType.SelectedItem.Key = "�Һ�"
    mnuEdit_Print_Case.Enabled = blnUsed And tbsType.SelectedItem.Key = "�Һ�"
    tbr.Buttons("Del").Enabled = blnUsed And tbsType.SelectedItem.Key = "�Һ�"
    
    mnuEdit_Incept.Enabled = blnUsed And tbsType.SelectedItem.Key = "����"
    tbr.Buttons("����").Enabled = blnUsed And tbsType.SelectedItem.Key = "����"
    
    mnuEdit_Cancel.Enabled = blnUsed And tbsType.SelectedItem.Key <> "�Һ�"
    tbr.Buttons("ȡ��").Enabled = blnUsed And tbsType.SelectedItem.Key <> "�Һ�"


    mnuEdit_View.Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
    '
    mnuEdit_CancelAuditing.Enabled = blnUsed And mnuEdit_CancelAuditing.Visible
    '�����:45507
    mnuEdit_BatchChangeNum.Enabled = InStr(1, mstrPrivs, ";������������;") > 0
    
    Call vsThis_RowColChange
End Sub

Private Sub Form_Load()
    Dim i As Integer, blnHavePrivs As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    
    '�Զ���ҵ
    strSQL = "zl1_Auto_Buildingregisterplan"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mstrVsType = tbsType.SelectedItem.Key
    
    mstr���ӷ� = ""
    mstr������ĿID = ""
    strSQL = "Select zl_Fun_RegCustomName As ���ӷ� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        mstr���ӷ� = Split(Nvl(rsTmp!���ӷ�) & "|", "|")(0)
        mstr������ĿID = Split(Nvl(rsTmp!���ӷ�) & "|", "|")(1)
    End If
    
    If mstr���ӷ� <> "" And mstr������ĿID <> "" Then
        mnuEdit_DelExtra.Caption = "��" & mstr���ӷ� & "(&E)"
        tbr.Buttons("Del").ButtonMenus("DelExtra").Text = "��" & mstr���ӷ�
        mnuEdit_DelExtra.Visible = True
        tbr.Buttons("Del").ButtonMenus("DelExtra").Visible = True
    Else
        mnuEdit_DelExtra.Visible = False
        tbr.Buttons("Del").ButtonMenus("DelExtra").Visible = False
    End If
    
    Call Form_Resize '����Bh���ò������¼�Form_Resize
    Call tbsType_Click
    Call RestoreWinState(Me, App.ProductName)
      
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    'Ȩ������
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1111_1")
    
    If InStr(mstrPrivs, ";LED������;") = 0 Then gblnLED = False
    
    '�����Һ�
    If InStr(";" & mstrPrivs & ";", ";���շѺ�;") = 0 And InStr(";" & mstrPrivs & ";", ";����Ѻ�;") = 0 Then
        mnuEdit_Add.Visible = False
        tbr.Buttons("Add").Visible = False
        mnuEdit_Print.Visible = False
    End If
    '52328
    mnuEdit_Print_Supplemental.Visible = (InStr(mstrPrivs, ";����Ѻ�;") > 0 Or InStr(mstrPrivs, ";���շѺ�;") > 0) And InStr(mstrPrivs, ";����Ʊ��;") > 0
    If InStr(";" & mstrPrivs & ";", ";�ش�Ʊ��;") = 0 Then
        mnuEdit_Print.Visible = False
    End If
    mnuEdit_Print_Slip.Visible = InStr(mstrPrivs, ";�Һ�ƾ����ӡ;") > 0
    If InStr(";" & mstrPrivs & ";", ";�˺�;") = 0 Then
        mnuEdit_Del.Visible = False
        tbr.Buttons("Del").Visible = False
    End If
    If InStr(";" & mstrPrivs & ";", ";����Ѻ�;") = 0 And InStr(mstrPrivs, ";���շѺ�;") = 0 _
        And InStr(";" & mstrPrivs & ";", ";�˺�;") = 0 Then
        mnuEdit_1.Visible = False
        tbr.Buttons("Fun_1").Visible = False
    End If
    If InStr(";" & mstrPrivs & ";", ";�˺����;") = 0 Then
        mnuEdit_CancelAuditing.Visible = False
    End If
    If InStr(";" & mstrPrivs & ";", ";����Ű�;") = 0 Then
    '����Ű�
        mnuEdit_BindPatNum.Enabled = False
    End If
    
    '�շ����ʹ���
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";����;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("����").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    
    'ԤԼ�Һ�
    If InStr(mstrPrivs, ";ԤԼ�Һ�;") = 0 Then
        mnuEdit_Bespeak.Visible = False
        tbr.Buttons("ԤԼ").Visible = False
        frmBookingDefer.Visible = False
    End If
    If InStr(mstrPrivs, ";����ԤԼ;") = 0 Then
        mnuEdit_Incept.Visible = False
        tbr.Buttons("����").Visible = False
    End If
    If InStr(mstrPrivs, ";ȡ��ԤԼ;") = 0 Then
        mnuEdit_Cancel.Visible = False
        mnuEdit_Clear.Visible = False
        tbr.Buttons("ȡ��").Visible = False
    End If
    If InStr(mstrPrivs, "'ԤԼ�Һ�;") = 0 _
        And InStr(mstrPrivs, ";����ԤԼ;") = 0 _
        And InStr(mstrPrivs, ";ȡ��ԤԼ;") = 0 Then
        mnuEdit_2.Visible = False
        tbr.Buttons("Fun_2").Visible = False
    End If
            
    Call SetHeader
    Call SetMenu(False)
    mbytCancel = 1: mstrFilter = ""
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
    '���ݲ���Ȩ�޶�λȱʡ�嵥
    If InStr(";" & mstrPrivs & ";", ";����Ѻ�;") > 0 Or InStr(mstrPrivs, ";���շѺ�;") > 0 Then
        'ȱʡ���Ǹ�ҳ
    ElseIf InStr(mstrPrivs, ";ԤԼ�Һ�;") > 0 Then
        tbsType.Tabs("ԤԼ").Selected = True
    ElseIf InStr(mstrPrivs, ";����ԤԼ;") > 0 Then
        tbsType.Tabs("����").Selected = True
    End If
    
    On Error GoTo errH
    InitActionType
    Call InitPara
    If mactionType = t_ʱ�� Then
        mnuEdit_Defer.Enabled = False
    End If
    
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, mlngModul, UserInfo.���, UserInfo.����)
    End If
    On Error GoTo errH
    
    '��ʼ����Ϣ�������ģ��
    Call InitMsgModule
    
    '���������˰�ش�ӡ����
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(1)
        End If
        On Error GoTo 0
    End If
    
    Call LoadPlugInMnu
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    vsThis.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    With tbsType
        .Top = Me.ScaleTop + cbrH + 15
        .Left = Me.ScaleLeft + 30
        .Width = Me.ScaleWidth - 60
    End With
    
    With vsThis
        .Top = Me.ScaleTop + cbrH + 350
        .Height = Me.ScaleHeight - cbrH - staH - 350
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Unload frmRegistFilter
    Unload frmRegistFind
    Call SaveWinState(Me, App.ProductName)
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
            Exit For
        End If
    Next
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
    
    '��ж��Ϣ����ģ��
    Call UnloadMsgModule
End Sub

Private Sub mnuViewGo_Click()
    If tbsType.SelectedItem.Key = "�Һ�" Then
        frmRegistFind.txtFact.Enabled = True
        frmRegistFind.txtFact.BackColor = Me.vsThis.BackColor
    Else
        frmRegistFind.txtFact.Text = ""
        frmRegistFind.txtFact.Enabled = False
        frmRegistFind.txtFact.BackColor = Me.BackColor
    End If
    If mbytCancel <> 2 Then
        frmRegistFind.lbl����Ա.Caption = "�Һ�Ա"
    Else
        frmRegistFind.lbl����Ա.Caption = "�˺�Ա"
    End If
    frmRegistFind.Show 1, Me
    If gblnOk Then Call SeekBill(frmRegistFind.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To vsThis.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmRegistFind
            If .txtNO.Text <> "" Then
                blnFill = blnFill And vsThis.TextMatrix(i, getColNum("���ݺ�")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And vsThis.TextMatrix(i, getColNum("����Ʊ��")) = .txtFact.Text
            End If
            If .cbo����Ա.ListIndex > 0 Then
                If mbytCancel <> 2 Then
                    blnFill = blnFill And vsThis.TextMatrix(i, getColNum("�Һ�Ա")) = NeedName(.cbo����Ա.Text)
                Else
                    blnFill = blnFill And vsThis.TextMatrix(i, getColNum("�˺�Ա")) = NeedName(.cbo����Ա.Text)
                End If
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(vsThis.TextMatrix(i, getColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If IsNumeric(.txt�����.Text) Then
                blnFill = blnFill And Val(vsThis.TextMatrix(i, getColNum("�����"))) = Val(.txt�����.Text)
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            vsThis.Row = i: vsThis.TopRow = i
            vsThis.Col = 0: vsThis.ColSel = vsThis.Cols - 1
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

Private Sub InitPara()
    '��ȡ�Һ���ز���
     With mTy_Para
            .bln�˺���� = Val(zlDatabase.GetPara("�˺����", glngSys, mlngModul, 0)) = 1
            .lngN��ȡ��ԤԼ = Val(zlDatabase.GetPara("N���ڲ���ȡ��ԤԼ��", glngSys, mlngModul, 0))
            .blnReuseRegNo = Val(zlDatabase.GetPara("�����������Һ�", glngSys, mlngModul, 1)) = 1
     End With
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    '�����:48911
    If tbsType.SelectedItem.Key = "�Һ�" Then
            '�����:51672
            strHead = "ҽ��,4,500|���ݺ�,1,850|����Ʊ��,1,850|�ű�,1,1600|����,1,650|����,1,1000|ҽ��,1,700|����,4,500|�����,1,900|����,1,650|���￨��,1,1000|�ֻ���,1,1100|�ѱ�,1,1000|�ܽ��,7,800|�Һŷ�,7,800|������,7,800|�Һ�ʱ��,4,1800|�Ǽ�ʱ��,4,1800|�˺�ʱ��,4," & IIf(mbytCancel = 1, 0, 1800) & "|" & IIf(mbytCancel = 2, "�˺�Ա", "�Һ�Ա") & ",1,650|�շ�Ա,1,650|�շѵ�,1,850|ժҪ,1,1800|ԤԼʱ��,1,1800|����,4,500|����ID,1,0|��¼״̬,1,0|�˺������,1,650|�˺����ʱ��,4,1800|ԤԼ����Ա,1,650|����,1,0|���ʷ���,1,0"
    Else
            strHead = "ͣ�ð���,1,800|���ݺ�,1,850|ԤԼʱ��,1,1550|�ű�,1,1600|����,1,650|����,1,1000|ҽ��,1,700|����,4,500|�����,1,900|����,1,650|���֤��,1,2000|��ϵ�绰,1,2650|�ֻ���,1,2650|�ѱ�,1,1000|���,7,800|ժҪ,1,2000|�Ǽ�ʱ��,4,1800|�Һ�Ա,1,650,|��¼״̬,2,0|�˺������,1,650|�˺����ʱ��,4,1800|ԤԼ����Ա,1,650|����ID,1,0|����,1,0|���ʷ���,1,0"
    End If
     For i = 0 To Me.vsThis.Cols - 2
        With vsThis
            .FixedAlignment(i) = flexAlignCenterCenter
        End With
    
    Next
    vsThis.FixedAlignment(vsThis.Cols - 2) = flexAlignLeftCenter
    vsThis.ColAlignment(vsThis.Cols - 2) = flexAlignLeftCenter
    With vsThis
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .ColKey(i) = Split(Split(strHead, "|")(i), ",")(0)
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
           ' .ColAlignmentFixed(i) = 4
        Next
      '   .ColHidden(getColNum("��¼״̬")) = True
        .RowHeight(0) = 320
        .ExtendLastCol = True
        
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
         Call vsThis_EnterCell
         zl_vsGrid_Para_Restore mlngModul, vsThis, Me.Caption, Me.tbsType.SelectedItem.Key, False, InStr(1, mstrPrivs, ";��������;") > 0
        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
    Dim strSQL          As String
    Dim strTime         As String
    Dim i               As Long
    Dim strFilter       As String
    Dim str������ü�¼  As String
    Dim strDate As String
    Dim strTmp          As String
    Dim blnChange       As Boolean
    Dim strPlanFilter   As String
    blnChange = tbsType.SelectedItem.Key <> mstrVsType
    mstrVsType = tbsType.SelectedItem.Key
    On Error GoTo errH
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�Һ�����,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        
        SQLCondition.Default = (mstrFilter = "")
        strFilter = mstrFilter
        
        '�����:48911
        If tbsType.SelectedItem.Key = "�Һ�" Then
            '�ѹһ��ѽ��յĺ�:�Ǽ�ʱ����ָ����Χ�ڵ�,
            If mstrFilter = "" Then
                'ȱʡ��ʾ��ǰ����Ա�����ڹҵĺ�
                mbytCancel = 1
                strFilter = " And A.�Ǽ�ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60 And A.����Ա����||''=[1]"
            End If
            '�����:49528
            '�����:51672
              strSQL = "  " & _
                "       Select D.�ֻ���,a.No As ���ݺ�, f.ʵ��Ʊ�� As ����Ʊ��, Decode(a.�ű�, Null, Null, '[' || a.�ű� || ']') || Max(Decode(f.���,1,c3.����,Null)) As �ű�, a.����, e.���� As ����," & vbNewLine & _
                "              a.ִ���� As ҽ��, Decode(Max(Nvl(f.���ӱ�־, 0)), 1, '��', Null) As ����, a.�����, a.����, d.���￨��, f.�ѱ�, " & vbNewLine & _
                "              To_Char(Decode(a.��¼״̬,2,-1*Sum(Nvl(f.ʵ�ս��, 0)),Sum(Nvl(f.ʵ�ս��, 0))), '99999999999999990.00') As ���,To_Char(Decode(a.��¼״̬,2,-1*Sum(Decode(Sign(Nvl(f.���ӱ�־, 0)), 1, 0, 1) * Nvl(f.ʵ�ս��, 0)),Sum(Decode(Sign(Nvl(f.���ӱ�־, 0)), 1, 0, 1) * Nvl(f.ʵ�ս��, 0))), '99999999999999990.00') As �Һŷ�," & _
                "               To_Char(Decode(a.��¼״̬,2,-1*Sum(Decode(Sign(Nvl(f.���ӱ�־, 0)), 1, 1, 0) * Nvl(f.ʵ�ս��, 0)),Sum(Decode(Sign(Nvl(f.���ӱ�־, 0)), 1, 1, 0) * Nvl(f.ʵ�ս��, 0))), '99999999999999990.00') As ������, a.����ʱ�� as �Һ�ʱ��,a.�Ǽ�ʱ��,Decode(A.��¼״̬,2,A.�Ǽ�ʱ��,Null) As �˺�ʱ�� ," & IIf(mbytCancel = 2, "a.����Ա���� as �˺�Ա", "a.����Ա���� as �Һ�Ա") & ", a.����Ա���� As �շ�Ա, a.�շѵ�,a.ժҪ, " & vbNewLine & _
                "              Decode(a.ԤԼ, 1, a.����ʱ��, Null) As ԤԼʱ��, Decode(a.����, 1, '��', Null) As ����, a.����id, a.��¼״̬, a.�˺������,A.�˺����ʱ��,a.ԤԼ����Ա ," & _
                "              Max(A.����) as ����,Max(F.���ʷ���) as ���ʷ���" & vbNewLine & _
                "       From ���˹Һż�¼ A, ������Ϣ D, �ٴ������¼ B, �ٴ������Դ B1, �ҺŰ��� B2, ���ű� E, ������ü�¼ F, �շ���ĿĿ¼ C, �շ���ĿĿ¼ C1, �շ���ĿĿ¼ C2, �շ���ĿĿ¼ C3 " & vbNewLine & _
                "       Where a.����id = d.����id(+) And a.�շѵ� Is Null And a.�����¼ID = b.ID(+) And a.�ű� = b1.����(+) And a.�ű� = b2.����(+) And b2.��Ŀid = c2.id(+) And f.�շ�ϸĿID=C3.ID(+) And b1.��Ŀid = c1.id(+)  And a.ִ�в���id = e.Id(+) And b.��Ŀid = c.Id(+) And a.��¼���� = 1 " & " And (e.վ��='" & gstrNodeNo & "' Or e.վ�� is Null) And " & vbNewLine & _
                "             a.��¼״̬ = f.��¼״̬ And a.No = f.No And F.��¼����=4 " & IIf(mbytCancel = 1, " And a.��¼״̬ = 1 ", IIf(mbytCancel = 2, " And a.��¼״̬ = 2 ", "")) & strFilter & vbNewLine & _
                "       Group By a.No, f.ʵ��Ʊ��, a.�ű�, a.����, e.����, a.ִ����, a.�����, a.����, d.���￨��,d.�ֻ���, f.�ѱ�, a.����ʱ��,a.�Ǽ�ʱ��,Decode(A.��¼״̬,2,A.�Ǽ�ʱ��,Null), a.����Ա����, a.ժҪ," & _
                "              Decode(a.ԤԼ, 1, a.����ʱ��, Null), a.����id, a.��¼״̬, a.����,a.�˺������,a.�˺����ʱ��,a.ԤԼ����Ա,a.����Ա����, a.�շѵ�" & vbNewLine

              strSQL = strSQL & " Union All " & _
                "       Select D.�ֻ���,a.No As ���ݺ�, f.ʵ��Ʊ�� As ����Ʊ��, Decode(a.�ű�, Null, Null, '[' || a.�ű� || ']') || Max(Decode(f.���,1,c3.����,Null)) As �ű�, a.����, e.���� As ����," & vbNewLine & _
                "              a.ִ���� As ҽ��, Decode(Max(Nvl(f.���ӱ�־, 0)), 1, '��', Null) As ����, a.�����, a.����, d.���￨��, f.�ѱ�, " & vbNewLine & _
                "              To_Char(Sum(Nvl(f.ʵ�ս��, 0)), '99999999999999990.00') As ���,To_Char(Sum(Decode(Sign(Nvl(f.���ӱ�־, 0)), 1, 0, 1) * Nvl(f.ʵ�ս��, 0)), '99999999999999990.00') As �Һŷ�," & _
                "               To_Char(Sum(Decode(Sign(Nvl(f.���ӱ�־, 0)), 1, 1, 0) * Nvl(f.ʵ�ս��, 0)), '99999999999999990.00') As ������, a.����ʱ�� as �Һ�ʱ��,a.�Ǽ�ʱ��,Decode(A.��¼״̬,2,A.�Ǽ�ʱ��,Null) As �˺�ʱ�� ," & IIf(mbytCancel = 2, "a.����Ա���� as �˺�Ա", "a.����Ա���� as �Һ�Ա") & ", Decode(Max(f.��¼״̬),0,Null,Max(f.����Ա����)) As �շ�Ա, a.�շѵ�,a.ժҪ, " & vbNewLine & _
                "              Decode(a.ԤԼ, 1, a.����ʱ��, Null) As ԤԼʱ��, Decode(a.����, 1, '��', Null) As ����, a.����id, a.��¼״̬, a.�˺������,A.�˺����ʱ��,a.ԤԼ����Ա ," & _
                "              Max(A.����) as ����,Max(F.���ʷ���) as ���ʷ���" & vbNewLine & _
                "       From ���˹Һż�¼ A, ������Ϣ D, �ٴ������¼ B, �ٴ������Դ B1, �ҺŰ��� B2, ���ű� E, ������ü�¼ F, �շ���ĿĿ¼ C, �շ���ĿĿ¼ C1, �շ���ĿĿ¼ C2, �շ���ĿĿ¼ C3 " & vbNewLine & _
                "       Where a.����id = d.����id(+) And a.�����¼ID = b.ID(+) And a.�ű� = b1.����(+) And a.�ű� = b2.����(+) And b2.��Ŀid = c2.id(+) And f.�շ�ϸĿID=C3.ID(+) And b1.��Ŀid = c1.id(+)  And a.ִ�в���id = e.Id(+) And b.��Ŀid = c.Id(+) And a.��¼���� = 1 " & " And (e.վ��='" & gstrNodeNo & "' Or e.վ�� is Null) And " & vbNewLine & _
                "             a.�շѵ� Is Not Null And a.�շѵ� = f.No And F.��¼����=1 And f.��¼״̬ <> 2 " & IIf(mbytCancel = 1, " And a.��¼״̬ = 1 ", IIf(mbytCancel = 2, " And a.��¼״̬ = 2 ", "")) & strFilter & vbNewLine & _
                "       Group By a.No, f.ʵ��Ʊ��, a.�ű�, a.����, e.����, a.ִ����, a.�����, a.����, d.���￨��,D.�ֻ���, f.�ѱ�, a.����ʱ��,a.�Ǽ�ʱ��,Decode(A.��¼״̬,2,A.�Ǽ�ʱ��,Null), a.����Ա����, a.ժҪ," & _
                "              Decode(a.ԤԼ, 1, a.����ʱ��, Null), a.����id, a.��¼״̬, a.����,a.�˺������,a.�˺����ʱ��,a.ԤԼ����Ա , a.�շѵ�" & vbNewLine

                If frmRegistFilter.mblnDateMoved Then
                      strSQL = strSQL & " Union All " & Replace(strSQL, "������ü�¼", "H������ü�¼")
                End If
                
                strSQL = _
                "       Select Decode(Nvl(����,0),0,'','��') As ҽ��,���ݺ�, ����Ʊ��, �ű�, ����, ����,ҽ��, ����, �����, ����, ���￨��,�ֻ���, �ѱ�, " & vbNewLine & _
                "              ���,�Һŷ�,������,�Һ�ʱ��,�Ǽ�ʱ��,�˺�ʱ�� ," & IIf(mbytCancel = 2, "�˺�Ա", "�Һ�Ա") & ", �շ�Ա, �շѵ�, ժҪ, " & vbNewLine & _
                "              ԤԼʱ��, ����, ����id, ��¼״̬, �˺������,�˺����ʱ��,ԤԼ����Ա, ����,���ʷ���" & vbNewLine & _
                "       From (" & strSQL & ")" & _
                "       Order By ���ݺ� Desc,�Һ�ʱ�� Desc"
          
        ElseIf tbsType.SelectedItem.Key = "ԤԼ" Then
            '��ԤԼ�ĺ�:����������Ч��Χ��(�Ǽ�ʱ��>=��ǰʱ��-����ԤԼ����),ԤԼʱ����ָ����Χ�ڵ�
            If mstrFilter = "" Then
                'ȱʡ��ʾ��ǰ����Ա�ҵ�ԤԼʱ��δʧЧ�ĵ���
                strFilter = " And A.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate)+1-1/24/60/60 + zl_Fun_GetAppointmentDays + Decode(Nvl(B1.ԤԼ����," & gintԤԼ���� & "),0,15,Nvl(B1.ԤԼ����," & gintԤԼ���� & "))" & _
                    " And A.����Ա����||''=[1]"
            End If
            strFilter = Replace(strFilter, "����ʱ��+0", "����ʱ��")
              strPlanFilter = Replace(strFilter, "And (F.�ѱ� = [11] or F.�ѱ� is Null)", "")
               strSQL = "" & _
                " Select decode(A.ͣ��,1,'��ͣ','')  As ͣ�ð���, A.���ݺ�,A.ԤԼʱ��,A.�ű�,To_Char(A.����,'99999') ����,D.���� As ����,A.ҽ��,A.����,A.�����,A.����,A.���֤��,A.��ϵ�绰,A.�ֻ���,A.�ѱ�,A.���,A.ժҪ,A.�Ǽ�ʱ��,A.�Һ�Ա,A.��¼״̬ ,a.�˺������,a.�˺����ʱ��,a.ԤԼ����Ա ,A.����ID" & vbNewLine & _
                " From (  Select Max(M.ͣ��) as ͣ��,A.NO as ���ݺ�,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as ԤԼʱ��," & vbNewLine & _
                "                   Decode(A.�ű�,NULL,NULL,'['||A.�ű�||']') || Nvl(C.����,Nvl(C1.����,C2.����)) as �ű�,A.����,A.ִ���� as ҽ��,a.��¼״̬,a.ִ�в���Id," & vbNewLine & _
                "                   Decode(Max(Decode(F.���ӱ�־,1,1,0)),1,'��',NULL) as ����," & vbNewLine & _
                "                   A.�����,A.����,F.�ѱ� as �ѱ�,D.���֤��,D.��ͥ�绰 as ��ϵ�绰,D.�ֻ���," & vbNewLine & _
                "                   To_Char(Sum(decode(f.��¼״̬,2,-1,1)*nvl(f.ʵ�ս��,0)), '9999990.00') as ���," & vbNewLine & _
                "                   A.ժҪ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��,A.����Ա���� as �Һ�Ա,a.�˺������,a.�˺����ʱ��,a.ԤԼ����Ա ,A.����ID,0 as ����,0 as ���ʷ���" & vbNewLine & _
                "           From ���˹Һż�¼ A, ������Ϣ D,�ٴ������¼ B, �ٴ������Դ B1, �ҺŰ��� B2, �շ���ĿĿ¼ C, �շ���ĿĿ¼ C1, �շ���ĿĿ¼ C2 , ������ü�¼ F, " & vbNewLine & _
                "               (   Select A.ID,Max(1) as ͣ�� From  ���˹Һż�¼ A, ������Ϣ D,�ٴ������¼ B,�շ���ĿĿ¼ C,�ٴ������Դ B1 " & vbNewLine & _
                "                    Where   A.�����¼ID=B.ID And B.��ԴID=B1.ID And B.��ĿID=C.ID(+) And a.����id = d.����id(+) " & vbNewLine & _
                "                               And A.��¼����=2  " & IIf(mbytCancel = 1, " And A.��¼״̬=1", "") & strPlanFilter & vbNewLine & _
                "                               And A.�Ǽ�ʱ��>=Sysdate - zl_Fun_GetAppointmentDays - Decode(Nvl(B1.ԤԼ����," & gintԤԼ���� & "),0,15,Nvl(B1.ԤԼ����," & gintԤԼ���� & "))" & vbNewLine & _
                "                               And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & vbNewLine & _
                "                               And A.����ʱ�� between B.ͣ�￪ʼʱ�� and B.ͣ����ֹʱ�� " & vbNewLine & _
                "                     Group by A.ID ) M" & vbNewLine & _
                "           Where A.�����¼ID=B.ID(+) And a.�ű�=b1.����(+) And b1.��Ŀid=c1.id(+) And a.�ű�=b2.����(+) And b2.��Ŀid=c2.id(+) And a.����id = d.����id(+) And b.��ĿID=c.ID(+) " & vbNewLine & _
                "                 And A.NO=F.NO(+)  And A.ID=M.ID(+)  And A.��¼����(+)=2 And F.��¼����(+)=4 and a.��¼״̬=decode(F.��¼״̬,0,1,F.��¼״̬)  " & IIf(mbytCancel = 1, " And a.��¼״̬=1", "") & strFilter & vbNewLine & _
                "                   And A.�Ǽ�ʱ��>=Sysdate - zl_Fun_GetAppointmentDays - Decode(Nvl(B1.ԤԼ����," & gintԤԼ���� & "),0,15,Nvl(B1.ԤԼ����," & gintԤԼ���� & "))" & " And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & vbNewLine & _
                "           Group by A.NO,A.����ʱ��,A.�ű�,a.����,Nvl(C.����,Nvl(C1.����,C2.����)),a.��¼״̬,a.ִ�в���Id," & vbNewLine & _
                "                       A.ִ����,A.�����,A.����,D.���֤��,D.��ͥ�绰,D.�ֻ���,A.ժҪ,A.�Ǽ�ʱ��,A.����Ա����,f.�ѱ�,a.�˺������,a.�˺����ʱ��,a.ԤԼ����Ա,A.����ID" & vbNewLine & _
                "            " & vbNewLine & _
                "   ) A, ���ű� D" & vbNewLine & _
                "   Where A.ִ�в���ID=D.ID " & _
                "   Order by A.�Ǽ�ʱ�� Desc "
         
            
        ElseIf tbsType.SelectedItem.Key = "����" Then
            'Ӧ���յĺ�:����������Ч��Χ��,ԤԼ�ڽ����,��ǰʱ����ԤԼ��ʱ��η�Χ֮�ڵ�
            If mstrFilter = "" Then
                'ȱʡ��ʾ��ǰ����Ա��ǰӦ�ý��յĺ�
                strFilter = " And A.����Ա����||''=[1]"
            End If
            strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")

            strTime = ""

            'ȡ���ڵ���������Ӧ���ŵ�ʱ���
            strSQL = "Decode(To_Char(SysDate,'D'),'1',B.����,'2',B.��һ,'3',B.�ܶ�,'4',B.����,'5',B.����,'6',B.����,'7',B.����,NULL)"
            
            '�ű��ܱ�,�����󲻳���
            strSQL = " " & _
            "      Select Decode(a.ͣ��, 1, '��ͣ', '') As ͣ�ð���, a.���ݺ�, a.ԤԼʱ��, a.�ű�, To_Char(a.����, '99999') ����, d.���� As ����, a.ҽ��, a.����, a.�����," & vbNewLine & _
            "             a.���� , a.���֤��, a.��ϵ�绰,a.�ֻ���, a.�ѱ�, a.���, a.ժҪ, a.�Ǽ�ʱ��, a.�Һ�Ա,A.��¼״̬ ,A.�˺������,a.�˺����ʱ��,a.ԤԼ����Ա ,A.����ID" & vbNewLine & _
            "      From (Select Max(m.ͣ��) As ͣ��, a.No As ���ݺ�," & vbNewLine & _
            "                   To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI') As ԤԼʱ��, Decode(a.�ű�, Null, Null, '[' || a.�ű� || ']') || Nvl(c.����,Nvl(c1.����,c2.����)) As �ű�," & vbNewLine & _
            "                   a.����, a.ִ���� As ҽ��, Decode(Max(Decode(F.���ӱ�־, 1, 1, 0)), 1, '��', Null) As ����, a.�����, a.����, d.���֤��," & vbNewLine & _
            "                   d.��ͥ�绰 As ��ϵ�绰,D.�ֻ���, F.�ѱ� As �ѱ�, To_Char(Sum(F.ʵ�ս��), '9999990.00') As ���, a.ժҪ, " & vbNewLine & _
            "                   To_Char(a.�Ǽ�ʱ��, 'YYYY-MM-DD HH24:MI:SS') As �Ǽ�ʱ��, a.����Ա���� As �Һ�Ա,a.�˺������,a.�˺����ʱ��,a.��¼״̬ ,a.ԤԼ����Ա,A.����ID,0 as ����,0 as ���ʷ���" & vbNewLine & _
            "            From ���˹Һż�¼ A, ������Ϣ D, �ٴ������¼ B, �ٴ������Դ B1, �ҺŰ��� B2, �շ���ĿĿ¼ C, �շ���ĿĿ¼ C1, �շ���ĿĿ¼ C2, ������ü�¼ F," & vbNewLine & _
            "                 (Select A.ID, Max(1) As ͣ��" & vbNewLine & _
            "                  From ���˹Һż�¼ A, ������Ϣ D, �ٴ������¼ B, �շ���ĿĿ¼ C " & vbNewLine & _
            "                  Where a.�����¼ID = b.ID(+) And b.��Ŀid = c.Id(+) And a.����id = d.����id(+) And a.��¼���� = 2 And a.��¼״̬ = 1 And " & vbNewLine & _
             IIf(SQLCondition.Default = False, "  a.����ʱ�� Between [1] And [2] ", "                        a.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60  ") & vbNewLine & _
            "                        And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & vbNewLine & vbNewLine & _
            "                         And a.����ʱ�� Between B.ͣ�￪ʼʱ�� And B.ͣ����ֹʱ��" & vbNewLine & _
            "                  Group By A.ID) M " & vbNewLine & _
            "           Where a.�����¼ID = b.ID(+) And a.�ű�=b1.����(+) And b1.��Ŀid=c1.id(+) And a.�ű�=b2.����(+) And b2.��Ŀid=c2.id(+) And b.��Ŀid = c.Id(+) And a.����id = d.����id(+) And a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = F.No(+) And" & vbNewLine & _
            "                 A.Id = m.ID(+) And F.��¼����=4   " & strFilter & vbNewLine & _
            IIf(strTime = "", "", "                 And " & strSQL & " IN(" & strTime & ")") & vbNewLine & _
            IIf(SQLCondition.Default = False, " and  a.����ʱ�� Between [1] And [2] ", "                 And a.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 ") & vbNewLine & _
            "                 And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null)" & vbNewLine & _
            "           Group By a.No, a.����ʱ��, F.�ѱ�,a.�ű�, Nvl(c.����,Nvl(c1.����,c2.����)), a.����, a.ִ����, a.�����, a.����, d.���֤��, d.��ͥ�绰,d.�ֻ���, a.ժҪ, a.�Ǽ�ʱ��, a.����Ա����,a.�˺������,a.�˺����ʱ��,a.��¼״̬,a.ԤԼ����Ա ,A.����ID" & vbNewLine & _
            "           Order By a.����ʱ�� Desc) A, ���˹Һż�¼ B, ���ű� D " & vbNewLine & _
            "     Where a.���ݺ� = b.No And b.��¼���� = 2 And b.��¼״̬ = 1 And b.ִ�в���id = d.ID "
        End If
        
        If SQLCondition.Default Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
        Else
            With SQLCondition
                Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, .Patientid, .Doctor, .FactB, .FactE, .DeptID, .FeeType, .ItemType, .PatiName)
            End With
        End If
    End If
    
    vsThis.Clear
    vsThis.Rows = 2
    
    If mrsList.EOF Then
        Call SetHeader
        stbThis.Panels(2).Text = "��ǰ������û���κ�����"
        Call SetMenu(False)
    Else
        Set vsThis.DataSource = mrsList
        Call SetHeader
        stbThis.Panels(2) = "�� " & mrsList.RecordCount & " ������"
        If tbsType.SelectedItem.Key = "�Һ�" Then
            stbThis.Panels(2) = stbThis.Panels(2) & ",���ϼ�:" & Format(GetBillSum, "0.00") & "Ԫ(�����۵�" & Format(GetHJSum, "0.00") & "Ԫ)"
        End If
        Call SetMenu(True)
    End If
    
    If tbsType.SelectedItem.Key = "ԤԼ" Then
        Call vsThis_RowColChange
    End If
    Call SetRowColor
    Call SetMenuEnable  '���ò˵�
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetHJSum() As Currency
    Dim i As Long, lngCol As Long
    If tbsType.SelectedItem.Key = "�Һ�" Then
        lngCol = getColNum("�ܽ��")
    Else
        lngCol = getColNum("���")
    End If
    For i = 1 To vsThis.Rows - 1
        If vsThis.TextMatrix(i, getColNum("�շѵ�")) <> "" Then
            GetHJSum = GetHJSum + Val(vsThis.TextMatrix(i, lngCol))
        End If
    Next
End Function

Private Sub vsThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strVal As String
    If tbsType.SelectedItem.Key = "����" Then
        With Me.vsThis
             strVal = .TextMatrix(NewRow, getColNum("ͣ�ð���"))
             If strVal = "" Then
                .ForeColorSel = -2147483634
             Else
                .ForeColorSel = vbRed
             End If
        End With
    Else
        vsThis.ForeColorSel = -2147483634
    End If
End Sub

Private Sub vsThis_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strVal As String
    If tbsType.SelectedItem.Key = "����" Then
        With Me.vsThis
             strVal = .TextMatrix(.Row, getColNum("ͣ�ð���"))
             If strVal = "" Then
                .ForeColorSel = -2147483634
             Else
                .ForeColorSel = vbRed
             End If
        End With
    Else
        vsThis.ForeColorSel = -2147483634
    End If
End Sub

Private Function SetRowColor()
    '--------------------------------
    '������ɫ
    '--------------------------------
    Dim X As Long, i As Long, strVal As String
    For X = 1 To Me.vsThis.Rows - 1
          With Me.vsThis
               strVal = .TextMatrix(X, getColNum("��¼״̬"))
               If strVal = "2" Then
                  .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = &HFF&
               ElseIf strVal = "3" Then
                    .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = &HFF0000
               Else
                    .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = &H80000008
               End If
               
          End With
     Next
    
    If tbsType.SelectedItem.Key = "����" Then
        For X = 1 To Me.vsThis.Rows - 1
              With Me.vsThis
                   strVal = .TextMatrix(X, getColNum("ͣ�ð���"))
                   If strVal <> "" Then
                        .Cell(flexcpForeColor, X, 0, X, .Cols - 1) = vbRed
                   End If
                   If X = 1 Then
                    If strVal <> "" Then
                         .ForeColorSel = vbRed
                    Else
                         .ForeColorSel = -2147483634
                    End If
                   End If
              End With
         Next
    End If
End Function

Private Function GetBillSum() As Currency
    Dim i As Long, lngCol As Long
    If tbsType.SelectedItem.Key = "�Һ�" Then
        lngCol = getColNum("�ܽ��")
    Else
        lngCol = getColNum("���")
    End If
    If mbytCancel = 3 Then
        For i = 1 To vsThis.Rows - 1
             If Val(vsThis.TextMatrix(i, vsThis.ColIndex("��¼״̬"))) <> 2 Then GetBillSum = GetBillSum + Val(vsThis.TextMatrix(i, lngCol))
        Next
    Else
        For i = 1 To vsThis.Rows - 1
            GetBillSum = GetBillSum + Val(vsThis.TextMatrix(i, lngCol))
        Next
    End If
End Function

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Sub PrintBill(BytMode As Byte)
'���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
'bytMode=0-�ش�,1-����
    Dim strNO As String, str�Һ�ʱ�� As String
    Dim lng����ID As Long, lng����ID As Long, intInsure As Integer
    Dim blnVirtualPrint As Boolean, lngShareUseID As Long
    strNO = vsThis.TextMatrix(vsThis.Row, getColNum("���ݺ�"))
    
    If strNO = "" Then
        MsgBox "��ǰû�м�¼�����ش�Ʊ�ݣ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    If CheckBillExistReplenishData(strNO) Then
        MsgBox "ѡ��ĹҺż�¼������ҽ��������㣬�������ش򲹴������", vbInformation, gstrSysName
        Exit Sub
    End If
    str�Һ�ʱ�� = vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�ʱ��"))
    
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNO, 4, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
        
    If BytMode = 0 Then
        If Not BillOperCheck(1, vsThis.TextMatrix(vsThis.Row, getColNum("�Һ�Ա")), _
            CDate(str�Һ�ʱ��), "�ش�") Then Exit Sub
    Else
        If Trim(vsThis.TextMatrix(vsThis.Row, getColNum("����Ʊ��"))) <> "" Then
            MsgBox "��ǰ�����Ѵ�ӡ��Ʊ��,���ܽ��в���", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    lng����ID = GetBill����ID(strNO, 4, lng����ID)
    intInsure = ExistInsure(strNO)
    If intInsure <> 0 Then
        blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure)
    End If
        
    Dim blnStartFactUseType  As Boolean, strUseType As String
    
    If gblnSharedInvoice Then
        '�Һ�������Ʊ��:42703
        blnStartFactUseType = zlStartFactUseType("1")
        If blnStartFactUseType Then
            strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        End If
    End If
    If Not RePrintBill(Me, IIf(BytMode = 0, 3, 4), strNO, lng����ID, intInsure, blnVirtualPrint, strUseType, True) Then Exit Sub
                      
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If

End Sub

 
 

Private Sub mnuViewRefeshOptionItem_Click(index As Integer)
    Dim i As Integer
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = index
    Next
End Sub

Private Sub tbsType_Click()
    If Visible Then
        Call SaveFlexState(vsThis, App.ProductName & "\" & Me.Name)
    End If
    vsThis.ForeColorSel = -2147483634
    If Val(vsThis.Tag) = tbsType.SelectedItem.index Then Exit Sub
    vsThis.Tag = tbsType.SelectedItem.index
    Call SetHeader
    
    If Visible Then
        Call RestoreFlexState(vsThis, App.ProductName & "\" & Me.Name)
    End If
    
    If Visible Or tbsType.SelectedItem.Key <> "�Һ�" Then
        '�л��嵥ʱ�ָ�ȱʡֵ
        Unload frmRegistFilter
        mbytCancel = 1: mstrFilter = ""
        vsThis.Clear 1
        vsThis.Rows = 1
        SetMenu False '����: 50358
    End If
    
    If Visible Then vsThis.SetFocus
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.Hwnd)
End Sub
 

Private Function CallBackBalanceInterface(ByVal cllBalance As Collection, _
    ByVal lng�ҺŽ���ID As Long, ByVal lng���ѽ���ID As Long, _
    ByVal dblMoney As Double, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û��˽ӿ�
    '���:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String, strSwapGlideNO As String, strSwapMemo As String, str������Ϣ As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim bln���ѿ� As Boolean, lng�����ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim lng�Һų���ID As Long, lng�˿�����ID As Long, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
    If cllBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO,����ID
    bln���ѿ� = Val(cllBalance(1)(2)) = 1
    lng�����ID = cllBalance(1)(0)
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str���� = cllBalance(1)(1)
    strSwapGlideNO = cllBalance(1)(3)
    strSwapMemo = cllBalance(1)(4)
    If lng���ѽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||5|" & lng���ѽ���ID
    If lng�ҺŽ���ID <> 0 Then str������Ϣ = str������Ϣ & "||4|" & lng�ҺŽ���ID
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 3)
    
    
    If lng���ѽ���ID <> 0 Then
        strSQL = " Select ����ID,���ʷ��� From סԺ���ü�¼  Where ��¼����=5 And NO =(Select Max(NO) From סԺ���ü�¼ where ����ID=[1] and  ��¼����=5  )  and ��¼״̬=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���ѽ���ID)
        If rsTemp.EOF Then
            strErrMsg = "δ�ҵ��˿���Ϣ�����ܼ���": Exit Function
        End If
        lng�˿�����ID = Val(Nvl(rsTemp!����ID))
    End If
    
    If lng�ҺŽ���ID <> 0 Then
        strSQL = "Select ����ID From ������ü�¼  Where ��¼����=4 And NO =(Select Max(NO) From ������ü�¼ where ����ID=[1] and  ��¼����=4  )  and ��¼״̬=2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ҺŽ���ID)
        If rsTemp.EOF Then
            strErrMsg = "δ�ҵ��˺���Ϣ�����ܼ���": Exit Function
        End If
        lng�Һų���ID = Val(Nvl(rsTemp!����ID))
    End If
    
    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    If lng�˿�����ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||5|" & lng�˿�����ID
    If lng�Һų���ID <> 0 Then strSwapExtendInfor = strSwapExtendInfor & "||4|" & lng�Һų���ID
    If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
    strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ���˽���
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
    '       strCardNo-����
    '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
    '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
    '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
    '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
    '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       strSwapExtendInfor-���������׵���չ��Ϣ
    '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, lng�����ID, bln���ѿ�, str����, str������Ϣ, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    If lng�˿�����ID <> 0 Then
        '�����:58536
        If Not bln���ѿ� Then
            Call zlAddUpdateSwapSQL(False, lng�˿�����ID, lng�����ID, bln���ѿ�, str����, strSwapGlideNO, strSwapMemo, cllUpdate)
        End If
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�˿�����ID, lng�����ID, bln���ѿ�, str����, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    If lng�Һų���ID <> 0 Then
        Call zlAddUpdateSwapSQL(False, lng�Һų���ID, lng�����ID, bln���ѿ�, str����, strSwapGlideNO, strSwapMemo, cllUpdate)
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�Һų���ID, lng�����ID, bln���ѿ�, str����, strSwapExtendInfor, cllThreeSwap)
        End If
    End If
    CallBackBalanceInterface = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function


Private Sub InitActionType()
    '-------------------------
    '��ȡ �Ƿ�����˷�ʱ�εĴ���ʽ
    '�ж�����Ϊ �ҺŰ����б��Ƿ�������
    '-------------------------
    Dim strSQL       As String
    Dim rsTmp        As ADODB.Recordset
    strSQL = "Select 1  dt From  �ٴ������¼ Where �Ƿ��ʱ��=1 And Rownum < 2"
    On Error GoTo Hd
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    mactionType = t_��ͨ
    If rsTmp.RecordCount <> 0 Then mactionType = t_ʱ��
    Set rsTmp = Nothing
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
 
Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, _
    Optional strInvoiceNO As String = "", Optional strUseType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '       strUserType-ʹ�����
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-11-19 16:32:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng����ID = GetInvoiceGroupID(IIf(gblnSharedInvoice, 1, 4), intNum, lng����ID, glng�Һ�ID, strInvoiceNO, strUseType)
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(strUseType) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & strUseType & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(strUseType) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & strUseType & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣģ��
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub UnloadMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ж��Ϣģ��
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    If mobjMsgModule Is Nothing Then Exit Sub
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub



