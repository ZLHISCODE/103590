VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frm������� 
   Caption         =   "�������"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9570
   Icon            =   "frm�������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TabStrip tabSelect 
      Height          =   300
      Left            =   15
      TabIndex        =   6
      Tag             =   "1"
      Top             =   765
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   529
      MultiRow        =   -1  'True
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���и���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "һ�㸶��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Ԥ���� "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ƻ�����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ܾ�����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   6150
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":08CA
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":0AEA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":0D0A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":0F26
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":1146
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":1366
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":1582
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":179E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":19B8
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":1B12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":1D2E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":1F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":26C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   6750
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":2E42
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":3062
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":3282
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":349E
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":36BE
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":38DE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":3AFA
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":3D16
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":3F30
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":408A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":42AA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":44CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�������.frx":4C44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6015
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm�������.frx":53BE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11800
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
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1376
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   9570
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   11040
      NewRow1         =   0   'False
      MinHeight2      =   0
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
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
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "CheckSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Ԥ��"
               Key             =   "Check"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "CheckBack"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frm�������.frx":5C52
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   2535
      Left            =   -15
      TabIndex        =   8
      Top             =   3165
      Width           =   4560
      _cx             =   8043
      _cy             =   4471
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483648
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm�������.frx":5F6C
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
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1650
      Left            =   -15
      TabIndex        =   7
      Top             =   1110
      Width           =   9480
      _cx             =   16722
      _cy             =   2910
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483648
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm�������.frx":61D1
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
   Begin VSFlex8Ctl.VSFlexGrid vsAddition 
      Height          =   2535
      Left            =   4785
      TabIndex        =   9
      Top             =   3150
      Width           =   4560
      _cx             =   8043
      _cy             =   4471
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483648
      BackColorAlternate=   -2147483628
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm�������.frx":63D2
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
   Begin VB.Label lblHsc_s 
      Height          =   2865
      Left            =   5535
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   3150
      Width           =   60
   End
   Begin VB.Label lblVsc_s 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   2520
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   2775
      Width           =   1425
   End
   Begin VB.Label lblRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
      Height          =   180
      Left            =   30
      TabIndex        =   3
      Top             =   2895
      Width           =   3330
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBillPrePrint 
         Caption         =   "����Ԥ��(&V)"
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditu 
         Caption         =   "��������"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Begin VB.Menu mnuEditAddPayment 
            Caption         =   "���(&P)"
         End
         Begin VB.Menu mnuEditMultAdd 
            Caption         =   "��������(&M)"
         End
         Begin VB.Menu mnuEditAddScheme 
            Caption         =   "�ƻ����(&S)"
         End
         Begin VB.Menu mnuEditAddImprest 
            Caption         =   "Ԥ���(&I)"
         End
         Begin VB.Menu mnuEditAddSign 
            Caption         =   "��Ǹ��(&B)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "Ԥ��(&P)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCheckBack 
         Caption         =   "����(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "����(&K)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSavePrint 
         Caption         =   "���̴�ӡ(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewVerifyPrint 
         Caption         =   "��˴�ӡ(&V)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frm�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msngDownX As Single, msngDownY As Single
Private mstrFilter As String  '��������
Private mstrPrivs As String
Private mstr���� As String      '��Ӧ������
Private mlngModule As Long
Private mblnFirst As Boolean
Private mstrOthers() As String    '0-��¼״̬,1-��ʼ����,2-��������,3-��Ӧ��ID,4-�����,5-������,6-��ʼ��Ʊ��,7-������Ʊ��,8-Ʒ��
Private mdtStartDate As Date
Private mdtEndDate As Date
Private mdtVrfyStartDate As Date
Private mdtVrfyEndDate As Date
Private mint����Flag As Integer
Private mint�豸Flag As Integer
Private mblnԤ�� As Boolean     'True��Ԥ��  False����Ԥ��
Private mint��ʾ��λ As Integer '0����С��λ��  1�����λ

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call zlControl.IsCtrlSetFocus(vsList)
    Call vsList_GotFocus
End Sub

Private Sub Form_Load()
'    Dim strStart As String, strEnd As String
    Dim strReg As String
    Dim strOthers(0 To 9) As String     '0-��¼״̬,1-��ʼ����,2-��������,3-��Ӧ��ID,4-�����,5-������,6-��ʼ��Ʊ��,7-������Ʊ��,8-Ʒ��
    mstrOthers = strOthers
    '����24925 by lesfeng 2010-02-08
    mint����Flag = 0
    mint�豸Flag = 0
    
    mblnFirst = True
    'Ȩ�޿���
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    mstr���� = "0000"
    Call Ȩ�޿���
   '�ָ�����
    mnuViewSavePrint.Checked = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1
    mnuViewVerifyPrint.Checked = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule)) = 1, 1, 0) = 1
    
    mint��ʾ��λ = Val(zlDatabase.GetPara("��ʾ��λѡ��", glngSys, mlngModule))
    
    'Ԥ��
    mblnԤ�� = IIf(Val(zlDatabase.GetPara("һ�㸶����Ҫ����Ԥ��", glngSys, mlngModule)) = 1, True, False)
    If mblnԤ�� Then
        mnuEditLine0.Visible = True
        mnuEditCheck.Visible = True
        mnuEditCheckBack.Visible = True
        tlbThis.Buttons("CheckSeparate").Visible = True
        tlbThis.Buttons("Check").Visible = True
        tlbThis.Buttons("CheckBack").Visible = True
    End If
    
    'by lesfeng 2009-12-2 �����Ż�
    mdtStartDate = Format(DateAdd("d", -15, zlDatabase.Currentdate), "yyyy-MM-dd")
    mdtEndDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    mdtVrfyStartDate = "1901-01-01"
    mdtVrfyEndDate = "1901-01-01"
    
    lblRange.Caption = "��ѯ��Χ:" & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
    mstrFilter = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [1] And [2]"
    
'    strStart = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-MM-dd")
'    strEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
'    lblRange.Caption = "��ѯ��Χ:" & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    
'    mstrFilter = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between To_Date('" & strStart & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & strEnd & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    
    RestoreWinState Me, App.ProductName
    
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    
    
    '��ʼ������ؼ�
    Call initGrid
    Call GetHeadData
End Sub

Private Sub mnuEditAddSign_Click()
    '����27930 by lesfeng 2010-03-23
    Dim blnReturn As Boolean
    
    If InStr(1, mstrPrivs, ";��Ǹ���;") = 0 Then Exit Sub
    
    frm����༭.ShowCard Me, g����, mstrPrivs, , , , blnReturn, 1
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditCheck_Click()
'Ԥ��
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln���� As Boolean
    Dim str��� As String
    Dim int��� As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) <> "1" And (Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("�ƻ����"))) = 0 Or GetMultiPayment(strNO) = True) Then
        'Ԥ��
        str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
        
        frm����༭.ShowCard Me, gԤ��, mstrPrivs, strNO, , , blnSuccess, IIf(str��� = "���", 1, 0)
        If blnSuccess = False Then Exit Sub
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditCheckBack_Click()
'Ԥ�����
    Dim strTmp As String, strNO As String
    Dim intRow As Integer
    
    If mnuEditCheckBack.Visible = False Then Exit Sub
    If vsList.Rows <= 1 Then Exit Sub
    
    On Error GoTo errHandle
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    If TestCheck(1, strNO) Then
        MsgBox "�õ����ѱ�ɾ����", vbInformation, gstrSysName
        intRow = vsList.Row
        mnuViewRefresh_Click
        Exit Sub
    End If
    If TestCheck(2, strNO) Then
        MsgBox "�õ����ѱ���ˣ�", vbInformation, gstrSysName
        intRow = vsList.Row
        mnuViewRefresh_Click
        Exit Sub
    End If
    
    If MsgBox("�Ƿ񽫵���Ԥ����ˣ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strTmp = "zl_�������_CancelCheck('" & strNO & "')"
    Call zlDatabase.ExecuteProcedure(strTmp, "Ԥ�����")
    mnuViewRefresh_Click
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditMultAdd_Click()
    '����:�������Ӹ��
    If frm����������������.ShowCard(Me, mstrPrivs) = False Then Exit Sub
    'ˢ�µ���
    Call mnuViewRefresh_Click
End Sub

Private Sub mnueditu_Click()
    Call frm�����������.���ò���(Me, glngModul, mstrPrivs)
    'Ԥ��
    mblnԤ�� = IIf(Val(zlDatabase.GetPara("һ�㸶����Ҫ����Ԥ��", glngSys, mlngModule)) = 1, True, False)
    mnuEditLine0.Visible = mblnԤ��
    mnuEditCheck.Visible = mblnԤ��
    mnuEditCheckBack.Visible = mblnԤ��
    tlbThis.Buttons("CheckSeparate").Visible = mblnԤ��
    tlbThis.Buttons("Check").Visible = mblnԤ��
    tlbThis.Buttons("CheckBack").Visible = mblnԤ��
    Call Form_Activate
    mint��ʾ��λ = Val(zlDatabase.GetPara("��ʾ��λѡ��", glngSys, mlngModule))
    
    With vsList
        If tabSelect.SelectedItem.Index = 1 Or tabSelect.SelectedItem.Index = 2 Then
            .ColHidden(.ColIndex("Ԥ����")) = Not mblnԤ��
            .ColHidden(.ColIndex("Ԥ������")) = Not mblnԤ��
        Else
            .ColHidden(.ColIndex("Ԥ����")) = True
            .ColHidden(.ColIndex("Ԥ������")) = True
        End If
    End With
    Call vsList_Click
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNO As String, lngԤ���� As Long, lng��¼״̬ As Long, lng��Ӧ��ID As Long
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    lngԤ���� = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")))
    lng��¼״̬ = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("��¼״̬")))
    lng��Ӧ��ID = Val(vsList.Cell(flexcpData, vsList.Row, vsList.ColIndex("��Ӧ������")))
    
    '2006-04-25:���˺�:�����Զ��屨������ģ��Ĺ���
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNO, "Ԥ����=" & lngԤ����, "��¼״̬=" & lng��¼״̬, "��Ӧ��=" & lng��Ӧ��ID)
End Sub

Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ĭ������
    '����:���˺�
    '����:2009-02-11 11:35:41
    '-----------------------------------------------------------------------------------------------------------
    Call zl_vsGrid_Para_Restore(mlngModule, vsList, Me.Caption, "�����ͷ�б�", True)
    Call zl_vsGrid_Para_Restore(mlngModule, vsDetail, Me.Caption, "������ϸ�б�", True)
    Call zl_vsGrid_Para_Restore(mlngModule, vsAddition, Me.Caption, "���ʽ�б�", True)
    Call vsDetail_LostFocus
    Call vsAddition_LostFocus
    Call vsList_LostFocus
    
    With vsList
        .Clear 1
        .Rows = 2
        .ColHidden(.ColIndex("��¼״̬")) = True: .ColWidth(.ColIndex("��¼״̬")) = 0
        .ColHidden(.ColIndex("Ԥ����")) = True: .ColWidth(.ColIndex("Ԥ����")) = 0
        .ColHidden(.ColIndex("Ԥ����")) = Not mblnԤ��
        .ColHidden(.ColIndex("Ԥ������")) = Not mblnԤ��
        
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("���ݺ�")) = "1||0"
        .ColData(.ColIndex("Ԥ����")) = "-1||0"
        .ColData(.ColIndex("��¼״̬")) = "-1||0"
        '����27930 by lesfeng 2010-03-23
        .ColData(.ColIndex("�ܸ����")) = "1||0"
    End With
    With vsDetail
        .Clear 1
        .Rows = 2
        .ColData(.ColIndex("Ʒ��")) = "1||0"
        .ColData(.ColIndex("��ⵥ��")) = "1||0"
        .ColData(.ColIndex("��Ʊ��")) = "1||0"
        .ColData(.ColIndex("��Ʊ���")) = "1||0"
        .ColHidden(.ColIndex("Ԥ��")) = mblnԤ��
    End With
    With vsAddition
        .Clear 1
        .Rows = 2
        .ColData(.ColIndex("���㷽ʽ")) = "1||0"
    End With
End Sub

Private Sub GetHeadData()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡͷ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-11 11:40:03
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, i As Long, lngRow As Long, str���� As String, strȨ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strStore As String
    
    Err = 0: On Error GoTo ErrHand:
    
    str���� = ""
    For i = 1 To Len(mstr����)
        If Mid(mstr����, i, 1) = 1 Then str���� = str���� & " or substr(b.����," & i & ",1)=1"
    Next
    If str���� <> "" Then str���� = " And (" & Mid(str����, 4) & ") "
    
    strȨ�� = " and " & Get����Ȩ��(mstrPrivs, "b.")
    
    
    Call zlCommFun.ShowFlash("�������������¼,���Ժ� ...", Me)
    DoEvents
    '����27930 by lesfeng 2010-03-23
    Screen.MousePointer = vbHourglass
    Select Case tabSelect.SelectedItem.Index
        Case 1
            strWhere = ""
        Case 2
            strWhere = "" & _
                " And A.Ԥ����<>1 And A.�ܸ���־<>1 And " & _
                " A.������� Not In (Select Distinct ������� " & _
                "                    From Ӧ����¼ " & _
                "                    Where ��¼����=-1 And ������� Is Not Null)"
        Case 3
            strWhere = " And A.Ԥ����=1"
        Case 4
            strWhere = "" & _
                " And A.�ܸ���־ = 0 And ( A.������� In (Select Distinct ������� " & _
                "                    From Ӧ����¼ " & _
                "                    Where ��¼����=-1 And ������� Is Not Null)  and a.Ԥ���� <>1) "
        Case 5
            strWhere = "" & _
                " And A.�ܸ���־ = 1 "
    End Select
    
    strStore = ",(Select NO, f_List2str(Cast(Collect(����) As t_Strlist)) ��Դ�ⷿ " & _
               "  From (Select Distinct a.No, c.���� " & _
               "        From �����¼ A, Ӧ����¼ B, ���ű� C " & _
               "        Where a.������� = b.������� And b.�ⷿid = c.Id And b.�ⷿid Is Not Null " & mstrFilter & _
               "        Order by c.����) " & _
               "  Group By NO ) C "
    
    strSQL = "" & _
        "   SELECT  a.no as ���ݺ�,b.id as ��Ӧ��ID, b.���� as ��Ӧ������,nvl(Ԥ����,0) as Ԥ���� ," & _
        "           ltrim(to_char(SUM (a.���),'9999999999999990.00')) AS ������, " & _
        "           a.������ AS ������,TO_CHAR (min(a.��������), 'yyyy-MM-dd') AS ��������," & _
        "           a.Ԥ����,TO_CHAR (min(a.Ԥ������), 'yyyy-MM-dd') AS Ԥ������," & _
        "           a.�����,TO_CHAR (min(a.�������), 'yyyy-MM-dd') AS �������," & _
        "           decode(a.�ܸ���־,1,'�ܸ�','����') as �ܸ����,a.��¼״̬,max(��Դ�ⷿ) ��Դ�ⷿ, a.ժҪ " & _
        "   FROM �����¼ a, ��Ӧ�� b " & _
        strStore & _
        "   Where a.��λid = b.id and a.NO=c.NO(+) " & zl_��ȡվ������(True, "b") & "  " & str���� & strWhere & mstrFilter & strȨ�� & _
        "   GROUP BY a.no,b.id,b.����,nvl(Ԥ����,0),a.������,a.Ԥ����,a.�����,'',a.��¼״̬, '',a.�ܸ���־, a.ժҪ " & _
        "   ORDER BY a.no desc "
    'by lesfeng 2009-12-2 �����Ż�
    'mstrOthers(0 To 9) '0-��¼״̬,1-��ʼ����,2-��������,3-��Ӧ��ID,4-�����,5-������,6-��ʼ��Ʊ��,7-������Ʊ��,8-Ʒ��,9-�ⷿID
    '������Χ: 1-��ʼ��������,2-������������
    '          3-��ʼ�������,4-�����������
    '          5-��ʼ����,6-��������,7-��Ӧ��ID,8-�����,9-������,10-��ʼ��Ʊ��,11-������Ʊ��,12-Ʒ��,13-�ⷿID
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(mdtStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtEndDate, "yyyy-mm-dd") & " 23:59:59"), _
                     CDate(Format(mdtVrfyStartDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(mdtVrfyEndDate, "yyyy-mm-dd") & " 23:59:59"), _
                     mstrOthers(1), mstrOthers(2), Val(mstrOthers(3)), mstrOthers(4), mstrOthers(5), mstrOthers(6), mstrOthers(7), mstrOthers(8), _
                     Val(mstrOthers(9)))
    With vsList
        .Redraw = flexRDNone
        .Rows = 2
        .Clear 1
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = .ForeColor
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        
        If tabSelect.SelectedItem.Index = 1 Or tabSelect.SelectedItem.Index = 2 Then
            .ColHidden(.ColIndex("Ԥ����")) = Not mblnԤ��
            .ColHidden(.ColIndex("Ԥ������")) = Not mblnԤ��
        Else
            .ColHidden(.ColIndex("Ԥ����")) = True
            .ColHidden(.ColIndex("Ԥ������")) = True
        End If
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Nvl(rsTemp!���ݺ�)
            .TextMatrix(lngRow, .ColIndex("��Ӧ������")) = Nvl(rsTemp!��Ӧ������)
            .Cell(flexcpData, lngRow, .ColIndex("��Ӧ������")) = Nvl(rsTemp!��Ӧ��ID)
            .TextMatrix(lngRow, .ColIndex("Ԥ����")) = Nvl(rsTemp!Ԥ����)
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(lngRow, .ColIndex("Ԥ����")) = Nvl(rsTemp!Ԥ����)
            .TextMatrix(lngRow, .ColIndex("Ԥ������")) = Nvl(rsTemp!Ԥ������)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("��¼״̬")) = Nvl(rsTemp!��¼״̬)
            .TextMatrix(lngRow, .ColIndex("�ܸ����")) = Nvl(rsTemp!�ܸ����)
            .TextMatrix(lngRow, .ColIndex("��Դ�ⷿ")) = Nvl(rsTemp!��Դ�ⷿ)
            .TextMatrix(lngRow, .ColIndex("ժҪ")) = Nvl(rsTemp!ժҪ)
            '������ر����ɫ
            If Val(Nvl(rsTemp!��¼״̬)) = 3 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000001
            ElseIf Val(Nvl(rsTemp!��¼״̬)) = 2 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &HFF
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = .ColIndex("���ݺ�")
    End With
    
    Full������ϸ
    Full������ϸ
    Call SetEnabled
    Call zlCommFun.StopFlash
    vsList.Redraw = flexRDBuffered
    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "��ǰ����" & rsTemp.RecordCount & "�ŵ���"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    vsList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
    Call zlCommFun.StopFlash
    staThis.Panels(2).Text = "��ǰ����" & 0 & "�ŵ���"
End Sub

Private Sub Full������ϸ()
    '-----------------------------------------------------------------------------------------------------------
    '����:���Ƶ�����ϸ����
    '����:
    '����:���˺�
    '����:2009-02-11 11:58:12
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, lng������� As Long, int״̬ As Integer
    Dim strNO As String, intԤ�� As Integer, lngRow As Long
    Dim int����Flag As Integer, int�豸Flag As Integer
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHand:
    With vsList
        int״̬ = Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
        intԤ�� = Val(.TextMatrix(.Row, .ColIndex("Ԥ����")))
    End With
    
    If strNO = "" Or int״̬ = 2 Or intԤ�� = 1 Then
        vsDetail.Clear 1: vsDetail.Rows = 2
        vsDetail.Cell(flexcpData, 1, 0, 1, vsDetail.Cols - 1) = ""
        Exit Sub
    End If
    
    
    strSQL = " Select ������� From �����¼ Where NO=[1] and ��¼״̬ in (1,3) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then
        lng������� = 0
    Else
        lng������� = Val(Nvl(rsTemp!�������))
    End If
    '����24925 by lesfeng 2010-02-08
    strTemp = "1,4,5"
    If GetShareSys(400) Then
        int����Flag = 1
    Else
        int����Flag = 0
        strTemp = strTemp & ",2"
    End If
    If GetShareSys(600) Then
        int�豸Flag = 1
    Else
        int�豸Flag = 0
        strTemp = strTemp & ",3"
    End If
    'ҩƷ�������������Լ�û�й�������ʡ��豸 ��Ҫ���Ӷ��豸��桢���ʿ�����Ȩ�ޣ�ͬʱ�����Ѿ�����Ȩ�޵����
    strSQL = IIf(mblnԤ��, "select a.*, decode(b.Ԥ��, 1, '��', null) Ԥ�� from (", "") & _
    "   Select Max(A.��ⵥ�ݺ�) ��ⵥ�ݺ�, Max(B1.�����) As �����, To_Char(Max(B1.�������), 'yyyy-mm-dd') As �������, " & _
    "          Max(A.ID) As ID, decode(a.��¼����,2,null,a.�ƻ����) �ƻ����, A.��Ʊ��, " & _
    "          To_Char(Sum(Nvl(case when a.��Ʊ��� = a.�ƻ���� or a.�ƻ���� is null and a.��¼����<>2 then a.��Ʊ��� else a.�ƻ���� end, 0)), '999999999990.99') As ��Ʊ���, " & _
    "          To_Char(Sum(Nvl(A.�ƻ����, 0)), '999999999990.99') As �ƻ����, To_Char(A.�ƻ�����, 'yyyy-mm-dd') As �ƻ�����, " & _
    "          Max(A.Ʒ��) Ʒ��, Max(A.���) ���, Max(A.����) ����, Max(A.����) ����," & _
    IIf(mint��ʾ��λ = 1, "decode(a.ϵͳ��ʶ,1,max(e.ҩ�ⵥλ),5,max(f.��װ��λ),max(a.������λ)) ������λ, To_Char(Round(Sum(Nvl(A.����, 0)) / decode(a.ϵͳ��ʶ,1,max(e.ҩ���װ),5,max(f.����ϵ��),1), 4), '999999999990.9999') As ����,", _
                          "max(A.������λ) ������λ, To_Char(Sum(Nvl(A.����, 0)), '999999999990.9999') As ����,") & _
    "          To_Char(Max(A.�ɹ���), '999999999990.9999') As �ɹ���, " & _
    "          To_Char(Sum(Nvl(A.�ɹ����, 0)), '999999999990.9999') As �ɹ����," & _
    IIf(mint��ʾ��λ = 1, "To_Char(round(Sum(Nvl(D.�������, 0)) / decode(a.ϵͳ��ʶ,1,max(e.ҩ���װ),5,max(f.����ϵ��),1), 4), '999999999990.9999')", _
                          "To_Char(Sum(Nvl(D.�������, 0)), '999999999990.9999')") & " As �������" & _
    "   From Ӧ����¼ A, " & _
    "        (Select B.ID, Max(�����) As �����, Max(�������) As ������� " & _
    "          From Ӧ����¼ B, (Select Distinct ID From Ӧ����¼ Where ������� = [1]) C " & _
    "          Where B.ID = C.ID Group By B.ID) B1," & _
    "        (Select ҩƷid,sum(��������) As ��������,Sum(ʵ������) As �������,Sum(ʵ�ʽ��) As ʵ�ʽ��,�ϴ�����,�ϴι�Ӧ��id " & _
    "          From ҩƷ��� Group By ҩƷid,�ϴ�����,�ϴι�Ӧ��id) D" & _
    IIf(mint��ʾ��λ = 1, ",ҩƷ��� E, �������� F ", "") & _
    "   Where A.������� = [1] And A.ID = B1.ID And A.��Ŀid = D.ҩƷid(+) And A.����= D.�ϴ�����(+) And A.��λid = D.�ϴι�Ӧ��id(+) " & _
    "     And nvl(A.ϵͳ��ʶ,4) In (" & strTemp & ")" & _
    IIf(mint��ʾ��λ = 1, " And a.��Ŀid=e.ҩƷid(+) and a.��Ŀid=f.����id(+) ", "") & _
    "   Group By A.ϵͳ��ʶ, A.��ⵥ�ݺ�, A.��Ŀid, A.���, decode(a.��¼����,2,null,a.�ƻ����), A.�ƻ�����, A.��Ʊ��"
    '���ʲ���
    If int����Flag = 1 Then
        strSQL = strSQL & " Union All " & _
        "   Select Max(A.��ⵥ�ݺ�) ��ⵥ�ݺ�, Max(B1.�����) As �����, To_Char(Max(B1.�������), 'yyyy-mm-dd') As �������, " & _
        "          Max(A.ID) As ID, decode(a.��¼����,2,null,a.�ƻ����) �ƻ����, A.��Ʊ��, " & _
        "          To_Char(Sum(Nvl(case when a.��Ʊ��� = a.�ƻ���� or a.�ƻ���� is null and a.��¼����<>2 then a.��Ʊ��� else a.�ƻ���� end, 0)), '999999999990.99') As ��Ʊ���, " & _
        "          To_Char(Sum(Nvl(A.�ƻ����, 0)), '999999999990.99') As �ƻ����, To_Char(A.�ƻ�����, 'yyyy-mm-dd') As �ƻ�����, " & _
        "          Max(A.Ʒ��) Ʒ��, Max(A.���) ���, Max(A.����) ����, Max(A.����) ����, " & _
        IIf(mint��ʾ��λ = 1, "Max(E.��װ��λ) ������λ, To_Char(round(Sum(Nvl(A.����, 0)) / max(e.����ϵ��), 4), '999999999990.9999') As ����, ", _
                              "Max(A.������λ) ������λ, To_Char(Sum(Nvl(A.����, 0)), '999999999990.9999') As ����, ") & _
        "          To_Char(Max(A.�ɹ���), '999999999990.9999') As �ɹ���, " & _
        "          To_Char(Sum(Nvl(A.�ɹ����, 0)), '999999999990.9999') As �ɹ����," & _
        IIf(mint��ʾ��λ = 1, "To_Char(round(Sum(Nvl(D.�������, 0)) / max(e.����ϵ��), 4), '999999999990.9999') As ������� ", _
                              "To_Char(Sum(Nvl(D.�������, 0)), '999999999990.9999') As ������� ") & _
        "   From Ӧ����¼ A, " & _
        "        (Select B.ID, Max(�����) As �����, Max(�������) As ������� " & _
        "          From Ӧ����¼ B, (Select Distinct ID From Ӧ����¼ Where ������� = [1]) C " & _
        "          Where B.ID = C.ID Group By B.ID) B1," & _
        "        (Select ����id,sum(��������) As ��������,Sum(ʵ������) As �������,Sum(ʵ�ʽ��) As ʵ�ʽ��,�ϴ�����,�ϴι�Ӧ��id " & _
        "          From ���ʿ�� Group By ����id,�ϴ�����,�ϴι�Ӧ��id) D" & _
        IIf(mint��ʾ��λ = 1, ",����Ŀ¼ E ", "") & _
        "   Where A.������� = [1] And A.ID = B1.ID And A.��Ŀid = D.����id(+) And A.����= D.�ϴ�����(+) And A.��λid = D.�ϴι�Ӧ��id(+) " & _
        "     And A.ϵͳ��ʶ = 2 " & _
        IIf(mint��ʾ��λ = 1, " And a.��Ŀid=e.ID ", "") & _
        "   Group By A.ϵͳ��ʶ, A.��ⵥ�ݺ�, A.��Ŀid, A.���, decode(a.��¼����,2,null,a.�ƻ����), A.�ƻ�����, A.��Ʊ��"
    End If
    '�豸����
    If int�豸Flag = 1 Then
        strSQL = strSQL & " Union All " & _
        "   Select Max(A.��ⵥ�ݺ�) ��ⵥ�ݺ�, Max(B1.�����) As �����, To_Char(Max(B1.�������), 'yyyy-mm-dd') As �������, " & _
        "          Max(A.ID) As ID, decode(a.��¼����,2,null,a.�ƻ����) �ƻ����, A.��Ʊ��, " & _
        "          To_Char(Sum(Nvl(case when a.��Ʊ��� = a.�ƻ���� or a.�ƻ���� is null and a.��¼����<>2 then a.��Ʊ��� else a.�ƻ���� end, 0)), '999999999990.99') As ��Ʊ���, " & _
        "          To_Char(Sum(Nvl(A.�ƻ����, 0)), '999999999990.99') As �ƻ����, To_Char(A.�ƻ�����, 'yyyy-mm-dd') As �ƻ�����, " & _
        "          Max(A.Ʒ��) Ʒ��, Max(A.���) ���, Max(A.����) ����, Max(A.����) ����, Max(A.������λ) ������λ, " & _
        "          To_Char(Sum(Nvl(A.����, 0)), '999999999990.9999') As ����, To_Char(Max(A.�ɹ���), '999999999990.9999') As �ɹ���, " & _
        "          To_Char(Sum(Nvl(A.�ɹ����, 0)), '999999999990.9999') As �ɹ����,To_Char(Sum(Nvl(D.�������, 0)), '999999999990.9999') As ������� " & _
        "   From Ӧ����¼ A, " & _
        "        (Select B.ID, Max(�����) As �����, Max(�������) As ������� " & _
        "          From Ӧ����¼ B, (Select Distinct ID From Ӧ����¼ Where ������� = [1]) C " & _
        "          Where B.ID = C.ID Group By B.ID) B1," & _
        "        (Select �豸id,sum(��������) As ��������,Sum(ʵ������) As �������,Sum(ʵ�ʽ��) As ʵ�ʽ��,����,�ϴι�Ӧ��id " & _
        "          From �豸��� Group By �豸id,����,�ϴι�Ӧ��id) D" & _
        "   Where A.������� = [1] And A.ID = B1.ID And A.��Ŀid = D.�豸id(+) And A.����= D.����(+) And A.��λid = D.�ϴι�Ӧ��id(+) " & _
        "     And A.ϵͳ��ʶ = 3 " & _
        "   Group By A.ϵͳ��ʶ, A.��ⵥ�ݺ�, A.��Ŀid, A.���, decode(a.��¼����,2,null,a.�ƻ����), A.�ƻ�����, A.��Ʊ��"
    End If
    
    If mblnԤ�� Then
        strSQL = strSQL & ") A, Ӧ����¼ B where a.ID=b.ID(+) " 'and b.��¼����(+) <> 2
        If lng������� <> 0 Then
            strSQL = strSQL & " and b.�������(+) = [1] "
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�������)
    
    With vsDetail
        If mblnԤ�� And tabSelect.SelectedItem.Index = 2 Then
            .ColHidden(.ColIndex("Ԥ��")) = False
        Else
            .ColHidden(.ColIndex("Ԥ��")) = True
        End If
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            If mblnԤ�� Then
                .TextMatrix(lngRow, .ColIndex("Ԥ��")) = Nvl(rsTemp!Ԥ��)
            End If
            .TextMatrix(lngRow, .ColIndex("��ⵥ��")) = Nvl(rsTemp!��ⵥ�ݺ�)
            .Cell(flexcpData, lngRow, .ColIndex("��ⵥ��")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("��Ʊ��")) = Nvl(rsTemp!��Ʊ��)
            .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Nvl(rsTemp!��Ʊ���)
            
            .TextMatrix(lngRow, .ColIndex("�ƻ����")) = Nvl(rsTemp!�ƻ����)
            .TextMatrix(lngRow, .ColIndex("�ƻ�����")) = Nvl(rsTemp!�ƻ�����)
            .TextMatrix(lngRow, .ColIndex("�ƻ����")) = Nvl(rsTemp!�ƻ����)
            .TextMatrix(lngRow, .ColIndex("Ʒ��")) = Nvl(rsTemp!Ʒ��)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("������λ")) = Nvl(rsTemp!������λ)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("�ɹ���")) = Nvl(rsTemp!�ɹ���)
            .TextMatrix(lngRow, .ColIndex("�ɹ����")) = Nvl(rsTemp!�ɹ����)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    vsDetail.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Full������ϸ()
    '-----------------------------------------------------------------------------------------------------------
    '����:��丶����ϸ
    '����:���˺�
    '����:2009-02-11 13:36:56
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, int״̬ As Integer, intԤ�� As Integer, strNO As String
    Dim lng������� As Long, lngRow As Long
    
    
    Err = 0: On Error GoTo ErrHand:
    With vsList
        int״̬ = Val(.TextMatrix(.Row, .ColIndex("��¼״̬")))
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
        intԤ�� = Val(.TextMatrix(.Row, .ColIndex("Ԥ����")))
    End With
    
    If strNO = "" Then
       vsAddition.Rows = 2: vsAddition.Clear 1:
       vsAddition.Cell(flexcpData, 1, 0, 1, vsAddition.Cols - 1) = ""
        Exit Sub
    End If
    '����27930 by lesfeng 2010-03-23
    If int״̬ <> 1 Then
        '����
        strSQL = "" & _
            "   Select Decode(Ԥ����,1,'��','��') as Ԥ����,to_char(���,'99999999999.99') as ���,���㷽ʽ,�������,Decode(Ԥ����,1,NO,'')  as ���Ԥ�����," & _
            "          decode(�ܸ���־,1,'�ܸ�','����') as �ܸ��� " & _
            "   From �����¼ " & _
            "   Where NO=[1] And ��¼״̬=[2]"
    ElseIf intԤ�� = 1 Then
        '����
        strSQL = "" & _
            "   Select Decode(Ԥ����,1,'��','��') as Ԥ����,to_char(���,'99999999999.99') as ���,���㷽ʽ,�������,Decode(Ԥ����,1,NO,'') as ���Ԥ�����, " & _
            "          decode(�ܸ���־,1,'�ܸ�','����') as �ܸ��� " & _
            "   From �����¼ " & _
            "   Where NO=[1] And ��¼״̬=[2]"
    Else
        '������
        strSQL = "Select ������� From �����¼ Where NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        
        If rsTemp.EOF Then
            lng������� = 0
        Else
            lng������� = Nvl(rsTemp!�������, 0)
        End If
        
        strSQL = "" & _
            "   Select Decode(Ԥ����,1,'��','��') as Ԥ����,to_char(���,'99999999999.99') as ���,���㷽ʽ,�������,Decode(Ԥ����,1,NO,'') as ���Ԥ�����," & _
            "          decode(�ܸ���־,1,'�ܸ�','����') as �ܸ��� " & _
            "   From �����¼ " & _
            "   Where �������=[3]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, int״̬, lng�������)
    With vsAddition
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            '.TextMatrix(lngRow, .ColIndex("�����־")) = Nvl(rsTemp!�����־)
            .TextMatrix(lngRow, .ColIndex("Ԥ����")) = Nvl(rsTemp!Ԥ����)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!���㷽ʽ)
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("���Ԥ�����")) = Nvl(rsTemp!���Ԥ�����)
            .TextMatrix(lngRow, .ColIndex("�ܸ����")) = Nvl(rsTemp!�ܸ���)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    vsAddition.Redraw = flexRDBuffered
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
        
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 5000 Then
            Me.Height = 5000
        End If
        If Me.Width < 4500 Then
            Me.Width = 4500
        End If
    End If

    cbrTool.Width = Me.ScaleWidth
    lblVsc_s.Left = 0
    lblVsc_s.Width = Me.ScaleWidth
    
    If lblVsc_s.Top > Me.ScaleHeight - 2000 Then lblVsc_s.Top = Me.ScaleHeight - 2000
    
    tabSelect.Top = IIf(cbrTool.Visible, cbrTool.Height + 30, 0)
    
    vsList.Top = tabSelect.Top + tabSelect.Height + 30
    vsList.Width = Me.ScaleWidth
    vsList.Height = lblVsc_s.Top - vsList.Top
    
    lblRange.Move 30, lblVsc_s.Top + 75, Me.ScaleWidth
    
    lblHsc_s.Top = lblVsc_s.Top + lblVsc_s.Height
    lblHsc_s.Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - lblHsc_s.Top
    
    If lblHsc_s.Left > Me.ScaleWidth - 2000 Then lblHsc_s.Left = Me.ScaleWidth - 2000
    
    vsDetail.Move 0, lblHsc_s.Top, lblHsc_s.Left, lblHsc_s.Height
    vsAddition.Move lblHsc_s.Left + lblHsc_s.Width, lblHsc_s.Top, Me.ScaleWidth - lblHsc_s.Left - lblHsc_s.Width, lblHsc_s.Height
    
    mnuViewToolButton.Checked = cbrTool.Visible
    mnuViewStatus.Checked = staThis.Visible
    mnuViewToolText.Checked = tlbThis.Buttons(1).Caption <> ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call zl_vsGrid_Para_Save(mlngModule, vsList, Me.Caption, "�����ͷ�б�", True)
    Call zl_vsGrid_Para_Save(mlngModule, vsDetail, Me.Caption, "������ϸ�б�", True)
    Call zl_vsGrid_Para_Save(mlngModule, vsAddition, Me.Caption, "���ʽ�б�", True)
End Sub

Private Sub lblHsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
End Sub

Private Sub lblHsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblHsc_s
            If .Left + X - msngDownX < 2000 Then Exit Sub
            If .Left + X - msngDownX > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + X - msngDownX
        End With
        
        Me.vsDetail.Width = lblHsc_s.Left
        Me.vsAddition.Left = lblHsc_s.Left + lblHsc_s.Width
        Me.vsAddition.Width = Me.ScaleWidth - Me.vsAddition.Left
    End If
End Sub

Private Sub lblVsc_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownY = Y
End Sub

Private Sub lblVsc_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With lblVsc_s
            If .Top + Y - msngDownY < 2000 Then Exit Sub
            If .Top + Y - msngDownY > ScaleHeight - 2000 Then Exit Sub
            .Top = .Top + Y - msngDownY
        End With
        Form_Resize
    End If
End Sub

Private Sub mnuEditAddImprest_Click()
    Dim strNO As String, blnSuccess As Boolean
    
    If InStr(1, mstrPrivs, ";Ԥ��;") = 0 Then Exit Sub
    strNO = ""
    frmDrugImprestCard.ShowCard Me, strNO, 1, , blnSuccess
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditAddScheme_Click()
    Dim blnReturn As Boolean
    
    '�ƻ�����
    frm�ƻ�����༭.ShowCard Me, False, g����, mstrPrivs, , , , blnReturn
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditDel_Click()
    Dim strBillNo As String, strTitle As String, intReturn As Integer
    Dim str��� As String
    Dim int��� As Integer
    
    With vsList
        If Val(.TextMatrix(.Row, .ColIndex("Ԥ����"))) = 1 Then
            strTitle = "Ԥ����"
            If InStr(1, mstrPrivs, ";Ԥ��;") = 0 Then Exit Sub
        Else
            '����27930 by lesfeng 2010-03-23
            str��� = Trim(vsList.TextMatrix(.Row, .ColIndex("�ܸ����")))
            If str��� = "���" Then
                strTitle = "�ܸ����"
                If InStr(1, mstrPrivs, ";��Ǹ���;") = 0 Then Exit Sub
            Else
                strTitle = "����"
            End If
        End If
        
        strBillNo = .TextMatrix(.Row, .ColIndex("���ݺ�"))
        
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & strBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        If intReturn <> vbYes Then Exit Sub
        gstrSQL = "zl_�����¼_delete('" & strBillNo & "')"
        
        Err = 0: On Error GoTo ErrHand:
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        mnuViewRefresh_Click
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditDisplay_Click()
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln���� As Boolean
    Dim bytRec As RecBillStatus
    Dim int��¼״̬  As Integer
    Dim str��� As String
    Dim int��� As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    If strNO = "" Then Exit Sub
    
    If Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����"))) = 1 Then
        int��¼״̬ = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("��¼״̬")))
        frmDrugImprestCard.ShowCard Me, strNO, 4, int��¼״̬, blnSuccess
    Else
        
        int��¼״̬ = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("��¼״̬")))
    
        Select Case int��¼״̬
        Case 1
            bytRec = ������¼
        Case 2
            bytRec = ������¼
        Case Else
            bytRec = ��������¼
        End Select
        
        bln���� = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("�ƻ����"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("��ⵥ��"))) <> ""
        If bln���� Or IsPlanPayment(strNO) = False Then
            '����27930 by lesfeng 2010-03-23
            str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
            int��� = 0
            If str��� = "���" Then int��� = 1
            frm����༭.ShowCard Me, g�鿴, mstrPrivs, strNO, , bytRec, blnSuccess, int���
        Else
            frm�ƻ�����༭.ShowCard Me, bln����, g�鿴, mstrPrivs, strNO, , bytRec, blnSuccess
        End If
    End If
End Sub

Private Function IsPlanPayment(ByVal strNO As String) As Boolean
'���ܣ��ж��Ƿ�Ϊ�ƻ������
'������strNO���ݺ�
'���أ�True�ƻ�����ݣ�False�Ǽƻ������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Count(1) Rec From Ӧ����¼ A, �����¼ B Where a.������� = b.������� And a.��¼���� = -1 And b.No = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�Ƿ�Ϊ�ƻ������", strNO)
    If rsTmp!rec > 0 Then
        IsPlanPayment = True
    Else
        IsPlanPayment = False
    End If
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub mnuEditModify_Click()
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln���� As Boolean
    Dim str��� As String
    Dim int��� As Integer
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) = "1" Then
        If InStr(1, mstrPrivs, ";Ԥ��;") = 0 Then Exit Sub
        strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
        frmDrugImprestCard.ShowCard Me, strNO, 2, , blnSuccess
        
    Else
        bln���� = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("�ƻ����"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("��ⵥ��"))) <> ""
        strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
        
        If bln���� Or IsPlanPayment(strNO) = False Then
            '����27930 by lesfeng 2010-03-23
            str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
            int��� = 0
            If str��� = "�ܸ�" Then int��� = 1
            frm����༭.ShowCard Me, g�޸�, mstrPrivs, strNO, , , blnSuccess, int���
        Else
            frm�ƻ�����༭.ShowCard Me, bln����, g�޸�, mstrPrivs, strNO, , , blnSuccess
        End If
        If blnSuccess = False Then Exit Sub
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditStrike_Click()
    Dim strNO As String
    Dim blnYes As Boolean
    Dim blnSuccess As Boolean
    Dim bln���� As Boolean
    Dim str��� As String
    Dim int��� As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    
    If Trim(strNO) = "" Then Exit Sub
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) = "1" Then
        ShowMsgbox "��ȷʵҪ�������ݺ�Ϊ��" & strNO & "���ĵ�����", True, blnYes
        If blnYes = False Then Exit Sub
        If StrikeSave = True Then Call mnuViewRefresh_Click
    Else
        bln���� = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("�ƻ����"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("��ⵥ��"))) <> ""
        If bln���� Or IsPlanPayment(strNO) = False Then
            '����27930 by lesfeng 2010-03-23
            str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
            int��� = 0
            If str��� = "���" Then int��� = 1
            frm����༭.ShowCard Me, gȡ��, mstrPrivs, strNO, , , blnSuccess, int���
        Else
            frm�ƻ�����༭.ShowCard Me, bln����, gȡ��, mstrPrivs, strNO, , , blnSuccess
        End If
        If blnSuccess = False Then Exit Sub
        mnuViewRefresh_Click
    End If
End Sub

Private Function StrikeSave() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-11 14:23:36
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    StrikeSave = False
    With vsList
        gstrSQL = "zl_�������_STRIKE('" & .TextMatrix(.Row, .ColIndex("���ݺ�")) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mnuEditVerify_Click()
    Dim strNO As String
    Dim blnSuccess As Boolean
    Dim bln���� As Boolean
    Dim str��� As String
    Dim int��� As Integer
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    If Trim(strNO) = "" Then Exit Sub
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) = "1" Then
        frmDrugImprestCard.ShowCard Me, strNO, 3, , blnSuccess
        If blnSuccess = True Then Call mnuViewRefresh_Click
    Else
        bln���� = Val(vsDetail.TextMatrix(1, vsDetail.ColIndex("�ƻ����"))) = 0 And Trim(vsDetail.TextMatrix(1, vsDetail.ColIndex("��ⵥ��"))) <> ""
        If bln���� Or IsPlanPayment(strNO) = False Then
            '����27930 by lesfeng 2010-03-23
            str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
            int��� = 0
            If str��� = "���" Then int��� = 1
            frm����༭.ShowCard Me, g���, mstrPrivs, strNO, , , blnSuccess, int���
        Else
            frm�ƻ�����༭.ShowCard Me, bln����, g���, mstrPrivs, strNO, , , blnSuccess
        End If
        If blnSuccess = False Then Exit Sub
        Call mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuFileBillPrePrint_Click()
    printbill 1
End Sub

Private Sub mnuFileBillPrint_Click()
    printbill 0
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    subPrint 1
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    '�˳�
    Unload Me
End Sub

Private Sub mnuEditAddPayment_Click()
    Dim blnReturn As Boolean
    frm����༭.ShowCard Me, g����, mstrPrivs, , , , blnReturn
    If blnReturn = False Then Exit Sub
    mnuViewRefresh_Click
End Sub

Private Sub mnuViewRefresh_Click()
    Call GetHeadData
    Call vsList_GotFocus
End Sub

Private Sub mnuViewSearch_Click()
'    Dim strStart As Date
'    Dim strEnd As Date
'    Dim strVerifyStart As Date
'    Dim strVerifyEnd As Date
    Dim strFind As String
    Dim strType As String
    Dim strOthers() As String
    
    strFind = FrmDrugPaymentSearch.GetSearch(Me, mstrPrivs, mdtStartDate, mdtEndDate, mdtVrfyStartDate, mdtVrfyEndDate, strType, strOthers)
    
    If strFind <> "" Then
        mstr���� = strType
        mstrFilter = strFind
        mstrOthers = strOthers
        
        GetHeadData
        
        If Format(mdtStartDate, "yyyy-mm-dd") = "1901-01-01" And Format(mdtVrfyStartDate, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" And Format(mdtVrfyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��") & "  ������� " & Format(mdtVrfyStartDate, "yyyy��MM��dd��") & "��" & Format(mdtVrfyEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(mdtStartDate, "yyyy��MM��dd��") & "��" & Format(mdtEndDate, "yyyy��MM��dd��")
        ElseIf Format(mdtVrfyStartDate, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:������� " & Format(mdtVrfyStartDate, "yyyy��MM��dd��") & "��" & Format(mdtVrfyEndDate, "yyyy��MM��dd��")
        End If
    End If
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbThis.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .Count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .Count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbThis.Height
    
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
'    ReportMan gcnOracle, Me
   ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim lngCol As Long
    
    Set objPrint = New zlPrint1Grd
    
        
    objPrint.Title.Text = "�������"
    '��Ҫ������صĿ��
    With vsList
        .Redraw = flexRDNone
        For lngCol = 0 To vsList.Cols - 1
           If .ColHidden(lngCol) Then
                .Cell(flexcpData, 0, lngCol) = .ColWidth(lngCol)
                .ColWidth(lngCol) = 0
           End If
        Next
    End With
    Set objPrint.Body = vsList
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    '�ָ�
    With vsList
        For lngCol = 0 To vsList.Cols - 1
           If .ColHidden(lngCol) Then
                .ColWidth(lngCol) = Val(.Cell(flexcpData, 0, lngCol))
           End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsAddition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If vsAddition.MouseRow <= 0 Then
        Call ShowColSet(2)
    End If
End Sub
 
Private Sub vsDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If vsDetail.MouseRow <= 0 Then
        Call ShowColSet(1)
    End If
End Sub

Private Sub vsList_Click()
    Full������ϸ
    Full������ϸ
    SetEnabled
End Sub

Private Sub vsList_DblClick()
    mnuEditDisplay_Click
End Sub

Private Sub ShowColSet(ByVal bytType As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʾ������
    '����:bytType:0-��ͷ,1-������,2-δ����Ϣ
    '����:���˺�
    '����:2009-02-11 15:31:27
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngTop As Long, strKey As String
    Dim vRect  As RECT, objVsGrid As VSFlexGrid
    Dim lngCol As Long
    
    Select Case bytType
    Case 0
        Set objVsGrid = vsList: strKey = "�����ͷ�б�"
    Case 1
        Set objVsGrid = vsDetail: strKey = "������ϸ�б�"
    Case 2
        Set objVsGrid = vsAddition: strKey = "���ʽ�б�"
    Case Else
        Exit Sub
    End Select
    lngCol = objVsGrid.MouseCol
    
    If lngCol < 0 Then Exit Sub
    vRect = zlControl.GetControlRect(objVsGrid.hwnd)
    lngLeft = vRect.Left + objVsGrid.ColPos(lngCol)
    lngTop = vRect.Top + objVsGrid.RowHeight(0) + 100
    Call frmVsColSel.ShowColSet(Me, Me.Caption, objVsGrid, lngLeft, lngTop, objVsGrid.RowHeight(0))
    Call zl_vsGrid_Para_Save(mlngModule, objVsGrid, Me.Caption, strKey, True)
End Sub

Private Sub vsList_GotFocus()
        zl_VsGridGotFocus vsList
End Sub

Private Sub vsList_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsList)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsList, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    If vsList.MouseRow <= 0 Then
        Call ShowColSet(0)
    Else
        Me.PopupMenu mnuEdit
    End If
End Sub

Private Sub vsList_RowColChange()
    Full������ϸ
    Full������ϸ
    SetEnabled
End Sub

Private Sub tabSelect_Click()
    If tabSelect.SelectedItem.Index = tabSelect.Tag Then Exit Sub
    tabSelect.Tag = tabSelect.SelectedItem.Index
    vsList.SetFocus
    GetHeadData
    Call vsList_GotFocus
End Sub

Private Sub tlbthis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Add"
            Select Case tabSelect.SelectedItem.Index
                Case 1, 2
                    mnuEditAddPayment_Click
                Case 3
                    mnuEditAddImprest_Click
                Case 4
                    mnuEditAddScheme_Click
                 '����27930 by lesfeng 2010-03-23
                Case 5
                    mnuEditAddSign_Click
            End Select
        Case "Modify"
            mnuEditModify_Click
        Case "Check"
            mnuEditCheck_Click
        Case "CheckBack"
            mnuEditCheckBack_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "PrintView"
            mnuFilePreView_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            Unload Me
    End Select
End Sub

Private Sub tlbthis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuViewTool
End Sub

Private Sub SetEnabled()
    Dim blnData As Boolean
    Dim blnVerify As Boolean    '����˵ĵ���
    Dim blnCancel As Boolean    '�Ѿ������˵ĵ���
    Dim blnԤ�� As Boolean
    Dim blnVrfy As Boolean
    Dim blnDelete As Boolean
    Dim blnStrike As Boolean
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim bln��� As Boolean
    Dim bln�ƻ����� As Boolean
    Dim strNO As String
    
    blnData = vsList.TextMatrix(1, vsList.ColIndex("���ݺ�")) <> ""
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    blnVerify = vsList.TextMatrix(vsList.Row, vsList.ColIndex("�������")) <> ""
    blnCancel = Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("��¼״̬"))) <> 1
    blnԤ�� = vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) = "1" Or tabSelect.SelectedItem.Index = 3
    bln�ƻ����� = vsDetail.TextMatrix(1, vsDetail.ColIndex("�ƻ����")) = "1" Or tabSelect.SelectedItem.Index = 4
    '����27930 by lesfeng 2010-03-23
    bln��� = vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")) = "�ܾ�" Or tabSelect.SelectedItem.Index = 5
    
    If blnԤ�� Then
        blnModify = InStr(1, mstrPrivs, ";Ԥ��;") <> 0
        blnDelete = blnModify
        blnVrfy = blnModify
        blnStrike = blnModify
        blnAdd = (tabSelect.SelectedItem.Index = 3 And blnModify) Or (InStr(1, mstrPrivs, ";�Ǽ�;") <> 0 And tabSelect.SelectedItem.Index = 1)
    Else
        '����27930 by lesfeng 2010-03-23
        If bln��� Then
            blnModify = InStr(1, mstrPrivs, ";�޸�;") <> 0 And InStr(1, mstrPrivs, ";��Ǹ���;") <> 0
            blnDelete = InStr(1, mstrPrivs, ";ɾ��;") <> 0 And InStr(1, mstrPrivs, ";��Ǹ���;") <> 0
            blnVrfy = InStr(1, mstrPrivs, ";���;") <> 0 And InStr(1, mstrPrivs, ";��Ǹ���;") <> 0
            blnStrike = InStr(1, mstrPrivs, ";����;") <> 0 And InStr(1, mstrPrivs, ";��Ǹ���;") <> 0
            blnAdd = (tabSelect.SelectedItem.Index = 5 And InStr(1, mstrPrivs, ";��Ǹ���;") <> 0 And InStr(1, mstrPrivs, ";�Ǽ�;") <> 0) _
            Or (InStr(1, mstrPrivs, ";�Ǽ�;") <> 0 And InStr(1, mstrPrivs, ";��Ǹ���;") <> 0 And tabSelect.SelectedItem.Index = 1)
        Else
            blnModify = InStr(1, mstrPrivs, ";�޸�;") <> 0
            blnDelete = InStr(1, mstrPrivs, ";ɾ��;") <> 0
            blnVrfy = InStr(1, mstrPrivs, ";���;") <> 0
            blnStrike = InStr(1, mstrPrivs, ";����;") <> 0
            blnAdd = InStr(1, mstrPrivs, ";�Ǽ�;") <> 0
        End If
    End If
    
    '����
    mnuFilePrint.Enabled = blnData
    mnuFilePreView.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    tlbThis.Buttons("Print").Enabled = blnData
    tlbThis.Buttons("PrintView").Enabled = blnData
    
    '��ɾ��
'    mnuEditAddPayment.Enabled = blnAdd
'    mnuEditAddImprest.Enabled = blnAdd
'    mnuEditAddScheme.Enabled = blnAdd
     tlbThis.Buttons("Add").Enabled = blnAdd
    
    'Ԥ��
    If mblnԤ�� Then
        If blnData Then
            mnuEditModify.Enabled = False
            tlbThis.Buttons("Modify").Enabled = False
            mnuEditDel.Enabled = False
            tlbThis.Buttons("Delete").Enabled = False
            
            mnuEditCheck.Enabled = False
            tlbThis.Buttons("Check").Enabled = False
            mnuEditCheckBack.Enabled = False
            tlbThis.Buttons("CheckBack").Enabled = False
            
            mnuEditVerify.Enabled = False
            tlbThis.Buttons("Verify").Enabled = False
            mnuEditStrike.Enabled = False
            tlbThis.Buttons("Strike").Enabled = False
            
            If blnVerify Then
                mnuEditStrike.Enabled = (Not blnCancel) And blnStrike
                tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
            ElseIf IsPlanPayment(strNO) Then    'vsDetail.TextMatrix(1, vsDetail.ColIndex("��ⵥ��")) = ""
                '�ƻ�
                mnuEditModify.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnModify
                tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
                mnuEditDel.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnDelete
                tlbThis.Buttons("Delete").Enabled = mnuEditDel.Enabled
                mnuEditVerify.Enabled = blnData And (Not blnVerify) And blnVrfy
                tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                mnuEditStrike.Enabled = blnData And blnVerify And (Not blnCancel) And blnStrike
                tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
            ElseIf GetBillCheck(0, strNO) Then
                '�Ƿ�ȫѡ
                mnuEditVerify.Enabled = (Not blnVerify) And blnVrfy
                tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
                mnuEditStrike.Enabled = blnVerify And (Not blnCancel) And blnStrike
                tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
                mnuEditCheckBack.Enabled = mnuEditVerify.Enabled
                tlbThis.Buttons("CheckBack").Enabled = mnuEditVerify.Enabled And InStr(mstrPrivs, ";����;") > 0
            ElseIf GetBillCheck(1, strNO) Then
                '�Ƿ�ѡ��
                mnuEditCheck.Enabled = InStr(mstrPrivs, ";Ԥ��;") > 0
                tlbThis.Buttons("Check").Enabled = mnuEditCheck.Enabled
                mnuEditCheckBack.Enabled = InStr(mstrPrivs, ";����;") > 0
                tlbThis.Buttons("CheckBack").Enabled = mnuEditCheckBack.Enabled
            Else
                mnuEditModify.Enabled = (blnCancel = False And blnModify And blnVerify = False)
                tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
                mnuEditDel.Enabled = (blnCancel = False And blnDelete And blnVerify = False)
                tlbThis.Buttons("Delete").Enabled = mnuEditDel.Enabled
                mnuEditCheck.Enabled = (bln�ƻ����� = False And blnԤ�� = False And bln��� = False And InStr(mstrPrivs, ";Ԥ��;") > 0)
                tlbThis.Buttons("Check").Enabled = mnuEditCheck.Enabled
                mnuEditVerify.Enabled = Not (bln��� = False And blnԤ�� = False And bln�ƻ����� = False)
                tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
            End If
        Else
            mnuEditModify.Enabled = False
            tlbThis.Buttons("Modify").Enabled = False
            mnuEditDel.Enabled = False
            tlbThis.Buttons("Delete").Enabled = False
            
            mnuEditCheck.Enabled = False
            tlbThis.Buttons("Check").Enabled = False
            mnuEditCheckBack.Enabled = False
            tlbThis.Buttons("CheckBack").Enabled = False
            
            mnuEditVerify.Enabled = False
            tlbThis.Buttons("Verify").Enabled = False
            mnuEditStrike.Enabled = False
            tlbThis.Buttons("Strike").Enabled = False
        End If
    Else
        mnuEditModify.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnModify
        tlbThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
        mnuEditDel.Enabled = blnData And (Not blnVerify) And (Not blnCancel) And blnDelete
        tlbThis.Buttons("Delete").Enabled = mnuEditDel.Enabled
        '���
        mnuEditVerify.Enabled = blnData And (Not blnVerify) And blnVrfy
        tlbThis.Buttons("Verify").Enabled = mnuEditVerify.Enabled
        '����
        mnuEditStrike.Enabled = blnData And blnVerify And (Not blnCancel) And blnStrike
        tlbThis.Buttons("Strike").Enabled = mnuEditStrike.Enabled
    End If
    
    mnuEditDisplay.Enabled = blnData
    mnuFileBillPrePrint.Enabled = blnData
    mnuFileBillPrint.Enabled = blnData
    
    Call Ȩ�޿���_���ݴ�ӡ
End Sub

Public Sub Ȩ�޿���_���ݴ�ӡ()
    Dim blnBillPrint As Boolean
    Dim strNO As String
    Dim bytBillType As Byte        '0-Ԥ��,1-����,2-�ƻ�����
    Dim str��� As String
    
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) = "1" Then
        strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
        bytBillType = 0
    Else
        bytBillType = 1
    End If
    If bytBillType = 0 Then
        blnBillPrint = InStr(mstrPrivs, ";Ԥ����֪ͨ����ӡ;") <> 0
    Else
        '����27930 by lesfeng 2010-03-23
        str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
        If str��� = "���" Then
            blnBillPrint = InStr(mstrPrivs, ";��Ǹ��;") <> 0
        Else
            blnBillPrint = InStr(mstrPrivs, ";����֪ͨ��;") <> 0
        End If
    End If
        
    mnuFileBillPrePrint.Visible = blnBillPrint
    mnuFileBillPrint.Visible = blnBillPrint
    mnuFileSp.Visible = blnBillPrint
End Sub

Public Sub printbill(ByVal bytPrint As Byte)
    'bytPrint-0 ��ӡ,1-Ԥ��
    '���ݴ�ӡ
    Dim blnBillPrint As Boolean
    Dim strNO As String
    Dim bytBillType As Byte        '0-Ԥ��,1-����,2-�ƻ�����
    Dim intStatus  As Integer
    Dim str��� As String
    
    strNO = vsList.TextMatrix(vsList.Row, vsList.ColIndex("���ݺ�"))
    intStatus = vsList.TextMatrix(vsList.Row, vsList.ColIndex("��¼״̬"))
       
    If vsList.TextMatrix(vsList.Row, vsList.ColIndex("Ԥ����")) = "1" Then
        bytBillType = 0
    Else
        bytBillType = 1
    End If
    
    If bytBillType = 0 Then
        blnBillPrint = InStr(mstrPrivs, ";Ԥ����֪ͨ����ӡ;") <> 0
        If blnBillPrint Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1323_2", Me, "���ݱ��=" & strNO, "��¼״̬=" & intStatus, IIf(bytPrint = 1, 1, 2)
        End If
    Else
        '����27930 by lesfeng 2010-03-23
        str��� = Trim(vsList.TextMatrix(vsList.Row, vsList.ColIndex("�ܸ����")))
        If str��� = "���" Then
            blnBillPrint = InStr(mstrPrivs, ";��Ǹ��;") <> 0
            If blnBillPrint Then
                ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_3", Me, "���ݱ��=" & strNO, "��¼״̬=" & intStatus, , IIf(bytPrint = 1, 1, 2)
            End If
        Else
            blnBillPrint = InStr(mstrPrivs, ";����֪ͨ��;") <> 0
            If blnBillPrint Then
                ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_1", Me, "���ݱ��=" & strNO, "��¼״̬=" & intStatus, , IIf(bytPrint = 1, 1, 2)
            End If
        End If
    End If
End Sub

Private Sub Ȩ�޿���()
    Dim blnAdd As Boolean
    Dim blnModify As Boolean
    Dim blnDelete As Boolean
    Dim blnVerify As Boolean
    Dim blnCancel As Boolean
    Dim blnAdvance As Boolean
    Dim bln��� As Boolean
    
    blnAdd = InStr(1, mstrPrivs, ";�Ǽ�;") <> 0
    blnModify = InStr(1, mstrPrivs, ";�޸�;") <> 0
    blnDelete = InStr(1, mstrPrivs, ";ɾ��;") <> 0
    blnVerify = InStr(1, mstrPrivs, ";���;") <> 0
    blnCancel = InStr(1, mstrPrivs, ";����;") <> 0
    blnAdvance = InStr(1, mstrPrivs, ";Ԥ��;") <> 0
    '����27930 by lesfeng 2010-03-23
    bln��� = InStr(1, mstrPrivs, ";��Ǹ���;") <> 0
    
    If blnAdd = False And blnAdvance = False And bln��� = False Then
        mnuEditAdd.Visible = False
    Else
        mnuEditAddPayment.Visible = blnAdd
        mnuEditAddScheme.Visible = blnAdd
        mnuEditMultAdd.Visible = blnAdd
        mnuEditAddImprest.Visible = blnAdvance
        mnuEditAddSign.Visible = blnAdd And bln���
    End If
    
    mnuEditModify.Visible = blnModify Or blnAdvance Or (blnModify And bln���)
    mnuEditDel.Visible = blnDelete Or blnAdvance Or (blnDelete And bln���)
    mnuEditLine1.Visible = mnuEditAdd.Visible Or mnuEditModify.Visible Or mnuEditDel.Visible
    
    mnuEditVerify.Visible = blnVerify Or blnAdvance Or (blnVerify And bln���)
    mnuEditStrike.Visible = blnCancel Or blnAdvance Or (blnCancel And bln���)
    mnuEditLine2.Visible = mnuEditVerify.Visible Or mnuEditStrike.Visible
    
    tlbThis.Buttons("Add").Visible = blnAdd Or blnAdvance Or (blnAdd And bln���)
    tlbThis.Buttons("Modify").Visible = mnuEditModify.Visible
    tlbThis.Buttons("Delete").Visible = blnDelete Or blnAdvance Or (blnDelete And bln���)
    tlbThis.Buttons("EditSeparate").Visible = mnuEditLine1.Visible
    
    tlbThis.Buttons("Verify").Visible = mnuEditVerify.Visible
    tlbThis.Buttons("Strike").Visible = mnuEditStrike.Visible
    tlbThis.Buttons("VerifySeparate").Visible = mnuEditLine2.Visible
End Sub

Private Sub mnuViewSavePrint_Click()
        mnuViewSavePrint.Checked = Not mnuViewSavePrint.Checked
        Call zlDatabase.SetPara("���̴�ӡ", IIf(mnuViewSavePrint.Checked, "1", "0"), glngSys, mlngModule)
End Sub

Private Sub mnuViewVerifyPrint_Click()
        mnuViewVerifyPrint.Checked = Not mnuViewVerifyPrint.Checked
        Call zlDatabase.SetPara("��˴�ӡ", IIf(mnuViewVerifyPrint.Checked, "1", "0"), glngSys, mlngModule)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub vsDetail_GotFocus()
    zl_VsGridGotFocus vsDetail
End Sub

Private Sub vsDetail_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsDetail)
End Sub

Private Sub vsDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsDetail, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsAddition_GotFocus()
    zl_VsGridGotFocus vsAddition
End Sub

Private Sub vsAddition_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsAddition)
End Sub

Private Sub vsAddition_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsAddition, OldRow, NewRow, OldCol, NewCol)
End Sub
'����24925 by lesfeng 2010-02-08
Private Function GetShareSys(ByVal intSys As Integer) As Boolean
    ' ��Ҫ�������豸 ����400 �豸600
    Dim strSQL As String, strTmp As String
    Dim rsTemp As New ADODB.Recordset
    Dim intShareSys As Integer
    
    GetShareSys = False
    If intSys = 400 Then
        Select Case mint����Flag
        Case 1
            GetShareSys = True
            Exit Function
        Case 2
            GetShareSys = False
            Exit Function
        End Select
    End If
    If intSys = 600 Then
        Select Case mint�豸Flag
        Case 1
            GetShareSys = True
            Exit Function
        Case 2
            GetShareSys = False
            Exit Function
        End Select
    End If
    
    On Error GoTo errH
    strSQL = "SELECT decode(�����,null,0,1) as ����� FROM zlsystems WHERE ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, intSys)
    If Not rsTemp.EOF Then
        intShareSys = IIf(IsNull(rsTemp!�����), 0, rsTemp!�����)
        If intShareSys = 1 Then
            GetShareSys = True
            If intSys = 400 Then mint����Flag = 1
            If intSys = 600 Then mint�豸Flag = 1
        Else
            If intSys = 400 Then mint����Flag = 2
            If intSys = 600 Then mint�豸Flag = 2
        End If
    Else
        If intSys = 400 Then mint����Flag = 2
        If intSys = 600 Then mint�豸Flag = 2
    End If
    rsTemp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetBillCheck(ByVal bytType As Byte, ByVal strNO As String) As Boolean
'���ܣ���ȡ����Ԥ���Ƿ�ѡ�������Ƿ�ȫѡ
'������
'   bytType=1 ����ȡ�Ƿ�ѡ��
'   bytType=0 ����ȡ�Ƿ�ȫѡ
'���أ�ѡ����ȫѡѡ��True����֮False
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select Sum(Rec) Rec, Sum(Reccheck) Reccheck " & _
              "From (Select Count(1) Rec, Case When a.Ԥ�� = 1 Then Count(a.Ԥ��) Else 0 End Reccheck " & _
              "  From Ӧ����¼ A, �����¼ B " & _
              "  Where a.������� = b.������� And a.��¼״̬ = 1 And b.No = [1] " & _
              "  Group By a.Ԥ��) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԤ��ѡ����¼��", strNO)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!rec) And IsNull(rsTmp!reccheck) Then Exit Function
        If bytType = 1 Then
            'ѡ��
            GetBillCheck = Nvl(rsTmp!reccheck, 0) > 0
        Else
            'ȫѡ
            GetBillCheck = (Nvl(rsTmp!rec, 0) - Nvl(rsTmp!reccheck, 0) = 0)
        End If
    End If
    rsTmp.Close
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function TestCheck(ByVal bytType As Byte, ByVal strNO As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '�����Ƿ�ɾ�������
    On Error GoTo errHandle
    
    If bytType = 1 Then
        gstrSQL = "Select id From �����¼ Where NO=[1] And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�ɾ��", strNO)
    Else
        gstrSQL = "Select id From �����¼ Where NO=[1] And ��¼״̬=1 And ������� is null And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�ͨ�����", strNO)
    End If
    TestCheck = (rsTemp.RecordCount = 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetMultiPayment(ByVal strPaymentNO As String) As Boolean
'���ܣ��жϸ���ݵ���ϸ�Ƿ���ڶ��ٸ������
'���أ�True���ڣ�False������
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select Count(1) Rec From �����¼ A, Ӧ����¼ B " & _
             "Where a.������� = b.������� And ��¼���� = 2 And a.��¼״̬ = 1 And a.No = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�Ϊ��θ���", strPaymentNO)
    GetMultiPayment = Nvl(rsTemp!rec) > 0
    rsTemp.Close
    
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
