VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceOperate 
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmAdviceOperate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   10875
   StartUpPosition =   3  '����ȱʡ
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   10875
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   450
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   450
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Ctrl+A)"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "ȫѡ"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Ctrl+R)"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ȫ��"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ִ��"
               Key             =   "ִ��"
               Description     =   "ִ��"
               Object.ToolTipText     =   "ִ��(Ctrl+E)"
               Object.Tag             =   "ִ��"
               ImageKey        =   "ִ��"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "������������(F12)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "ˢ��"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "���¶�ȡ����(F5)"
               Object.Tag             =   "ˢ��"
               ImageKey        =   "ˢ��"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�(ALT+X)"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   10875
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   510
      Width           =   10875
      Begin VB.CommandButton cmdAlley 
         Caption         =   "����ʷ/����״̬"
         Height          =   350
         Left            =   9285
         TabIndex        =   4
         Top             =   20
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.ComboBox cboҽ�� 
         Height          =   300
         Left            =   9285
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   45
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblҽ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ֹͣҽ��(&D)"
         Height          =   180
         Left            =   8250
         TabIndex        =   2
         Top             =   105
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblPati 
         BackStyle       =   0  'Transparent
         Caption         =   "����: סԺ��: ����: ����:"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   105
         Width           =   6825
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6735
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6930
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   1560
      TabIndex        =   12
      Top             =   6885
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Align           =   1  'Align Top
      Height          =   1590
      Left            =   0
      TabIndex        =   1
      Top             =   5220
      Visible         =   0   'False
      Width           =   10875
      _cx             =   19182
      _cy             =   2805
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VB.PictureBox picUD 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   10875
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5175
      Visible         =   0   'False
      Width           =   10875
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Align           =   1  'Align Top
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   10875
      _cx             =   19182
      _cy             =   7541
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
      BackColorSel    =   16764057
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceOperate.frx":058A
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
      Begin MSComctlLib.ImageList img16 
         Left            =   2235
         Top             =   855
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
               Picture         =   "frmAdviceOperate.frx":0625
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":0BBF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":1159
               Key             =   "ǩ��"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   2835
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":14AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":17A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":1A9F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":1D99
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceOperate.frx":2093
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   360
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":238D
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":25A7
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":27C1
            Key             =   "ִ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":29DB
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":2BF5
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":2E0F
            Key             =   "ˢ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3509
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   960
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3723
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":393D
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3B57
            Key             =   "ִ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3D71
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":3F8B
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":41A5
            Key             =   "ˢ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceOperate.frx":489F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6825
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceOperate.frx":4AB9
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12859
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":534D
            Text            =   "ͨ��"
            TextSave        =   "ͨ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":5937
            Text            =   "����"
            TextSave        =   "����"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":5F21
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":655B
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu mnuPass 
      Caption         =   "Pass"
      Visible         =   0   'False
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩ���ٴ���Ϣ�ο�(&C)"
         Index           =   0
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩƷ˵����(&D)"
         Index           =   1
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "�й�ҩ��(&N)"
         Index           =   2
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "������ҩ����(&S)"
         Index           =   3
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "����ֵ(&T)"
         Index           =   4
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ר����Ϣ(&P)"
         Index           =   6
         Begin VB.Menu mnuPassSpec 
            Caption         =   "ҩ��-ҩ���໥����(&D)"
            Index           =   0
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "ҩ��-ʳ���໥����(&F)"
            Index           =   1
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����ע�������(&M)"
            Index           =   3
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����ע�������(&T)"
            Index           =   4
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "����֢(&C)"
            Index           =   6
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "������(&S)"
            Index           =   7
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&G)"
            Index           =   9
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��ͯ��ҩ(&P)"
            Index           =   10
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&E)"
            Index           =   11
         End
         Begin VB.Menu mnuPassSpec 
            Caption         =   "��������ҩ(&L)"
            Index           =   12
         End
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҽҩ��Ϣ����(&I)"
         Index           =   8
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҩƷ�����Ϣ(&M)"
         Index           =   10
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "��ҩ;�������Ϣ(&R)"
         Index           =   11
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ҽԺҩƷ��Ϣ(&F)"
         Index           =   12
      End
   End
End
Attribute VB_Name = "frmAdviceOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���ܣ�
'0-ҽ������:
'    ֻ��ѡ��Ҫ���ϵ�ҽ��
'1-ֹͣҽ��:
'    ��Ҫָ����ֹʱ��(ȱʡΪ��ǰ,������ЧȱʡΪ�������,Ԥ���Ĳ���)
'    ��ʿͣʱ��Ҫָ��ֹͣҽ��
'2-ȷ��ֹͣ:
'    ֻ��ѡ����Ҫȷ�ϵ�ҽ��
'3-У��ҽ��:
'    ��¼��ҽ�������޸�У��ʱ��(�ǲ�¼��ȱʡΪ��ǰ���ɸ�,��¼��ȱʡΪ����ʱ��+1m)
'4-�����Ƽ���Ŀ:
'    ��ɾ��ÿ��ҽ���ļƼ���Ŀ
'5-��ͣҽ��
'    ѡ����Ҫ��ͣ��ҽ��
'6-����ҽ��
'    ѡ����Ҫ���õ�ҽ��
Public mstrPrivs As String
Public mlngҽ��ID As Long '����ȱʡ��λ
Public mint���� As Integer '0-ҽ������,1-ֹͣҽ��,2-ȷ��ֹͣ,3-ҽ��У��,4-�����Ƽ���Ŀ,5-��ͣҽ��,6-����ҽ��
Public mlng����ID As Long
Public mlng����ID As Long
Public mlng��ҳID As Long
Public mbln��ʿվ As Boolean
Public mblnOK As Boolean

Private mrsPrice As ADODB.Recordset
Private mrsDept As ADODB.Recordset
Private mstrLike As String
Private mblnReturn As Boolean
Private mint���� As Integer
Private mstrRollNotify As String '������Ҫ���г����ջ����ѵĲ���(����ID,��ҳID;...)
Private mlngPassPati As Long 'Pass:�ϴ��Ѵ���PASS�Ĳ���ID

'��������
Private mblnFirst As Boolean
Private mstr����IDs As String
Private mint��Ч As Integer
Private mint��� As Integer
Private mblnPauseLast As Boolean

'������
Private Const COL_ID = 0
Private Const COL_���ID = 1
Private Const COL_��ID = 2
Private Const COL_��� = 3
Private Const COL_������� = 4
Private Const COL_������� = 5
Private Const COL_���� = 6 '1-��ҩ�䷽,2-�������
'Pass��ʾ��
Private Const COL_��ʾ = 7
'������
Private Const COL_ѡ�� = 8 '
Private Const COL_���� = 9 '
'�ɼ���
Private Const COL_���� = 10
Private Const COL_סԺ�� = 11
Private Const COL_���� = 12
Private Const COL_Ӥ�� = 13
Private Const COL_��Ч = 14
Private Const COL_����ʱ�� = 15
Private Const COL_��ʼʱ�� = 16
Private Const COL_ҽ������ = 17
Private Const COL_Ƥ�� = 18
Private Const COL_���� = 19
Private Const COL_���� = 20
Private Const COL_Ƶ�� = 21
Private Const COL_�÷� = 22
Private Const COL_ҽ������ = 23
Private Const COL_ִ��ʱ�� = 24
Private Const COL_��ֹʱ�� = 25 '
Private Const COL_ִ�п��� = 26
Private Const COL_ִ������ = 27
Private Const COL_�ϴ�ִ�� = 28 '
Private Const COL_��־ = 29
Private Const COL_����ҽ�� = 30
Private Const COL_У�Ի�ʿ = 31 '
Private Const COL_У��ʱ�� = 32 '
Private Const COL_ͣ��ҽ�� = 33 '
Private Const COL_ͣ��ʱ�� = 34 '
'����
Private Const COL_����ID = 35
Private Const COL_��ҳID = 36
Private Const COL_�������� = 37
Private Const COL_ִ�п���ID = 38
Private Const COL_���˿���ID = 39
Private Const COL_�շ�ϸĿID = 40
Private Const COL_������λ = 41
Private Const COL_ǰ��ID = 42
Private Const COL_ǩ��ID = 43
Private Const COL_������Ա = 44

'�Ƽ��嵥����ֵ
Private Const COLP_ҽ��ID = 0 '���Ӵ�ű����Ϣ
Private Const COLP_���ID = 1 '���Ӵ�ű����Ϣ
Private Const COLP_������� = 2 '���Ӵ�ű����Ϣ
Private Const COLP_������ĿID = 3
Private Const COLP_�շ�ϸĿID = 4
Private Const COLP_�̶� = 5
Private Const COLP_�Ƽ�ҽ�� = 6
Private Const COLP_��� = 7 '�շ��������
Private Const COLP_�շ���Ŀ = 8
Private Const COLP_��λ = 9
Private Const COLP_���� = 10
Private Const COLP_���� = 11
Private Const COLP_ִ�п��� = 12
Private Const COLP_�������� = 13
Private Const COLP_���� = 14
Private Const COLP_�շ���� = 15
Private Const COLP_ִ�п���ID = 16
Private Const COLP_�������� = 17

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.Value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.Value = vNewValue
        txtPer.Text = CInt(psb.Value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub cmdAlley_Click()
'���ܣ��Բ��˹���ʷ/����״̬���в鿴
    'Pass
    Call AdviceCheckWarn(21, vsAdvice.Row)
End Sub

Private Function ResetCond() As Boolean
'���ܣ����÷�������
    Dim blnSeek As Boolean
    Me.Refresh
    With frmAdviceOperateCond
        .mstrPrivs = mstrPrivs
        .mint���� = mint����
        .mlng����ID = mlng����ID
        .mlng����ID = mlng����ID
        .Show 1, Me
        If .mblnOK Then
            mlng����ID = .mlng����ID
            mstr����IDs = .mstr����IDs
            mint��Ч = .mint��Ч
            mint��� = .mint���
            mblnPauseLast = .mblnPauseLast
                        
            'ֻѡ���˵�ǰ���˲Ŷ�λ��ǰҽ��
            If UBound(Split(mstr����IDs, ";")) = 0 Then
                If Val(Split(mstr����IDs, ",")(0)) = mlng����ID Then blnSeek = True
            End If
            Call RefreshData(IIF(blnSeek, mlngҽ��ID, 0), True)
        End If
        ResetCond = .mblnOK
    End With
End Function

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        If tbr.Buttons("����").Visible Then
            If Not ResetCond Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫѡ"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫ��"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ִ��"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("�˳�"))
    ElseIf KeyCode = vbKeyF1 Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    ElseIf KeyCode = vbKeyF5 Then
        Call tbr_ButtonClick(tbr.Buttons("ˢ��"))
    ElseIf KeyCode = vbKeyF12 Then
        If tbr.Buttons("����").Visible Then
            Call tbr_ButtonClick(tbr.Buttons("����"))
        End If
    ElseIf KeyCode = vbKeyF7 Then '�л����뷨
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Call InitAdviceTable
    Call SetAdviceCol '������һ��������,�Ա���ȷ�ָ����Ի�
    If mint���� = 3 Or mint���� = 4 Then
        Call InitPriceTable
    End If
    Call RestoreWinState(Me, App.ProductName, mint����)
    
    mblnOK = False
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    Select Case mint����
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not (mint���� = 3 Or mint���� = 4) Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    
    '�������ÿ�����,ȱʡ��������
    mblnFirst = True
    mblnPauseLast = False
    mint��Ч = 0: mint��� = 0
    mstr����IDs = mlng����ID & "," & mlng��ҳID
    If mbln��ʿվ And InStr(",3,5,6,", mint����) > 0 Then
        If mint���� = 3 Then
            tbr.Buttons("����").Enabled = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ҽ��У��", 0)) <> 0
        ElseIf mint���� = 5 Or mint���� = 6 Then
            tbr.Buttons("����").Enabled = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ҽ����ͣ", 0)) <> 0
        End If
    Else
        tbr.Buttons("����").Enabled = False
    End If
    tbr.Buttons("����").Visible = tbr.Buttons("����").Enabled 'Enabled�����ж�
    
    If mint���� = 0 Then
        Caption = "����ҽ������"
        tbr.Buttons("ִ��").Caption = "����"
        tbr.Buttons("ִ��").ToolTipText = "����ѡ���ҽ��(Ctrl+E)"
    ElseIf mint���� = 1 Then
        Caption = "����ҽ��ֹͣ"
        tbr.Buttons("ִ��").Caption = "ֹͣ"
        tbr.Buttons("ִ��").ToolTipText = "ֹͣѡ���ҽ��(Ctrl+E)"
        If mbln��ʿվ Then
            lblҽ��.Visible = True
            cboҽ��.Visible = True
        End If
    ElseIf mint���� = 2 Then
        Caption = "ȷ��ҽ��ֹͣ"
        tbr.Buttons("ִ��").Caption = "ȷ��"
        tbr.Buttons("ִ��").ToolTipText = "ȷ��ѡ���ҽ��(Ctrl+E)"
    ElseIf mint���� = 3 Then
        Caption = "����ҽ��У��"
        tbr.Buttons("ִ��").Caption = "У��"
        tbr.Buttons("ִ��").ToolTipText = "ȷ��ѡ���ҽ��(Ctrl+E)"
        
        stbThis.Panels(3).Visible = True
        stbThis.Panels(4).Visible = True
        
        picUD.Visible = True
        vsPrice.Visible = True
        
        '���˹���ʷ/����״̬���ü��
        mlngPassPati = 0
        If gblnPass And InStr(mstrPrivs, "������ҩ���") > 0 Then 'Pass
            cmdAlley.Visible = True
            vsAdvice.ColHidden(COL_��ʾ) = False
            cmdAlley.Enabled = PassGetState("AlleyEnable") = 1
        End If
    ElseIf mint���� = 4 Then
        Caption = "�����Ƽ���Ŀ"
        tbr.Buttons("ִ��").Caption = "ȷ��"
        tbr.Buttons("ִ��").ToolTipText = "ȷ��ѡ����Ŀ�ļ�Ŀ(Ctrl+E)"
        
        picUD.Visible = True
        vsPrice.Visible = True
    ElseIf mint���� = 5 Then
        Caption = "����ҽ����ͣ"
        tbr.Buttons("ִ��").Caption = "��ͣ"
        tbr.Buttons("ִ��").ToolTipText = "��ͣѡ���ҽ��(Ctrl+E)"
    ElseIf mint���� = 6 Then
        Caption = "����ҽ������"
        tbr.Buttons("ִ��").Caption = "����"
        tbr.Buttons("ִ��").ToolTipText = "����ѡ���ҽ��(Ctrl+E)"
    End If
        
    '��ȡ������Ϣ
    If mint���� = 3 Or mint���� = 4 Then
        strSQL = "Select ID,���� From ���ű�"
        Set mrsDept = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrsDept, strSQL, Me.Caption)
    End If
    
    '��ʾ������Ϣ��һ�����˲��������
    strSQL = _
        " Select A.סԺ��,A.����,A.�Ա�,A.����,B.��Ժ����," & _
        " B.סԺҽʦ,B.��Ժ����ID,C.���� as ����" & _
        " From ������Ϣ A,������ҳ B,���ű� C" & _
        " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    lblPati.Caption = "����:" & rsTmp!���� & "��סԺ��:" & Nvl(rsTmp!סԺ��) & _
        "������:" & Nvl(rsTmp!��Ժ����) & "������:" & Nvl(rsTmp!����)
    
    '��ѡ��ͣ��ҽ��:ȱʡΪ���˵�סԺҽʦ���˿��ҵĵ�һ��ҽ��
    'Ŀǰ��֧������ֹͣҽ��,��˿϶����Դ���ĵ�ǰ����Ϊ׼��ȡ
    If mint���� = 1 And mbln��ʿվ Then
        Call Get����ҽ��(rsTmp!��Ժ����ID, True, Nvl(rsTmp!סԺҽʦ), 0, cboҽ��)
        If cboҽ��.ListIndex = -1 And cboҽ��.ListCount > 0 Then cboҽ��.ListIndex = 0
    End If
    
    '��ʾҽ������
    If Not tbr.Buttons("����").Enabled Then Call RefreshData(mlngҽ��ID, True)
End Sub

Private Sub RefreshData(Optional ByVal lngҽ��ID As Long, Optional ByVal blnNotify As Boolean)
'���ܣ�ˢ������
'������lngҽ��ID=����ҽ����λ
'      blnNotify=�Ƿ���������ҽ��
    Dim blnChange As Boolean, i As Long
    Dim strPatis As String, arrPatis As Variant
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strMsg As String, strTmp As String
    
    '��ʾҽ������
    Call LoadAdvice(strPatis)
    
    '��ȡ�Ƽ�����
    If mint���� = 3 Or mint���� = 4 Then
        Call InitPriceRecordset
        Screen.MousePointer = 11
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            Progress = i / (vsAdvice.Rows - 1) * 100
            blnChange = False
            Call LoadPrice(i, blnChange)
            If blnChange And mint���� = 4 Then Call SelectRow(i)
        Next
        Progress = 0: Screen.MousePointer = 0
    End If
    
    If lngҽ��ID <> 0 Then
        i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
        If i <> -1 Then vsAdvice.Row = i
    End If
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    '����ҽ������
    If blnNotify And InStr(",3,4,6,", mint����) > 0 And strPatis <> "" Then
        arrPatis = Split(strPatis, ";")
        For i = 0 To UBound(arrPatis)
            lng����ID = Split(arrPatis(i), ",")(0)
            lng��ҳID = Split(arrPatis(i), ",")(1)
            strTmp = ExistsSpecAdvice(lng����ID, lng��ҳID)
            If strTmp <> "" Then
                strTmp = Replace(Replace(strTmp, "��������", ""), vbCrLf & vbCrLf, vbCrLf)
                strMsg = strMsg & vbCrLf & strTmp
            End If
        Next
        If strMsg <> "" Then MsgBox Mid(strMsg, 3), vbInformation, gstrSysName & " - ������"
    End If
End Sub

Private Sub SelectRow(ByVal lngRow As Long)
'���ܣ�ʹָ����ѡ��(����һ����ҩ)
    With vsAdvice
        If mint���� = 3 Then
            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages(1).Picture
            .Cell(flexcpData, lngRow, COL_ѡ��) = 1
        Else
            .TextMatrix(lngRow, COL_ѡ��) = -1 'ֱ�Ӷ�TextMatrixʱ,��Ҫ��True
        End If
    End With
    Call vsAdvice_AfterEdit(lngRow, COL_ѡ��)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsAdvice.Height = Me.ScaleHeight - cbr.Height - stbThis.Height - picPati.Height - IIF(picUD.Visible, picUD.Height + vsPrice.Height, 0)
    If cboҽ��.Visible Then
        lblPati.Width = Me.ScaleWidth - lblҽ��.Width - cboҽ��.Width - lblPati.Left - 350
        cboҽ��.Left = Me.ScaleWidth - cboҽ��.Width - 200
        lblҽ��.Left = cboҽ��.Left - lblҽ��.Width - 45
    ElseIf cmdAlley.Visible Then
        lblPati.Width = Me.ScaleWidth - cmdAlley.Width - lblPati.Left - 350
        cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 200
    Else
        lblPati.Width = Me.ScaleWidth - lblPati.Left
    End If
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mint����)
    
    Set mrsPrice = Nothing
    Set mrsDept = Nothing
    mstrPrivs = ""
    mlngҽ��ID = 0
    mint���� = 0
    mlng����ID = 0
    mlng����ID = 0
    mlng��ҳID = 0
    mbln��ʿվ = False
End Sub

Private Sub picUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsPrice.Height - y < 500 Then Exit Sub
        vsAdvice.Height = vsAdvice.Height + y
        vsPrice.Height = vsPrice.Height - y
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", _
            IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
        mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    Select Case Button.Key
        Case "ȫѡ"
            If vsAdvice.ColHidden(COL_ѡ��) Then Exit Sub
            If vsAdvice.Rows = vsAdvice.FixedRows + 1 And Val(vsAdvice.TextMatrix(vsAdvice.FixedRows, COL_ID)) = 0 Then Exit Sub
            
            If mint���� = 3 Then
                For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
                    If vsAdvice.Cell(flexcpData, i, COL_ѡ��) = Empty Then '�������ʵĲ���
                        Set vsAdvice.Cell(flexcpPicture, i, COL_ѡ��) = img16.ListImages(1).Picture
                        vsAdvice.Cell(flexcpData, i, COL_ѡ��) = 1
                    End If
                Next
            Else
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = True
            End If
        Case "ȫ��"
            If mint���� = 3 Then
                Set vsAdvice.Cell(flexcpPicture, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = Nothing
                vsAdvice.Cell(flexcpData, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = Empty
            Else
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = False
            End If
        Case "ˢ��"
            Call RefreshData(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
        Case "����"
            Call ResetCond
        Case "ִ��"
            If Not CheckValid Then Exit Sub
            If Not CheckSignValid Then Exit Sub
            If ExecuteOperate Then
                'ҽ��У��ʱ��鲢���ѳ����ջ�(�Զ�)ֹͣ��ҽ��
                If mint���� = 3 And mstrRollNotify <> "" Then
                    Call ShowRollNotify
                End If
                
                mblnOK = True: Unload Me
            End If
        Case "����"
            ShowHelp App.ProductName, Me.Hwnd, Me.Name
        Case "�˳�"
            Unload Me
    End Select
End Sub

Private Sub ShowRollNotify()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim lng����ID As Long, lng��ҳID As Long, i As Long
    
    On Error GoTo errH
    
    For i = 0 To UBound(Split(mstrRollNotify, ";"))
        '�����볬���ջ���һ�£���ֻ������ǰ״̬Ϊ(�Զ�)ֹͣ�ġ�
        strSQL = "(A.ִ��ʱ�䷽�� is NULL And (Nvl(A.Ƶ�ʴ���,0)=0 Or Nvl(A.Ƶ�ʼ��,0)=0 Or A.Ƶ�ʼ�� is NULL))"
        strSQL = _
            " Select A.����,A.ҽ������ From ����ҽ����¼ A,������ĿĿ¼ E" & _
            " Where A.������ĿID=E.ID And A.����ID=[1] And A.��ҳID=[2]" & _
            " And Not(A.�������='H' And E.��������='1') And Not(A.�������='Z' And E.��������='4')" & _
            " And Nvl(A.ִ������,0)<>0 And A.�ܸ����� is NULL And Nvl(A.ҽ����Ч,0)=0" & _
            " And ((Not " & strSQL & " And A.ִ����ֹʱ��<A.�ϴ�ִ��ʱ��)" & _
            " Or (" & strSQL & " And Trunc(A.ִ����ֹʱ��)<Trunc(A.�ϴ�ִ��ʱ��)+1))" & _
            " And A.ҽ��״̬=8 And (A.���ID is Null Or A.������� IN('5','6'))" & _
            " And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3 And Not Exists(" & _
                " Select ID From ����ҽ����¼ X" & _
                " Where ������� IN('5','6') And X.���ID=A.ID" & _
                " And ����ID=[1] And ��ҳID=[2])" & _
            " Order by A.���"
        lng����ID = Split(Split(mstrRollNotify, ";")(i), ",")(0)
        lng��ҳID = Split(Split(mstrRollNotify, ";")(i), ",")(1)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If Not rsTmp.EOF Then
            strMsg = strMsg & vbCrLf & vbCrLf & "����""" & Nvl(rsTmp!����) & """��ҽ����"
            Do While Not rsTmp.EOF
                strMsg = strMsg & vbCrLf & "��" & rsTmp!ҽ������
                rsTmp.MoveNext
            Loop
        End If
    Next
    If strMsg <> "" Then
        MsgBox "������ֹͣ�Ĳ���ҽ�������ڷ��ͣ�" & strMsg & vbCrLf & vbCrLf & "����ҽ�������ڻ�ʿ����վ��ʹ��""���ڷ����ջ�""���д���", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    '��ʾ�Ƽ���Ŀ
    If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
        If (mint���� = 3 Or mint���� = 4) And Not mrsPrice Is Nothing Then
            Call ShowPrice(NewRow)
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = COL_ҽ������ Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ�һ����ҩ��һ������
    Dim lngBegin As Long, lngEnd As Long
    Dim strTmp As String, vPause As Date, i As Long
        
    With vsAdvice
        If Col = COL_���� And Not mblnReturn Then
            '�ǻس�����ת��ȷ��
            strTmp = .TextMatrix(Row, Col)
            If strTmp <> "" Then strTmp = GetFullDate(strTmp)
            If Not IsDate(strTmp) Then
                .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
            End If
            
            If mint���� = 1 Then '�����ֹʱ��
                If IsDate(.Cell(flexcpData, Row, COL_�ϴ�ִ��)) Then
                    If .TextMatrix(Row, COL_ִ��ʱ��) = "" And Format(.TextMatrix(Row, COL_�ϴ�ִ��), "HH:mm") = "00:00" Then
                        '"������"����,ֹͣ���첻����
                        If Format(strTmp, "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, Row, COL_�ϴ�ִ��)) + 1, "yyyy-MM-dd") Then
                            .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                        End If
                    Else
                        If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                            .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                        End If
                    End If
                End If
                If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
                If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
            ElseIf mint���� = 2 Then  '���ȷ��ֹͣʱ��
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
            ElseIf mint���� = 3 Then
                '����С�ڿ���ʱ��,��ʼʱ���С��(��ʼʱ����Ըĳɱȿ���ʱ��С)
                If Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                Else
                    If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
            ElseIf mint���� = 5 Then '�����ͣʱ��
                'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
                If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                End If
                'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
                If .TextMatrix(Row, COL_�ϴ�ִ��) <> "" Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
                'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
                If .TextMatrix(Row, COL_��ֹʱ��) <> "" Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
                'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 7)
                If vPause <> CDate(0) Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
            ElseIf mint���� = 6 Then '�������ʱ��
                'Ӧ>��ͣʱ��
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 6)
                If vPause <> CDate(0) Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
                'Ӧ<=ִ����ֹʱ��
                If .TextMatrix(Row, COL_��ֹʱ��) <> "" Then
                    If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                        .TextMatrix(.Row, .Col) = CStr(.Cell(flexcpData, .Row, .Col)): Exit Sub
                    End If
                End If
            End If
            .TextMatrix(Row, Col) = strTmp
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            'һ����ҩһ�����
            mblnReturn = True
            Call vsAdvice_AfterEdit(Row, Col)
        Else
            'һ����ҩ��һ��ѡ�������
            If (Col = COL_ѡ�� Or Col = COL_����) And InStr(",5,6,", .TextMatrix(Row, COL_�������)) > 0 Then
                If RowInһ����ҩ(Row, lngBegin, lngEnd) Then
                    For i = lngBegin To lngEnd
                        If i <> Row Then
                            .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                            .Cell(flexcpData, i, Col) = .Cell(flexcpData, Row, Col)
                            Set .Cell(flexcpPicture, i, Col) = .Cell(flexcpPicture, Row, Col)
                        End If
                    Next
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Or Col = COL_���� Or Col = COL_��ʾ Then Cancel = True 'Pass
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If mint���� = 3 And .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ӥ��: lngRight = COL_��ʼʱ��
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_Ƶ��: lngRight = COL_�÷�
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
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
'���ܣ���λ����һ���뵥Ԫ������У�Ա�־
    Dim blnGroup As Boolean, i As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        With vsAdvice
            If .ColHidden(COL_ѡ��) And .ColHidden(COL_����) Then
                If .Row + 1 <= .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    .Row = .FixedRows
                End If
            Else
                If .Col = COL_ѡ�� Then
                    If Not .ColHidden(COL_����) Then
                        .Col = COL_����
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            .Row = .FixedRows
                        End If
                    End If
                ElseIf .Col = COL_���� Then
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    .Col = COL_ѡ��
                Else
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    If Not .ColHidden(COL_ѡ��) Then .Col = COL_ѡ��
                End If
            End If
            Call .ShowCell(.Row, .Col)
        End With
    ElseIf KeyAscii = 32 Then
        With vsAdvice
            If mint���� = 3 And .Col = COL_ѡ�� Then
                KeyAscii = 0
                
                If .Cell(flexcpData, .Row, .Col) = Empty Then
                    Set .Cell(flexcpPicture, .Row, .Col) = img16.ListImages(1).Picture
                    .Cell(flexcpData, .Row, .Col) = 1
                ElseIf .Cell(flexcpData, .Row, .Col) = 1 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = img16.ListImages(2).Picture
                    .Cell(flexcpData, .Row, .Col) = 2
                ElseIf .Cell(flexcpData, .Row, .Col) = 2 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = Nothing
                    .Cell(flexcpData, .Row, .Col) = Empty
                End If
            
                If InStr(",5,6,", .TextMatrix(.Row, COL_�������)) > 0 Then
                    If .Row - 1 >= .FixedRows Then
                        blnGroup = Val(.TextMatrix(.Row - 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID))
                    End If
                    If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                        blnGroup = Val(.TextMatrix(.Row + 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID))
                    End If
                    If blnGroup Then
                        For i = .Row - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                        For i = .Row + 1 To .Rows - 1
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strTmp As String, vPause As Date
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        mblnReturn = True
        With vsAdvice
            '����������Ч��
            If .EditText <> "" Then .EditText = GetFullDate(.EditText)
            If Not IsDate(.EditText) Then
                MsgBox "������һ����Ч��" & .TextMatrix(0, Col) & " ��", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
            End If
            
            If mint���� = 1 Then '�����ֹʱ��
                '������ڿ�ʼִ��ʱ��
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "�����ִ����ֹʱ��������ҽ���Ŀ�ʼִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '����С�ڿ���ʱ��
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "�����ִ����ֹʱ�䲻ӦС�ڿ���ʱ�� " & Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                '��ӦС���ϴ�ִ��ʱ��
                If IsDate(.Cell(flexcpData, Row, COL_�ϴ�ִ��)) Then
                    If .TextMatrix(Row, COL_ִ��ʱ��) = "" And Format(.TextMatrix(Row, COL_�ϴ�ִ��), "HH:mm") = "00:00" Then

                        '"������"����,ֹͣ���첻����
                        If Format(.EditText, "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, Row, COL_�ϴ�ִ��)) + 1, "yyyy-MM-dd") Then
                            strTmp = .EditText 'MsgBoxһ��,EditText�Ϳ���,����Ҫ��¼
                            If MsgBox("�Գ����Գ�����ִ����ֹ����Ӧ�����ϴ�ִ������ " & Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd") & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                            End If
                        End If
                    Else
                        If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                            strTmp = .EditText 'MsgBoxһ��,EditText�Ϳ���,����Ҫ��¼
                            If MsgBox("�����ִ����ֹʱ��С��ҽ�����ϴ�ִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                            End If
                        End If
                    End If
                End If
            ElseIf mint���� = 2 Then  '���ȷ��ֹͣʱ��
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "ȷ��ֹͣҽ����ʱ�䲻��С��ҽ����ִ����ֹʱ�� " & Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
            ElseIf mint���� = 3 Then  '���У��ʱ��(��¼�Ĳ��ܸ�)
                '����С�ڿ���ʱ��,��ʼʱ���С��(��ʼʱ����Ըĳɱȿ���ʱ��С)
                If Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                        MsgBox "�����У��ʱ�䲻��С�ڿ���ʱ�� " & Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                Else
                    If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                        MsgBox "�����У��ʱ�䲻��С��ҽ���Ŀ�ʼִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            ElseIf mint���� = 5 Then '�����ͣʱ��
                'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "ҽ������ͣʱ��Ӧ���ڵ��ڿ�ʼִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                End If
                'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
                If .TextMatrix(Row, COL_�ϴ�ִ��) <> "" Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                        MsgBox "ҽ������ͣʱ��Ӧ�����ϴ�ִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
                If .TextMatrix(Row, COL_��ֹʱ��) <> "" Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                        MsgBox "ҽ������ͣʱ��ӦС��ִ����ֹʱ�� " & Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 7)
                If vPause <> CDate(0) Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        MsgBox "ҽ������ͣʱ��Ӧ�����ϴ���ͣ�������ʱ�� " & Format(vPause, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            ElseIf mint���� = 6 Then '�������ʱ��
                'Ӧ>��ͣʱ��
                vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 6)
                If vPause <> CDate(0) Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                        MsgBox "ҽ��������ʱ��Ӧ�����ϴ���ͣʱ�� " & Format(vPause, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
                
                'Ӧ<=ִ����ֹʱ��
                If .TextMatrix(Row, COL_��ֹʱ��) <> "" Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                        MsgBox "ҽ��������ʱ��ӦС�ڵ���ִ����ֹʱ�� " & Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Sub
                    End If
                End If
            End If
            .TextMatrix(Row, Col) = IIF(.EditText = "" And strTmp <> "", strTmp, .EditText)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            Call vsAdvice_AfterEdit(Row, Col) 'һ����ҩ��һ������:��ʾ�󲻻��Զ�ִ�и��¼�
            
            '����Ϊ��ͬʱ��(У��,��ͣ,����)
            If Row = .FixedRows And .Rows > .FixedRows + 1 Then
                If mint���� = 3 Then
                    If MsgBox("Ҫ�������е�ҽ���������ʱ��У����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call SetSameTime(Row)
                    End If
                ElseIf (mint���� = 5 Or mint���� = 6) Then
                    If MsgBox("Ҫ��������ҽ���������ʱ��" & IIF(mint���� = 5, "��ͣ", "����") & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        Call SetSameTime(Row)
                    End If
                End If
            End If
            Call vsAdvice_KeyPress(13) '��λ��һ�����뵥Ԫ
        End With
    Else
        If InStr("0123456789-: " & Chr(8) & Chr(27) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mblnReturn = False
    If Col <> COL_ѡ�� And Col <> COL_���� Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(Row, COL_ID)) = 0 Then
        Cancel = True
    ElseIf mint���� = 1 And Col = COL_���� And vsAdvice.TextMatrix(Row, COL_����) = "1" Then
        Cancel = True 'ֹͣҽ��ʱ,��ҩ�䷽(����)����ֹʱ�䲻���޸�
    ElseIf mint���� = 3 Then
        If Col = COL_���� And Not (vsAdvice.TextMatrix(Row, COL_��־) = "��¼" Or InStr(mstrPrivs, "�޸�У��ʱ��") > 0) Then
            Cancel = True 'У��ҽ��ʱ,�ǲ�¼��У��ʱ�䲻�ɸ���
        ElseIf Col = COL_ѡ�� Then
            Cancel = True '����ֱ�ӱ༭
        End If
    End If
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ID;���ID;��ID;���;�������;�������;��ҩ;,240,4;,300,4;,1530,1;" & _
        "����,750,1;סԺ��,750,1;����,500,1;Ӥ��,500,1;��Ч,500,4;����ʱ��,1080,1;��ʼʱ��,1080,1;" & _
        "ҽ������,3000,1;,375,4;����,850,1;����,850,1;Ƶ��,1000,1;�÷�,1000,1;ҽ������,1000,1;ִ��ʱ��,1000,1;" & _
        "��ֹʱ��,1080,1;ִ�п���,850,1;ִ������,850,1;�ϴ�ִ��,1080,1;��־,500,4;" & _
        "����ҽ��,850,1;У�Ի�ʿ,850,1;У��ʱ��,1080,1;ͣ��ҽ��,850,1;ͣ��ʱ��,1080,1;" & _
        "����ID;��ҳID;��������;ִ�п���ID;���˿���ID;�շ�ϸĿID;������λ;ǰ��ID;ǩ��ID;������Ա"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColHidden(COL_��ʾ) = True 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ҽ��ID;���ID;�������;������ĿID;�շ�ϸĿID;�̶�;" & _
        "�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2500,1;��λ,500,4;����,650,1;����,850,7;" & _
        "ִ�п���,1000,1;��������,850,1;����,450,4;�շ����;ִ�п���ID;��������"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub SetAdviceCol()
'���ܣ�����һЩ�ɼ��м��༭����,Ӧ�ڱ������װ������
    With vsAdvice
        .TextMatrix(0, COL_ѡ��) = ""
        .Editable = flexEDKbdMouse
        
        '������������еĿɼ���
        If mint���� = 0 Then
            'ҽ������
            .ColHidden(COL_����) = True
            .ColHidden(COL_�ϴ�ִ��) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 1 Then
            'ֹͣҽ��
            .TextMatrix(0, COL_����) = "��ֹʱ��"
            .ColHidden(COL_��ֹʱ��) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 2 Then
            'ȷ��ֹͣ
            .TextMatrix(0, COL_����) = "ȷ��ʱ��"
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 3 Then
            'ҽ��У��
            .TextMatrix(0, COL_����) = "У��ʱ��"
            .ColHidden(COL_�ϴ�ִ��) = True
            .ColHidden(COL_У�Ի�ʿ) = True
            .ColHidden(COL_У��ʱ��) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .Cell(flexcpPictureAlignment, .FixedRows, COL_ѡ��, .Rows - 1, COL_ѡ��) = 4
        ElseIf mint���� = 4 Then
            '�����Ƽ���Ŀ
            .ColHidden(COL_����) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 5 Then
            '��ͣҽ��
            .TextMatrix(0, COL_����) = "��ͣʱ��"
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 6 Then
            '����ҽ��
            .TextMatrix(0, COL_����) = "����ʱ��"
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        End If
        
        '���ö�����
        If Not .ColHidden(COL_����) Then
            .FrozenCols = COL_���� + 1 - .FixedCols
            .SheetBorder = vbBlack
        ElseIf Not .ColHidden(COL_ѡ��) Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
            .SheetBorder = vbBlack
        End If
        
        '�������б�ʶ
        .Cell(flexcpBackColor, .FixedRows, COL_ѡ��, .Rows - 1, COL_����) = &HC0FFC0
    End With
End Sub

Private Function GetWhere() As String
'���ܣ����ݴ��幦�ܲ���ҽ��������
'˵��������"����ҽ����¼"����Ϊ"A"
    Dim strSQL As String
    
    If mint���� = 0 Then
        'ҽ������:��У��,��δ���͹�����������������ͣ�ĳ���Ҳ����ֱ�����ϡ�
        '��ʱ����ҽ��У�Ժ��Զ�ֹͣ������Ҳ��������
        strSQL = " And (A.ҽ��״̬ Not IN(1,2,4,8,9) And A.�ϴ�ִ��ʱ�� is NULL Or A.ҽ����Ч=1 And A.������ĿID is Null And A.ҽ��״̬=8)"
    ElseIf mint���� = 1 Then
        'ֹͣҽ��:����,����ͣ��Ҳ����ֱ��ֹͣ,������ҩ�䷽(�и���,�Զ�ͣ)
        strSQL = " And A.ҽ��״̬ Not IN(1,2,4,8,9) And Nvl(A.ҽ����Ч,0)=0 And A.�ܸ����� is NULL"
    ElseIf mint���� = 2 Then
        'ȷ��ֹͣ:ֹͣ״̬�ĳ���(�����Զ�ֹͣ����ҩ�䷽����)
        strSQL = " And A.ҽ��״̬=8 And Nvl(A.ҽ����Ч,0)=0"
    ElseIf mint���� = 3 Then
        'ҽ��У��:���¿��ģ�����ҽ�������ʸ�Ļ�����˵�ҽ������У�ԡ�
        strSQL = " And A.ҽ��״̬=1 And Exists(" & _
            "Select M.���� From ��Ա�� M,ִҵ��� N" & _
            " Where M.����=Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1)" & _
            " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
            " )"
    ElseIf mint���� = 4 Then
        '�����Ƽ���Ŀ
        strSQL = " And A.ҽ��״̬ Not IN(1,2,4,8,9)"
    ElseIf mint���� = 5 Then
        '��ͣҽ��:����,������ҩ�䷽(�и���,��׼��ͣ)
        strSQL = " And A.ҽ��״̬ IN(3,5,7) And Nvl(A.ҽ����Ч,0)=0 And A.�ܸ����� is NULL"
    ElseIf mint���� = 6 Then
        '����ҽ��
        strSQL = " And A.ҽ��״̬=6"
    End If
    GetWhere = strSQL
End Function

Private Function LoadAdvice(strPatis As String) As Boolean
'���ܣ����ݵ�ǰ�������ö�ȡ����ʾҽ���嵥
'������str����IDs=���ڷ���ʵ�������ݵĲ��˴�:"����ID,��ҳID,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsPause As New ADODB.Recordset
    Dim str��ҩ As String, str��ҩ As String
    Dim strSQL As String, strWhere As String
    Dim vCurDate As Date, bln��ҩ;�� As Boolean
    Dim lng����ID As Long, lng��ҳID As Long
    Dim i As Long, j As Long, k As Long
    Dim strӤ�� As String, str����s As String, strTmp As String
    
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
        
    '----------------------------------------------------------------------
    strPatis = ""
    With vsAdvice
        .Rows = .FixedRows
        .ColHidden(COL_����) = True
        .ColHidden(COL_סԺ��) = True
        .ColHidden(COL_����) = True
        .ColHidden(COL_Ӥ��) = True
    End With
    
    '----------------------------------------------------------------------
    strWhere = GetWhere
    'ҽ������ʱ����ʾҽ����,��ʿ�༭ʱ��ʾ����
    strWhere = strWhere & IIF(Not mbln��ʿվ, " And A.ǰ��ID is NULL", "")
    
    'У�Ե�ҽ����Χ����
    If mint���� = 3 And InStr(mstrPrivs, "ȫԺҽ��У��") = 0 Then
        strWhere = strWhere & " And A.����ҽ�� In(" & _
            " Select Distinct B.����" & _
            " From ������Ա A,��Ա�� B,��Ա����˵�� C" & _
            " Where A.��ԱID=B.ID And B.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            "   And A.����ID In(" & _
            "     Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
            "     Where A.��ԱID=(Select ��ԱID From �ϻ���Ա�� Where �û���=User)" & _
            "       And A.����ID=B.����ID)" & _
            ")"
    End If
    
    '��������ʱ���õ�����
    If mint��Ч <> 0 Then
        strWhere = strWhere & " And Nvl(A.ҽ����Ч,0)=" & mint��Ч - 1
    End If
    If mint��� <> 0 Then
        If mint��� = 1 Then
            'ҩƷ��
            strWhere = strWhere & _
                " And (A.������� IN('5','6','7')" & _
                " Or (A.�������='E' And A.���ID is Not NULL)" & _
                " Or Exists(Select ID From ����ҽ����¼ S Where ������� IN('5','6','7') And S.���ID=A.ID And ����ID=[1])" & _
                " )"
        ElseIf mint��� = 2 Then
            '������
            strWhere = strWhere & _
                " And Not A.������� IN('5','6','7')" & _
                " And Not(A.�������='E' And A.���ID is Not NULL)" & _
                " And Not Exists(Select ID From ����ҽ����¼ S Where ������� IN('5','6','7') And S.���ID=A.ID And ����ID=[1])"
        End If
    End If
    
    vCurDate = zlDatabase.Currentdate
    
    '----------------------------------------------------------------------
    For k = 0 To UBound(Split(mstr����IDs, ";"))
        lng����ID = Split(Split(mstr����IDs, ";")(k), ",")(0)
        lng��ҳID = Split(Split(mstr����IDs, ";")(k), ",")(1)
        
        'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨
        strSQL = _
            "Select /*+ RULE */ A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
                " Nvl(A.�������,'*') as �������,C.�������,NULL as ��ҩ,A.�����,NULL as ѡ��,NULL as ����," & _
                " P.����,P.סԺ��,P.��ǰ���� as ����,Decode(Nvl(A.Ӥ��,0),0,'����','Ӥ��'||A.Ӥ��) as Ӥ��,Decode(Nvl(A.ҽ����Ч,0),0,'����','����') as ��Ч," & _
                " To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,A.ҽ������,A.Ƥ�Խ�� as Ƥ��," & _
                " Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ),'5',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,'6',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,A.�ܸ�����||B.���㵥λ)) as ����," & _
                " Decode(A.��������,NULL,NULL,A.��������||B.���㵥λ) as ����," & _
                " A.ִ��Ƶ�� as Ƶ��,Decode(A.�������,'E',Decode(Instr('246',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�," & _
                " A.ҽ������,A.ִ��ʱ�䷽�� as ִ��ʱ��,To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��," & _
                " Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���," & _
                " Decode(Instr('567E',Nvl(A.�������,'*')),0,NULL,A.ִ������) as ִ������," & _
                " To_Char(A.�ϴ�ִ��ʱ��,'YYYY-MM-DD HH24:MI') as �ϴ�ִ��," & _
                " Decode(A.������־,1,'����',2,'��¼','��ͨ') as ��־," & _
                " A.����ҽ��,A.У�Ի�ʿ,To_Char(A.У��ʱ��,'YYYY-MM-DD HH24:MI') as У��ʱ��," & _
                " A.ͣ��ҽ��,To_Char(A.ͣ��ʱ��,'YYYY-MM-DD HH24:MI') as ͣ��ʱ��,A.����ID,A.��ҳID," & _
                " B.��������,A.ִ�п���ID,A.���˿���ID,A.�շ�ϸĿID,B.���㵥λ as ������λ,A.ǰ��ID,S.ǩ��ID,S.������Ա" & _
            " From ����ҽ����¼ A,������Ϣ P,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B,����ҽ��״̬ S,����ҽ����¼ X" & _
            " Where A.����ID=P.����ID And A.������ĿID=B.ID" & IIF(InStr(",0,1,2,3,", mint����) > 0, "(+)", "") & _
                " And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+)" & _
                " And A.�շ�ϸĿID=D.ҩƷID(+) And A.���ID=X.ID(+)" & _
                " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
                " And A.ID=S.ҽ��ID And S.��������=1 And A.����ID=[1] And A.��ҳID=[2]" & _
                " And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & strWhere & _
            " Order by Nvl(A.Ӥ��,0),���,��ID,A.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        
        If Not rsTmp.EOF Then
            strPatis = strPatis & ";" & lng����ID & "," & lng��ҳID
            If InStr(str����s & ",", "," & rsTmp!���˿���ID & ",") = 0 Then
                str����s = str����s & "," & rsTmp!���˿���ID
            End If
            
            '��ͣҽ��ʱ��ȡҽ�����ϴ�����ʱ��(��һ����)
            '����ҽ��ʱ��ȡҽ������ͣʱ��
            If mint���� = 5 Or mint���� = 6 Then
                strSQL = "Select B.ҽ��ID,Max(B.����ʱ��) as �ϴ�ʱ��" & _
                    " From ����ҽ����¼ A,����ҽ��״̬ B" & _
                    " Where A.ID=B.ҽ��ID And B.��������=" & IIF(mint���� = 5, 7, 6) & _
                    " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
                    " And A.����ID=[1] And A.��ҳID=[2] And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & strWhere & _
                    " Group by B.ҽ��ID"
                Set rsPause = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
            End If
            
            With vsAdvice
                .Redraw = flexRDNone
                Do While Not rsTmp.EOF
                    '�������
                    strTmp = ""
                    For i = 0 To rsTmp.Fields.Count - 1
                        strTmp = strTmp & vbTab & Nvl(rsTmp.Fields(i).Value)
                    Next
                    .AddItem Mid(strTmp, 2): i = .Rows - 1
                    
                    '�Ƿ���ʾӤ����
                    If InStr(strӤ�� & ",", "," & .TextMatrix(i, COL_Ӥ��) & ",") = 0 Then
                        If strӤ�� <> "" Then .ColHidden(COL_Ӥ��) = False
                        strӤ�� = strӤ�� & "," & .TextMatrix(i, COL_Ӥ��)
                    End If
                    
                    '����֮��ļ����
                    If .TextMatrix(i, COL_סԺ��) <> .TextMatrix(i - 1, COL_סԺ��) And i - 1 >= .FixedRows Then
                        .CellBorderRange i - 1, .FixedCols, i - 1, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                    End If
                    
                    '��ҩ����ҩ��һЩ����
                    bln��ҩ;�� = False
                    If .TextMatrix(i, COL_�������) = "E" Then
                        If Val(.TextMatrix(i - 1, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                            If InStr(",5,6,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                                bln��ҩ;�� = True
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        '��ʾ��ҩ�ĸ�ҩ;��
                                        .TextMatrix(j, COL_�÷�) = .TextMatrix(i, COL_�÷�)
                                        '��ʾ��ҩ��ִ������
                                        If Val(.TextMatrix(j, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                            .TextMatrix(j, COL_ִ������) = "�Ա�ҩ"
                                        ElseIf Val(.TextMatrix(j, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                            .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ"
                                        Else
                                            .TextMatrix(j, COL_ִ������) = ""
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                                If .TextMatrix(i - 1, COL_�������) = "7" Then
                                    .TextMatrix(i, COL_����) = "1" '��ҩ�䷽
                                ElseIf .TextMatrix(i - 1, COL_�������) = "C" Then
                                    .TextMatrix(i, COL_����) = "2" '�������
                                End If
                                
                                '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                                .TextMatrix(i, COL_ִ�п���) = .TextMatrix(i - 1, COL_ִ�п���)
                                
                                If .TextMatrix(i - 1, COL_�������) = "7" Then
                                    '��ʾ��ҩ�䷽ִ������
                                    If Val(.TextMatrix(i - 1, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                        .TextMatrix(i, COL_ִ������) = "�Ա�ҩ"
                                    ElseIf Val(.TextMatrix(i - 1, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                        .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(i, COL_ִ������) = ""
                                    End If
                                Else
                                    .TextMatrix(i, COL_ִ������) = ""
                                End If
                                
                                'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ;ͬʱ�жϼ�������
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        .RemoveItem j: i = .Rows - 1
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Else
                            .TextMatrix(i, COL_ִ������) = ""
                        End If
                    End If
                                                                    
                    '����ɼ��еĵ�һЩ��ʶ
                    If Not bln��ҩ;�� And .TextMatrix(i, COL_�������) <> "7" Then
                        '����С��������,��δ�뵽�취
                        If Left(.TextMatrix(i, COL_����), 1) = "." Then
                            .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                        End If
                        If Left(.TextMatrix(i, COL_����), 1) = "." Then
                            .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                        End If
                    
                        'ʱ����MM-DD HH:MI��ʽ��ʾ,��CellData�����ж�
                        .Cell(flexcpData, i, COL_��ʼʱ��) = .TextMatrix(i, COL_��ʼʱ��)
                        .Cell(flexcpData, i, COL_����ʱ��) = .TextMatrix(i, COL_����ʱ��)
                        .Cell(flexcpData, i, COL_�ϴ�ִ��) = .TextMatrix(i, COL_�ϴ�ִ��)
                        .Cell(flexcpData, i, COL_��ֹʱ��) = .TextMatrix(i, COL_��ֹʱ��)
                        .Cell(flexcpData, i, COL_У��ʱ��) = .TextMatrix(i, COL_У��ʱ��)
                        .Cell(flexcpData, i, COL_ͣ��ʱ��) = .TextMatrix(i, COL_ͣ��ʱ��)
                        .TextMatrix(i, COL_��ʼʱ��) = Format(.TextMatrix(i, COL_��ʼʱ��), "MM-dd HH:mm")
                        .TextMatrix(i, COL_����ʱ��) = Format(.TextMatrix(i, COL_����ʱ��), "MM-dd HH:mm")
                        .TextMatrix(i, COL_�ϴ�ִ��) = Format(.TextMatrix(i, COL_�ϴ�ִ��), "MM-dd HH:mm")
                        .TextMatrix(i, COL_��ֹʱ��) = Format(.TextMatrix(i, COL_��ֹʱ��), "MM-dd HH:mm")
                        .TextMatrix(i, COL_У��ʱ��) = Format(.TextMatrix(i, COL_У��ʱ��), "MM-dd HH:mm")
                        .TextMatrix(i, COL_ͣ��ʱ��) = Format(.TextMatrix(i, COL_ͣ��ʱ��), "MM-dd HH:mm")
                        
                        If mint���� = 1 Then
                            'ͣ��ʱȱʡ��ҽ����ֹʱ��
                            If .Cell(flexcpData, i, COL_��ֹʱ��) = "" Then
                                'ȱʡִ����ֹʱ��
                                If gbln����ҽ��������Ч Then
                                    .TextMatrix(i, COL_����) = Format(vCurDate + 1, "yyyy-MM-dd 00:00")
                                Else
                                    .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                End If
                                '������͹�,ȱʡ���ϴ�ִ��ʱ��
                                If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                                    If .TextMatrix(i, COL_ִ��ʱ��) = "" And Format(.TextMatrix(i, COL_�ϴ�ִ��), "HH:mm") = "00:00" Then
                                        '"������"�ĳ���,ֹͣ���ղ�����
                                        If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) + 1, "yyyy-MM-dd") Then
                                            .TextMatrix(i, COL_����) = Format(CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) + 1, "yyyy-MM-dd HH:mm")
                                        End If
                                    Else
                                        If .TextMatrix(i, COL_����) < CStr(.Cell(flexcpData, i, COL_�ϴ�ִ��)) Then
                                            .TextMatrix(i, COL_����) = CStr(.Cell(flexcpData, i, COL_�ϴ�ִ��))
                                        End If
                                    End If
                                End If
                            Else
                                .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ֹʱ��)
                            End If
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 2 Then
                            .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            'Ӧ>=��ֹʱ��
                            If .TextMatrix(i, COL_����) < .Cell(flexcpData, i, COL_��ֹʱ��) Then
                                .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_��ֹʱ��))), "yyyy-MM-dd HH:mm")
                            End If
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 3 Then
                            'У��ʱ��ȱʡУ��ʱ��
                            If .TextMatrix(i, COL_��־) = "��¼" Then
                                .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_����ʱ��))), "yyyy-MM-dd HH:mm")
                            Else
                                .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            End If
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 5 Then
                            If mblnPauseLast Then
                                If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                                    'ȱʡ���ϴ�ִ��ʱ��֮����ͣ
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��))), "yyyy-MM-dd HH:mm")
                                Else
                                    '�����ϴ�ִ��ʱ�����Կ�ʼʱ��Ϊ׼
                                    .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ʼʱ��)
                                End If
                            Else
                                '��ͣҽ��ʱ��:��ͣ����,ҽ����ͣ����Ч,���õ���Ч��
                                .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            End If
                            
                            'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
                            If .TextMatrix(i, COL_����) < .Cell(flexcpData, i, COL_��ʼʱ��) Then
                                .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ʼʱ��)
                            End If
                            'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
                            If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                                If .TextMatrix(i, COL_����) <= .Cell(flexcpData, i, COL_�ϴ�ִ��) Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
                            If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                                If .TextMatrix(i, COL_����) >= .Cell(flexcpData, i, COL_��ֹʱ��) Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", -1, CDate(.Cell(flexcpData, i, COL_��ֹʱ��))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
                            rsPause.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_����) <= Format(rsPause!�ϴ�ʱ��, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, rsPause!�ϴ�ʱ��), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 6 Then
                            '����ҽ��ʱ��
                            .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            
                            'Ӧ>��ͣʱ��
                            rsPause.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_����) <= Format(rsPause!�ϴ�ʱ��, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, rsPause!�ϴ�ʱ��), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            'Ӧ<=ִ����ֹʱ��
                            If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                                If .TextMatrix(i, COL_����) > .Cell(flexcpData, i, COL_��ֹʱ��) Then
                                    .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ֹʱ��)
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        End If
                        
                        '�и�
                        If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                        
                        '���龫ҩƷ��ʶ
                        If .TextMatrix(i, COL_�������) <> "" Then
                            If InStr(",����ҩ,����ҩ,����ҩ,", .TextMatrix(i, COL_�������)) > 0 Then
                                .Cell(flexcpFontBold, i, COL_ҽ������) = True
                            End If
                        End If
                        
                        'Ƥ�Խ����ʶ
                        If .TextMatrix(i, COL_Ƥ��) = "(+)" Then
                            .Cell(flexcpForeColor, i, COL_Ƥ��) = vbRed
                        ElseIf .TextMatrix(i, COL_Ƥ��) = "(-)" Then
                            .Cell(flexcpForeColor, i, COL_Ƥ��) = vbBlue
                        End If
                        
                        'Pass:�����������ʾ��ʾ��
                        If .TextMatrix(i, COL_��ʾ) <> "" Then
                            Set .Cell(flexcpPicture, i, COL_��ʾ) = imgPass.ListImages(Val(.TextMatrix(i, COL_��ʾ)) + 1).Picture
                            .Cell(flexcpData, i, COL_��ʾ) = .TextMatrix(i, COL_��ʾ) '���ڵ�ҩ����
                            .TextMatrix(i, COL_��ʾ) = ""
                        End If
                        
                        '����ǩ����ʶ
                        If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 Then
                            Set .Cell(flexcpPicture, i, COL_ҽ������) = img16.ListImages("ǩ��").Picture
                        End If
                    End If
                    
                    If bln��ҩ;�� Then .RemoveItem i
                    
                    Progress = rsTmp.AbsolutePosition / rsTmp.RecordCount * 100
                    
                    rsTmp.MoveNext
                Loop
            End With
        End If
    Next
        
    '----------------------------------------------------------------------
    '������Ϣ��ʾ
    If strPatis <> "" Then strPatis = Mid(strPatis, 2)
    If UBound(Split(strPatis, ";")) = 0 Then
        'ֻ��һ�����˵����ݵ����
        lng����ID = Split(strPatis, ",")(0)
        lng��ҳID = Split(strPatis, ",")(1)
        If lng����ID <> mlng����ID Then '���ǵ�ǰ��������ȡ����ʾ
            strSQL = _
                " Select A.סԺ��,A.����,A.�Ա�,A.����,B.��Ժ����," & _
                " B.סԺҽʦ,B.��Ժ����ID,C.���� as ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                " And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
            lblPati.Caption = "����:" & rsTmp!���� & "��סԺ��:" & Nvl(rsTmp!סԺ��) & _
                "������:" & Nvl(rsTmp!��Ժ����) & "������:" & Nvl(rsTmp!����)
        End If
    ElseIf UBound(Split(strPatis, ";")) > 0 Then
        '�ж���������ݵ����
        vsAdvice.ColHidden(COL_����) = False
        vsAdvice.ColHidden(COL_סԺ��) = False
        vsAdvice.ColHidden(COL_����) = False
                
        strSQL = "Select ���� From ���ű� Where ID IN(" & Mid(str����s, 2) & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        str����s = ""
        Do While Not rsTmp.EOF
            str����s = str����s & "," & rsTmp!����
            rsTmp.MoveNext
        Loop
        lblPati.Caption = "����(" & Mid(str����s, 2) & ") " & UBound(Split(strPatis, ";")) + 1 & " �����˵�ҽ��"
    ElseIf UBound(Split(strPatis, ";")) = -1 Then
        'û���κβ������ݵ����
        lblPati.Caption = ""
    End If
    
    '----------------------------------------------------------------------
    If vsAdvice.Rows = vsAdvice.FixedRows Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        vsPrice.Rows = vsPrice.FixedRows
        vsPrice.Rows = vsPrice.FixedRows + 1
    Else
        '����ǩ��ͼ�����
        vsAdvice.Cell(flexcpPictureAlignment, vsAdvice.FixedRows, COL_ҽ������, vsAdvice.Rows - 1, COL_ҽ������) = 0
        '�Զ������и�
        vsAdvice.AutoSize COL_ҽ������
    End If
    Call SetAdviceCol
    vsAdvice.Row = vsAdvice.FixedRows
    If Not vsAdvice.ColHidden(COL_ѡ��) Then
        vsAdvice.Col = COL_ѡ��
    Else
        vsAdvice.Col = COL_ҽ������
    End If
    vsAdvice.Redraw = flexRDDirect
    
    Progress = 0: Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    strPatis = ""
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Function CheckValid() As Boolean
'���ܣ�ȷ��ǰ���Ϸ���
    Dim str���� As String, str���� As String
    Dim str���� As String, strTmp As String
    Dim curDate As Date, i As Long, k As Long
    Dim strPatis As String
    
    mstrRollNotify = ""
    curDate = zlDatabase.Currentdate
    
    With vsAdvice
        '�Ƿ��п��Բ����ļ�¼
        If .Rows = .FixedRows + 1 And Val(.TextMatrix(.FixedRows, COL_ID)) = 0 Then
            If mint���� = 0 Then
                'ҽ������
                strTmp = "��ǰû�п������ϵ�ҽ����"
            ElseIf mint���� = 1 Then
                'ֹͣҽ��
                strTmp = "��ǰû�п���ֹͣ��ҽ����"
            ElseIf mint���� = 2 Then
                'ȷ��ֹͣ
                strTmp = "��ǰû�б�ֹͣ��ҽ����"
            ElseIf mint���� = 3 Then
                'ҽ��У��
                strTmp = "��ǰû���¿���ҽ����"
            ElseIf mint���� = 4 Then
                '�����Ƽ���Ŀ
                strTmp = "��ǰû��ͨ��У�Ե���Чҽ����"
            ElseIf mint���� = 5 Then
                '��ͣҽ��
                strTmp = "��ǰû�п�����ͣ��ҽ����"
            ElseIf mint���� = 6 Then
                '����ҽ��
                strTmp = "��ǰû����ͣ����Ҫ���õ�ҽ����"
            End If
            If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            Exit Function
        End If
        
        '�Ƿ���ѡ��
        str���� = "": str���� = "": str���� = ""
        If Not .ColHidden(COL_ѡ��) Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And (Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) <> Empty) Then
                    k = k + 1
                    If InStr(strPatis & ",", "," & .TextMatrix(i, COL_����ID)) = 0 Then
                        strPatis = strPatis & "," & .TextMatrix(i, COL_����ID)
                    End If
                    
                    If mint���� = 1 Then
                        '�ռ����ڷ��͵�ҽ��
                        If IsDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) Then
                            If .TextMatrix(i, COL_ִ��ʱ��) = "" And Format(.TextMatrix(i, COL_�ϴ�ִ��), "HH:mm") = "00:00" Then
                                '"������"����,ֹͣ���첻����
                                If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) + 1, "yyyy-MM-dd") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������)
                                End If
                            Else
                                If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������)
                                End If
                            End If
                        End If
                        
                        '�ռ�����ֹͣ��ҽ��
                        If CDate(.TextMatrix(i, COL_����)) - curDate > 7 Then
                            str���� = str���� & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������) & "��ֹͣʱ�䣺" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm")
                        End If
                    ElseIf mint���� = 2 Then
                        '�ռ����ڷ��͵�ҽ��
                        If IsDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) Then
                            If .TextMatrix(i, COL_ִ��ʱ��) = "" And Format(.TextMatrix(i, COL_�ϴ�ִ��), "HH:mm") = "00:00" Then
                                '"������"����,ֹͣ���첻����
                                If Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) + 1, "yyyy-MM-dd") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������)
                                End If
                            Else
                                If Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������)
                                End If
                            End If
                        End If
                    ElseIf mint���� = 3 Then
                        '�ռ�����ҽ��,ͨ��У�ԵĲ��ж�
                        If .Cell(flexcpData, i, COL_ѡ��) = 1 And _
                            .TextMatrix(i, COL_�������) = "Z" And .TextMatrix(i, COL_��������) = "4" Then
                            If InStr(str���� & ";", ";" & .TextMatrix(i, COL_����ID) & "," & .TextMatrix(i, COL_��ҳID) & ";") = 0 Then
                                str���� = str���� & ";" & .TextMatrix(i, COL_����ID) & "," & .TextMatrix(i, COL_��ҳID)
                            End If
                        End If
                    End If
                End If
            Next
            If k = 0 Then
                MsgBox "û��ѡ���κ�ҽ������ѡ����Ҫ" & tbr.Buttons("ִ��").Caption & "��ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
                
        'ҽ��
        If mint���� = 1 And mbln��ʿվ Then
            If cboҽ��.ListIndex = -1 Then
                MsgBox "��ѡ��ֹͣҽ����ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    strTmp = ""
    strPatis = IIF(UBound(Split(Mid(strPatis, 2), ",")) > 0, "��ѡ���˶�����˵�ҽ��������ϸ���м���Ա�����ֲ��" & vbCrLf & vbCrLf, "")
    If mint���� = 0 Then
        'ҽ������
        strTmp = "ȷʵҪ�����Ѿ�ѡ���ҽ����"
    ElseIf mint���� = 1 Then
        'ֹͣҽ��
        If str���� <> "" Then '����Ƿ�����Ҫ�˻س�ǰ��ҩ�����
            If MsgBox("����Ҫֹͣ��ҽ�������ڷ��ͣ�" & vbCrLf & str���� & _
                vbCrLf & vbCrLf & "����ҽ�������ڻ�ʿ����վ��ʹ��""���ڷ����ջ�""���д���" & _
                vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str���� <> "" Then
            If MsgBox("����ҽ����ֹͣʱ�䳬����ǰʱ��̫�ã�" & vbCrLf & str���� & _
                vbCrLf & vbCrLf & "���ֹͣʱ�䲻��ȷ�������ҽ���ķ��ͺͼƷѲ���Ӱ�졣" & _
                vbCrLf & "ȷʵҪ��ָ����ʱ��ֹͣ��Щҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str���� = "" And str���� = "" Then
            strTmp = "ȷʵҪֹͣ�Ѿ�ѡ���ҽ����"
        End If
    ElseIf mint���� = 2 Then
        'ȷ��ֹͣ
        If str���� <> "" Then
            If MsgBox("����ֹͣ��ҽ�������ڷ��ͣ�" & vbCrLf & str���� & _
                vbCrLf & vbCrLf & "����ҽ�������ڻ�ʿ����վ��ʹ��""���ڷ����ջ�""���д���" & _
                vbCrLf & "ȷ���Ѿ�ѡ���ҽ��ֹͣ��", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            strTmp = "ȷ���Ѿ�ѡ���ҽ��ֹͣ��"
        End If
    ElseIf mint���� = 3 Then
        'ҽ��У��
        If str���� <> "" Then
            If MsgBox(strPatis & "ҪУ�Ե�ҽ���а�������ҽ����У�Ժ��ֹͣ��������ҽ����" & _
                vbCrLf & "ȷʵҪ����У����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            mstrRollNotify = Mid(str����, 2)
        Else
            strTmp = strPatis & "ȷʵҪ���Ѿ�ѡ���ҽ������У�Դ�����"
        End If
    ElseIf mint���� = 5 Then
        '��ͣҽ��
        strTmp = strPatis & "ȷʵҪ��ͣ�Ѿ�ѡ���ҽ����"
    ElseIf mint���� = 6 Then
        '����ҽ��
        strTmp = strPatis & "ȷʵҪ�����Ѿ�ѡ���ҽ����"
    End If
    If strTmp <> "" Then
        If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    CheckValid = True
End Function

Private Function CheckSignValid() As Boolean
'���ܣ�1.���δǩ����ҽ�����ܽ���У��
'      2.һ��ǩ����ҽ������һ��ͨ��У��
    Dim colҽ��ID As New Collection, strҽ��ID As String
    Dim colǩ��ID As New Collection, strǩ��ID As String
    Dim strסԺ As String, strҽ�� As String
    Dim lngǩ��ID As Long, strTmp As String
    Dim int״̬ As Integer, i As Long, j As Long
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNurse As String
    
    If mint���� <> 3 Then CheckSignValid = True: Exit Function
    
    With vsAdvice
        '��ȡ��ʿ��Ա�б�ֻ�ǻ�ʿ������ҽ��
        If Mid(gstrESign, 2, 1) = "1" Or Mid(gstrESign, 3, 1) = "1" Then
            strNurse = ""
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                    If .Cell(flexcpData, i, COL_ѡ��) = 1 And Val(.TextMatrix(i, COL_ǩ��ID)) = 0 Then
                        If InStr(strNurse & ",", "," & .TextMatrix(i, COL_������Ա) & ",") = 0 Then
                            strNurse = strNurse & "," & .TextMatrix(i, COL_������Ա)
                        End If
                    End If
                End If
            Next
            If strNurse <> "" Then
                strSQL = "Select /*+ Rule*/ A.����" & _
                    " From ��Ա�� A,(Select * From Table(Cast(f_Str2List([1]) As zlTools.t_StrList))) B" & _
                    " Where A.����=B.Column_Value" & _
                    " And Exists(Select 1 From ��Ա����˵�� X Where X.��ԱID=A.ID And X.��Ա����='��ʿ')" & _
                    " And Not Exists(Select 1 From ��Ա����˵�� Y Where Y.��ԱID=A.ID And Y.��Ա����='ҽ��')"
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strNurse, 2))
                On Error GoTo 0
                
                strNurse = ""
                Do While Not rsTmp.EOF
                    strNurse = strNurse & "," & rsTmp!����
                    rsTmp.MoveNext
                Loop
                strNurse = strNurse & ","
            End If
        End If
        
        For i = .FixedRows To .Rows - 1
            'flexcpData:0-������,1-У��,2-����
            If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                '1.�ռ�δǩ����ҽ������
                If .Cell(flexcpData, i, COL_ѡ��) = 1 And Val(.TextMatrix(i, COL_ǩ��ID)) = 0 Then
                    '����Ϊʹ��ǩ���ĳ���
                    If InStr(strNurse, "," & .TextMatrix(i, COL_������Ա) & ",") = 0 Then '��ʿ¼���ҽ��������ǩ�����
                        If Val(.TextMatrix(i, COL_ǰ��ID)) = 0 And Mid(gstrESign, 2, 1) = "1" Then
                            If UBound(Split(strסԺ, vbCrLf)) < 10 Then
                                strסԺ = strסԺ & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������)
                            ElseIf InStr(strסԺ, "�� ��") = 0 Then
                                strסԺ = strסԺ & vbCrLf & "�� ��"
                            End If
                        ElseIf Val(.TextMatrix(i, COL_ǰ��ID)) <> 0 And Mid(gstrESign, 3, 1) = "1" Then
                            If UBound(Split(strҽ��, vbCrLf)) < 10 Then
                                strҽ�� = strҽ�� & vbCrLf & "��" & .TextMatrix(i, COL_ҽ������)
                            ElseIf InStr(strҽ��, "�� ��") = 0 Then
                                strҽ�� = strҽ�� & vbCrLf & "�� ��"
                            End If
                        End If
                    End If
                End If
                
                '2.�ռ���ǩ��ҽ����У��״̬
                lngǩ��ID = Val(.TextMatrix(i, COL_ǩ��ID))
                If lngǩ��ID <> 0 Then
                    j = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID))) '��ID
                    int״̬ = .Cell(flexcpData, i, COL_ѡ��)
                    If int״̬ = 2 Then int״̬ = 0 '�������ʵ�ͬ�ڲ�У��
                    If InStr(strǩ��ID & ",", "," & lngǩ��ID & ",") > 0 Then
                        '�ռ�����ǩ���ڽ����ϵ�У��״̬
                        strTmp = Split(colǩ��ID("_" & lngǩ��ID), "=")(1)
                        If InStr(strTmp, int״̬) = 0 Then
                            colǩ��ID.Remove "_" & lngǩ��ID
                            colǩ��ID.Add lngǩ��ID & "=" & strTmp & int״̬, "_" & lngǩ��ID
                        End If
                        
                        '�ռ�����ǩ���Ѷ��������ҽ��(��ID)
                        strTmp = colҽ��ID("_" & lngǩ��ID)
                        If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                            colҽ��ID.Remove "_" & lngǩ��ID
                            colҽ��ID.Add strTmp & "," & j, "_" & lngǩ��ID
                        End If
                    Else
                        strǩ��ID = strǩ��ID & "," & lngǩ��ID
                        colǩ��ID.Add lngǩ��ID & "=" & int״̬, "_" & lngǩ��ID
                        colҽ��ID.Add j, "_" & lngǩ��ID
                    End If
                End If
            End If
        Next
        
        '�����ǩ��ҽ��У�����
        strTmp = "": strҽ��ID = Mid(strҽ��ID, 2)
        For i = 1 To colǩ��ID.Count
            lngǩ��ID = Split(colǩ��ID(i), "=")(0)
            strǩ��ID = Split(colǩ��ID(i), "=")(1)
            
            '����һ��ǩ����δ��������δУ��ҽ��
            strҽ��ID = colҽ��ID("_" & lngǩ��ID)
            strҽ��ID = ExistOtherSignAdvice(lngǩ��ID, strҽ��ID)
            If strҽ��ID <> "" Then
                If InStr(strǩ��ID, "0") = 0 Then
                    strǩ��ID = strǩ��ID & "0"
                    strTmp = strTmp & strҽ��ID
                End If
            End If
            
            If Not (strǩ��ID = "1" Or strǩ��ID = "0") Then
                '���ǩ�������ݲ���"��Ҫͨ��У�Ի򶼲�ͨ��У��(��������)"�����
                j = .FindRow(CStr(lngǩ��ID), , COL_ǩ��ID)
                Do While j <> -1
                    If Val(.TextMatrix(j, COL_ID)) <> 0 And Not .RowHidden(j) Then
                        If InStr(",0,2,", .Cell(flexcpData, j, COL_ѡ��)) > 0 Then
                            strTmp = strTmp & vbCrLf & .TextMatrix(j, COL_����) & "��" & IIF(Len(.TextMatrix(j, COL_ҽ������)) > 40, Left(.TextMatrix(j, COL_ҽ������), 40) & "...", .TextMatrix(j, COL_ҽ������))
                        End If
                    End If
                    j = .FindRow(CStr(lngǩ��ID), j + 1, COL_ǩ��ID)
                Loop
                Exit For '��ֻ��ʾ��һ��
            End If
        Next
    End With
    
    '1.û��ǩ����ҽ��������У�ԣ���סԺҽ����ҽ��ҽ���ֱ���м��
    If strסԺ <> "" Then
        MsgBox "����ҽ��ҽ����û��ǩ�������ܽ���У�ԣ�" & vbCrLf & strסԺ, vbInformation, gstrSysName
        Exit Function
    End If
    If strҽ�� <> "" Then
        MsgBox "����ҽ��ҽ����û��ǩ�������ܽ���У�ԣ�" & vbCrLf & strҽ��, vbInformation, gstrSysName
        Exit Function
    End If
    
    '2.һ��ǩ����ҽ������һ��ͨ��У��
    If strTmp <> "" Then
        MsgBox "����ҽ������������Ҫͨ��У�Ե�ҽ��һ��ǩ��������ǰ����Ϊ��У�Ի�У�����ʣ�" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "һ��ǩ����ҽ������һ��ͨ��У�ԣ���������ҽ����У��״̬��", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckSignValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExistOtherSignAdvice(ByVal lngǩ��ID As Long, ByVal strҽ��ID As String) As String
'���ܣ�����Ƿ����ĳ���¿�ҽ��ǩ���б���û�ж�ȡ�������ϵ�ҽ��(��ΪҪһ��ͨ��У��,�����,��Щҽ��Ҳ��ûУ�Ե�)
'���أ�δ��ȡ�������δУ��ҽ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.����,B.ҽ������ From ����ҽ��״̬ A,����ҽ����¼ B" & _
        " Where A.ҽ��ID=B.ID And A.��������=1 And B.ҽ��״̬ IN(1,2)" & _
        " And (B.���ID is Null Or B.������� IN('5','6'))" & _
        " And Not Exists(Select 1 From ����ҽ����¼ S Where ������� IN('5','6') And S.���ID=B.ID)" & _
        " And Instr([2],','||Nvl(B.���ID,B.ID)||',')=0 And A.ǩ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngǩ��ID, "," & strҽ��ID & ",")
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & vbCrLf & Nvl(rsTmp!����) & "��" & IIF(Len(Nvl(rsTmp!ҽ������)) > 40, Left(Nvl(rsTmp!ҽ������), 40) & "...", Nvl(rsTmp!ҽ������))
        rsTmp.MoveNext
    Loop
    ExistOtherSignAdvice = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lngҽ��ID As Long, ByVal lng��ĿID As Long, ByVal lngCol As Long)
'���ܣ���λ������ʾָ��ҽ����ָ���Ƽ���
'������lngRow=ҽ���к�,lngҽ��ID=�Ƽ�ҽ��ID
'      lng��ĿID=�Ƽ���ĿID,lngCol=�Ƽ۱����ʾ��
    Dim k As Long
    
    With vsAdvice
        .Row = lngRow: .Col = COL_ҽ������ '�������Զ�ShowPrice,mrsPrice�����仯
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_ҽ��ID)) = lngҽ��ID _
                And Val(vsPrice.TextMatrix(k, COLP_�շ�ϸĿID)) = lng��ĿID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function ExecuteOperate() As Boolean
    Dim arrSQL As Variant, lng���ID As Long
    Dim blnExe As Boolean, i As Long, j As Long
    Dim lngҽ��ID As Long, lngִ�п���ID As Long
    Dim strOper As String, blnVarZero As Boolean
    Dim strҽ��ID As String, intRule As Integer
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim strSource As String, strSign As String
    Dim colStopTime As New Collection
    
    Screen.MousePointer = 11
    
    '����SQL
    arrSQL = Array()
    With vsAdvice
        If mint���� <> 4 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) <> Empty Then
                    'һ��ҽ��ֻУ��һ��,��һ����ҩ��,����ҽ��ֻ��һ����ʾ��
                    blnExe = False
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> lng���ID Then blnExe = True
                    Else
                        blnExe = True
                    End If
                    If blnExe Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        '(��ID)ʹ�����IDΪNULL��ҽ����ID(��ҩ;��,��ҩ�÷�,�����Ŀ,��Ҫ����,������ҽ��)
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            lngҽ��ID = Val(.TextMatrix(i, COL_���ID))
                        Else
                            lngҽ��ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If mint���� = 0 Then      'ҽ������
                            'ҽ������ҽ������ǩ��
                            If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 Then
                                strҽ��ID = strҽ��ID & "," & lngҽ��ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_����(" & lngҽ��ID & ")"
                        ElseIf mint���� = 1 Then  'ֹͣҽ��
                            'ҽ��ֹͣҽ������ǩ��
                            If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 Then
                                strҽ��ID = strҽ��ID & "," & lngҽ��ID
                                '��¼ֹͣҽ����ִ����ֹʱ�䣺��������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                                colStopTime.Add Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm:00"), "_" & lngҽ��ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_ֹͣ(" & lngҽ��ID & "," & _
                                "To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                "'" & IIF(mbln��ʿվ, NeedName(cboҽ��.Text), UserInfo.����) & "')"
                        ElseIf mint���� = 2 Then  'ȷ��ֹͣ
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_ȷ��ֹͣ(" & lngҽ��ID & "," & _
                            "To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        ElseIf mint���� = 3 Then  'ҽ��У��
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_У��(" & lngҽ��ID & "," & _
                                IIF(.Cell(flexcpData, i, COL_ѡ��) = 1, 3, 2) & "," & _
                                "To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        ElseIf mint���� = 5 Then  '��ͣҽ��
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_��ͣ(" & lngҽ��ID & ",To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        ElseIf mint���� = 6 Then  '����ҽ��
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_����(" & lngҽ��ID & ",To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                        End If
                    End If
                End If
                lng���ID = Val(.TextMatrix(i, COL_���ID))
            Next
        End If
        
        'ҽ���Ƽ۲���
        lng���ID = 0
        If mint���� = 3 Or mint���� = 4 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) = 1 Then
                    'һ����ҩ��ֻ�账��һ��
                    blnExe = False
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> lng���ID Then blnExe = True
                    Else
                        blnExe = True
                    End If
                    
                    If blnExe Then
                        'ɾ����Ӧ�ļƼ�
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            arrSQL(UBound(arrSQL)) = "zl_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_���ID)) & ")"
                        Else
                            arrSQL(UBound(arrSQL)) = "zl_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        End If
                        
                        '�����µļƼ�
                        '������һ����ѭ����Щ,��Ϊ���ж��Ƿ�Ҫ���漰����Ϸ���,������Filter
                        If Val(vsAdvice.TextMatrix(i, COL_���ID)) <> 0 Then
                            mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                " Or ҽ��ID=" & Val(vsAdvice.TextMatrix(i, COL_���ID))
                        Else
                            mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                " Or ���ID=" & vsAdvice.TextMatrix(i, COL_ID)
                        End If
                        For j = 1 To mrsPrice.RecordCount
                            '֮�д����շ�ϸĿIDΪ�յ����ü�¼(�������ȷ����ѡ�ļƼ�ҽ��)
                            If Not IsNull(mrsPrice!�շ�ϸĿID) And InStr(",5,6,7,", mrsPrice!�������) = 0 Then
                                If Nvl(mrsPrice!����, 0) <> 0 Then '��������Ϊ0���Զ����˵�
                                    blnVarZero = False
                                    If Nvl(mrsPrice!����, 0) = 0 Then
                                        blnVarZero = ItemIsVarPrice(mrsPrice!�շ�ϸĿID)
                                    End If
                                    If blnVarZero Then
                                        Call SeekPriceRow(i, mrsPrice!ҽ��ID, mrsPrice!�շ�ϸĿID, COLP_����)
                                        Screen.MousePointer = 0
                                        MsgBox "����Ϊ��۵��շ���Ŀȷ��һ���շѼ۸�", vbInformation, gstrSysName
                                        vsPrice.SetFocus: Exit Function
                                    End If
                                    
                                    '�Ƽ�ִ�п���:ֻ�����ҩ��ҩƷ�����ļƼ�
                                    If InStr(",5,6,7,", mrsPrice!�շ����) > 0 _
                                        Or mrsPrice!�շ���� = "4" And Nvl(mrsPrice!����, 0) = 1 Then
                                        lngִ�п���ID = Nvl(mrsPrice!ִ�п���ID, 0)
                                        
                                        '���ı�������ִ�п���
                                        If lngִ�п���ID = 0 And mrsPrice!�շ���� = "4" Then
                                            Call SeekPriceRow(i, mrsPrice!ҽ��ID, mrsPrice!�շ�ϸĿID, COLP_ִ�п���)
                                            Screen.MousePointer = 0
                                            MsgBox "����""" & vsPrice.TextMatrix(vsPrice.Row, COLP_�շ���Ŀ) & """û��ȷ��ִ�п��ң����ֹ�������ȷ��ִ�п��ҡ�" & vbCrLf & _
                                                "�������ȷ����ȷ��ִ�п��ң��뵽""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                            vsPrice.SetFocus: Exit Function
                                        End If
                                    Else
                                        lngִ�п���ID = 0
                                    End If
                                    
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = "zl_����ҽ���Ƽ�_Insert(" & mrsPrice!ҽ��ID & "," & _
                                        mrsPrice!�շ�ϸĿID & "," & mrsPrice!���� & "," & Nvl(mrsPrice!����, 0) & "," & _
                                        Nvl(mrsPrice!����, 0) & "," & ZVal(lngִ�п���ID) & ")"
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    End If
                End If
                lng���ID = Val(.TextMatrix(i, COL_���ID))
            Next
        End If
    End With
    
    '���ϻ�ֹͣʱ�ĵ���ǩ��
    If (mint���� = 0 Or mint���� = 1) And strҽ��ID <> "" Then
        strOper = Decode(mint����, 0, "����", 1, "ֹͣ")
        
        '��ʿ�������ϡ�ֹͣҽ����ǩ����ҽ��
        If mbln��ʿվ Then
            MsgBox "��Ҫ" & strOper & "��ҽ���а���ҽ����ǩ����ҽ����ֻ����ҽ����" & strOper & "��ǩ����", vbInformation, gstrSysName
            Screen.MousePointer = 0: Exit Function
        End If
        
        'ҽ��ֹͣ,����ʱ����Ҫǩ��
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox strOper & "��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ�����" & strOper & "��", vbInformation, gstrSysName
            Else
                MsgBox strOper & "��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ������" & strOper & "��", vbInformation, gstrSysName
            End If
            Screen.MousePointer = 0: Exit Function
        End If
        
        '��ȡǩ��ҽ��Դ��
        strҽ��ID = Mid(strҽ��ID, 2) '��ID,����Ϊ��ϸID
        intRule = ReadAdviceSignSource(Decode(mint����, 0, 4, 1, 8), mlng����ID, mlng��ҳID, strҽ��ID, 0, False, strSource, , colStopTime)
        If intRule = 0 Then Screen.MousePointer = 0: Exit Function
        If strSource = "" Then
            Screen.MousePointer = 0
            MsgBox "���ܶ�ȡ��Ҫ" & strOper & "����ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
        If strSign <> "" Then
            lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & "," & Decode(mint����, 0, 4, 1, 8) & "," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��ID & "')"
        Else
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    
    'ִ��SQL
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    Screen.MousePointer = 0
    ExecuteOperate = True
    Exit Function
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitPriceRecordset()
'˵�����༭ʱ,���Ƽ�ҽ�����շ���Ŀ�������,�ż����¼��
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "�������", adVarChar, 1
    mrsPrice.Fields.Append "������ĿID", adBigInt
    mrsPrice.Fields.Append "�շ����", adVarChar, 1, adFldIsNullable
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "����", adInteger '�����Ƿ��������
    mrsPrice.Fields.Append "����", adInteger
    mrsPrice.Fields.Append "�̶�", adInteger '���е��շѹ�ϵ���Ƿ�̶�����
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub ShowDefaultRow()
'���ܣ����ڿ��ԼƼ۵�ҽ��,ȱʡ����һ�в�����ȱʡ�Ƽ�ҽ��
'˵����ComboList="#ҽ��ID1;�Ƽ�ҽ��1|#ҽ��ID2;�Ƽ�ҽ��2|..."
'      ���ڵ�һ����ʾ�Ƽ۱�ͻس�������ʱ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrCombo As Variant, lngRow As Long
    Dim lngҽ��ID As Long, str�Ƽ�ҽ�� As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If .ColData(COLP_�Ƽ�ҽ��) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_�Ƽ�ҽ��), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_ҽ��ID)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_�շ�ϸĿID)) <> 0 Then
                '��һ����ʾʱȱʡ����һ��
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '���ǵ�һ����ʾʱȱʡ�Ƽ�ҽ������һ����ͬ
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_�̶�)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_ҽ��ID)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lngҽ��ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str�Ƽ�ҽ�� = Replace(arrCombo(i), "#" & lngҽ��ID & ";", "")
                If blnHave Then
                    If lngҽ��ID = Val(.TextMatrix(lngRow - 1, COLP_ҽ��ID)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            'ģ��ѡ������Ƽ�ҽ��
            strSQL = "Select ���ID,�������,������ĿID From ����ҽ����¼ Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COLP_ҽ��ID) = lngҽ��ID
                .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
                .TextMatrix(lngRow, COLP_���ID) = Nvl(rsTmp!���ID)
                .TextMatrix(lngRow, COLP_������ĿID) = rsTmp!������ĿID
                .TextMatrix(lngRow, COLP_�������) = rsTmp!�������
                .Cell(flexcpData, lngRow, COLP_�Ƽ�ҽ��) = .TextMatrix(lngRow, COLP_�Ƽ�ҽ��)
                
                'ֻ��һ���Ƽ�ҽ��ʱ����ͣ��
                If UBound(arrCombo) = 0 Then
                    .Col = COLP_�շ���Ŀ
                Else
                    .Col = COLP_�Ƽ�ҽ��
                End If
            End If
        End If
        Call .ShowCell(.Row, .Col)
        If blnFirst Then .TopRow = .Row '��һ����ʾʱ,ShowCell��Ȼ��������
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset, strSQL As String, i As Long
    Dim lngԭ��ID As Long, lngҽ��ID As Long, lng�շ�ϸĿID As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_�Ƽ�ҽ�� Then
            '�������ComboData,TextMatrixȡֵ��ΪComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lngҽ��ID = .ComboData
                lngԭ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                
                '���üƼ�ҽ���Ƿ�������ͬ�շ�ϸĿ
                If lng�շ�ϸĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�Ѿ��������շ���Ŀ""" & .TextMatrix(Row, COLP_�շ���Ŀ) & """��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                                                                
                'ԭ����ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                If lngԭ��ID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                                                                
                '������ݣ�mrsPrice�п�����ɾ��,����Ҫ�����ݿ��
                strSQL = "Select ���ID,�������,������ĿID From ����ҽ����¼ Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                If rsTmp.EOF Then
                    MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�����Ѿ���������ɾ��,���˳����½��롣", vbInformation, gstrSysName
                    Exit Sub
                End If
                .TextMatrix(Row, COLP_ҽ��ID) = lngҽ��ID
                .TextMatrix(Row, COLP_���ID) = Nvl(rsTmp!���ID)
                .TextMatrix(Row, COLP_������ĿID) = rsTmp!������ĿID
                .TextMatrix(Row, COLP_�������) = rsTmp!�������
                
                '��¼������
                If lng�շ�ϸĿID <> 0 Then
                    '��ѡ���ҽ���Ƿ��д�������޸ĺ����Ŀ�Ƿ����
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ����=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_����) = IIF(blnHaveSub, "��", "")
                    
                    If lngԭ��ID = 0 Then
                        mrsPrice.AddNew '����
                    Else '����
                        mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    End If
                    mrsPrice!ҽ��ID = lngҽ��ID
                    mrsPrice!���ID = rsTmp!���ID
                    mrsPrice!������ĿID = rsTmp!������ĿID
                    mrsPrice!������� = rsTmp!�������
                    If lngԭ��ID = 0 Then
                        mrsPrice!�շ�ϸĿID = lng�շ�ϸĿID
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!�̶� = 0
                    End If
                    mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            End If
        ElseIf Col = COLP_�շ���Ŀ Or Col = COLP_ִ�п��� Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), "0.00000")
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str��ĿIDs As String, blnCancel As Boolean
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim vPoint As POINTAPI
    
    With vsPrice
        If Col = COLP_�շ���Ŀ Then
            '����ѡ�����е���Ŀ
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_ҽ��ID)) = Val(.TextMatrix(Row, COLP_ҽ��ID)) _
                    And Val(.TextMatrix(Row, COLP_ҽ��ID)) <> 0 And i <> Row Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                End If
            Next
            str��ĿIDs = Mid(str��ĿIDs, 2)
            
            strSQL = _
                " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
                " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',7,'��������') as ����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as �Ƿ���ID,Null as ���ID,-NULL as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7)"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,-ID as ID,Nvl(-�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as �Ƿ���ID,Null as ���ID,-NULL as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as �Ƿ���ID,Null as ���ID,-NULL as ��������ID" & _
                " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select ĩ��,ID,�ϼ�ID,����,����,��λ,���,����,���,��������,˵��," & _
                " Decode(Nvl(�Ƿ���,0),1,Decode(Instr('567',���ID),0,Sum(ԭ��)||'-'||Sum(�ּ�),'ʱ��'),Sum(�ּ�)) as �۸�," & _
                " Sum(ԭ��) as ԭ��ID,Sum(�ּ�) as �ּ�ID,�Ƿ��� as �Ƿ���ID,���ID,��������ID" & _
                " From (" & _
                " Select Distinct 1 as ĩ��,A.ID,Decode(Instr('567',A.���),0,A.����ID,-E.����ID) as �ϼ�ID,A.����,A.����," & _
                " A.���㵥λ as ��λ,A.���,A.����,A.��� as ���ID,C.���� as ���,A.��������,A.˵��,B.ԭ��,B.�ּ�,A.�Ƿ���," & _
                " -NULL as ��������ID" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,������ĿĿ¼ E" & _
                " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.��� Not IN('4','J','1') And A.���=C.���� And A.ID=D.ҩƷID(+) And D.ҩ��ID=E.ID(+)"
            If DeptExist("���ϲ���", 2) Then
                strSQL = strSQL & " Union ALL" & _
                    " Select Distinct 1 as ĩ��,A.ID,-E.����ID as �ϼ�ID,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,A.��� as ���ID,C.���� as ���,A.��������,A.˵��," & _
                    " B.ԭ��,B.�ּ�,A.�Ƿ���,D.�������� as ��������ID" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,�������� D,������ĿĿ¼ E" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And A.ID Not IN(" & str��ĿIDs & ")", "") & _
                    " And A.���='4' And A.���=C.���� And A.ID=D.����ID And D.����ID=E.ID"
            End If
            strSQL = strSQL & " ) Group by ĩ��,ID,�ϼ�ID,����,����,��λ,���,����,���,��������,˵��,�Ƿ���,���ID,��������ID"
            
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "�շ���Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str��ĿIDs & ",")
            If Not rsTmp Is Nothing Then
                'ҽ��������
                If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ҳID))) Then
                    .SetFocus: Exit Sub
                End If
                
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                Call SetItemInput(Row, rsTmp, lngҽ��ID, lngԭ��ĿID)
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ��շ���Ŀ�����ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_ִ�п��� Then
            vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_�շ����) = "4" Then
                '�������õ�����
                strSQL = _
                    " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                    " And B.������� IN(2,3) And B.����ID=C.ID" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (A.������Դ is NULL Or A.������Դ=2)" & _
                    " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                    " And A.�շ�ϸĿID=[1]" & _
                    " Order by B.�������,C.����"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                'ҩƷ
                'ҩƷ��ϵͳָ���Ĵ���ҩ������
                If Not Check�ϰల��(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                    Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!ִ�п���ID = rsTmp!ID
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��������(ѡ��)�Ƽ���Ŀ�Ƿ�ҽ������
'���أ����δ���룬������ʾѡ�񲻼������򷵻��档
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int���� As Integer
    
    If gintҽ������ = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then int���� = Nvl(rsTmp!����, 0)
    If int���� <> 0 Then
        If Not ItemExistInsure(rsInput!ID, int����) Then
            If gintҽ������ = 1 Then
                If MsgBox("��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gintҽ������ = 2 Then
                MsgBox "��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lngҽ��ID As Long, ByVal lngԭ��ĿID As Long)
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng����ID As Long, lng��ҳID As Long
    Dim lng�к� As Long, dbl���� As Double
    Dim blnHaveSub As Boolean, dbl���� As Double
    Dim rsTmp As ADODB.Recordset
    
    With vsPrice
        '�������
        .TextMatrix(lngRow, COLP_�շ����) = rsInput!���ID
        .TextMatrix(lngRow, COLP_�շ�ϸĿID) = rsInput!ID
        .TextMatrix(lngRow, COLP_���) = rsInput!���
        .TextMatrix(lngRow, COLP_�շ���Ŀ) = rsInput!����
        If Not IsNull(rsInput!����) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & "(" & rsInput!���� & ")"
        End If
        If Not IsNull(rsInput!���) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & " " & rsInput!���
        End If
        
        '����Ǽ���ҩƷ�Ƽ�(��ҩ��),�����۵�λ����
        .TextMatrix(lngRow, COLP_����) = 1 'ȱʡ�Ƽ�����Ϊ1
        .TextMatrix(lngRow, COLP_��λ) = Nvl(rsInput!��λ)
                
        '���ۼ��㴦��:ҩ���Ƽ۲����������ﴦ��,��ҩ��ҩƷ�Ƽ۰��ۼ۴���
        .Cell(flexcpData, lngRow, 0) = 0
        .Cell(flexcpData, lngRow, 1) = 0
        .Cell(flexcpData, lngRow, 2) = 0
        
        'ִ�п���
        lng�к� = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
        If lng�к� = -1 Then
            Set rsTmp = GetItemField("����ҽ����¼", lngҽ��ID)
            lng����ID = rsTmp!����ID
            lng��ҳID = Nvl(rsTmp!��ҳID, 0)
            lngִ�п���ID = Nvl(rsTmp!ִ�п���ID, 0)
            lng���˿���ID = Nvl(rsTmp!���˿���ID, 0)
            dbl���� = Nvl(rsTmp!�ܸ�����, 0)
            If dbl���� = 0 Then dbl���� = 1
        Else
            lng����ID = Val(vsAdvice.TextMatrix(lng�к�, COL_����ID))
            lng��ҳID = Val(vsAdvice.TextMatrix(lng�к�, COL_��ҳID))
            lngִ�п���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))
            lng���˿���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID))
            dbl���� = Val(vsAdvice.TextMatrix(lng�к�, COL_����))
            If dbl���� = 0 Then dbl���� = 1
        End If
            
        '��ҩ���͸������õ�����ר����ִ�п���
        If InStr(",5,6,7,", rsInput!���ID) > 0 Or rsInput!���ID = "4" And Nvl(rsInput!��������ID, 0) = 1 Then
            lngִ�п���ID = Get�շ�ִ�п���ID(lng����ID, lng��ҳID, rsInput!���ID, rsInput!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID)
            '��¼�����Ƿ��������
            If rsInput!���ID = "4" Then
                .TextMatrix(lngRow, COLP_��������) = Nvl(rsInput!��������ID, 0)
            End If
        End If
        If lngִ�п���ID <> 0 Then
            mrsDept.Filter = "ID=" & lngִ�п���ID
            If Not mrsDept.EOF Then
                .TextMatrix(lngRow, COLP_ִ�п���) = mrsDept!����
            End If
        End If
        .TextMatrix(lngRow, COLP_ִ�п���ID) = lngִ�п���ID
                
        '����
        If InStr(",5,6,7,", rsInput!���ID) > 0 Then
            If Nvl(rsInput!�Ƿ���ID, 0) = 0 Then
                dbl���� = Nvl(rsInput!�ּ�ID, 0)
            Else 'δȷ���Ƽ�ҽ��ʱ,ҩƷ�޷�����۸�
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, dbl����, , True) '��ȱʡ�Ƽ�����Ϊ1�����۵�λ����
            End If
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, "0.00000")
        ElseIf rsInput!���ID = "4" And Nvl(rsInput!��������ID, 0) = 1 And Nvl(rsInput!�Ƿ���ID, 0) = 1 Then
            '�������õ�ʱ�����ĺ�ҩƷһ������
            dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, dbl����, , True)
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, "0.00000")
        Else
            If Nvl(rsInput!�Ƿ���ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_����) = Format(Nvl(rsInput!�ּ�ID, 0), "0.00000")
            Else
                .Cell(flexcpData, lngRow, 0) = 1
                .Cell(flexcpData, lngRow, 1) = Nvl(rsInput!ԭ��ID, 0)
                .Cell(flexcpData, lngRow, 2) = Nvl(rsInput!�ּ�ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_��������) = Nvl(rsInput!��������)
        .TextMatrix(lngRow, COLP_�̶�) = "0"
        
        '��������ָ�
        .Cell(flexcpData, lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_ִ�п���) = .TextMatrix(lngRow, COLP_ִ�п���)
        
        '��¼������
        If lngҽ��ID <> 0 Then
            If lngԭ��ĿID = 0 Then
                '��ǰҽ���Ƿ��д��������������Ŀ�Ƿ����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ����=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_����) = IIF(blnHaveSub, "��", "")

                mrsPrice.AddNew '����
            Else '����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
            End If
            If lngԭ��ĿID = 0 Then
                mrsPrice!ҽ��ID = lngҽ��ID
                mrsPrice!���ID = Val(.TextMatrix(lngRow, COLP_���ID))
                mrsPrice!������� = .TextMatrix(lngRow, COLP_�������)
                mrsPrice!������ĿID = Val(.TextMatrix(lngRow, COLP_������ĿID))
                mrsPrice!���� = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!�շ���� = rsInput!���ID
            mrsPrice!�շ�ϸĿID = rsInput!ID
            If lngִ�п���ID <> 0 Then
                mrsPrice!ִ�п���ID = lngִ�п���ID
            Else
                mrsPrice!ִ�п���ID = Null
            End If
            mrsPrice!���� = Nvl(rsInput!��������ID, 0)
            mrsPrice!���� = 1
            mrsPrice!���� = Val(.TextMatrix(lngRow, COLP_����))
            mrsPrice!�̶� = 0
            mrsPrice.Update
            Call SelectRow(vsAdvice.Row)
        End If
    End With
End Sub

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_�Ƽ�ҽ�� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_�̶�)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_ҽ��ID)) <> 0 And Val(.TextMatrix(.Row, COLP_�շ�ϸĿID)) <> 0 Then
                    'ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(.Row, COLP_ҽ��ID)) & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_�Ƽ�ҽ��) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("ȷ��Ҫɾ����ǰ�Ƽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(.Row, COLP_ҽ��ID)) & " And �շ�ϸĿID=" & Val(.TextMatrix(.Row, COLP_�շ�ϸĿID))
                    mrsPrice.Delete
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_�շ���Ŀ Or .Col = COLP_ִ�п���) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str��ĿIDs As String
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim StrInput As String, strMatch As String
    Dim vPoint As POINTAPI
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COLP_�Ƽ�ҽ�� Then
                '����ʱ�س�
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '��ȻEnterNextCell����Ҫ�˳�
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շ�����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շѵ���������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��������뷶Χ
                strTmp = CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, "0.00000")
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_�շ���Ŀ And .EditText <> "" Then
                '����ѡ�����е���Ŀ
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COLP_ҽ��ID)) = Val(.TextMatrix(Row, COLP_ҽ��ID)) _
                        And Val(.TextMatrix(Row, COLP_ҽ��ID)) <> 0 And i <> Row Then
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                    End If
                Next
                str��ĿIDs = Mid(str��ĿIDs, 2)
                
                '��ͬ������ƥ�䷽ʽ
                StrInput = UCase(.EditText)
                strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
                If IsNumeric(StrInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
                ElseIf zlCommFun.IsCharAlpha(StrInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
                ElseIf zlCommFun.IsCharChinese(StrInput) Then
                    strMatch = " And C.���� Like [2] And C.����=[3]"
                End If
                
                strSQL = ""
                If Not DeptExist("���ϲ���", 2) Then strSQL = " And A.���<>'4'"
                strSQL = _
                    " Select A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,A.˵��," & _
                    " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(A.ԭ��)||'-'||Sum(A.�ּ�),'ʱ��'),Sum(A.�ּ�)) as �۸�," & _
                    " Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,B.�������� as ��������ID" & _
                    " From (" & _
                    " Select Distinct 1 as ĩ��,A.ID,A.��� as ���ID,D.���� as ���,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,A.��������,A.˵��,B.ԭ��,B.�ּ�,A.�Ƿ���" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ���� C,�շ���Ŀ��� D" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.�շ�ϸĿID And A.���=D.���� And A.��� Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,�������� B" & _
                    " Where A.ID=B.����ID(+)" & _
                    " Group by A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,A.˵��,A.�Ƿ���,A.���ID,B.��������" & _
                    " Order by A.���,A.����"
                vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ���Ŀ", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    StrInput & "%", mstrLike & StrInput & "%", mint���� + 1, "," & str��ĿIDs & ",")
                If Not rsTmp Is Nothing Then
                    'ҽ��������
                    If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ҳID))) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                        .SetFocus: Exit Sub
                    End If
                    
                    lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    Call SetItemInput(Row, rsTmp, lngҽ��ID, lngԭ��ĿID)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õ��շ���Ŀ��", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            ElseIf Col = COLP_ִ�п��� And .EditText <> "" Then 'ִ�п���
                vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_�շ����) = "4" Then
                    '�������õ�����
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And A.�շ�ϸĿID=[1] And (C.���� Like [3] Or C.���� Like [4] Or C.���� Like [4])" & _
                        " Order by B.�������,C.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                    'ҩƷ��ϵͳָ���Ĵ���ҩ������
                    If Not Check�ϰల��(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                        Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!����
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    
                    '���¼�¼��
                    lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                        mrsPrice!ִ�п���ID = rsTmp!ID
                        mrsPrice.Update
                        Call SelectRow(vsAdvice.Row)
                    End If
                    
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_���� Or Col = COLP_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_�Ƽ�ҽ�� Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_�շ���Ŀ Or NewCol = COLP_ִ�п��� Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not CellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = COLP_���� Or Col = COLP_���� Or Col = COLP_ִ�п��� Then
        If vsPrice.TextMatrix(Row, COLP_�շ���Ŀ) = "" Then
            Cancel = True '������ȷ���շ���Ŀ
        End If
    End If
    
    If Col = COLP_���� Or Col = COLP_���� Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ��жϼ۱��е�Ԫ���Ƿ���Ա༭
    CellEditable = vsPrice.Editable
    With vsPrice
        If lngCol = COLP_ִ�п��� Then
            '�������õ�����,��ҩ��ҩƷ�Ƽ۵�ִ�п��ҿ����޸�
            If Not (.TextMatrix(lngRow, COLP_�շ����) = "4" And Val(.TextMatrix(lngRow, COLP_��������)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�շ����)) > 0 And InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�������)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Or .TextMatrix(lngRow, COLP_�������) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_�̶�)) <> 0 Then
            '�̶������н������޸ı��
            If Not (.Cell(flexcpData, lngRow, 0) = 1 And lngCol = COLP_����) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_���� Then
                If .Cell(flexcpData, lngRow, 0) <> 1 Then CellEditable = False
            ElseIf lngCol <> COLP_�Ƽ�ҽ�� And lngCol <> COLP_���� And lngCol <> COLP_�շ���Ŀ Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ���۱�����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    
    With vsPrice
        '��ǰ��Ԫ�����δ��������,���˳�
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_���� And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_�Ƽ�ҽ��) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '��ǰ�����û���ҵ���һ���ɱ༭��Ԫ,�������Ƽ�ҽ��,������һ����
            If CStr(.ColData(COLP_�Ƽ�ҽ��)) <> "" Then
                '��ǰ��δ��������,��λ����������Ԫ
                If .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = "" Then
                    .Col = COLP_�Ƽ�ҽ��
                ElseIf .TextMatrix(lngRow, COLP_����) = "" Then
                    .Col = COLP_����
                ElseIf .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Then
                    .Col = COLP_�շ���Ŀ
                ElseIf .Cell(flexcpData, lngRow, 0) = 1 And Val(.TextMatrix(lngRow, COLP_����)) = 0 Then
                    .Col = COLP_����
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_�Ƽ�ҽ��
                    
                    'ȱʡѡ��Ƽ�ҽ��(�������)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '���ɱ༭ʱ���ⶨһ��
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Function LoadPrice(ByVal lngRow As Long, Optional blnChange As Boolean) As Boolean
'���ܣ���ȡָ��ҽ���ļƼ�,�����ݵ�ǰ�������շѹ�ϵ���и���
'���أ�blnChange=�Ƿ���ݵ�ǰ�������շѹ�ϵ�����еļƼ����ݽ����˵���
    Dim rsMan As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim rsAdd As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim blnLoad As Boolean, lng������ĿID As Long
    Dim dblPrice As Double, blnSubItem As Boolean
    Dim lngִ�п���ID As Long
    
    On Error GoTo errH
    
    With vsAdvice
        '�Ѿ���ȡ����,�����ظ���ȡ
        If .TextMatrix(lngRow, COL_ID) = "" Then LoadPrice = True: Exit Function
        If .RowData(lngRow) = 1 Then LoadPrice = True: Exit Function
                            
        'ҩƷ�ļƼ�(�����������ʾ������Ϊ�������,ҩƷ�̶�Ϊ1��ʵ��ҩƷ������ʾʱ����)
        'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��,����ҩ:���ܰ������ҽ��,����1��סԺ��װ�ĵ���
            strSQL = _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,C.��� as �շ����,C.ID as �շ�ϸĿID," & _
                " 1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.סԺ��װ as ����,0 as ����," & _
                " A.ִ�п���ID,0 as ��������,C.����ʱ��" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (A.�շ�ϸĿID is NULL Or A.�շ�ϸĿID=B.ҩƷID)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(2,3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        ElseIf .TextMatrix(lngRow, COL_����) = "1" Then
            '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
            strSQL = _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,C.��� as �շ����,C.ID as �շ�ϸĿID," & _
                " 1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.סԺ��װ as ����,0 as ����," & _
                " A.ִ�п���ID,0 as ��������,C.����ʱ��" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.�������='7' And A.���ID=[1]" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID And C.������� IN(2,3)" & _
                " And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        End If
        
        '��ȡ���мƼۣ���ҩƷ��ļƼ�,�������ҽ���Ƽ�
        blnLoad = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��ҩ;��:һ����ҩ��ֻ��ȡһ��������
            If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                If .TextMatrix(lngRow - 1, COL_���ID) = .TextMatrix(lngRow, COL_���ID) Then
                    blnLoad = False
                End If
            End If
        End If
        If blnLoad Then
            '��ҩ�ĸ�ҩ;������ҩ�䷽�ļ巨���÷�����鼰��λ����������������,������Ŀ
            '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�У���ҽ������ȡ
            '��Union��ʽ������������
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,B.��� as �շ����,A.�շ�ϸĿID," & _
                "   A.����,A.����,Nvl(A.����,0) as ����,A.ִ�п���ID,C.��������,B.����ʱ��" & _
                " From (" & _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,B.�շ�ϸĿID,B.����,B.����,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B" & _
                " Where A.������� Not IN('5','6','7') And A.ID=B.ҽ��ID(+) And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And A.ID=[1]" & _
                " Union ALL" & _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,B.�շ�ϸĿID,B.����,B.����,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B" & _
                " Where A.������� Not IN('5','6','7') And A.ID=B.ҽ��ID(+) And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And A.ID=[2]" & _
                " Union ALL" & _
                " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,B.�շ�ϸĿID,B.����,B.����,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B" & _
                " Where A.������� Not IN('5','6','7') And A.ID=B.ҽ��ID(+) And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And A.���ID=[1]" & _
                " ) A,�շ���ĿĿ¼ B,�������� C" & _
                " Where A.�շ�ϸĿID=B.ID(+) And A.�շ�ϸĿID=C.����ID(+)" & _
                " Order by ���,����"
        End If
        Set rsMan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
        
        '�����շѹ�ϵ���շ����������ж����Ƿ�仯
        strSQL = "Select C.������ĿID,C.�շ���ĿID,C.�շ�����,C.���ж���,C.������Ŀ" & _
            " From ����ҽ����¼ A,����ҽ���Ƽ� B,�����շѹ�ϵ C" & _
            " Where A.ID=B.ҽ��ID And A.������ĿID=C.������ĿID And B.�շ�ϸĿID=C.�շ���ĿID" & _
            " And (A.ID=[1] Or A.ID=[2] Or A.���ID=[1])"
        Set rsCur = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
        
        '����ҩƷ�����еļƼ�
        For i = 1 To rsMan.RecordCount
            mrsPrice.AddNew '��δ����Ƽ۹�ϵ��ҲҪ��������ȷ���ɼƼ�ҽ��(�ü�¼����)
            mrsPrice!ҽ��ID = rsMan!ID
            mrsPrice!���ID = rsMan!���ID
            mrsPrice!������� = rsMan!�������
            mrsPrice!������ĿID = rsMan!������ĿID
            mrsPrice!�̶� = IIF(InStr(",5,6,7,", rsMan!�������) > 0, 1, 0)
            
            '������ҽ���Ƽ�ʱ,����ԭ������,�����ѳ�������Ŀ,����δ����(�Ա���������)
            If Not IsNull(rsMan!�շ�ϸĿID) _
                And Format(Nvl(rsMan!����ʱ��, "3000-01-01"), "yyyy-MM-dd") = "3000-01-01" Then
                mrsPrice!�շ���� = rsMan!�շ����
                mrsPrice!�շ�ϸĿID = rsMan!�շ�ϸĿID
                mrsPrice!ִ�п���ID = rsMan!ִ�п���ID
                mrsPrice!���� = Nvl(rsMan!��������, 0)
                mrsPrice!���� = rsMan!����
                
                'ҩƷ(��������ʾ)�����Ϊʱ�ۣ���ʾʱ���㣻�������ȡ�����¼۸�
                '��ҩƷ�����Ϊ���,��ȡ��ǰ����(�����)����������ȡ���¼۸�
                mrsPrice!���� = rsMan!����
                mrsPrice!���� = Nvl(rsMan!����, 0)
                        
                '�����շѹ�ϵ���շ����������ж����Ƿ�仯
                If InStr(",5,6,7,", rsMan!�������) = 0 Then '������ҩ����ҩƷ�Ƽ�
                    rsCur.Filter = "������ĿID=" & rsMan!������ĿID & " And �շ���ĿID=" & rsMan!�շ�ϸĿID
                    If Not rsCur.EOF Then
                        If Nvl(rsCur!���ж���, 0) <> 0 And Nvl(rsMan!����, 0) <> Nvl(rsCur!�շ�����, 0) Then
                            mrsPrice!���� = rsCur!�շ����� '����˹��ж��ղ�ȡ�����õ�����
                            blnChange = True
                        End If
                        mrsPrice!���� = Nvl(rsCur!������Ŀ, 0)
                        mrsPrice!�̶� = Nvl(rsCur!���ж���, 0)
                    End If
                    '�۸�ȡ���µ�(�Ǳ��)
                    dblPrice = CalcPrice(rsMan!�շ�ϸĿID)
                    If dblPrice <> 0 Then mrsPrice!���� = Format(dblPrice, "0.00000")
                End If
            End If
            mrsPrice.Update
            If mrsPrice!���� = 1 Then blnSubItem = True '���ڴ�����Ŀ
            
            '�����շѹ�ϵ�������˵Ķ���(��δУ��֮ǰ,����ҽ���Ƽ�û������,��ʱҲ��������ӵ�)
            If InStr(",5,6,7,", rsMan!�������) = 0 Then '������ҩ����ҩƷ�Ƽ�
                lng������ĿID = rsMan!������ĿID
                blnLoad = False: rsMan.MoveNext
                If rsMan.EOF Then
                    blnLoad = True
                ElseIf rsMan!������ĿID <> lng������ĿID Then
                    blnLoad = True
                End If
                rsMan.MovePrevious
                If blnLoad Then
                    strSQL = _
                        " Select A.������ĿID,C.��� as �շ����,A.�շ���ĿID,A.�շ�����,A.���ж���,Nvl(A.������Ŀ,0) as ������Ŀ," & _
                        " C.���,B.���˿���ID,B.ִ�п���ID,E.��������,Sum(Decode(Nvl(C.�Ƿ���,0),1,NULL,D.�ּ�)) as ����" & _
                        " From �����շѹ�ϵ A,����ҽ����¼ B,�շ���ĿĿ¼ C,�շѼ�Ŀ D,�������� E" & _
                        " Where A.������ĿID=B.������ĿID And B.ID=[1] And C.ID=E.����ID(+)" & _
                        " And A.�շ���ĿID Not IN(Select �շ�ϸĿID From ����ҽ���Ƽ� Where ҽ��ID=[1])" & _
                        " And A.�շ���ĿID=C.ID And A.�շ���ĿID=D.�շ�ϸĿID And C.������� IN(2,3)" & _
                        " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                        " Group by A.������ĿID,C.���,A.�շ���ĿID,A.�շ�����,A.���ж���,Nvl(A.������Ŀ,0),C.���,B.���˿���ID,B.ִ�п���ID,E.��������" & _
                        " Order by ������Ŀ"
                    Set rsAdd = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMan!ID))
                    If Not rsAdd.EOF Then
                        For j = 1 To rsAdd.RecordCount
                            '��ҩ���͸������õ�����ר����ִ�п���
                            lngִ�п���ID = Nvl(rsAdd!ִ�п���ID, 0)
                            If InStr(",5,6,7,", rsAdd!���) > 0 Or rsAdd!��� = "4" And Nvl(rsAdd!��������, 0) = 1 Then
                                lngִ�п���ID = Get�շ�ִ�п���ID(Val(.TextMatrix(lngRow, COL_����ID)), Val(.TextMatrix(lngRow, COL_��ҳID)), rsAdd!���, rsAdd!�շ���ĿID, 4, Nvl(rsAdd!���˿���ID, 0), 0, 2, lngִ�п���ID)
                            End If
                            
                            mrsPrice.AddNew
                            mrsPrice!ҽ��ID = rsMan!ID
                            mrsPrice!���ID = rsMan!���ID
                            mrsPrice!������� = rsMan!�������
                            mrsPrice!������ĿID = rsMan!������ĿID
                            mrsPrice!�շ���� = rsAdd!�շ����
                            mrsPrice!�շ�ϸĿID = rsAdd!�շ���ĿID
                            If lngִ�п���ID <> 0 Then
                                mrsPrice!ִ�п���ID = lngִ�п���ID
                            Else
                                mrsPrice!ִ�п���ID = Null
                            End If
                            mrsPrice!���� = Nvl(rsAdd!��������, 0)
                            mrsPrice!���� = rsAdd!�շ�����
                            mrsPrice!���� = rsAdd!����
                            mrsPrice!���� = Nvl(rsAdd!������Ŀ, 0)
                            mrsPrice!�̶� = Nvl(rsAdd!���ж���, 0)
                            mrsPrice.Update
                            
                            If mrsPrice!���� = 1 Then blnSubItem = True '���ڴ�����Ŀ
                            If Nvl(mrsPrice!����, 0) <> 0 Then blnChange = True '�б仯
                            
                            rsAdd.MoveNext
                        Next
                        
                        'ȷ���˶�Ӧ�շ���Ŀ,ɾ�����շ���Ŀ�����ü�¼
                        mrsPrice.Filter = "ҽ��ID=" & rsMan!ID
                        Do While Not mrsPrice.EOF
                            If IsNull(mrsPrice!�շ�ϸĿID) Then
                                mrsPrice.Delete
                                mrsPrice.Update
                            End If
                            mrsPrice.MoveNext
                        Loop
                    End If
                    
                    '�Դ��ڴ���ļƼ۽��д�����ֻ֤��һ������
                    If blnSubItem Then
                        j = 0
                        strSQL = _
                            " Select Sum(Decode(������Ŀ,1,1,0)) as ������," & _
                            " Max(Decode(������Ŀ,1,NULL,�շ���ĿID)) as ����ID" & _
                            " From �����շѹ�ϵ Where ������ĿID=[1]"
                        Set rsAdd = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMan!������ĿID))
                        If Not rsMan.EOF Then j = Nvl(rsAdd!������, 0)
                        If j = 0 Then
                            '������мƼ�û�д�����Ŀ����ȡ�����д�������
                            mrsPrice.Filter = "ҽ��ID=" & rsMan!ID
                            Do While Not mrsPrice.EOF
                                If mrsPrice!���� = 1 Then
                                    mrsPrice!���� = 0
                                    mrsPrice.Update
                                    blnChange = True
                                End If
                                mrsPrice.MoveNext
                            Loop
                        Else
                            '������ڴ�����Ŀ�������������ȫ������Ϊ����
                            mrsPrice.Filter = "ҽ��ID=" & rsMan!ID
                            Do While Not mrsPrice.EOF
                                If mrsPrice!�շ�ϸĿID = Val(Nvl(rsAdd!����ID, 0)) Then 'Ϊʲôһ��Ҫ��Val?
                                    If mrsPrice!���� = 1 Then
                                        mrsPrice!���� = 0 '����϶�����ֻ��һ��
                                        mrsPrice.Update
                                        blnChange = True
                                    End If
                                Else
                                    If mrsPrice!���� = 0 Then
                                        mrsPrice!���� = 1
                                        mrsPrice.Update
                                        blnChange = True
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Loop
                        End If
                    End If
                    blnSubItem = False '�µ�һ��ҽ����ʼ�ж�
                End If
            End If
            
            rsMan.MoveNext
        Next
        .RowData(lngRow) = 1
    End With
    
    LoadPrice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowPrice(ByVal lngRow As Long)
'���ܣ���ʾ��ǰҽ���еļƼ�����(�������ҽ���ļƼ���Ŀ),ͬʱ����һЩ�༭����
    Dim rs������Ŀ As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    Dim strҽ��IDs As String, str�շ�ϸĿIDs As String
    Dim strSQL As String, strTmp As String
    Dim str�Ƽ�ҽ�� As String, i As Long, j As Long
    Dim blnNoFirst As Boolean, lngBegin As Long
    Dim blnAllFixed As Boolean, blnHavePrice As Boolean
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    
    On Error GoTo errH
    
    With vsPrice
        .Redraw = False
        '�����Ŀ���
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        .Editable = flexEDNone
        
        '�Ƿ�һ����ҩ�еķǵ�һҩƷ��
        If RowInһ����ҩ(lngRow, lngBegin, 0) Then
            If lngRow > lngBegin Then blnNoFirst = True
        End If
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            If blnNoFirst Then
                'һ����ҩʱ����һ����ʾ��ҩ;���ļƼ�
                mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
            Else
                mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                    " Or ҽ��ID=" & Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
            End If
        Else
            mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                " Or ���ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
        End If
        
        If Not mrsPrice.EOF Then
'            If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
'                mrsPrice.Sort = "�������" 'һ����ҩʱ��ʾ˳��Ҫ��ҩƷ��ǰ
'            Else
'                mrsPrice.Sort = ""
'            End If
                        
            '��ȡ������Ŀ,�շ�ϸĿ,�۸���Ϣ
            For i = 1 To mrsPrice.RecordCount
                strҽ��IDs = strҽ��IDs & "," & mrsPrice!ҽ��ID
                If Not IsNull(mrsPrice!�շ�ϸĿID) Then
                    str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & mrsPrice!�շ�ϸĿID
                End If
                mrsPrice.MoveNext
            Next
            strҽ��IDs = Mid(strҽ��IDs, 2)
            str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 2)
                        
            strSQL = "Select B.ID,C.���� as �������,B.����,B.�걾��λ" & _
                " From ����ҽ����¼ A,������ĿĿ¼ B,������Ŀ��� C" & _
                " Where A.ID IN(" & strҽ��IDs & ") And A.������ĿID=B.ID And B.���=C.����"
            Call zlDatabase.OpenRecordset(rs������Ŀ, strSQL, Me.Caption) 'In
            
            '��ȡ�Ƿ��ۼ���۷�Χ����Ŀ��Ϣ
            If str�շ�ϸĿIDs <> "" Then
                strSQL = _
                    " Select A.ID,C.���� as �������,A.����,A.����,A.���," & _
                    " A.����,A.���㵥λ,D.סԺ��λ,A.��������,A.�Ƿ���,D.סԺ��װ" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ��� C,ҩƷ��� D" & _
                    " Where A.���=C.���� And A.ID=D.ҩƷID" & _
                    " And A.��� IN('5','6','7') And A.ID IN(" & str�շ�ϸĿIDs & ")"
                strSQL = strSQL & " Union ALL " & _
                    " Select A.ID,C.���� as �������,A.����,A.����,A.���,A.����," & _
                    " A.���㵥λ,NULL as סԺ��λ,A.��������,A.�Ƿ���,-NULL as סԺ��װ" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ��� C" & _
                    " Where A.���=C.���� And A.��� Not IN('5','6','7')" & _
                    " And A.ID IN(" & str�շ�ϸĿIDs & ")"
                
                strSQL = _
                    " Select A.ID,A.�������,A.����,A.����,A.���,A.����,A.���㵥λ," & _
                    " A.סԺ��λ,A.��������,A.�Ƿ���,A.סԺ��װ,Sum(B.ԭ��) as ԭ��,Sum(B.�ּ�) as �ּ�" & _
                    " From (" & strSQL & ") A,�շѼ�Ŀ B Where A.ID=B.�շ�ϸĿID" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " Group by A.ID,A.�������,A.����,A.����,A.���,A.����,A.���㵥λ,A.סԺ��װ,A.��������,A.�Ƿ���,A.סԺ��λ"

                strSQL = _
                    " Select A.ID,A.�������,A.����,Nvl(B.����,A.����) as ����,A.���,A.����," & _
                    " A.���㵥λ,A.סԺ��λ,A.��������,A.�Ƿ���,A.ԭ��,A.�ּ�,A.סԺ��װ" & _
                    " From (" & strSQL & ") A,�շ���Ŀ���� B" & _
                    " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIF(gbln��Ʒ��, 3, 1)
                Call zlDatabase.OpenRecordset(rs�շ�ϸĿ, strSQL, Me.Caption) 'In
            End If
                        
            'ȷ����ʾ����
            If str�շ�ϸĿIDs <> "" Then
                .Rows = .FixedRows + UBound(Split(str�շ�ϸĿIDs, ",")) + 1
            End If
                                    
            '��ʾÿ������
            j = .FixedRows
            blnAllFixed = True: blnHavePrice = False
            mrsPrice.MoveFirst
            For i = 1 To mrsPrice.RecordCount
                'ȷ���Ƽ�ҽ������
                rs������Ŀ.Filter = "ID=" & mrsPrice!������ĿID
                If InStr(",5,6,7,", mrsPrice!�������) > 0 Then
                    str�Ƽ�ҽ�� = "ҩƷҽ��-" & rs������Ŀ!����
                ElseIf mrsPrice!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    str�Ƽ�ҽ�� = "��ҩ;��-" & rs������Ŀ!����
                ElseIf mrsPrice!������� = "E" And InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_����))) > 0 Then
                    If vsAdvice.TextMatrix(lngRow, COL_����) = "2" Then
                        str�Ƽ�ҽ�� = "�ɼ�����-" & rs������Ŀ!����
                    ElseIf Not IsNull(mrsPrice!���ID) Then
                        str�Ƽ�ҽ�� = "��ҩ�巨-" & rs������Ŀ!����
                    Else
                        str�Ƽ�ҽ�� = "��ҩ�÷�-" & rs������Ŀ!����
                    End If
                ElseIf Not IsNull(mrsPrice!���ID) Then
                    If mrsPrice!������� = "C" Then
                        str�Ƽ�ҽ�� = "������Ŀ-" & rs������Ŀ!����
                    ElseIf mrsPrice!������� = "D" Then
                        str�Ƽ�ҽ�� = "��鲿λ-" & rs������Ŀ!�걾��λ
                    ElseIf mrsPrice!������� = "F" Then
                        str�Ƽ�ҽ�� = "��������-" & rs������Ŀ!����
                    ElseIf mrsPrice!������� = "G" Then
                        str�Ƽ�ҽ�� = "������Ŀ-" & rs������Ŀ!����
                    End If
                Else
                    str�Ƽ�ҽ�� = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!����
                End If
                
                '����ѡ��ļƼ�ҽ��(������δ�����շѹ�ϵ��)
                If mrsPrice!�̶� = 0 Then
                    If InStr(strTmp, "|#" & mrsPrice!ҽ��ID & ";" & str�Ƽ�ҽ��) = 0 Then
                        strTmp = strTmp & "|#" & mrsPrice!ҽ��ID & ";" & str�Ƽ�ҽ��
                    End If
                End If
                '����Ѷ�Ӧ���շѹ�ϵȫΪ�̶�,������ѡ��
                If InStr(",5,6,7,", mrsPrice!�������) = 0 Then
                    If Not IsNull(mrsPrice!�շ�ϸĿID) Then
                        '����ҩƷ,�Ƿ��мƼ۹�ϵ
                        blnHavePrice = True
                        '����ҩƷ�ļƼ۹�ϵ�Ƿ�ȫ��Ϊ�̶�
                        blnAllFixed = blnAllFixed And (mrsPrice!�̶� <> 0)
                    End If
                End If
                
                '��δ�����շѹ�ϵ�Ĳ���ʾ,������ѡ��
                If Not IsNull(mrsPrice!�շ�ϸĿID) Then
                    rs�շ�ϸĿ.Filter = "ID=" & mrsPrice!�շ�ϸĿID
                    
                    '��ʾ�Ƽ۵�ҽ������
                    .TextMatrix(j, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
                    .TextMatrix(j, COLP_ҽ��ID) = mrsPrice!ҽ��ID
                    .TextMatrix(j, COLP_���ID) = Nvl(mrsPrice!���ID)
                    .TextMatrix(j, COLP_�������) = mrsPrice!�������
                    .TextMatrix(j, COLP_������ĿID) = mrsPrice!������ĿID
                        
                    '��ʾ����Ƽ۵���Ŀ
                    .TextMatrix(j, COLP_�շ����) = mrsPrice!�շ����
                    .TextMatrix(j, COLP_�շ�ϸĿID) = mrsPrice!�շ�ϸĿID
                    .TextMatrix(j, COLP_���) = rs�շ�ϸĿ!�������
                    .TextMatrix(j, COLP_�շ���Ŀ) = rs�շ�ϸĿ!����
                    If Not IsNull(rs�շ�ϸĿ!����) Then
                        .TextMatrix(j, COLP_�շ���Ŀ) = .TextMatrix(j, COLP_�շ���Ŀ) & "(" & rs�շ�ϸĿ!���� & ")"
                    End If
                    If Not IsNull(rs�շ�ϸĿ!���) Then
                        .TextMatrix(j, COLP_�շ���Ŀ) = .TextMatrix(j, COLP_�շ���Ŀ) & " " & rs�շ�ϸĿ!���
                    End If
                    
                    '��ҩ��ҩƷ���ۼ۵�λ����
                    If InStr(",5,6,7,", mrsPrice!�������) Then
                        .TextMatrix(j, COLP_��λ) = Nvl(rs�շ�ϸĿ!סԺ��λ)
                    Else
                        .TextMatrix(j, COLP_��λ) = Nvl(rs�շ�ϸĿ!���㵥λ)
                    End If
                    'ҩ��ȱʡΪ1,��ҩ��ҩƷ������(�ۼ۵�λ)
                    .TextMatrix(j, COLP_����) = FormatEx(mrsPrice!����, 5)
                    
                    'ҩ��ҩƷΪ��1��סԺ��λ����ļ۸�
                    .TextMatrix(j, COLP_����) = Format(Nvl(mrsPrice!����), "0.00000")
                    
                    'ִ�п���
                    lngִ�п���ID = Nvl(mrsPrice!ִ�п���ID, 0)
                    '��ҩ��ҩƷ��������õ����ļƼۿ�������ִ�п���
                    If mrsPrice!�շ���� = "4" And Nvl(mrsPrice!����, 0) = 1 _
                        Or InStr(",5,6,7,", mrsPrice!�շ����) > 0 And InStr(",5,6,7,", mrsPrice!�������) = 0 Then
                        '�Ե�ǰֵ��Ϊȱʡ����ȡ��Ч��ִ�п���
                        lng���˿���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���˿���ID))
                        lngִ�п���ID = Get�շ�ִ�п���ID(Val(vsAdvice.TextMatrix(lngRow, COL_����ID)), Val(vsAdvice.TextMatrix(lngRow, COL_��ҳID)), _
                            mrsPrice!�շ����, rs�շ�ϸĿ!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID)
                        '��¼�Ƿ��������
                        If mrsPrice!�շ���� = "4" Then
                            .TextMatrix(j, COLP_��������) = Val(Nvl(mrsPrice!����, 0))
                        End If
                        .Editable = flexEDKbdMouse
                    End If
                    If lngִ�п���ID <> 0 Then
                        mrsDept.Filter = "ID=" & lngִ�п���ID
                        If Not mrsDept.EOF Then
                            .TextMatrix(j, COLP_ִ�п���) = mrsDept!����
                        End If
                    End If
                    .TextMatrix(j, COLP_ִ�п���ID) = lngִ�п���ID
                                        
                    '��۵Ĵ���
                    If Nvl(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                        If InStr(",5,6,7,", mrsPrice!�շ����) > 0 Then
                            If InStr(",5,6,7,", mrsPrice!�������) > 0 Then
                                'ҩ��ҩƷ����1��סԺ��λ��ʱ��
                                .TextMatrix(j, COLP_����) = CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rs�շ�ϸĿ!סԺ��װ, 1))
                                .TextMatrix(j, COLP_����) = Format(Val(.TextMatrix(j, COLP_����)) * Nvl(rs�շ�ϸĿ!סԺ��װ, 1), "0.00000")
                            Else
                                '��ҩ��ҩƷ�����۵�λ����
                                .TextMatrix(j, COLP_����) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, mrsPrice!����), "0.00000")
                            End If
                        ElseIf mrsPrice!�շ���� = "4" And Nvl(mrsPrice!����, 0) = 1 Then
                            'ʱ�����ļ۸��ҩƷһ������
                            .TextMatrix(j, COLP_����) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, mrsPrice!����), "0.00000")
                        Else
                            '��¼��������ļ۸�Χ
                            .Cell(flexcpData, j, 0) = 1 '��ʶΪ���(ҩƷ����)
                            .Cell(flexcpData, j, 1) = Nvl(rs�շ�ϸĿ!ԭ��, 0)
                            .Cell(flexcpData, j, 2) = Nvl(rs�շ�ϸĿ!�ּ�, 0)
                            'Ҳ����ǰ���˱��,���ڱ�۷�Χ����
                            If .TextMatrix(j, COLP_����) <> "" Then
                                If CheckScope(Nvl(rs�շ�ϸĿ!ԭ��, 0), Nvl(rs�շ�ϸĿ!�ּ�, 0), Nvl(mrsPrice!����, 0)) <> "" Then
                                    .TextMatrix(j, COLP_����) = ""
                                End If
                            End If
                            '��ҩƷ���,��ʹ�̶�Ҳ���Ա༭
                            .Editable = flexEDKbdMouse
                        End If
                    End If

                    .TextMatrix(j, COLP_��������) = Nvl(rs�շ�ϸĿ!��������)
                    .TextMatrix(j, COLP_�̶�) = mrsPrice!�̶�
                    .TextMatrix(j, COLP_����) = IIF(Nvl(mrsPrice!����, 0) = 0, "", "��")
                    
                    '��¼���ڻָ�����
                    .Cell(flexcpData, j, COLP_�Ƽ�ҽ��) = .TextMatrix(j, COLP_�Ƽ�ҽ��)
                    .Cell(flexcpData, j, COLP_�շ���Ŀ) = .TextMatrix(j, COLP_�շ���Ŀ)
                    .Cell(flexcpData, j, COLP_����) = .TextMatrix(j, COLP_����)
                    .Cell(flexcpData, j, COLP_����) = .TextMatrix(j, COLP_����)
                    .Cell(flexcpData, j, COLP_ִ�п���) = .TextMatrix(j, COLP_ִ�п���)
                    
                    '��ʶ�̶�����Ϊ��ɫ
                    If mrsPrice!�̶� <> 0 Then
                        .Cell(flexcpBackColor, j, .FixedCols, j, .Cols - 1) = &HE0E0E0
                    End If
                    
                    j = j + 1
                End If
                
                mrsPrice.MoveNext
            Next
            
            '���ñ༭����
            '------------------------------------------------------------------
            '��Ҫ�Ƽ۵�ҽ��ѡ��
            If strTmp <> "" And Not (blnHavePrice And blnAllFixed) Then
                .ColData(COLP_�Ƽ�ҽ��) = Mid(strTmp, 2)
                .Editable = flexEDKbdMouse '����ѡ������Ա༭
            Else
                .ColData(COLP_�Ƽ�ҽ��) = ""
            End If
        End If
        .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
        
        'ȱʡѡ��Ƽ�ҽ��(�������)
        Call ShowDefaultRow
        .Redraw = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub SetSameTime(ByVal lngRow As Long)
'���ܣ���������ҽ����Ϊ��ͬ��У��,��ͣ,����ʱ��
    Dim strTime As String, vPause As Date
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        strTime = Format(.TextMatrix(lngRow, COL_����), "yyyy-MM-dd HH:mm")
        For i = .FixedRows To .Rows - 1
            If i <> lngRow Then
                blnDo = True
                If mint���� = 3 Then
                    'Ӧ>=����ʱ��
                    If strTime < Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                ElseIf mint���� = 5 Then
                    'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
                    If strTime < Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                    'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
                    If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                        If strTime <= Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                    'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
                    If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                        If strTime >= Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                    'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
                    vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 7)
                    If vPause <> CDate(0) Then
                        If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                ElseIf mint���� = 6 Then
                    'Ӧ>��ͣʱ��
                    vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 6)
                    If vPause <> CDate(0) Then
                        If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                    'Ӧ<=ִ����ֹʱ��
                    If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                        If strTime > Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                    End If
                End If
                If blnDo Then
                    .TextMatrix(i, COL_����) = strTime
                    .Cell(flexcpData, i, COL_����) = strTime
                End If
            End If
        Next
    End With
End Sub

Private Function GetPauseTime(ByVal lngҽ��ID As Long, ByVal int״̬ As Integer) As Date
'���ܣ���ȡָ��ҽ������ͣʱ��(��ҽ����ǰӦ����ͣ)���ϴ�����ʱ��(�����)
'������int״̬=6-��ͣ,7-����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(����ʱ��) as �ϴ�ʱ�� From ����ҽ��״̬ Where ��������=[2] And ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, int״̬)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!�ϴ�ʱ��) Then
            GetPauseTime = rsTmp!�ϴ�ʱ��
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceCheckWarn(ByVal lngCmd As Long, ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ��ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        21-����״̬/����ʷ����(ֻ��)
'      lngRow=��ǰҩƷҽ�����к�:lngCmd=0ʱ��Ҫ,�ಡ����������ʱ��Ҫ��ǰ������
'���أ����PASS�˵�ʱ������>=0��ʾ���Ե����˵�,��������-1
'˵������ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strSQL As String, i As Long, k As Long
    
    AdviceCheckWarn = -1
    If Not (lngRow >= vsAdvice.FixedRows) Then Exit Function '����Ҫȷ������������
    
    On Error GoTo errH
    Screen.MousePointer = 11
        
    '����PASS����״̬
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If
    
    '���벡�˾�����Ϣ(PASS��Ҫ�Ļ�������,ͬһ���˿ɲ��ظ�����)
    '-------------------------------------------------------------
    lng����ID = Val(vsAdvice.TextMatrix(lngRow, COL_����ID))
    lng��ҳID = Val(vsAdvice.TextMatrix(lngRow, COL_��ҳID))
    If lng����ID <> mlngPassPati Then
        strSQL = _
            " Select A.����,A.�Ա�,A.��������,B.��Ժ����,B.��Ժ����," & _
            " C.���� as ������,C.���� as ������,D.��� as ҽ����,D.���� as ҽ����" & _
            " From ������Ϣ A,������ҳ B,���ű� C,��Ա�� D" & _
            " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
            " And B.סԺҽʦ=D.����(+) And A.����ID=[1] And B.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    
        Call PassSetPatientInfo(lng����ID, lng��ҳID, rsTmp!����, Nvl(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
            rsTmp!������ & "/" & rsTmp!������, IIF(Not IsNull(rsTmp!ҽ����), Nvl(rsTmp!ҽ����) & "/" & Nvl(rsTmp!ҽ����), ""), _
            IIF(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy-MM-dd")))
        mlngPassPati = lng����ID
    End If
    
    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With vsAdvice
            If Val(.TextMatrix(lngRow, COL_ID)) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, COL_�������)) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                'ȡҩƷ����
                strҩƷ = .TextMatrix(lngRow, COL_ҽ������)
                If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
                If InStr(strҩƷ, "(") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, "(") - 1)
                'ȡҩƷ��ҩ;��
                str�÷� = .TextMatrix(lngRow, COL_�÷�)
                
                '�����ѯҩƷ��Ϣ
                Call PassSetQueryDrug(.TextMatrix(lngRow, COL_�շ�ϸĿID), strҩƷ, .TextMatrix(lngRow, COL_������λ), str�÷�)
                    
                '���ò˵�����״̬
                Call SetPassMenuState
                
                AdviceCheckWarn = 1 '��ʾ���Ե����˵�
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    
    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    
    'Pass
    If Button = 2 Then
        With vsAdvice
            lngRow = .MouseRow
            If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
                If Not .RowHidden(lngRow) Then .Row = lngRow
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Pass
    If Button = 2 And gblnPass And InStr(mstrPrivs, "������ҩ���") > 0 And mint���� = 3 Then
        If AdviceCheckWarn(0, vsAdvice.Row) >= 0 Then PopupMenu mnuPass, 2
    End If
End Sub

Private Sub SetPassMenuState()
'���ܣ�����Pass�˵�����״̬
    'Pass
    'һ���˵�
    'ҩ���ٴ���Ϣ�ο�
    mnuPassItem(0).Enabled = PassGetState("CPRRes") = 1
    'ҩƷ˵����
    mnuPassItem(1).Enabled = PassGetState("Directions") = 1
    '�й�ҩ��
    mnuPassItem(2).Enabled = PassGetState("Chp") = 1
    '������ҩ����
    mnuPassItem(3).Enabled = PassGetState("CPERes") = 1
    '����ֵ
    mnuPassItem(4).Enabled = PassGetState("CheckRes") = 1
    'ר����Ϣ
    'mnuPassItem(6).Enabled = PassGetState("") = 1
    'ҽҩ��Ϣ����
    mnuPassItem(8).Enabled = PassGetState("MEDInfo") = 1
    'ҩƷ�����Ϣ
    mnuPassItem(10).Enabled = PassGetState("MATCH-DRUG") = 1
    '��ҩ;�������Ϣ
    mnuPassItem(11).Enabled = PassGetState("MATCH-ROUTE") = 1
    'ҽԺҩƷ��Ϣ
    mnuPassItem(12).Enabled = PassGetState("HisDrugInfo") = 1
    
    '���˲˵�
    'ҩ��-ҩ���໥����
    mnuPassSpec(0).Enabled = PassGetState("DDIM") = 1
    'ҩ��-ʳ���໥ʹ��
    mnuPassSpec(1).Enabled = PassGetState("DFIM") = 1
    '����ע�����������
    mnuPassSpec(3).Enabled = PassGetState("MatchRes") = 1
    '����ע�����������
    mnuPassSpec(4).Enabled = PassGetState("TriessRes") = 1
    '����֢
    mnuPassSpec(6).Enabled = PassGetState("DDCM") = 1
    '������
    mnuPassSpec(7).Enabled = PassGetState("SIDE") = 1
    '��������ҩ
    mnuPassSpec(9).Enabled = PassGetState("GERI") = 1
    '��ͯ��ҩ
    mnuPassSpec(10).Enabled = PassGetState("PEDI") = 1
    '��������ҩ
    mnuPassSpec(11).Enabled = PassGetState("PREG") = 1
    '��������ҩ
    mnuPassSpec(12).Enabled = PassGetState("LACT") = 1
End Sub

Private Sub mnuPassItem_Click(Index As Integer)
'���ܣ�ִ��PASS����
    'Pass
    Select Case Index
    Case 0 'ҩ���ٴ���Ϣ�ο�
        Call PassDoCommand(101)
    Case 1 'ҩƷ˵����
        Call PassDoCommand(102)
    Case 2 '�й�ҩ��
        Call PassDoCommand(107)
    Case 3 '������ҩ����
        Call PassDoCommand(103)
    Case 4 '����ֵ
        Call PassDoCommand(104)
    Case 8 'ҽҩ��Ϣ����
        Call PassDoCommand(106)
    Case 10 'ҩƷ�����Ϣ
        Call PassDoCommand(13)
    Case 11 '��ҩ;�������Ϣ
        Call PassDoCommand(14)
    Case 12 'ҽԺҩƷ��Ϣ
        Call PassDoCommand(105)
    End Select
End Sub

Private Sub mnuPassSpec_Click(Index As Integer)
'���ܣ�ִ��ר��PASS����
    'Pass
    Select Case Index
    Case 0 'ҩ��-ҩ���໥����
        Call PassDoCommand(201)
    Case 1 'ҩ��-ʳ���໥ʹ��
        Call PassDoCommand(202)
    Case 3 '����ע�������
        Call PassDoCommand(203)
    Case 4 '����ע�������
        Call PassDoCommand(204)
    Case 6 '����֢
        Call PassDoCommand(205)
    Case 7 '������
        Call PassDoCommand(206)
    Case 9 '��������ҩ
        Call PassDoCommand(207)
    Case 10 '��ͯ��ҩ
        Call PassDoCommand(208)
    Case 11 '��������ҩ
        Call PassDoCommand(209)
    Case 12 '��������ҩ
        Call PassDoCommand(210)
    End Select
End Sub
